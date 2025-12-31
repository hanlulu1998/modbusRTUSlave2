import pandas as pd
import yaml
import asyncio
import os
import sys
from datetime import datetime

from pymodbus.server import StartAsyncSerialServer
from pymodbus.datastore import (
    ModbusSequentialDataBlock,
    ModbusServerContext,
    ModbusDeviceContext,
)

import logging
from logging.handlers import TimedRotatingFileHandler


# =========================================================
# Modbus 日志（仅文件，不输出到控制台）
# =========================================================
def setup_modbus_logging(enable: bool):
    if not enable:
        return

    log_dir = os.path.join(os.getcwd(), "log")
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(log_dir, f"{datetime.now():%Y-%m-%d}.log")

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"
    )

    file_handler = TimedRotatingFileHandler(
        log_file,
        when="midnight",
        backupCount=7,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logger = logging.getLogger("pymodbus")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.addHandler(file_handler)
    logger.propagate = False  # ❗关键：不传到 root logger

    for name in (
        "pymodbus.server",
        "pymodbus.transport",
        "pymodbus.framer",
        "pymodbus.factory",
    ):
        sub = logging.getLogger(name)
        sub.setLevel(logging.DEBUG)
        sub.propagate = True


# =========================================================
# 基础检查
# =========================================================
def check_value_dir(value_dir: str):
    if not os.path.exists(value_dir):
        print("❌ 配置错误：Excel 数据目录不存在")
        print(f"   value_dir = {value_dir}")
        sys.exit(1)

    if not os.path.isdir(value_dir):
        print("❌ 配置错误：value_dir 不是文件夹")
        print(f"   value_dir = {value_dir}")
        sys.exit(1)


# =========================================================
# 配置 / Excel 处理
# =========================================================
def load_config(path="config.yaml"):
    return yaml.safe_load(open(os.path.join(os.getcwd(), path), encoding="utf-8"))


def load_values_from_excel(path, data_base="dec"):
    df = pd.read_excel(path, header=None)

    values = []
    for v in df.iloc[:, 0]:
        if pd.isna(v):
            values.append(0)
        elif data_base == "hex":
            values.append(int(str(v), 16))
        else:
            values.append(int(v))

    values = [values[0]] + values
    return values


def build_block(start_addr: int, values: list):
    return ModbusSequentialDataBlock(start_addr, values)


def parse_hex_address(value, default=0):
    if value is None:
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        return int(value, 16)
    raise ValueError(f"非法 start_address 值: {value}")


# =========================================================
# 启动界面打印
# =========================================================
def print_banner():
    print("=" * 72)
    print("🟢 Modbus RTU 从站软件".center(72))
    print("-" * 72)
    print("版本      : 1.0")
    print("作者      : 韩露露")
    print("联系邮箱  : hanlulu1998@outlook.com")
    print("=" * 72)
    print()


def print_help():
    print("【使用说明】")
    print("  1) 本程序为 Modbus RTU 从站（Slave）")
    print("  2) Excel 文件用于定义寄存器数据内容")
    print("  3) Excel 第 1 行数据 → 对应 start_address 中配置的起始地址")
    print("  4) 修改 Excel 后需重启程序才能生效")
    print("  5) 使用 Modbus Poll / 上位机 / PLC 作为主站读取数据")
    print()

    print("【config.yaml 参数说明】")
    print()
    print("  serial:                  # 串口参数")
    print("    port      : COM2       # 串口号")
    print("    baudrate  : 9600       # 波特率")
    print("    bytesize  : 8          # 数据位")
    print("    parity    : N          # 校验位 (N/E/O)")
    print("    stopbits  : 1          # 停止位")
    print("    timeout   : 5          # 超时时间(秒)")
    print()
    print("  modbus:")
    print("    slave_id  : 1          # Modbus 从站地址")
    print()
    print("  value_enable:            # 是否启用各类寄存器")
    print("    co        : on/off     # Coil")
    print("    di        : on/off     # Discrete Input")
    print("    hr        : on/off     # Holding Register")
    print("    ir        : on/off     # Input Register")
    print()
    print("  value_dir:")
    print("    default                # Excel 文件所在目录（必须存在）")
    print()
    print("  data_base:")
    print("    dec                    # dec=十进制, hex=十六进制")
    print()
    print("  start_address:           # 寄存器起始地址（16进制）")
    print("    co        : 0x0000")
    print("    di        : 0x0000")
    print("    hr        : 0x0000")
    print("    ir        : 0x0000")
    print("    # Excel 第 1 行即对应以上地址")
    print()
    print("  enable_logging: on/off   # Modbus 日志开关")
    print("    on   → 日志保存到 log/YYYY-MM-DD.log")
    print("    off  → 不记录 Modbus 通讯日志")
    print()
    print("=" * 72)
    print()


def wait_for_start():
    while True:
        cmd = input("▶ 回车启动 Modbus RTU 从站，输入 q 退出：").strip().lower()
        if cmd == "":
            return
        if cmd == "q":
            print("👋 用户选择退出，程序结束")
            sys.exit(0)
        print("⚠ 无效输入，请直接回车或输入 q")


# =========================================================
# 主程序
# =========================================================
async def main():
    print_banner()
    print_help()

    wait_for_start()

    cfg = load_config()

    enable_logging = cfg.get("enable_logging", True)
    if enable_logging:
        setup_modbus_logging(True)
        print("[LOG ] Modbus 通讯日志已启用（log/ 目录）")
    else:
        print("[LOG ] Modbus 通讯日志未启用")

    enable = cfg["value_enable"]
    slave_id = cfg["modbus"]["slave_id"]
    serial = cfg["serial"]

    data_base = cfg.get("data_base", "dec")
    value_dir_name = cfg.get("value_dir", ".")
    start_address = cfg.get("start_address", {})

    value_dir = os.path.join(os.getcwd(), value_dir_name)
    check_value_dir(value_dir)

    device = {}

    if enable.get("co"):
        start = parse_hex_address(start_address.get("co"), 0)
        device["co"] = build_block(
            start, load_values_from_excel(os.path.join(value_dir, "co.xlsx"), data_base)
        )
        print(f"[LOAD] CO | start=0x{start:04X}")

    if enable.get("di"):
        start = parse_hex_address(start_address.get("di"), 0)
        device["di"] = build_block(
            start, load_values_from_excel(os.path.join(value_dir, "di.xlsx"), data_base)
        )
        print(f"[LOAD] DI | start=0x{start:04X}")

    if enable.get("hr"):
        start = parse_hex_address(start_address.get("hr"), 0)
        device["hr"] = build_block(
            start, load_values_from_excel(os.path.join(value_dir, "hr.xlsx"), data_base)
        )
        print(f"[LOAD] HR | start=0x{start:04X}")

    if enable.get("ir"):
        start = parse_hex_address(start_address.get("ir"), 0)
        device["ir"] = build_block(
            start, load_values_from_excel(os.path.join(value_dir, "ir.xlsx"), data_base)
        )
        print(f"[LOAD] IR | start=0x{start:04X}")

    dev_ctx = ModbusDeviceContext(**device)
    context = ModbusServerContext(devices={slave_id: dev_ctx}, single=False)

    print()
    print("🚀 Modbus RTU 从站启动成功")
    print("-" * 40)
    print(f"  Slave ID : {slave_id}")
    print(f"  串口     : {serial['port']}")
    print(f"  波特率   : {serial['baudrate']}")
    print(f"  数据进制 : {data_base.upper()}")
    print(f"  Excel目录: {value_dir}")
    print("-" * 40)
    print("⌛ 等待 Modbus 主站请求...\n")

    await StartAsyncSerialServer(
        context=context,
        framer="rtu",
        port=serial["port"],
        baudrate=serial["baudrate"],
        bytesize=serial["bytesize"],
        parity=serial["parity"],
        stopbits=serial["stopbits"],
        timeout=serial["timeout"],
    )


if __name__ == "__main__":
    asyncio.run(main())
