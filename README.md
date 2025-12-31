# Modbus RTU 从站模拟程序

这是一个基于 **Python + pymodbus** 的 **Modbus RTU 从站（Slave）模拟程序**。
程序通过 **Excel 文件** 定义各类寄存器的初始数据，适用于：

- 上位机 / Modbus Poll 调试
- PLC 通讯联调
- 串口 Modbus RTU 从站仿真

---

## ✨ 功能特性

- ✅ 支持 RTU 串口模式
- ✅ 支持 CO / DI / HR / IR 四类寄存器
- ✅ Excel 定义寄存器数据
- ✅ 起始地址可配置（16 进制）
- ✅ 支持十进制 / 十六进制数据
- ✅ Modbus 通讯日志按天保存
- ✅ 控制台仅显示程序运行信息（无 Modbus 噪声日志）
- ✅ 启动前人工确认（回车启动，`q` 退出）

---

## 📦 目录结构

```text
.
├─ main.py
├─ config.yaml
├─ README.md
├─ default/
│  ├─ co.xlsx
│  ├─ di.xlsx
│  ├─ hr.xlsx
│  └─ ir.xlsx
└─ log/
   └─ 2025-01-01.log