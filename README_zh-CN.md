<div align="center">

<img src="exp_logo.ico" width="80" alt="FHX Interlock Extractor">

# FHX Interlock Extractor

[![GitHub stars](https://img.shields.io/github/stars/mobosa/fhx-interlock-extractor?style=social)](https://github.com/mobosa/fhx-interlock-extractor/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/mobosa/fhx-interlock-extractor?style=social)](https://github.com/mobosa/fhx-interlock-extractor/network/members)
[![GitHub issues](https://img.shields.io/github/issues/mobosa/fhx-interlock-extractor)](https://github.com/mobosa/fhx-interlock-extractor/issues)
[![GitHub license](https://img.shields.io/github/license/mobosa/fhx-interlock-extractor)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://www.python.org/)

**从 DeltaV FHX 文件中提取联锁信息并生成 Excel 报告**

[English](README.md) | [简体中文](README_zh-CN.md)

</div>

---

## 🔍 概述

FHX Interlock Extractor 是一款专为 **Emerson DeltaV 系统工程师** 设计的工具，用于从 FHX 配置文件中提取联锁信息并生成结构化的 Excel 报告。它能解析 DeltaV FHX 文件，提取 Permissive、Interlock 和 Force SP 配置。

> 💡 支持任何 FHX 类型（Library、Control Strategies、Setup 等），自动识别处理。

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| **Permissive 提取** | 提取所有 Permissive 块的条件和动作 |
| **Interlock 提取** | 提取 Interlock 配置，包括触发器和响应 |
| **Force SP 提取** | 提取 Force SP 定义和故障状态 |
| **Excel 导出** | 生成结构化 Excel，每种类型独立工作表 |
| **GUI 界面** | 现代 Apple 风格图形界面，操作简便 |
| **CLI 支持** | 命令行界面，支持批处理 |

## 🚀 快速开始

### 下载打包好的 exe（推荐）

1. 从 [Releases](https://github.com/mobosa/fhx-interlock-extractor/releases) 下载 `FHX_Interlock_Extractor.exe`
2. 双击运行，无需安装 Python

### 从源码运行

```bash
pip install openpyxl customtkinter
python fhx_gui.py
```

## 📖 工作流程

```
┌─────────────────────────────────────────────────────────────┐
│  1. 选择 FHX 文件                                           │
│     → 浏览并选择 DeltaV FHX 配置文件                         │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  2. 提取联锁信息                                             │
│     → 工具解析 FHX 并提取所有联锁块                          │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  3. 生成 Excel 报告                                          │
│     → 导出结构化 Excel，包含独立工作表                        │
└─────────────────────────────────────────────────────────────┘
```

## 💻 CLI 命令行模式

```bash
# 提取联锁信息并导出 Excel
python fhx_interlock_extractor.py <input.fhx> [-o output.xlsx]
```

| 参数 | 说明 |
|------|------|
| `input.fhx` | **必填。** DeltaV FHX 配置文件 |
| `-o, --output` | 可选。输出 Excel 文件路径 |

## 📊 Excel 输出格式

导出的 Excel 包含以下工作表：

| Sheet | 内容 |
|-------|------|
| **Permissive** | Permissive 块定义及条件 |
| **Interlock** | Interlock 配置及触发器 |
| **Force SP** | Force SP 定义及状态 |

## 🛠️ 打包 exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=exp_logo.ico fhx_gui.py
```

生成的 exe 在 `dist/` 目录下。

## 📁 项目结构

```
fhx-interlock-extractor/
├── core.py                    # 后端逻辑（FHX 解析、Excel 生成）
├── fhx_gui.py                 # GUI 入口（customtkinter）
├── fhx_interlock_extractor.py # CLI 入口
├── requirements.txt           # Python 依赖
├── exp_logo.ico               # 程序图标
└── README_zh-CN.md            # 本文档（中文）
```

## 📋 依赖

| 包 | 用途 |
|---|------|
| `openpyxl` | Excel 读写 |
| `customtkinter` | 现代化 GUI 框架 |
| Python 3.8+ | 运行环境（仅源码运行时需要） |

## 👤 作者

**Jared.Ji** — Jared.Ji@emerson.com

---

<div align="center">

**为 Emerson DeltaV 工程师用心打造 ❤️**

</div>
