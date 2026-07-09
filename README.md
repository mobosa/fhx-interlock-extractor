<div align="center">

<img src="exp_logo.ico" width="80" alt="FHX Interlock Extractor">

# FHX Interlock Extractor

[![GitHub stars](https://img.shields.io/github/stars/mobosa/fhx-interlock-extractor?style=social)](https://github.com/mobosa/fhx-interlock-extractor/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/mobosa/fhx-interlock-extractor?style=social)](https://github.com/mobosa/fhx-interlock-extractor/network/members)
[![GitHub issues](https://img.shields.io/github/issues/mobosa/fhx-interlock-extractor)](https://github.com/mobosa/fhx-interlock-extractor/issues)
[![GitHub license](https://img.shields.io/github/license/mobosa/fhx-interlock-extractor)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://www.python.org/)

**Extract interlock information from DeltaV FHX files and generate Excel reports**

[English](README.md) | [简体中文](README_zh-CN.md)

</div>

---

## 🔍 Overview

FHX Interlock Extractor is a tool designed for **Emerson DeltaV system engineers** to extract interlock information from FHX configuration files and generate structured Excel reports. It parses DeltaV FHX files to extract Permissive, Interlock, and Fault Handler configurations.

> 💡 Supports any FHX type (Library, Control Strategies, Setup, etc.) with automatic detection.

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| **Permissive Extraction** | Extract all Permissive blocks with conditions and actions |
| **Interlock Extraction** | Extract Interlock configurations with triggers and responses |
| **Fault Handler Extraction** | Extract Fault Handler definitions and fault states |
| **Excel Export** | Generate structured Excel with separate sheets for each type |
| **GUI Interface** | Modern Apple-style GUI for easy operation |
| **CLI Support** | Command-line interface for batch processing |

## 🚀 Quick Start

### Download Pre-built Executable (Recommended)

1. Download `FHX_Interlock_Extractor.exe` from [Releases](https://github.com/mobosa/fhx-interlock-extractor/releases)
2. Double-click to run — no Python installation required

### Run from Source

```bash
pip install openpyxl customtkinter
python fhx_gui.py
```

## 📖 Workflow

```
┌─────────────────────────────────────────────────────────────┐
│  1. Select FHX File                                         │
│     → Browse and select the DeltaV FHX configuration file   │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  2. Extract Interlock Information                           │
│     → Tool parses FHX and extracts all interlock blocks     │
└─────────────────────────────────────────────────────────────┘
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  3. Generate Excel Report                                   │
│     → Export structured Excel with separate worksheets      │
└─────────────────────────────────────────────────────────────┘
```

## 💻 CLI Mode

```bash
# Extract interlock info and export to Excel
python fhx_interlock_extractor.py <input.fhx> [-o output.xlsx]
```

| Argument | Description |
|----------|-------------|
| `input.fhx` | **Required.** DeltaV FHX configuration file |
| `-o, --output` | Optional. Output Excel file path |

## 📊 Excel Output Format

The exported Excel contains worksheets:

| Sheet | Content |
|-------|---------|
| **Permissive** | Permissive block definitions with conditions |
| **Interlock** | Interlock configurations with triggers |
| **Fault Handler** | Fault Handler definitions and states |

## 🛠️ Build Executable

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=exp_logo.ico fhx_gui.py
```

Generated exe is in the `dist/` directory.

## 📁 Project Structure

```
fhx-interlock-extractor/
├── core.py                    # Backend logic (FHX parsing, Excel generation)
├── fhx_gui.py                 # GUI entry point (customtkinter)
├── fhx_interlock_extractor.py # CLI entry point
├── requirements.txt           # Python dependencies
├── exp_logo.ico               # App icon
└── README.md                  # This file (English)
```

## 📋 Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl` | Excel read/write |
| `customtkinter` | Modern GUI framework |
| Python 3.8+ | Runtime (source code only) |

## 👤 Author

**Jared.Ji** — Jared.Ji@emerson.com

---

<div align="center">

**Made with ❤️ for Emerson DeltaV Engineers**

</div>
