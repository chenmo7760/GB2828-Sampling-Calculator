# GB/T 2828.1-2003 Sampling Inspection Calculator

[中文文档](README.md) | English

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?logo=microsoft-excel)](https://www.microsoft.com/excel)
[![Standard](https://img.shields.io/badge/Standard-GB%2FT%202828.1--2003-blue)](http://www.gb688.cn/bzgk/gb/)

> Automated calculation tool for sampling inspection based on GB/T 2828.1-2003 / ISO 2859-1 standard for single sampling plans under normal inspection.

## ✨ Features

- 🎯 **4 Excel Custom Functions** - Easy to use in any cell
- 📊 **Batch Calculation** - Support for multiple rows of data
- 🔍 **Automatic Detection** - Smart identification of 100% inspection scenarios
- ✅ **Fully Tested** - 28 test cases covering all scenarios
- 📝 **Well Documented** - Complete Chinese documentation with examples
- 🚀 **5-Minute Setup** - Quick start guide included

## 🚀 Quick Start

### Step 1: Import VBA Code
1. Open Excel and press `Alt + F11` to open VBA editor
2. Insert → Module
3. Copy all code from `抽样计算.vba` and paste into the module window
4. Close VBA editor (`Alt + Q`)

### Step 2: Use Functions
Enter formula in any cell:
```excel
=获取样本量(150, "Ⅱ", 1.5)
```
Press Enter to see result: `20`

### Step 3: Save File
Save as `.xlsm` format (Excel Macro-Enabled Workbook)

## 📚 Core Functions

| Function | Description | Example | Result |
|----------|-------------|---------|--------|
| `获取样本量()` | Get sample size | `=获取样本量(150,"Ⅱ",1.5)` | 20 |
| `获取Ac值()` | Get acceptance number | `=获取Ac值(150,"Ⅱ",1.5)` | 1 |
| `获取Re值()` | Get rejection number | `=获取Re值(150,"Ⅱ",1.5)` | 2 |
| `获取检验类型()` | Get inspection type | `=获取检验类型(150,"Ⅱ",1.5)` | 抽检 |

## 📖 Parameters

### Batch Size (PL)
- **Type**: Integer
- **Range**: 2 ~ 500,000+
- **Example**: `150`, `5000`, `100000`

### Inspection Level
- **Special Levels**: `"S-1"`, `"S-2"`, `"S-3"`, `"S-4"`
- **General Levels**: `"Ⅰ"`, `"Ⅱ"`, `"Ⅲ"` (or `"I"`, `"II"`, `"III"`)
- **Note**: Must be enclosed in quotes

### AQL (Acceptable Quality Limit)
Available values (21 standard values):
```
0.01, 0.015, 0.025, 0.04, 0.065,
0.1,  0.15,  0.25,  0.4,  0.65,
1.0,  1.5,   2.5,   4.0,  6.5,
10,   15,    25,    40,   65,    100
```

## 💡 Usage Examples

### Example 1: Single Calculation
```excel
A1: Batch Size     B1: 150
A2: Inspection     B2: Ⅱ
A3: AQL           B3: 1.5
A4: Sample Size   B4: =获取样本量(B1, B2, B3)
A5: Ac            B5: =获取Ac值(B1, B2, B3)
A6: Re            B6: =获取Re值(B1, B2, B3)
```

### Example 2: Batch Processing
Create a table with formulas that auto-calculate:

| # | Batch | Level | AQL | Sample | Ac | Re |
|---|-------|-------|-----|--------|----|----|
| 1 | 50    | Ⅱ     | 1.5 | =获取样本量(B2,C2,D2) | =获取Ac值(B2,C2,D2) | =获取Re值(B2,C2,D2) |
| 2 | 500   | Ⅱ     | 2.5 | =获取样本量(B3,C3,D3) | =获取Ac值(B3,C3,D3) | =获取Re值(B3,C3,D3) |

Drag formulas down to calculate multiple rows.

## 📁 Project Structure

```
.
├── 抽样标准GB2828.xlsm       # Excel workbook with VBA functions
├── 抽样标准GB2828.xlsx       # Excel workbook (no macros)
├── 抽样计算.vba              # VBA source code (main)
├── 抽样计算_改进版.vba        # VBA source code (improved)
├── 工作表事件代码.vba         # Worksheet event handlers
├── README.md                # Documentation (Chinese)
├── README_EN.md             # Documentation (English)
├── 快速参考.md               # Quick reference card
├── 更新说明_v1.1.md          # Update notes v1.1
├── re.md                    # Original requirements
├── LICENSE                  # MIT License
└── .gitignore              # Git ignore file
```

## 🧪 Testing

Includes 28 comprehensive test cases covering:
- ✓ Basic functionality (5 cases)
- ✓ Boundary values (4 cases)
- ✓ Different inspection levels (7 cases)
- ✓ Different AQL values (6 cases)
- ✓ 100% inspection scenarios (3 cases)
- ✓ Large batch sizes (3 cases)

## 📋 Typical Scenarios

### Scenario 1: Product Shipment Inspection
```
Batch: 500 units
Inspection Level: Ⅱ (General)
AQL: 1.5 (Allow minor defects)
→ Sample: 50, Ac=2, Re=3
```

### Scenario 2: Critical Component Inspection
```
Batch: 1000 units
Inspection Level: Ⅲ (Strict)
AQL: 0.4 (Strict requirement)
→ Sample: 80, Ac=2, Re=3
```

### Scenario 3: Small Batch Inspection
```
Batch: 10 units
Inspection Level: Ⅱ
AQL: 1.5
→ Sample: 10, 100% inspection required
```

## 🔗 Related Standards

- **GB/T 2828.1-2003**: Sampling procedures for inspection by attributes -- Part 1: Sampling schemes indexed by acceptance quality limit (AQL) for lot-by-lot inspection
- **ISO 2859-1**: Sampling procedures for inspection by attributes -- Part 1: Sampling schemes indexed by acceptance quality limit (AQL) for lot-by-lot inspection

## 📝 Changelog

### v1.1 (2025-11-22)
- ✅ Adapted table shift (2 rows down)
- ✅ New: Auto-output AC to B5, RE to B6
- ✅ New: Highlight selected cells
- ✅ Fix: Correct sample size update when encountering "上"/"下"

### v1.0 (2025-11-21)
- ✅ Core calculation functionality
- ✅ 4 custom functions
- ✅ Handle "上"/"下" arrow logic
- ✅ 100% inspection detection
- ✅ Complete test suite (28 cases)
- ✅ Detailed documentation

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!

## ⭐ Support

If this tool helps you, please give it a star ⭐️

---

**Version**: 1.1  
**Created**: 2025-11-21  
**Standard**: GB/T 2828.1-2003 / ISO 2859-1  
**Inspection Type**: Normal inspection, single sampling

