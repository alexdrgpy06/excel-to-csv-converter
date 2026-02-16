# Excel to CSV Automation Engine 📊✨

> **Architect: Alejandro Ramírez**

A professional-grade Python CLI utility designed for high-fidelity conversion of Excel workbooks into standard CSV formats. Optimized for automated data pipelines and batch processing.

## 🚀 Overview
This tool eliminates the manual overhead of spreadsheet extraction by providing a robust, programmable interface for Excel-to-CSV conversion. It handles multi-sheet workbooks, custom output mapping, and provides detailed logging for integration into larger systems.

## ✨ Key Features
- **High-Performance Parsing**: Leverages the `pandas` engine for rapid data extraction.
- **Multi-Sheet Support**: Convert specific sheets or the entire workbook in a single operation.
- **Automated Naming**: Intelligent output path generation based on input filenames and sheet names.
- **Robust Error Handling**: Descriptive failure reports for missing files or corrupted workbooks.
- **Integration Ready**: Clean CLI interface with support for verbose logging and explicit output targeting.

## 🛠️ Tech Stack
- **Language**: Python 3.x
- **Core Library**: Pandas
- **Excel Engine**: Openpyxl

## 📦 Usage
### Basic Conversion
```bash
python excel_to_csv.py data.xlsx
```

### Convert All Sheets
```bash
python excel_to_csv.py data.xlsx --all
```

### Specific Target
```bash
python excel_to_csv.py source.xlsx -o results.csv -s "Q4_Reports"
```

## 📜 Dependencies
Install the required libraries before initialization:
```bash
pip install pandas openpyxl
```

---
*Built with precision by Alejandro Ramírez.*
