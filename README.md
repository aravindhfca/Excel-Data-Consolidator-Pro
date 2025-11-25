# Excel Data Consolidator Pro

A powerful desktop application for merging and consolidating multiple Excel and CSV files with advanced data processing capabilities.

## 🎯 Features

- **Multi-Format Support**: Merge `.xlsx`, `.xlsm`, `.xls`, `.csv`, and `.xlsb` files
- **Flexible Sheet Selection**: Choose specific sheets or merge all sheets from multiple files
- **Smart Header Handling**: Skip, include, or use first row as headers
- **Data Cleaning**:
  - Skip empty rows
  - Remove duplicate rows
  - Custom row start and column start positions
- **Enhanced Data**:
  - Add source filename column
  - Add row ID numbering
  - Add sheet name column
- **Multiple Export Formats**: Export as `.xlsx` or `.csv`
- **Data Preview**: Real-time preview of merged data before export
- **Error Logging**: Automatic error tracking for troubleshooting
- **Dark/Light Theme**: Toggle between themes
- **Configuration Persistence**: Remembers your settings

## 📋 Requirements

- Python 3.7+
- tkinter (usually included with Python)
- openpyxl
- pyxlsb (optional, for `.xlsb` support)

## 🚀 Installation

1. **Clone the repository**:
```bash
git clone https://github.com/aravindhfca/excel-consolidator-pro.git
cd excel-consolidator-pro
```

2. **Install dependencies**:
```bash
pip install openpyxl
```

3. **Optional: For XLSB support**:
```bash
pip install pyxlsb
```

4. **Run the application**:
```bash
python "import tkinter as tk.py"
```

## 💡 Usage Guide

### Basic Workflow

1. **Select Files**
   - Click "Select Excel/CSV Files" button
   - Choose multiple files to merge
   - Use "Remove" or "Clear" buttons to manage selections

2. **Configure Settings**
   - **Sheet Name**: Select specific sheet to merge (dropdown)
   - **Header Row**: Choose how to handle headers
     - `Skip First`: Remove first row from all files
     - `Include All`: Keep all rows as-is
     - `First Only`: Keep first file's header, skip for others
   - **Start Row/Column**: Specify where data begins

3. **Add Enhancements** (Optional)
   - Enable "Add source filename column"
   - Enable "Add numerical Row IDs"
   - Choose positions (First/Last column)
   - Add sheet names as a column

4. **Select Export Format**
   - Excel (`.xlsx`) - with formatting and summary sheets
   - CSV (`.csv`) - plain text format

5. **Preview & Export**
   - Click "Preview Data" to see first 100 rows
   - Click "Merge & Export" to save merged file

### Advanced Features

#### Auto-Detect Headers
Click "Auto Headers" to automatically detect column names from the first file and display them in the preview.

#### Detect Column Headers
Use "Detect Headers" to populate the column headers listbox for custom header management.

#### Select Specific Sheets
Click "Select Sheets" to choose which sheet tabs to merge (useful for multi-sheet workbooks).

#### Data Cleaning Options
- **Skip empty rows**: Ignore rows where all cells are empty
- **Remove duplicate rows**: Eliminate exact row duplicates

## 📊 Output Formats

### Excel Format (.xlsx)
Creates three sheets:
- **MergedData**: Main consolidated data
- **Summary**: File-by-file row counts
- **Metadata**: Merge configuration and statistics

### CSV Format (.csv)
Plain text format with comma-separated values.

## 🛠️ Configuration

Settings are automatically saved to `~/.excel_merger_config.json`:
- Theme preference
- Recent settings

Error logs are saved to `~/excel_merger_error_log.txt`

## 📝 Examples

### Example 1: Merge Multiple Monthly Reports
```
Input: Jan_Report.xlsx, Feb_Report.xlsx, Mar_Report.xlsx
Config:
  - Skip First row (headers)
  - Add source filename column
  - Export as .xlsx
Output: Consolidated_Reports.xlsx (all data in one sheet)
```

### Example 2: Consolidate CSV Files
```
Input: data1.csv, data2.csv, data3.csv
Config:
  - Include all rows
  - Add Row IDs
  - Export as .csv
Output: merged_data.csv
```

### Example 3: Multi-Sheet Workbook
```
Input: workbook1.xlsx, workbook2.xlsx
Config:
  - Select specific sheets (e.g., "GSTR-2B")
  - Include sheet name column
  - Remove duplicates
Output: consolidated.xlsx
```

## 🐛 Troubleshooting

### XLSB Files Not Supported
**Solution**: Install pyxlsb:
```bash
pip install pyxlsb
```

### Encoding Issues with CSV
The app uses UTF-8 encoding by default. If you have encoding issues:
- Ensure your CSV files are UTF-8 encoded
- Check the error log at `~/excel_merger_error_log.txt`

### Headers Not Detected
- Ensure the first file has headers in the first row
- Try manually selecting headers from the listbox
- Check file format is supported

### Preview Shows No Data
- Verify your file selection
- Check "Start Row" and "Start Column" settings
- Ensure "Skip empty rows" isn't filtering all data

## 📜 Application UI Layout

```
┌─────────────────────────────────────────────────────────┐
│ 📊 Excel Data Consolidator Pro      🌓 [Theme] [GitHub] │
├─────────────────────────────┬──────────────────────────┤
│ Left Panel                  │ Right Panel              │
│                             │                          │
│ 📁 File Selection           │ 📋 Data Preview          │
│  - Select Files             │  - Live preview          │
│  - File List                │  - Scrollable table      │
│  - Clear/Remove buttons     │  - Row counter           │
│                             │                          │
│ ⚙️ Configuration            │                          │
│  - Sheet name dropdown      │                          │
│  - Header handling          │                          │
│  - Start row/column         │                          │
│  - Column headers list      │                          │
│                             │                          │
│ 🎯 Options                  │                          │
│  - Skip empty rows          │                          │
│  - Remove duplicates        │                          │
│  - Add source filename      │                          │
│  - Sheet/Row ID options     │                          │
│  - Export format            │                          │
│                             │                          │
│ [👁️ Preview] [✓ Merge]     │                          │
├─────────────────────────────┴──────────────────────────┤
│ Status Bar | Progress Bar                              │
└─────────────────────────────────────────────────────────┘
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## 📄 License

This project is licensed under the MIT License - see LICENSE file for details.

## 👨‍💻 Author

**CA Aravindh K**
- GitHub: [@aravindhfca](https://github.com/aravindhfca)
- Designed and developed the Excel Data Consolidator Pro

## 🙏 Acknowledgments

- Built with Python and tkinter
- Uses openpyxl for Excel file handling
- Uses pyxlsb for XLSB format support (optional)

## 📧 Support

For issues, questions, or suggestions:
1. Check the error log: `~/excel_merger_error_log.txt`
2. Review existing issues on GitHub
3. Create a new issue with detailed information
