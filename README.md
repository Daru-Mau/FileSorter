# Excel-Based File Renamer

A Windows application for renaming files based on matching data in Excel spreadsheets.

## Features

- **Excel Integration**: Look up file names in Excel spreadsheets
- **Smart Matching**: Match files to Excel entries by name
- **Detailed View**: View detailed Excel information for each file
- **Easy Renaming**: Rename files using IDs from Excel data
- **Search Capabilities**: Search for files by name

## Requirements

- Python 3.6+
- Tkinter (included with Python)
- pandas
- openpyxl

## Installation

1. Clone this repository:

```
git clone https://github.com/yourusername/ExcelFileRenamer.git
cd ExcelFileRenamer
```

2. Create a virtual environment (recommended):

```
python -m venv .venv
.venv\Scripts\activate
```

3. Install required packages:

```
pip install pandas openpyxl
```

## Usage

1. Run the application:

```
python src/excel_file_renamer.py
```

2. Using the application:
   - Select source folder containing files to rename
   - Select an Excel file containing the reference data
   - Configure column mappings:
     - Name Column: Column in Excel that contains file names to match
     - ID Column: Column in Excel that contains IDs to use for renaming
     - Date Column: Optional column with date information
   - Click "Scan Files" to load files from the folder
   - Select a file from the list to search for matches in Excel
   - When a match is found, review the information and click "Rename Selected File"

## Examples

### Workflow Example

1. **Load Excel File**: Load your Excel file that contains file names and their corresponding IDs
2. **Configure Columns**:
   - Name Column = "FileName" (or whatever column has the file names)
   - ID Column = "ID" (or whatever column has the ID numbers)
3. **Select a File**: Click on a file in the left panel
4. **Verify Match**: The application will find the matching entry in Excel
5. **Rename**: Click "Rename Selected File" to rename the file using the ID

### Matching Strategies

The application uses two matching strategies:

- **Exact Match**: Finds entries where the Excel name exactly matches the filename
- **Partial Match**: If no exact match is found, searches for partial matches

## License

MIT License
