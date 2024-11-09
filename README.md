# Excel Processor

This Python-based tool allows you to filter and save data from Excel files (`.xls` and `.xlsx` formats) by specifying a column to filter. It reads an input Excel file, filters data based on a specific column, and generates separate Excel files for each unique value in that column.

## Features
- **Load Excel Files**: Load Excel files and process them in the specified format (`.xls` or `.xlsx`).
- **Filter by Column**: Filter the data based on unique values from a column of your choice.
- **Save Filtered Files**: Save the filtered data to new Excel files, named based on the filter column values.
- **Supports `.xls` and `.xlsx` Formats**: Choose between `.xls` or `.xlsx` output files.

## Requirements
To use the `Excel Processor` script, you'll need the following Python packages:

- `pandas`
- `pyexcel`
- `pyexcel-xls`

### Installation

To install the required dependencies, you can use `pip`:

```bash
pip install pandas pyexcel pyexcel-xls
```

## Usage

1. **Run the Executable**:
   - Simply execute the `excel_processor.exe` on your system. If you're using the Python script, you can run it directly with Python.

2. **Input Excel File**:
   - The script will prompt you to input the path to the Excel file you wish to process.
   
3. **Specify Filter Column**:
   - Enter the column you want to filter by (e.g., `Region`, `Salesperson`, etc.). The script will generate separate files for each unique value in that column.

4. **Choose Output Format**:
   - You will be prompted to choose the file format for the output. You can choose between `.xls` and `.xlsx` formats.

5. **Generate Files**:
   - The script will create separate Excel files for each unique value in the filter column. The new files will be named according to the format `file_name_value.extension`.

6. **Exit**:
   - You can type 'c' to exit the script at any time.

## Example

Here’s an example of how the process works:

1. Input the Excel file: `data.xlsx`.
2. Choose the column to filter by: `Department`.
3. Choose the file format: `.xlsx`.
4. The script generates files like:
   - `filtered_data_Sales.xlsx`
   - `filtered_data_Marketing.xlsx`
   - `filtered_data_Engineering.xlsx`

## How to Build the Executable

If you'd like to build the executable yourself, use PyInstaller with the following command:

## Downloud the Executable
https://www.mediafire.com/file/92oar099dbgchri/excel_processor.exe/file



```bash
pyinstaller --onefile --icon="your_icon.ico" excel_processor.py
```

This will generate a standalone `.exe` file that you can run on any Windows machine.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

