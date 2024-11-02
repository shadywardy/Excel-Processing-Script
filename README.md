# Excel Processing Script

Welcome to the Excel Processing Script! This handy tool is designed to streamline your data management tasks by allowing you to filter Excel files and create separate files based on unique values in specified columns.

## Features

- **Easy Input**: Load your Excel file with a simple input prompt.
- **Dynamic Filtering**: Choose any column to filter the data dynamically.
- **Unique Value Extraction**: Automatically generates new Excel files for each unique value in the specified column.
- **User-Friendly**: Prompts you at each step, ensuring a smooth user experience.
- **Exit Options**: Easily exit the program at any time by typing 'exit' or 'c'.

## Installation

To get started, you’ll need Python and the `pandas` library. If you don’t have them installed, you can set them up using the following commands:

```bash
pip install pandas
```

## Usage

1. **Run the Script**: Launch the script using Python.
2. **Input the Excel File Path**: When prompted, enter the full path to your Excel file. You can also type `'exit'` to quit the program.
3. **Choose a Filter Column**: Enter the name of the column you wish to filter by. If you want to quit, simply type `'c'`.
4. **Check for Valid Column**: If the column does not exist, you’ll be prompted to enter a valid column name.
5. **Receive Output**: The script will generate separate Excel files for each unique value in the specified column, saving them in the current directory.
6. **Exit**: Type `'exit'` or `'c'` to stop the script at any time.

### Example

Here’s an example of how to use the script:

```bash
python excel_processor.py
```

When prompted:
```
Please input the Excel file path (or type 'exit' to quit): "D:\YourPath\YourExcelFile.xlsx"
Inputting file, Please Wait: D:\YourPath\YourExcelFile.xlsx
Please enter the filter column (or type 'c' to quit): "Company"
```

The script will create separate Excel files for each unique company listed in the "Company" column.

## Output

For each unique value in the chosen column, the script generates an Excel file named after that value, ensuring any invalid characters in the filename are replaced. For example, if the unique value is "ABC Corp", it will create a file named `ABC_Corp.xlsx`.

## Executable Version

If you prefer not to run the script from the command line, an executable version is also available for download. Simply click the link below to download the `.exe` file:

[Download Excel Processor Executable](https://www.mediafire.com/file/adbi587f55gtrdy/excel_processor_exe_file.7z/file)

## Contribution

This script is built by **Shady Wardy**. If you have suggestions or improvements, feel free to fork the repository and submit a pull request!

## Contact

For any inquiries, you can reach me at:  
**Email**: [shadywardy@gmail.com](mailto:shadywardy@gmail.com)

Thank you for using the Excel Processing Script. Enjoy organizing your data!

