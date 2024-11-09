import pandas as pd
import os
import pyexcel as p
import pyexcel_xls


def process_excel_file():
    # Load the Excel file
    input_file = input("Please input the Excel file path (or type 'c' to quit): ").strip('"')  # Remove surrounding quotes
    
    if input_file.lower() == 'c':
        return False  # Exit the function to stop the script

    print(f"Inputing file , Please Wait: {input_file}")

    # Check if the file exists
    if not os.path.exists(input_file):
        print(f"Error: The file '{input_file}' does not exist.")
        return True  # Continue to ask for another file

    df = pd.read_excel(input_file)

    while True:
        # Specify the column you want to filter by
        filter_column = input("Please enter the filter column: or type 'c' to exit: ")
        
        if filter_column.lower() == 'c':
                return False  # Exit the function to stop the script
        
        file_name=input("Please Input file name (file_name-filter_column): ")
        
        # Choose the file extension for saving
        extension = input("For extension, please press:\n1 for .xls\n2 for .xlsx\n")
        
        # Validate and set the file extension
        if extension == '1':
            file_extension = '.xls'
        elif extension == '2':
            file_extension = '.xlsx'
        else:
            print("Invalid input. Please enter '1' for .xls or '2' for .xlsx.")
            continue
        
        print(f"Selected extension: {file_extension}")
        
        
        

        # Loop until the user enters a valid column name
        while filter_column not in df.columns:
            print(f"Error: The column '{filter_column}' does not exist in the Excel file.")
            filter_column = input("Please enter a valid filter column (or type 'c' to quit): ")

            if filter_column.lower() == 'c':
                return False  # Exit the function to stop the script

        # Get unique values in the column
        unique_values = df[filter_column].unique()

        # Loop over each unique value in the specified column and generate separate Excel files
        for value in unique_values:
            # Filter the dataframe for the current value
            filtered_df = df[df[filter_column] == value]
            
            records = filtered_df.to_dict(orient='records')

            output_file = f'{file_name}_{value}{file_extension}'.replace('/', '_').replace('\\', '_')  # Replace any invalid characters

                

            
            # Save the filtered data to an Excel file in the selected format
            if file_extension == '.xls':
                p.save_as(records=records, dest_file_name=output_file)
            else:
                filtered_df.to_excel(output_file, index=False)  # .xlsx format with pandas

            print(f"Excel file generated for {value}: {output_file}")

    return True  # Continue to allow for another file input


# Main loop to process multiple files
while True:
    if not process_excel_file():
        break  # Exit the loop and end the script

print("Thank you for using the Excel processing script. Goodbye! /n this script build by shady wardy , Email:shadywardy@gmail.com ")

# Add this line at the end to keep the window open
input("Press Enter to exit...")

