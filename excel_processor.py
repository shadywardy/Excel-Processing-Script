import pandas as pd
import os

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
        filter_column = input("Please enter the filter column (or type 'c' to quit): ")

        if filter_column.lower() == 'c':
            return False  # Exit the function to stop the script


        # Loop until the user enters a valid column name
        while filter_column not in df.columns:
            print(f"Error: The column '{filter_column}' does not exist in the Excel file.")
            filter_column = input("Please enter a valid filter column (or type 'exit' to quit): ")

            if filter_column.lower() == 'exit':
                return False  # Exit the function to stop the script

        # Get unique values in the column
        unique_values = df[filter_column].unique()

        # Loop over each unique value in the specified column and generate separate Excel files
        for value in unique_values:
            # Filter the dataframe for the current value
            filtered_df = df[df[filter_column] == value]

            # Define the output file name based on the unique value, ensuring it's a valid filename
            output_file = f'{value}.xlsx'.replace('/', '_').replace('\\', '_')  # Replace any invalid characters
            
            # Save the filtered dataframe to a new Excel file
            filtered_df.to_excel(output_file, index=False)

            print(f"Excel file generated for {value}: {output_file}")

    return True  # Continue to allow for another file input


# Main loop to process multiple files
while True:
    if not process_excel_file():
        break  # Exit the loop and end the script

print("Thank you for using the Excel processing script. Goodbye! /n this script build by shady wardy , Email:shadywardy@gmail.com ")

# Add this line at the end to keep the window open
input("Press Enter to exit...")

