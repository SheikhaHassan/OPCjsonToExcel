import pandas as pd


def merge_excel_sheets(input_file, output_file):
    # Read all sheets from the input Excel file into a dictionary of DataFrames
    sheets_dict = pd.read_excel(input_file, sheet_name=None)

    # Concatenate all DataFrames into one DataFrame
    combined_df = pd.concat(sheets_dict.values(), ignore_index=True)

    # Write the combined DataFrame to a new Excel file
    combined_df.to_excel(output_file, index=False)



input_excel_file = "final.xlsx"
output_excel_file = "final1.xlsx"
merge_excel_sheets(input_excel_file, output_excel_file)
