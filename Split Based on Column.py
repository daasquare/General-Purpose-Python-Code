import pandas as pd

def split_excel_sheet_by_column(input_file, output_file, sheet_name, column_name):
    # Read the specified sheet from the Excel file into a DataFrame
    df = pd.read_excel(input_file, sheet_name)

    # Get unique values in the specified column
    unique_values = df[column_name].unique()

    # Create a writer to save multiple DataFrames to different sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Iterate through unique values and save each subset to a separate sheet
        for value in unique_values:
            subset_df = df[df[column_name] == value]
            subset_df.to_excel(writer, sheet_name=str(value), index=False)

if __name__ == "__main__":
    # Replace 'input_file.xlsx', 'output_file.xlsx', 'sheet_name', and 'column_name'
    input_file_path = '/Users/akin.o.akinjogbin/Downloads/Module  FMEA.xlsx'
    output_file_path = '/Users/akin.o.akinjogbin/Downloads/Module  FMEA_split.xlsx'
    input_sheet_name = 'Module FMEA'
    column_to_split = 'Item/Component/ User'

    split_excel_sheet_by_column(input_file_path, output_file_path, input_sheet_name, column_to_split)

