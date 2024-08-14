import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

def generate_contingency_table(dataset, output_filename="", maximum_entries=2**20, format_width=True,
                               sl_no_required=True, frequency_required=True, percentage_required=True,
                               cumulative_percentage_required=False, string_length_required=True):
    """
    Generate frequency tables for each column in a DataFrame and save them to an Excel file.

    This function helps in understanding the data by generating frequency tables for each column in the input DataFrame.
    Each column of the DataFrame is written to a separate sheet in an output Excel file, which includes various statistics
    like frequency, percentage, cumulative percentage, and string length.

    Parameters:
    ----------
    dataset : pandas.DataFrame
        The input DataFrame containing the data to analyze.

    output_filename : str, optional
        The name of the output Excel file (should end in ".xlsx"). If not provided, the function will generate a default
        filename in the format "frequency_table_<timestamp>.xlsx".

    maximum_entries : int, optional
        Maximum number of unique entries to include in the output for each column. Default is 2^20 (1,048,576).

    format_width : bool, optional
        If True, auto-formats the width of the columns in the output Excel file. Default is True.

    sl_no_required : bool, optional
        If True, includes a serial number column in the output. Default is True.

    frequency_required : bool, optional
        If True, includes a frequency column in the output. Default is True.

    percentage_required : bool, optional
        If True, includes a percentage column in the output. Default is True.

    cumulative_percentage_required : bool, optional
        If True, includes a cumulative percentage column in the output. Default is False.

    string_length_required : bool, optional
        If True, includes a string length column in the output. Default is True.

    Returns:
    -------
    None
        The function does not return anything; it writes the results directly to the specified Excel file.

    Notes:
    -----
    - Excel sheets have a limit of 1,048,576 rows. If the number of unique entries in a column exceeds this limit,
      only the top `maximum_entries` entries will be written to the output file.
    - The Excel sheet name for each column is limited to 31 characters. If a column name is longer, it will be truncated.

    Example:
    -------
    generate_contingency_table(dataset=iris, output_filename="frequency_table_iris.xlsx")
    """
    
    # Setting default output filename
    if output_filename == "":
        output_filename = f"frequency_table_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    if not output_filename.endswith(".xlsx"):
        output_filename = f"{output_filename}.xlsx"
    
    print(f"Writing frequency table of dataset to {output_filename}\n")
    
    # Creating the Excel writer object
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        
        for col in dataset.columns:
            sheet_name = col if len(col) <= 31 else f"{col[:15]}{col[-16:]}"
            
            # Calculating frequency table
            frequency_table = dataset[col].value_counts().reset_index()
            frequency_table.columns = [col, 'frequency']
            frequency_table = frequency_table.sort_values(by='frequency', ascending=False)
            
            # Calculating additional columns
            if percentage_required:
                frequency_table['percentage'] = 100 * frequency_table['frequency'] / frequency_table['frequency'].sum()
            if cumulative_percentage_required:
                frequency_table['cumulative_percentage'] = frequency_table['percentage'].cumsum()
            if sl_no_required:
                frequency_table['sl_no'] = range(1, len(frequency_table) + 1)
            if string_length_required:
                frequency_table['string_length'] = frequency_table[col].astype(str).apply(len)
            
            # Reordering columns
            if sl_no_required:
                column_order = ['sl_no', col, 'frequency']
            else:
                column_order = [col, 'frequency']
                
            if percentage_required:
                column_order.append('percentage')
            if cumulative_percentage_required:
                column_order.append('cumulative_percentage')
            if string_length_required:
                column_order.append('string_length')
                
            frequency_table = frequency_table[column_order]
            
            # Truncate the table if it exceeds maximum_entries
            frequency_table = frequency_table.head(maximum_entries)
            
            # Write the data to Excel
            frequency_table.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet = writer.sheets[sheet_name]
            
            # Auto-adjust column widths
            if format_width:
                for i, col in enumerate(frequency_table.columns):
                    max_length = max(frequency_table[col].astype(str).map(len).max(), len(col))
                    sheet.column_dimensions[get_column_letter(i + 1)].width = max_length + 2

            print(f"Generated frequency table for column {col}")
    
    print(f"\nFrequency table saved to {output_filename}")

