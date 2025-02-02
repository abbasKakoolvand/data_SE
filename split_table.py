import pandas as pd
# Path to the input Excel file
input_file = "s50.xlsx"
sheet_name = "main"  # Name of the sheet containing the tables
# Read the Excel file
df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
# Function to split tables, clean them, and merge single-row tables
def split_and_merge_tables(dataframe):
    tables = []
    current_table = []
    for _, row in dataframe.iterrows():
        if row.isnull().all():  # Check if the row is entirely blank
            if current_table:  # If there's data in the current table, save it
                table_df = pd.DataFrame(current_table)
                # Remove empty columns (columns with all NaN values)
                table_df = table_df.dropna(axis=1, how='all')
                tables.append(table_df)
                current_table = []
        else:
            current_table.append(row.values)
    # Append the last table if it exists
    if current_table:
        table_df = pd.DataFrame(current_table)
        table_df = table_df.dropna(axis=1, how='all')  # Remove empty columns
        tables.append(table_df)
    # Merge single-row tables into the next table
    merged_tables = []
    i = 0
    while i < len(tables):
        if len(tables[i]) == 1 and i + 1 < len(tables):  # Single-row table
            # Merge the single-row table into the next table
            next_table = tables[i + 1]
            merged_table = pd.concat([tables[i], next_table], ignore_index=True)
            merged_tables.append(merged_table)
            i += 2  # Skip the next table since it's merged
        else:
            merged_tables.append(tables[i])
            i += 1
    return merged_tables
# Split the data into individual tables, clean them, and merge single-row tables
tables = split_and_merge_tables(df)
# Save each table into a separate sheet in a new Excel file
output_file = "D:\PycharmProjects\data_SE\output_tables_split.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for i, table in enumerate(tables, start=1):
        sheet_name = f"Table_{i}"  # Name each sheet as "Table_1", "Table_2", etc.
        table.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
print(f"All tables have been saved to {output_file}")