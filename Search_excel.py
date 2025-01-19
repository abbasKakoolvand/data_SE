import pandas as pd

text_to_find = "DBName".lower()
df = pd.read_excel("warehouse SE/TablesDataEDO_2.xlsx", sheet_name="TablesandColumnsinDW")
# columns_to_keep = ['DataBaseName', 'TableSchema', 'TableName', 'TableDescription', 'TableStatus', 'ColumnName', 'ColumnDescription', 'ColumnDataType', 'ColumnPosition']  # Replace with your actual column names
columns_to_keep = ['DataBaseName', 'TableSchema', 'TableName', 'ColumnName']  # Replace with your actual column names
# Filter the DataFrame to keep only the specified columns
filtered_df = df[columns_to_keep]
for index, row in filtered_df.iterrows():
    # print(row.astype(str).values)
    if text_to_find in str(row.astype(str).values).lower():
        print(f'Text found in row index: {index}')
        print(str(row.astype(str).values))
