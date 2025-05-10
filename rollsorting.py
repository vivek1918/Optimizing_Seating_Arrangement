import pandas as pd

# Load the Excel file with all sheets
excel_file = 'cleaned_output_file.xlsx'
sheets = pd.read_excel(excel_file, sheet_name=None)

sheet_names = list(sheets.keys())

# --- Clean date columns in the first sheet ---
first_sheet_name = sheet_names[0]
df_first = sheets[first_sheet_name].copy()

# Convert datetime columns to date only
for col in df_first.columns:
    if pd.api.types.is_datetime64_any_dtype(df_first[col]):
        df_first[col] = df_first[col].dt.date

sheets[first_sheet_name] = df_first

# --- Sort the second sheet by roll number ---
second_sheet_name = sheet_names[1]
df_second = sheets[second_sheet_name]

# Sort by 'rollno' column (case insensitive)
df_sorted = df_second.sort_values(by='rollno', key=lambda x: x.str.upper())

sheets[second_sheet_name] = df_sorted

# Save all sheets back to a new file
with pd.ExcelWriter('sorted_output_file.xlsx', engine='openpyxl') as writer:
    for sheet_name, sheet_df in sheets.items():
        sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
