import pandas as pd

def clean_excel_file(input_file, output_file):
    # Read all sheets from the Excel file
    xl = pd.read_excel(input_file, sheet_name=None)

    cleaned_data = {}
    
    for sheet_name, df in xl.items():
        cleaned_df = df.copy()

        # Apply strip and cleanup to each cell if it's a string
        for col in cleaned_df.columns:
            if pd.api.types.is_datetime64_any_dtype(cleaned_df[col]):
                # Convert datetime to date (remove time part)
                cleaned_df[col] = cleaned_df[col].dt.date

            elif cleaned_df[col].dtype == 'object':
                cleaned_df[col] = cleaned_df[col].astype(str).apply(lambda x: ' '.join(x.strip().split()))        
        cleaned_data[sheet_name] = cleaned_df

    # Write cleaned data to new Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, cleaned_df in cleaned_data.items():
            cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)

# Example usage
input_path = 'C:/Users/Vivek Vasani/Desktop/Intern Proj - Optimal Seating Arrangement/Intern Proj/input_data_tt.xlsx'
output_path = 'cleaned_output_file.xlsx'
clean_excel_file(input_path, output_path)
