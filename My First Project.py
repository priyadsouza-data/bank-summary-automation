import pandas as pd     # For working with Excel files and dataframes
import os               # To help with file path operations

# STEP 1: Load the main working file where we want to fill in data

working_file = pd.read_excel("Working File.xlsx")

# Clean and standardize column names
working_file.columns = working_file.columns.str.strip().str.upper()

# STEP 2: Create an empty list to collect all the data from the 5 customer files
bank_data = []

# STEP 3: Define the folder where all the bank files (Customer1.xlsx to Customer5.xlsx) are saved
folder_path = "Bank details"
#folder_path = r"C:\Users\Priya\Documents\Bank details"

# STEP 4: Read each customer file and extract needed data
for i in range(1, 7):
    file_path = os.path.join(folder_path, f"Customer{i}.xlsx")
    
    try:
        df = pd.read_excel(file_path, header=None)
        
        # Extract ECR
        try:
            ecr_line = str(df.iloc[4, 0])
            ecr_value = float(ecr_line.split(":")[1].replace('%', '').strip())
        except: 
            ecr_value = None

        # Extract account number
        try:
            acc_line = str(df.iloc[1, 0])
            acc_number = acc_line.split(":")[1].strip()
        except:
            acc_number = None

        # Extract account name
        try:
            accna_line = str(df.iloc[0, 0])
            acc_name = accna_line.split(":")[1].strip()
        except:
            acc_name = None
        # Read actual data table
        data = pd.read_excel(file_path, skiprows=8)
        try:
            # Read header row from index 6 (7th row in Excel)
            header_row = [str(cell).strip().upper() for cell in df.iloc[6].tolist()]
            

            # Find the index of 'AFP CODE'
            afp_col_index = header_row.index('AFP CODE')
            

            # Extract values from 8th to 18th row → index 7 to 17 (inclusive)
            afp_code_series = df.iloc[7:18, afp_col_index].reset_index(drop=True)
            

        except:C:
            afp_code_series = pd.Series([None] * 11)  # fallback 11 None values

        #Extract balance information
        try:
            # Read header row (7th row in Excel → index 6 in Python)
            header_row = [str(cell).strip().upper() for cell in df.iloc[6].tolist()]

            # Find the index of 'BALANCE INFORMATION' (case-insensitive)
            balance_col_index = header_row.index('BALANCE INFORMATION')

            # Extract values from row 8 to 18 (index 7 to 17), from the identified column
            balance_values = df.iloc[7:18, balance_col_index].reset_index(drop=True)

        except:
            balance_values = pd.Series([None] * 11)  # fallback 11 None values

        
        # Add metadata to each row
        data['EARNINGS CREDIT RATE'] = ecr_value
        data['ACCOUNT NUMBER'] = acc_number
        data['ACCOUNT NAME'] = acc_name
        data['AFP CODE'] = afp_code_series
        data['BALANCE INFORMATION'] = balance_values

        bank_data.append(data)

    except Exception as e:
        print(f"❌ Failed to process {file_path}: {e}")


# STEP 5: Combine all 5 customer DataFrames into one single DataFrame
bank_df = pd.concat(bank_data, ignore_index=True)
bank_df.columns = bank_df.columns.str.strip().str.upper()

columns_to_keep = [
    'ACCOUNT NAME', 'ACCOUNT NUMBER',
    'AFP CODE', 'BALANCE INFORMATION', 'EARNINGS CREDIT RATE'
]
bank_df = bank_df[[col for col in columns_to_keep if col in bank_df.columns]]

print(bank_df)

# STEP 6: Clean ECR column
if 'EARNINGS CREDIT RATE' in bank_df.columns:
    bank_df['EARNINGS CREDIT RATE'] = (
        bank_df['EARNINGS CREDIT RATE']
        .astype(str)
        .str.replace('%', '', regex=False)
        .astype(float)
    )

# STEP 7: Create wide format for AFP codes using BALANCE INFORMATION
afp_valid_codes = ['000331', '000000', '000010', '000040']
afp_df = pd.pivot_table(
    bank_df,
    index=['ACCOUNT NAME', 'ACCOUNT NUMBER'],
    columns='AFP CODE',
    values='BALANCE INFORMATION',
    aggfunc='sum'
).reset_index()

# STEP 8: Merge AFP summary back into bank_df
bank_df = bank_df.drop_duplicates(subset=['ACCOUNT NAME', 'ACCOUNT NUMBER']).reset_index(drop=True)
bank_df = pd.merge(bank_df, afp_df, on=['ACCOUNT NAME', 'ACCOUNT NUMBER'], how='left')

# STEP 9: BUSINESS RULE:Set AFP CODE - 000040 to 0 where ECR is 0
if '000040' in bank_df.columns:
    bank_df.loc[bank_df['EARNINGS CREDIT RATE'] == 0, '000040'] = 0

# STEP 10: Rename columns for consistency with working file
bank_df.rename(columns=lambda x: x.upper(), inplace=True)

# Map AFP codes to readable names for final merge
afp_rename_map = {
    '000331': 'AFP CODE - 000331',
    '000000': 'AFP CODE - 000000',
    '000010': 'AFP CODE - 000010',
    '000040': 'AFP CODE - 000040'
}
bank_df.rename(columns=afp_rename_map, inplace=True)

print("Working file columns:", working_file.columns.tolist())
print("Bank DF columns:", bank_df.columns.tolist())

working_file['ACCOUNT NUMBER'] = working_file['ACCOUNT NUMBER'].astype(str).str.strip()
bank_df['ACCOUNT NUMBER'] = bank_df['ACCOUNT NUMBER'].astype(str).str.strip()

# STEP 11: Merge bank data into working file on ACCOUNT NUMBER
merged = pd.merge(
    working_file,
    bank_df[
        [
            'ACCOUNT NUMBER', 'ACCOUNT NAME', 
            'AFP CODE - 000331', 'AFP CODE - 000000',
            'AFP CODE - 000010', 'AFP CODE - 000040',
            'EARNINGS CREDIT RATE'
        ]
    ],
    on='ACCOUNT NUMBER',
    how='left',
    suffixes=('', '_from_bank')
)

# STEP 12: Fill missing values from bank file
for col in ['ACCOUNT NAME', 
            'AFP CODE - 000331', 'AFP CODE - 000000',
            'AFP CODE - 000010', 'AFP CODE - 000040',
            'EARNINGS CREDIT RATE']:
    merged[col] = merged[col].combine_first(merged[col + '_from_bank'])

# STEP 13: Drop '_from_bank' columns
merged.drop(columns=[col + '_from_bank' for col in [
    'ACCOUNT NAME', 
    'AFP CODE - 000331', 'AFP CODE - 000000',
    'AFP CODE - 000010', 'AFP CODE - 000040',
    'EARNINGS CREDIT RATE']], inplace=True)

# STEP 13.1: Identify new customers not present in working file
existing_accounts = set(working_file['ACCOUNT NUMBER'].astype(str).str.strip())
bank_df_accounts = set(bank_df['ACCOUNT NUMBER'].astype(str).str.strip())

# Find new accounts to be added
new_accounts = bank_df_accounts - existing_accounts
new_entries = bank_df[bank_df['ACCOUNT NUMBER'].astype(str).str.strip().isin(new_accounts)]

# Keep only the columns that are present in the working file
new_entries_filtered = new_entries[
    [col for col in merged.columns if col in new_entries.columns]
]

# Reorder columns as per the working file
new_entries_filtered = new_entries_filtered.reindex(columns=merged.columns)

# Append new entries to the merged DataFrame
merged = pd.concat([merged, new_entries_filtered], ignore_index=True)

# STEP 14: Save the final file
merged.to_excel("Working File.xlsx", index=False)

# STEP 15: Confirmation
print("✅ Updated working file saved successfully.")