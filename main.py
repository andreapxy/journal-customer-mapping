import glob
import pandas as pd
import re
from collections import defaultdict

# Define the directory path where you want to search for Excel files
directory_path = 'excel-files'

# Define the pattern for Excel files you want to process through
file_pattern = f'{directory_path}/*RECEPTING*.xlsx'

# Define the regex pattern for "MAP" followed by 8 digits
customer_ref_pattern = r'^MAP\d{8}$'

# Use glob to find Excel files matching the pattern in the directory
excel_files = glob.glob(file_pattern)

# Print the list of matching file names
print(f'Processing using the below excel files...')
count = 1
for file_path in excel_files:
  print(f'{count}. {file_path}')
  count += 1

customer_dict = defaultdict(list)
# Iterate through each file path in the list
for file_path in excel_files:
  # Read the Excel file into a DataFrame
  xls = pd.ExcelFile(file_path)

  # Get the list of sheet names in the Excel file
  sheet_names = xls.sheet_names
  print(f'\nProcessing {len(sheet_names)} sheets in {file_path}...')

  # Iterate through each sheet and print its contents
  for sheet_name in sheet_names:
    # Read the current sheet into a DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    for index, row in df.iterrows():
      try:
        customer_ref = df.iloc[index, 6]
        if isinstance(customer_ref, str) and re.match(customer_ref_pattern, customer_ref):
          print(f"'{customer_ref}' matches the pattern. Adding to dictionary: {df.iloc[index, 1]}")
          customer_dict[customer_ref].append(df.iloc[index, 1])
      except re.error as e:
        pass

output_data = {
  'CustomerRef': [],
  'Keyword': [],
}

for key in customer_dict:
  for value in customer_dict[key]:
    output_data['CustomerRef'].append(key)
    output_data['Keyword'].append(value)

# Create a DataFrame from the dictionary
df = pd.DataFrame(output_data)

# Specify the path where you want to save the Excel file
excel_file_path = 'output.xlsx'

# Write the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)

print(f'Excel file "{excel_file_path}" has been created.')