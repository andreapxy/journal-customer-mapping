import pandas as pd
import re

inv_excel = "open transactions.xlsx"
input_csv = 'input.csv'
results_excel = 'results.csv'

# Read the Excel file into a pandas DataFrame
df_inv_excel = pd.read_excel(inv_excel)

in_customer_section = False
customer_inv_tuple_list = []

for index, row in df_inv_excel.iterrows():
  if isinstance(df_inv_excel.iloc[index, 0], str) and df_inv_excel.iloc[index, 0] == 'Date':
    customer_account = df_inv_excel.iloc[index - 1, 0]
    customer_inv_tuple_list.append((customer_account, customer_account))
    in_customer_section = True
  else:
    if pd.isnull(df_inv_excel.iloc[index, 0]) and pd.isnull(df_inv_excel.iloc[index, 1]) and pd.isnull(df_inv_excel.iloc[index, 2]):
      in_customer_section = False
      continue
    if in_customer_section == True and not pd.isnull(df_inv_excel.iloc[index, 2]):
      if df_inv_excel.iloc[index, 2].startswith('SVP') or df_inv_excel.iloc[index, 2].startswith('MOP') or df_inv_excel.iloc[index, 2].startswith('PFTI'):
        customer_inv_tuple_list.append((customer_account, df_inv_excel.iloc[index, 2]))

# Read the Bank Transactions CSV file into a pandas DataFrame
df_2_csv = pd.read_csv(input_csv)

# Create a regular expression pattern to match the MAP, MOP, and SVP keywords
pattern = re.compile(r'\b(MAP|MOP|SVP|PFTI)\s*(\d+)\b', re.IGNORECASE)

# Create a new column called 'MAPID' and set it to ''
new_column_values = []

for index, row in df_2_csv.iterrows():
  # Get the description from the current row
  description = df_2_csv.iloc[index, 1]
  # Search the description for the pattern
  matches = pattern.finditer(description)
  # Set the result to '' by default
  result = ''
  # Append the result to the list of new column values
  new_column_values.append(result)
  # If the pattern was found, set the result to the match; otherwise, set the result to ''
  for match in matches:
    keyword = match.group(1).upper()
    number = match.group(2)
    
    # Ensure the number is 8 digits long
    if len(number) < 8:
        number = '0' * (8 - len(number)) + number
    result = f'{keyword}{number}'

    # Iterate through the list of customer invoice tuples, and if the invoice number matches, set the result to the customer account
    for customer_inv_tuple in customer_inv_tuple_list:
      if customer_inv_tuple[1] == result:
        new_column_values[-1] = customer_inv_tuple[0]
        break
        
    # If the result is not '' and the keyword is MAP, break out of the loop
    # This is to treat MAP as priority over MOP and SVP
    if new_column_values[-1] != '' and keyword == 'MAP':
      break

# Set the 'MAPID' column to the list of new column values
df_2_csv['MAPID'] = new_column_values

# Specify the path where you want to save the Excel file
output_file_path = f'{results_excel}'

# Write the DataFrame to an Excel file
df_2_csv.to_csv(output_file_path, index=False)

print(f'Excel file "{output_file_path}" has been created.')