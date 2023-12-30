import pandas as pd
import re
import datetime

inv_excel = "open transactions.xlsx"
lookup_excel = "lookup.xlsx"
input_csv = 'input.csv'
results_excel = 'results.csv'
results_journal_template = 'results_journal_template.csv'

# Read the Excel file into a pandas DataFrame
df_inv_excel = pd.read_excel(inv_excel)
df_lookup_excel = pd.read_excel(lookup_excel)

lookup_dict = {}
for index, row in df_lookup_excel.iterrows():
    map_key = str(df_lookup_excel.iloc[index, 0]).strip()
    if map_key not in lookup_dict:
        lookup_dict[map_key] = str(df_lookup_excel.iloc[index, 1]).strip()
    else:
        raise ValueError("Error: Duplicate key found in lookup.xlsx")

in_customer_section = False
customer_inv_tuple_list = []

for index, row in df_inv_excel.iterrows():
    if isinstance(df_inv_excel.iloc[index, 0], str) and df_inv_excel.iloc[index, 0] == 'Date':
        customer_account = df_inv_excel.iloc[index - 1, 0]
        customer_inv_tuple_list.append((customer_account, customer_account))
        in_customer_section = True
    else:
        if pd.isnull(df_inv_excel.iloc[index, 0]) and pd.isnull(df_inv_excel.iloc[index, 1]) and pd.isnull(
            df_inv_excel.iloc[index, 2]):
            in_customer_section = False
            continue
        if in_customer_section and not pd.isnull(df_inv_excel.iloc[index, 2]):
            if df_inv_excel.iloc[index, 2].startswith('SVP') or df_inv_excel.iloc[index, 2].startswith('MOP') or \
                df_inv_excel.iloc[index, 2].startswith('PFTI'):
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
    # Append the result to the list of new column values
    new_column_values.append('')

    # Skip the row if the description starts with 'POS '
    if description.startswith('POS '):
        continue

    # Replace 'SUP' with 'SVP' in the description
    updated_description = re.sub(r'(?i)SUP', 'SVP', description)

    # Search the description for the pattern
    matches = pattern.finditer(updated_description)
    # Set the result to '' by default
    result = ''

    # If the pattern was found, set the result to the match; otherwise, set the result to ''
    for match in matches:
        keyword = match.group(1).upper()
        number = match.group(2)

        # Ensure the number is 8 digits long
        if len(number) < 8:
            number = '0' * (8 - len(number)) + number
        result = f'{keyword}{number}'

        # Iterate through the list of customer invoice tuples, and if the invoice number matches, set the result to
        # the customer account
        for customer_inv_tuple in customer_inv_tuple_list:
            if customer_inv_tuple[1] == result:
                new_column_values[-1] = customer_inv_tuple[0]
                break

        # If the result is not '' and the keyword is MAP, break out of the loop
        # This is to treat MAP as priority over MOP and SVP
        if new_column_values[-1] != '' and keyword == 'MAP':
            break

    # If the result is '', search the description for 1-8 digit numbers
    if new_column_values[-1] == '':
        pattern_digits = re.compile(r'\b\d{1,8}\b')
        matches_digits = pattern_digits.findall(description)
        for match in matches_digits:
            if len(match) <= 8:
                number = '0' * (8 - len(match)) + match
            result = f'SVP{number}'
            # Iterate through the list of customer invoice tuples, and if the invoice number matches, set the result
            # to the customer account
            for customer_inv_tuple in customer_inv_tuple_list:
                if customer_inv_tuple[1] == result:
                    new_column_values[-1] = customer_inv_tuple[0]
                    break

    if new_column_values[-1] == '':
        for key, value in lookup_dict.items():
            if value.lower() in description.lower():
                new_column_values[-1] = key
                break

# Set the 'MAPID' column to the list of new column values
df_2_csv['MAPID'] = new_column_values

# create a array with 10 empty strings
empty_array = [''] * 10

# Create a new dictionary to store the results in desired format
results_dict = {
    'Date': df_2_csv['Process date'].tolist(),
    'Voucher': [''] * len(df_2_csv),
    'Company': ['AUJW'] * len(df_2_csv),
    'Account': new_column_values,
    'Name': [''] * len(df_2_csv),
    'Description': [f"DD {datetime.datetime.strptime(date, '%d/%m/%Y').strftime('%d%m%y')}" for date in df_2_csv['Process date'].tolist()],
    'Debit': [0] * len(df_2_csv),
    'Credit': df_2_csv[' Credit'].tolist(),
    'Currency': ['AUD'] * len(df_2_csv),
    'Offset account type': ['Bank'] * len(df_2_csv),
    'Offset Non-ledger account': ['1'] * len(df_2_csv),
    'Offset main account': [''] * len(df_2_csv),
    'Method of payment': [''] * len(df_2_csv),
    'Payment reference': [''] * len(df_2_csv),
    'Line number': [''] * len(df_2_csv)
}

# Specify the path where you want to save the Excel file
output_file_path = f'{results_excel}'
output_journal_template_path = f'{results_journal_template}'

output_df = pd.DataFrame(results_dict)
output_df.to_csv(output_journal_template_path, index=False)

# Write the DataFrame to an Excel file
df_2_csv.to_csv(output_file_path, index=False)

print(f'Excel file "{output_file_path}" has been created.')
