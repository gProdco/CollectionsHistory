import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)
from tabulate import tabulate

###identify and user input file location
print('filepath example >> C:\\users\\name\\OneDrive - Prodco Analytics\\Desktop\\Collection Status.xlsx')
user_input = input('Please enter the location of your collection status document: ')
emails_df = pd.read_excel(user_input, sheet_name = None)

### delete old sheets
target_sheet = '4-8-2024'
sheet_names = list(emails_df.keys())
target_index = sheet_names.index(target_sheet)
emails_df = {key: emails_df[key] for key in sheet_names[target_index:]}

### create master sheet
master_df = pd.concat(emails_df.values(), ignore_index=True)
master_df = master_df.drop(columns = ['Current', '< 1 Month', '1 Month', '2 Months', '3 Months', 'Older', 'Unnamed: 12'])
master_df = master_df.drop(master_df.columns[-1], axis=1)
master_df = master_df.dropna(subset=['1st', '2nd', '3rd'], how='all')

### user inputs client, pull sent emails for client

while True:
    lookup_client = input('Which client would you like collections emails data for?: ')
    lookup_client_df = master_df[master_df['Name'].str.lower() == lookup_client.strip().lower()]
    if not lookup_client_df.empty:
        lookup_client_df = lookup_client_df.copy()
        lookup_client_df.loc[:, 'EmailSent'] = None
        lookup_client_df.loc[:, 'EmailDate'] = None
        lookup_client_df = lookup_client_df.iloc[::-1].reset_index(drop=True)
        break
    else:
        print('Client not found, please try again')

for index, row in lookup_client_df.iterrows():
    if pd.notna(row['3rd']):
        lookup_client_df.at[index, 'EmailSent'] = '3rd'
        lookup_client_df.at[index, 'EmailDate'] = row['3rd']
    elif pd.notna(row['2nd']):
        lookup_client_df.at[index, 'EmailSent'] = '2nd'
        lookup_client_df.at[index, 'EmailDate'] = row['2nd']
    elif pd.notna(row['1st']):
        lookup_client_df.at[index, 'EmailSent'] = '1st'
        lookup_client_df.at[index, 'EmailDate'] = row['1st']

# Show the updated DataFrame
cols = [col for col in lookup_client_df.columns if col != 'Comments']
cols.append('Comments')
lookup_client_df = lookup_client_df[cols]
lookup_client_df = lookup_client_df.drop(['1st', '2nd', '3rd'], axis=1)
lookup_client_df['EmailDate'] = pd.to_datetime(lookup_client_df['EmailDate'], errors='coerce')
lookup_client_df['EmailDate'] = lookup_client_df['EmailDate'].dt.date
lookup_client_df['EmailDateShift'] = lookup_client_df['EmailDate'].shift()
lookup_client_df = lookup_client_df[lookup_client_df['EmailDate'] != lookup_client_df['EmailDateShift']]
lookup_client_df = lookup_client_df.drop(columns=['EmailDateShift'])
lookup_client_df = lookup_client_df.reset_index(drop=True)
lookup_client_df.index = lookup_client_df.index + 1

print(tabulate(lookup_client_df, headers='keys', tablefmt='fancy_grid', showindex=False))
