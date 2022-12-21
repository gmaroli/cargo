#!/usr/local/bin/python3
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
import os


# Get the File Name

# File path—to be updated

# /Users/gmaroli/work/Cargo/data/vipiska_ind11.xlsx

def get_file_path():
    filename = input("Please provide the filename: ")
    cwd = os.getcwd()
    print('Current working directoy: ' + cwd)
    return cwd + '/data/' + filename


# vipiska_ind11.xlsx
input_file_name = get_file_path()
# print(input_file_name)


# Read the Excel File and extract only the columns needed
df_raw = pd.read_excel(input_file_name, usecols=['No',
                                                 'Reference',
                                                 'Number of pieces of freight',
                                                 'Consignee\'s contact name',
                                                 'Phone number',
                                                 'Postcode',
                                                 'City',
                                                 'Delivery address or depot\'s '
                                                 'address',
                                                 'Actual weight, kg'])

# print(df_raw.dtypes)

df_raw = df_raw.where(pd.notnull(df_raw), None)

# Replace blank postcode with 0 and blank city names with 'No City'
df_raw['City'] = df_raw['City'].fillna('NoCity')
df_raw = df_raw.fillna(0)

# Convert Postcode field into interger from float
df_raw['Postcode'] = df_raw['Postcode'].astype(int)

# Cleanse Reference column—to have only numbers
df_raw['Reference'] = df_raw['Reference'].str.replace('NE', '').str.replace('IND', '')

# print(df_raw)

# extract data by state
# Step 1 : Get list of all states in the sheet.
df_states = df_raw['City'].unique()

"""
# Function to create an empty Excel Sheet
def create_workbook(path):
    currentdate = datetime.now().strftime('%Y%m%d%H%M%S')

    # split the file name to append the current timestamp
    path = path.split('.')[0] + '_' + currentdate + '.xlsx'

    workbook = Workbook()
    workbook.save(path)
    return path


excel_file_name = create_workbook("Manifest.xlsx")

df_1 = pd.DataFrame([])
for i in range(0, len(df_states)):
    state = df_states[i]
    df_1 = df_raw.query('City=="' + state + '"')
    df_1 = df_1.reset_index(drop=True)
    df_1 = df_1.sort_values(by=['Reference'])
    with pd.ExcelWriter(excel_file_name, mode='a', if_sheet_exists='replace') as writer:
        df_1.to_excel(writer, sheet_name=state)
# print(df_1)
"""


# NEW VERSION TO WRITE TO INDIVIDUAL EXCEL FILE
def create_workbook(path):
    currentdate = datetime.now().strftime('%Y%m%d%H%M%S')

    # split the file name to append the current timestamp
    # path = path + '.xlsx'
    path = path.split('.')[0] + '_' + currentdate + '.xlsx'

    workbook = Workbook()
    workbook.save(path)
    return path


# for state in df_states:
#     create_workbook(state)

df_1 = pd.DataFrame([])
for i in range(0, len(df_states)):
    state = df_states[i]
    print(state)
    excel_file_name = create_workbook(state)
    df_1 = df_raw.query('City=="' + state + '"')
    df_1 = df_1.reset_index(drop=True)
    df_1 = df_1.sort_values(by=['Reference'])
    # print(df_1['City'])
    with pd.ExcelWriter(excel_file_name, mode='a', if_sheet_exists='replace') as writer:
        df_1.to_excel(writer, sheet_name=state)
