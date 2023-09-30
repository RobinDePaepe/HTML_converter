import pandas as pd 
from bs4 import BeautifulSoup
import re
import html_text
import time 
import os
import sys



# Get a list of all files in the directory
directory_path = os.getcwd()
files_in_directory = os.listdir(directory_path)
print(f'Files in directory: {files_in_directory}')

# selecting the correct file
filename = f'{input("Please type the file you would like to convert the HTML text of: ")}'
print(f'Selected file: {filename}')

if filename in files_in_directory:
    print(f"The file '{filename}' was found in the directory.")
else:
    print(f"The file '{filename}' was not found in the directory.")
    sys.exit(1)
    
xls = pd.ExcelFile(filename)

# List all sheet names
sheet_names = xls.sheet_names
print("Available sheets:")
for sheet_name in sheet_names:
    print(sheet_name)

sheet = input('Specify the sheet which contains the data you want to convert: ')

# List all columns in the selected sheet
df = pd.read_excel(filename, sheet_name=sheet)
columns = df.columns.tolist()
print(f"\nColumns in '{sheet}':")
for column in columns:
    print(column)
    
col = input('Specify the collumn which contains the html text: ')

## Actual converting script
start_time = time.time()
df = pd.read_excel(filename)
print('Number of e-mails to convert:', len(df[col]))
df[col] = df[col].fillna('')

df['plain_text_column'] = df[col].apply(html_text.extract_text)
df['len_plain_text'] = df['plain_text_column'].apply(len)

#df['cleaned_text'] = df['plain_text_column'].replace('\n', '\n\n', regex=True)
#print(df[0:1].cleaned_text.values[0])

writer = pd.ExcelWriter(f'{filename[:-5]}_converted.xlsx', engine='openpyxl')
df.to_excel(writer, index=False, sheet_name=sheet)
workbook = writer.book
worksheet = writer.sheets[sheet]
#Set the column width to make sure the text fits properly


#Save the Excel file
writer._save()
print(f"HTML formatted to clean text! Please check the {filename[:-5]}_converted.xlsx' file.")
end_time = time.time()
script_runtime = end_time - start_time
print(f"Scripted executed in {script_runtime:.2f} seconds")