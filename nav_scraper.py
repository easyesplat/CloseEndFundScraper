import requests
from bs4 import BeautifulSoup
import openpyxl
import constants

# Make a request to the page
page = requests.get('https://middlefield.com/funds/navs/')

# Parse the HTML content of the page
soup = BeautifulSoup(page.content, 'html.parser')

# Find the table containing the NAV data
nav_table = soup.find('table', class_='funds-table')

# Extract the table headers and rows
headers = [th.text.strip() for th in nav_table.find_all('th')]
headers.pop()
rows = []
for tr in nav_table.find_all('tr'):
    row = [td.text.strip() for td in tr.find_all('td')]
    if row:
        row.pop()
        rows.append(row)

# Print the table headers and rows
print(headers)
print(rows)

# Uploading data to Excel
workbook = openpyxl.load_workbook(constants.FILE_PATH + 'Middlefield_Analysis.xlsx')
worksheet = workbook['Middlefield']

for row_ind in range(len(rows) + 1):
    for col_ind in range(len(headers)):
        value = ''
        if row_ind == 0:
            value = headers[col_ind]
        else:
            value = rows[row_ind - 1][col_ind] 
        worksheet.cell(row = row_ind + 1, column = col_ind + 1).value = value

workbook.save(constants.FILE_PATH + 'Middlefield_Analysis.xlsx')