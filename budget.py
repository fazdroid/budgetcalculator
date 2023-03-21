from pathlib import Path  #core python module
import win32com.client  #pip install pywin32

import requests
import pdfplumber
import re
from collections import namedtuple
import pandas as pd
import sqlite3
import PyPDF2
import glob
import os
import numpy as np


# Create output folder
output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Connect to folder
inbox = outlook.GetDefaultFolder(6)

# Get the latest message with the specific subject
message = None
for item in inbox.Items:
    if item.Subject == "XXXX e-Statement":
        if message is None or item.CreationTime > message.CreationTime:
            message = item

# Check if there is any message with the specific subject
if message is not None:
    # Get subject, body, and attachments
    subject = message.Subject
    body = message.body
    attachments = message.Attachments

    # Create separate folder for the message
    target_folder = output_dir / str(subject)
    target_folder.mkdir(parents=True, exist_ok=True)

    # Write body to text file
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    # Save attachments
    for attachment in attachments:
        attachment.SaveAsFile(target_folder / str(attachment))
else:
    print("No message with the specific subject found.")


#creating a new SQL database
conn = sqlite3.connect('bank.sqlite')
cur = conn.cursor()

#creating the table in the SQL database
cur.execute('DROP TABLE IF EXISTS Bank')
cur.execute('''CREATE TABLE Bank (Date TEXT, Description TEXT, Debit FLOAT, Balance FLOAT)''')


def read_pdf(file_path, password):
    with pdfplumber.open(file_path, password=password) as pdf:
        pages = []
        for page in pdf.pages:
            pages.append(page.extract_text())
    return pages

folder_path = 'C:/Users/fayaz/OneDrive/Desktop/Code/BUDGET CALCULATOR/Output/ADCB e-Statement'
pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
if pdf_files:
    latest_pdf = max(pdf_files, key=os.path.getctime)
    password = "XXXXXXX"
    pdf_pages = read_pdf(latest_pdf, password)
else:
    print("No pdf files found in the specified folder.")


#with pdfplumber.open(pdffile,password="11296797") as pdf:

#iterates through second page to the last page
for text in pdf_pages[1:]:
    for row in text.split('\n'):
        #print(row)
        if re.search('^[0-9]+/', row):
            items = row.split()
            # perform the rest of your operations here with the 'items' list
            #print(items)

            #if the second item (i.e. the one after the date) is B/F, write the starting balance
            if items[1]=="B/F":
                startingbalance = float(items[3].replace(',',''))

                #startingbalance has to have a comma after because on its own its an integer, with a comma it is a tuple of length one
                cur.execute('''INSERT OR IGNORE INTO Bank (Description, Balance) VALUES (?, ?)''', ("Starting Balance:", startingbalance))
                #print("Starting Balance:", startingbalance, '\n') - commented out for SQL

                #set a count to 0 here - if this count was anywhere outside of this if loop, it would keep resetting. Here the count set to 0 is contained, and this loop is never iterated through again as there is no other line with "B/F"
                count = 0

            #or else, if the length of the list is greater than 4 elements:
            elif len(items)>=5:

                #before anything, if the second item is ":CORR", skip it
                if items[2] ==":CORR":
                    continue

                #if the second item is not ":CORR", then:
                else:

                    #take the first, fourth, fifth, second to last and last elements of the list and store them in their respective variables
                    date = items[0]
                    desc1 = items[3]
                    desc2 = items[4]

                    #debit and balance needs to be turned into a float, but the comma also needs to be removed first
                    debit = float(items[-2].replace(',',''))
                    balance = float(items[-1].replace(',',''))

                    desc = desc1 + ' ' + desc2

                    #putting all the variables in the table under the relevant column headers
                    cur.execute('''INSERT OR IGNORE INTO Bank (Date, Description, Debit, Balance) VALUES (?, ?, ?, ?)''', (date, desc, debit, balance))

                    conn.commit()

                    #print('Date:',date)
                    #print('Description:',desc1, desc2)
                    #commented both out for SQL

                    #only for the first iteration (when the count is 0)
                    if count==0:
                        if startingbalance < balance:
                            #so if the balance increased from the start, money went in, therefore it's credit, not debit

                            #print('Credit:',debit) - commented out for SQL



                            #created new variable amount so we can use this to compare the last value in the loop to determine if its debit or Credit
                            #plus debit because money is going in
                            amount = startingbalance + debit
                        else:
                            #if money didn't go in, it came out, therefore debit

                            #print('Debit:',debit) - commented out for SQL

                            #minus debit as money came out
                            amount = startingbalance - debit
                        #count will be added to, making it 1, and therefore this loop will never be iterated through again (as the count will never be 0 again)
                        count = count + 1

                    #if the count is not 0, i.e. not the first iteration
                    else:
                        #after loop starts again, storing the new balance value, it compares it with the amount variable stored from the first loop
                        if amount > balance:
                            #if the amount (basically the previous balance) is greater than the balance now, then money went down, therefore it's debit

                            #print('Debit:',debit) - commented out for SQL

                            #then update the variable amount with the current balance
                            amount = balance
                        #if the opposite is true, it means the current balance is greater than the previous balance (called amount), meaning money went up and is therefore credit
                        elif amount < balance:

                            #print('Credit:',debit) - commented out for SQL

                            #again update the variable with the current balance
                            amount = balance

                    #print('Balance:',balance)
                    #print('\n')
                    #commented both out for SQL


            #if it's not the first iteration, or the length is less than 4, just print the whole row
            else:
                #print(row, '\n') - commented out for SQL
                continue


#code from here is to transfer sql database into excel

# Connect to the SQL database
conn = sqlite3.connect('bank.sqlite')

# Read the data from the SQL database into a pandas dataframe
df = pd.read_sql_query("SELECT * from bank", conn)

# Write the data to an Excel spreadsheet
#can also change directory of excel sheet here (e.g. df.to_excel(r'C:\Users\YourUsername\Documents\bank_statement_files\bank_statement.xlsx', index=False))
df.to_excel('budgetexcel.xlsx', index=False)

# Close the database connection
conn.close()





# Read the data from the Excel sheet into a pandas dataframe
df = pd.read_excel('budgetexcel.xlsx')

df['Category'] = np.select([
    df['Description'].str.contains('ADNOC|ENOC|EMARAT'),
    df['Description'].str.contains('SPINNEYS|Carrefour|GRANDIOSE'),
    df['Description'].str.contains('KINGS|MEDICNA'),
    df['Description'].str.contains('EMICOOL|DEWA|SmartDXBGo|DU|BROTHERS'),
    df['Description'].str.contains('Amazon|DAY'),
    df['Description'].str.contains('TALABAT|MCDONALDS|STARBUCKS'),
    df['Description'].str.contains('Virgin')
], [
    'Petrol',
    'Groceries',
    'Medical',
    'Bills',
    'Shopping',
    'Food',
    'Phone'
], default='Other')

categories = ['Petrol', 'Groceries', 'Medical', 'Bills', 'Shopping', 'Food', 'Phone', 'Other']

# Create a new column to hold the category totals
df['Category Total'] = 0

# Loop through each category and calculate the total
for category in categories:
    total = df.loc[df['Category'] == category, 'Debit'].sum()
    df.loc[df['Category'] == category, 'Category Total'] = total

# Write the updated dataframe to the Excel file
df.to_excel('budgetexcel.xlsx', index=False)
