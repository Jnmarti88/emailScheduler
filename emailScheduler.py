import openpyxl
import win32com.client as win32
import os
import datetime

# Read from the Excel spreadsheet
workbook = openpyxl.load_workbook('ExampleFile.xlsx') #name of excel file that has Store number, city, store, emails seperated by ;, date, time and day, can be changed to fit variables
worksheet = workbook.active

# Connect to Outlook
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')

# Choose the folder where you want to save the draft email
folder = namespace.Folders['EMAILHERE@EMAILDOMAINHERE.com'].Folders['Drafts'] #outlook email here

# Get the current user's default email account
account = namespace.Accounts[0]

# Get the signature file path
signature_path = os.path.join(os.getenv('APPDATA'), 'Microsoft/Signatures/default.htm') #default file name and path for outlook saved signature

# Read the signature file contents
with open(signature_path, 'r') as f:
    signature = f.read()

# Get the data from each row and create a draft email, I think the day variable is kinda useless and could be removed.
for row in worksheet.iter_rows(min_row=2, values_only=True):
    store_number, city, store, email, date, time, day = row

    # Format the date, there has to be a better way to do this, on excel I already formatted but still printed yyyy-mm-dd hh:mm:ss
    date = date.strftime('%m/%d/%y')

    # Format the date
    day_of_week = datetime.datetime.strptime(date, '%m/%d/%y').strftime('%A')

    # Create the email message
    message = outlook.CreateItem(0)
    message.Subject = 'Monthly audit'
    message.To = email

    # maybe have it refer to a file with message or part of exel sheet?
    message_body = f'Message here'

    message.Body = message_body

    # Get the existing HTML body and replace with just the signature contents
    html_body = message.HTMLBody
    html_body = html_body.replace('<body>', '').replace('</body>', '')
    html_body += signature
    message.HTMLBody = html_body

    # Save the draft email in the specified folder
    message.Save()
