import re
from datetime import datetime
import warnings

# Avoid warnings
warnings.simplefilter("ignore")

try:
    import win32com.client as win32
    from pdfminer.high_level import extract_text
    from tkinter import filedialog as fd
    from tkinter import *
    import pandas as pd
    import openpyxl
    import xlsxwriter

except:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'pywin32'])
    pip(['install', 'pdfminer.six'])
    pip(['install', 'tk'])
    pip(['install', 'pandas'])
    pip(['install', 'openpyxl'])
    pip(['install', 'XlsxWriter'])

# Create the e-mail
def emailer(subject, body, attachment):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Attachments.Add(attachment)
    mail.GetInspector
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + \
        body + mail.HTMLbody[index + 1:]
    mail.Display()

# Open the File Dialog Box
root = Tk()
root.update()
pdfFileLocation = fd.askopenfile()
root.destroy()

# Extract the text in the PDF file
text = extract_text(pdfFileLocation.name)

# Access the text in the PDF file and look for the pattern below
partNumRegex = re.compile(r'[^0-9]\d{3}\s?\d{2}\s?\d{2}-?\d{2}[^0-9]')
partNums = re.findall(partNumRegex, text)

# Create a Python list
listPartNums = [partNum for partNum in partNums]

# Use Pandas to save the list as an Excel File
mail_attachment = f'N:\\Avdelning\\P\\PSA\\Inventory Management\\Python\\output_files\\{str(datetime.now().strftime(r"%d-%m-%y %H-%M"))} PDF Analysis.xlsx'

# Save the Excel file on a shared folder
pd.DataFrame(partNums, columns=["Pn"]).to_excel(mail_attachment)

mail_subject = 'PDF Analysis completed'
mail_body = f"""
    <p style=color:rgb(47,84,150);> Dear User,
    <br></br><br></br>
    Attached you will find an Excel file with the extracted data from the given PDF,
    <br></br><br></br>
    Given PDF: {pdfFileLocation.name},
    <br></br><br></br>
    Thanks,
"""

emailer(mail_subject, mail_body, mail_attachment)