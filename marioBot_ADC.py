from datetime import datetime
import warnings

# Avoid warnings
warnings.simplefilter("ignore")

try:
    import win32com.client as win32
    import pandas as pd
    import numpy as np
    import pyodbc
    import pyautogui as bot
    import tk


except ImportError:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'pywin32'])
    pip(['install', 'pandas'])
    pip(['install', 'numpy'])
    pip(['install', 'pyodbc'])
    pip(['install', 'pyautogui'])
    pip(['install', 'tk'])

    import win32com.client as win32
    import pandas as pd
    import numpy as np
    import pyodbc
    import pyautogui as bot

# Create the e-mail
def emailer(subject, body, recipient, attachment):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Attachments.Add(attachment)
    mail.GetInspector
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + \
        body + mail.HTMLbody[index + 1:]
    mail.Display()

# Log in to access Rex screen (via Mario) and Rex DB (via SQL)
NDC_USER = bot.prompt(text='Please write your Rex NDC user.', title='User ID')
NDC_PWD = bot.password(text='Please write your Rex NDC password.', title='Password')

# Files locations
main_path = '\\\\SW03000.se.hvwan.net\\Dept\\P\\PSA\\Inventory Management\\Python\\'

# Read the SQL file
with open(main_path + 'sql_code\\sql05&ADC.sql', 'r') as sql:
    query = sql.read()
    sql.close()

# List of customers that should receive the availability list of ADC
dict_customers = {
    "AustraliaNewZeeland": ['     1396', '     4853'],
    "Brasil": '     2808',
    "AndeanRegion": ['     6684' , '     6734' , '     6742'],
    "India": '   100180',
    "SouthAfrica": '     6452',
    "SouthEastAsia": ['   368548', '   300285', '   316190'
                    , '   368019', '   218156', '   222045'
                    , '   323279', '   223597', '   100175'
                    , '   363861', '   311894', '   304642']
}

# SouthEastAsia =   '   368548' -- Hong Kong, '   300285' -- Woochang South Korea, '   218156' -- Singapure,
#                   '   316190' -- Kyung South Korea, '   368019' -- Thailand, '   222045' -- Fiji,
#                   '   323279' -- Vanuatu, '   223597' -- Timor Leste, '   100175' -- China,
#                   '   363861' -- Indonesia, '   311894' -- Taiwan, '   304642' -- Vietnam.

# Send one e-mail per customer in dict_customers
for customer_name, customer_number in dict_customers.items():
    query_sql = query\
        .replace('REPLACE_1', str(customer_name))\
        .replace('REPLACE_2', str(customer_number))\
        .replace("'[", '').replace("]'", '')

    # Connect to Rex and query the database, and close the connection to Rex
    with pyodbc.connect(driver='{iSeries Access ODBC Driver}', system='rex-ndc.hvwan.net', uid=NDC_USER, pwd=NDC_PWD) as connx:
        df = pd.read_sql(query_sql, connx)
    connx.close()

    # Avoid Nones in the columns with Availability
    df[["Avail. 05 (Q)", "Avail. ADC (Q)"]] = df[["Avail. 05 (Q)", "Avail. ADC (Q)"]].replace(np.nan, "Not set up")

    name_excel = main_path + "output_files\\" + str(datetime.now().strftime(r"%d-%m-%y")
                                        ) + f' - ADC_&_05_{customer_name}.xlsx'

    # Export dataset to XLSX
    with pd.ExcelWriter(name_excel, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name="ADC_&_05_result", index=False)

        # Auto-adjust columns' width
        for column in df:
            column_width = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets["ADC_&_05_result"].set_column(col_idx, col_idx, column_width)

    # Create an e-mail with the information subject, body, attachment and receiver
    mail_subject = f"ADC and 05 availability {customer_name} " + str(datetime.now().strftime(r"%d-%m-%y"))
    mail_body = """
    <p style=color:rgb(47,84,150);> Dear Sales Company,
    <br></br><br></br>
    Attached you will find the spare parts availability for 05 and ADC,
    <br></br><br></br>
    This list includes only parts coming from Loncin, Weima, and other Chinese suppliers,
    <br></br><br></br>
    Our goal is to deplete the stock in 05, and use ADC as the main warehouse for these assortments,
    <br></br><br></br>
    Thanks for your cooperation,
    </p>"""

    if customer_name == "AustraliaNewZeeland":
        mail_receiver = "chris.briggs@husqvarnagroup.com; glenda.murray@husqvarnagroup.com>; debbie.martin@husqvarnagroup.com;" + \
        " vanessa.dell@husqvarnagroup.com; erika.pasztuhov@husqvarnagroup.com;"
    elif customer_name == "Brasil":
        mail_receiver = "fabio.almeida@husqvarnagroup.com; rafhael.diniz@husqvarnagroup.com; ammar.abdo@husqvarnagroup.com;"
    elif customer_name == "AndeanRegion":
        mail_receiver = "oscar.urrea@husqvarnagroup.com; max.torres@husqvarnagroup.com; maritza.basto@husqvarnagroup.com;" + \
            "connie.cantellano@husqvarnagroup.com; carlos.leon@husqvarnagroup.com; ammar.abdo@husqvarnagroup.com;" + \
            "Angel.Torres@husqvarnagroup.com"
    elif customer_name == "India":
        mail_receiver = "arun.m@husqvarnagroup.com; elin.oskarsson@husqvarnagroup.com;"
    elif customer_name == "SouthAfrica":
        mail_receiver = "yolanda.noppe@husqvarnagroup.com; jenny.krantz@husqvarnagroup.com;"
    elif customer_name == "SouthEastAsia":
        mail_receiver = "julie.ang@husqvarnagroup.com; cheewooi.geh@husqvarnagroup.com; jim.andersson@husqvarnagroup.com; merve.turan@husqvarnagroup.com"

    mail_attachment = name_excel
    emailer(mail_subject, mail_body, mail_receiver, mail_attachment)

    print(customer_number, customer_name, '----- done')