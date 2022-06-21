"""
This script uses PyAutoGui to control the mouse and keyboard and release orders to South Africa and the UK
The logic is written on a PDF as a flowchart, as a better way to communicate (not shared)

Author: Mauricio Montilla
"""

from datetime import datetime
from re import X
import time
import os
import warnings

# Avoid warnings
warnings.simplefilter("ignore")

try:
    from PIL.ImageOps import grayscale
    import win32com.client as win32
    import pandas as pd
    import numpy as np
    import pyodbc
    import pyautogui as bot
    import pyperclip
    import keyboard
    import cv2
    import openpyxl
    import xlsxwriter
    import tk

except ImportError:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'Pillow'])
    pip(['install', 'pywin32'])
    pip(['install', 'pandas'])
    pip(['install', 'numpy'])
    pip(['install', 'pyodbc'])
    pip(['install', 'pyautogui'])
    pip(['install', 'pyperclip'])
    pip(['install', 'keyboard'])
    pip(['install', 'opencv-python'])
    pip(['install', 'openpyxl'])
    pip(['install', 'XlsxWriter'])
    pip(['install', 'tk'])

    from PIL.ImageOps import grayscale
    import win32com.client as win32
    import pandas as pd
    import numpy as np
    import pyodbc
    import pyautogui as bot
    import pyperclip
    import keyboard
    import cv2

# Measure the time that Mario needs
start_time = time.time()
time.sleep(2)

# Set the precision of the Data Frames to only two decimals and remove warning
pd.set_option('display.precision', 2)
pd.set_option('float_format', '{:.0f}'.format)
pd.options.mode.chained_assignment = None

# Log in to access Rex screen (via Mario) and Rex DB (via SQL)
NDC_USER = bot.prompt(text='Please write your Rex NDC user.', title='User ID')
NDC_PWD = bot.password(text='Please write your Rex NDC password.', title='Password')

# Files locations
main_path = '\\\\SW03000.se.hvwan.net\\Dept\\P\\PSA\\Inventory Management\\Python\\'
image_8_se = main_path + 'images\\G35_SA\\8_supplier_se.png'
image_8_ch = main_path + 'images\\G35_SA\\8_supplier_ch.png'
image_8_uk = main_path + 'images\\G35_SA\\8_supplier_uk.png'
image_8_uk_pro = main_path + 'images\\G35_SA\\8_supplier_uk_pro.png'
image_9 = main_path + 'images\\G35_SA\\9_in_G35.png'
image_10 = main_path + 'images\\G35_SA\\10_back_to_first_page.png'
image_11_se = main_path + 'images\\G35_SA\\11_supplier_air_se.png'
image_11_ch = main_path + 'images\\G35_SA\\11_supplier_air_ch.png'
image_11_uk = main_path + 'images\\G35_SA\\11_supplier_air_uk.png'
image_12 = main_path + 'images\\G35_SA\\12_exeption_handling.png'

# Locate the center of the image that it gets, try during 5 seconds
def bot_control(image, conf, grayscale=True):
    cords = None
    counter = 0

    while cords == None:
        cords = bot.locateCenterOnScreen(
            image, grayscale=grayscale, confidence=conf)

        # Wait five seconds, if we done find the image break
        counter += 1
        if counter == 5: break
        time.sleep(1)

    if cords != None:
        bot.moveTo(cords, duration=2), time.sleep(1)
    else:
        print('Mario cannot find ' + image)

    return cords

# Write down the decision made by Mario
def decision(key, wait):
    bot.typewrite(str(key)), time.sleep(wait)
    bot.press('enter'), time.sleep(wait)

# Interact with Rex to move to the next part in the list
def next_part_number(key, wait):
    decision(key, wait)
    bot.press('enter'), time.sleep(wait)

# Select everything except part numbers
def select_item(key, wait):
    bot.click(), time.sleep(wait)
    bot.hotkey('shift', 'tab')
    decision(key, wait)

# Create the mail report at the end of the interaction
def emailer(subject, body, recipient, attachment_1, attachment_2):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.Attachments.Add(attachment_1)
    mail.Attachments.Add(attachment_2)

    mail.GetInspector
    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + \
        body + mail.HTMLbody[index + 1:]
    mail.Display()

# Read the SQL file
with open(main_path + 'sql_code\\sqlReplen_G35.sql', 'r') as sql:
    query = sql.read()
    sql.close()

# Connect to Rex and query the database
with pyodbc.connect(driver='{iSeries Access ODBC Driver}', system='rex-ndc.hvwan.net', uid=NDC_USER, pwd=NDC_PWD) as connx:
    df = pd.read_sql(query, connx)

# Close the connection to Rex
connx.close()

# Avoid Nones and extra spaces in the output
df = df.replace(np.nan, '0')
df['Supplier'] = df['Supplier'].str.strip()
df['Air'] = df['Air'].str.strip()
df['Type'] = df['Type'].str.strip()

# Select the right supplier, in this case 05 or ADC for 139A
# Select the right supplier, in this case 05 for DON1
if bot_control(image_8_se, 0.80) != None:
    var_supplier = '9577S'
    list_of_supppliers = ['9577S', '9588S']
elif bot_control(image_8_ch, 0.80) != None:
    var_supplier = '9588S'
    list_of_supppliers = ['9577S', '9588S']
elif bot_control(image_8_uk, 0.80) != None:
    var_supplier = '50888'
    list_of_supppliers = ['50888' , '50888PRO']
elif bot_control(image_8_uk_pro, 0.80) != None:
    var_supplier = '50888PRO'
    list_of_supppliers = ['50888' , '50888PRO']

# This list will be a future Excel file with all the decisions made
replenishment_list = []
count_parts = 0

# To organize G35
bot.press('enter'), time.sleep(1)

try: # In case of a stop, we still want the final report
    for x_supplier in range(df['Supplier'].nunique()):
        # Filter the data by supplier, and limit calculates how many loops will perform
        df_loop = df[df['Supplier'] == var_supplier]
        start = 0
        limit = len(df_loop)

        # End the loop, Mario is not in the right place
        while bot.locateCenterOnScreen(image_9, confidence=0.7) != None:
            # At this stage, we are sure that we are in the right screen of G35
            bot.click(350, 485, duration=1), time.sleep(1)
            bot.press('ctrl')
            keyboard.press('shift')
            bot.mouseDown(button='left')

            # The drag cannot be fast, due to Rex interface
            bot.dragTo(745, 475, duration=2, button='left')
            keyboard.release('shift')
            bot.hotkey('ctrl', 'c')


            # Check which part number we are making decisions on
            pn = pyperclip.paste()
            pn = pn[-12:].replace(' ', '').replace('-', '')
            pn = int(pn)

            # Given the case that one part does not exist in Rex NDC but it is in G35 SA
            try:

                # If a part is duplicated should be skipped from the analisis
                var_dup = df_loop[df_loop['Pn'].duplicated()]['Pn'].isin([pn]).sum()

                if var_dup > 0:
                    # Get the variables such as availability, stock on hand, weight per pc, etc.
                    var_air = df_loop[df_loop['Pn'] == pn]['Air'].iloc[0]
                    var_typ = df_loop[df_loop['Pn'] == pn]['Type'].iloc[0]
                    var_sug = df_loop[df_loop['Pn'] == pn]['Sug (Q)'].iloc[0]
                    var_dem = 0
                    var_bos = 0
                    var_avl = 0
                    var_wgt = 0
                    var_stk = 0
                    var_rep = 0
                    var_rpn = 0

                else:
                    # Get the variables such as availability, stock on hand, weight per pc, etc.
                    var_air = df_loop[df_loop['Pn'] == pn]['Air'].item()
                    var_typ = df_loop[df_loop['Pn'] == pn]['Type'].item()
                    var_sug = df_loop[df_loop['Pn'] == pn]['Sug (Q)'].item()
                    var_dem = df_loop[df_loop['Pn'] == pn]['Roll. Dem. Y (Q)'].item()
                    var_bos = df_loop[df_loop['Pn'] == pn]['Bos'].item()
                    var_avl = df_loop[df_loop['Pn'] == pn]['Avail CDC (Q)'].item(
                    ) + df_loop[df_loop['Pn'] == pn]['Rns (Q)'].item() + 0.01
                    var_wgt = df_loop[df_loop['Pn'] == pn]['Weight per pc'].item()
                    var_stk = df_loop[df_loop['Pn'] == pn]['Stk per pc'].item()
                    var_rep = df_loop[df_loop['Pn'] == pn]['Repl. Code'].item()
                    var_rpn = df_loop[df_loop['Pn'] == pn]['New Part'].item()

            except:
                # Get the variables such as availability, stock on hand, weight per pc, etc.
                var_air = df_loop[df_loop['Pn'] == pn]['Air'].item()
                var_typ = df_loop[df_loop['Pn'] == pn]['Type'].item()
                var_sug = df_loop[df_loop['Pn'] == pn]['Sug (Q)'].item()
                var_dem = 0
                var_bos = 0
                var_avl = -1000
                var_wgt = 0
                var_stk = 0
                var_rep = 0
                var_rpn = 0

            # Skip and report the line if there are issues with the data
            if var_avl == -1000:
                var_note = '0. No data from Rex NDC, still in G35.'
                replenishment_list.append(
                    [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, 0, 0, var_rep, var_rpn, var_note])
                next_part_number(key=2, wait=1)

            # Skip and report if the part is obsolete or replaced
            elif var_rep in ( '12', '11' , '22' , '21'):
                var_note = '1. Replaced or obsolete, erased from G35.'
                replenishment_list.append(
                    [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                next_part_number(key=4, wait=1)

            # Skip and report if the line is repeated in G35
            elif var_dup > 0:
                var_note = '3. Mario cannot handle duplicated values, still in G35.'
                replenishment_list.append(
                    [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                next_part_number(key=2, wait=1)

            # Skip and report if we dont have enough stock in the shipping warehouse
            elif (var_sug / var_avl) >= 0.151:
                var_note = '2. CDC has not enough stock, still in G35.'
                replenishment_list.append(
                    [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                next_part_number(key=2, wait=1)

            # Ship by air to South Africa, the line is below 5 kilos, and we have back orders there
            # If not ship by boat. This line of code excludes the UK shipments
            elif var_bos > 0 and (var_supplier in ['9577S' , '9588S']):

                # Check weight and air supplier
                if (var_wgt * var_sug) <= 5.0 and not(var_air in ['9577A', '9588A']):
                    var_note = '4. Add air supplier to fix the back orders, erased from G35.'
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    next_part_number(key=4, wait=1)

                # Check weight
                elif (var_wgt * var_sug) <= 5.0 and var_wgt > 0:
                    decision(key=7, wait=1)

                    # Choose the right air supplier according to which CDC
                    if var_supplier == '9577S' and bot_control(image_11_se, 0.8) != None:
                        var_note = '5. Shipped by air from 05 to fix the back orders.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        select_item(key=6, wait=1)

                    # Choose the right air supplier according to which CDC
                    elif var_supplier == '9588S' and bot_control(image_11_ch, 0.8) != None:
                        var_note = '5. Shipped by air from ADC to fix the back orders.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        select_item(key=6, wait=1)

                # The part is too heavy, ship by air
                elif (var_wgt * var_sug) > 5.0 or var_wgt == 0:
                    var_note = '6. Shipped by boat, we have back orders but the line is too heavy.'
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    decision(key=6, wait=1)

                else:
                    var_note = '10. No decision was made, still in G35, check the logic. '
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    next_part_number(key=2, wait=1)

            # Ship by air to South Africa, if the line is below 1.5 kilos.
            # If not ship by boat. This line of code excludes the UK shipments
            elif var_supplier in ['9577S' , '9588S']:

                # Check weight and air supplier
                if (var_wgt * var_sug) <= 1.500 and not(var_air in ['9577A', '9588A']):
                    var_note = '4. Add air supplier to replenish, erased from G35.'
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    next_part_number(key=4, wait=1)

                # Check weight
                elif (var_wgt * var_sug) <= 1.500 and var_wgt > 0:
                    decision(key=7, wait=1)

                    # Choose the right air supplier according to which CDC
                    if var_supplier == '9577S' and bot_control(image_11_se, 0.8) != None:
                        var_note = '7. Shipped by air from 05 to replenish.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        select_item(key=6, wait=1)

                    # Choose the right air supplier according to which CDC
                    elif var_supplier == '9588S' and bot_control(image_11_ch, 0.8) != None:
                        var_note = '7. Shipped by air from ADC to replenish.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        select_item(key=6, wait=1)

                # The part is too heavy, so ship by boat
                elif (var_wgt * var_sug) > 1.500 or var_wgt == 0:
                    var_note = '8. Shipped by boat to replenish.'
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    decision(key=6, wait=1)

                else:
                    var_note = '10. No decision was made, still in G35, check the logic. '
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    next_part_number(key=2, wait=1)

            else:

                # Ship by air to the UK, the line is below 10 kilos, and we have back orders there
                # If not ship by truck. This line of code excludes South Africa shipments
                if var_bos > 0 and (var_supplier in ['50888' , '50888PRO']): 

                    # Check weight and air supplier
                    if (var_wgt * var_sug) <= 10.0 and not(var_air in ['50888DS']):
                        var_note = '4. Add air supplier to fix the back orders, erased from G35.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        next_part_number(key=4, wait=1)

                    # Check weight
                    elif (var_wgt * var_sug) <= 10.0 and var_wgt > 0:
                        decision(key=7, wait=1)

                        # Choose the right air supplier according to which CDC
                        if bot_control(image_11_uk, 0.8) != None:
                            var_note = '5. Shipped by air from 05 to fix the back orders.'
                            replenishment_list.append(
                                [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                            select_item(key=6, wait=1)

                    # The part is too heavy, ship by truck
                    elif (var_wgt * var_sug) > 10.00 or var_wgt == 0:
                        var_note = '6. Shipped by truck, we have back orders but the line is too heavy.'
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        decision(key=6, wait=1)

                    else:
                        var_note = '8. No decision was made, still in G35, check the logic. '
                        replenishment_list.append(
                            [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                        next_part_number(key=2, wait=1)


                # Ship by truck to UK
                elif var_supplier in ['50888' , '50888PRO']:

                    var_note = '7. Shipped by truck to replenish.'
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    decision(key=6, wait=1)

                else:
                    var_note = '8. No decision was made, still in G35, check the logic. '
                    replenishment_list.append(
                        [var_supplier, var_typ, pn, var_sug, var_dem, var_bos, (var_avl - 0.01), (var_wgt * var_sug), var_rep, var_rpn, var_note])
                    next_part_number(key=2, wait=1)

            # Follow up advance of the script
            start += 1
            print(start, limit, pn, var_note)

            # Finish the loop
            if start == limit:
                list_of_supppliers.remove(var_supplier)
                bot.press('F3'), time.sleep(1)

                count_parts += start

                # If we are shipping to SA, select the next supplier (9577S or 9577A)
                # If we are shipping to UK, select the next supplier (50888 or 50888PRO)
                if bot_control(image_8_se, 0.80) != None and '9577S' in list_of_supppliers:
                    var_supplier = '9577S'
                    select_item(key=1, wait=1)
                    break
                elif bot_control(image_8_ch, 0.80) != None and '9588S' in list_of_supppliers:
                    var_supplier = '9588S'
                    select_item(key=1, wait=1)
                    break
                elif bot_control(image_8_uk, 0.80) != None and '50888' in list_of_supppliers:
                    var_supplier = '50888'
                    select_item(key=1, wait=1)
                    break
                elif bot_control(image_8_uk_pro, 0.80) != None and '50888PRO' in list_of_supppliers:
                    var_supplier = '50888PRO'
                    select_item(key=1, wait=1)
                    break
                else:
                    print('Mario finishes the interaction with Rex, and now prepares a report.')
                    break

except: # In case of a stop, we still want the final report
    count_parts += start
    pass

# Create a data frame to store the result of Mario's analysis
df_result = pd.DataFrame(replenishment_list, columns=[
                         'Supplier', 'Type', 'Pn', 'Sug (Q)', 'Roll. Dem. Y (Q)', 'BOs', 'Avail CDC (Q)', 'Weight per line', 'Repl. Code', 'Repl. Pn', 'Action'])

name_excel = main_path + 'output_files\\' + \
    str(datetime.now().strftime(r'%d-%m-%y %H-%M')) + ' - G35_result.xlsx'

# Export dataset to XLSX
with pd.ExcelWriter(name_excel) as writer:
    df_result.to_excel(writer, sheet_name='G35_result', index=False)

    # Auto-adjust columns' width
    for column in df_result:
        column_width = max(df_result[column].astype(str).map(len).max(), len(column))
        col_idx = df_result.columns.get_loc(column)
        writer.sheets["G35_result"].set_column(col_idx, col_idx, column_width)

# Measure the time that Mario needs
end_time = time.time()
duration = end_time - start_time

# Create an e-mail with the information about Mario's anaysis
mail_subject = f'In Rex G35, Mario has processed {count_parts} part numbers in {str(duration/60)[:4]} minutes'
mail_body = f'''
<p style=color:rgb(47,84,150);>Dear Planners,
<br></br><br></br>
Mario has taken care of {count_parts} part numbers in Rex G35; please check the attachment,
<br></br><br></br>
There is still job to do: (1) several parts are still in G35, (2) several parts need to be replaced (attached), among others.
<br></br><br></br>
Please take a look, and feedback is welcome.
</p>
'''
# Send an e-mail and PDF with the right receiver
if var_supplier in ['50888' , '50888PRO']:
    mail_receiver = ""
    mail_attachment_2 = main_path + 'documentation\\Mariobot_UK.pdf'
elif var_supplier in ['9577S' , '9588S']:
    mail_receiver = ""
    mail_attachment_2 = main_path + 'documentation\\Mariobot_SA.pdf'


mail_attachment_1 = name_excel
emailer(mail_subject, mail_body, mail_receiver, mail_attachment_1, mail_attachment_2)
