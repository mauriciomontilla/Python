"""
This script uses PyAutoGui to control the mouse and keyboard and updates the promise weeks in the ERP system

Author: Mauricio Montilla
"""

from datetime import datetime
import time
import warnings

# Avoid warnings
warnings.simplefilter("ignore")

try:
    import pandas as pd
    import numpy as np
    import pyautogui as bot
    import keyboard
    from PIL import Image
    import cv2

except:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'pandas'])
    pip(['install', 'numpy'])
    pip(['install', 'pyautogui'])
    pip(['install', 'keyboard'])
    pip(['install', 'Pillow'])
    pip(['install', 'opencv-python'])

    import pandas as pd
    import numpy as np
    import pyautogui as bot
    import keyboard

# Remove warnings
pd.options.mode.chained_assignment = None
time.sleep(1)

# Locate the center of the image that it gets
def bot_control(image, conf, grayscale=True):
    cords = None
    counter = 0

    while cords == None:
        cords = bot.locateCenterOnScreen(
            image, grayscale=grayscale, confidence=conf)
        counter += 1
        if counter == 5: break

    if cords != None:
        bot.moveTo(cords, duration=2), time.sleep(1)
    else:
        print("Mario cannot find " + image)

    return cords

# Update the promised weeks accordingly
def update_line(line, new_week, new_day, information, wait):
    bot.hotkey("shift", "tab"), time.sleep((wait/2))
    bot.press("4"), time.sleep((wait/2))
    bot.hotkey("shift", "tab"), time.sleep(wait)
    bot.typewrite(line[:5])
    bot.press("enter"), time.sleep(wait)
    bot.press("tab", presses=13, interval=0.2), time.sleep(wait)
    bot.typewrite(str(new_week)), time.sleep(wait)
    bot.press("tab"), time.sleep((wait/2))
    bot.typewrite(str(new_day)), time.sleep((wait/2))
    bot.press("tab", presses=4, interval=0.2), time.sleep(wait)
    bot.typewrite(str(information))
    bot.press("enter", presses=2, interval=1), time.sleep(wait)

# Files locations
main_path = '\\\\SW03000.se.hvwan.net\\Dept\\P\\PSA\\Inventory Management\\Python\\'
image_1 = main_path + "images\\G01_SE\\7_routine.png"
image_2 = main_path + "images\\G01_SE\\8_inside_G01.png"

# Current week to filter the dataframes
now = datetime.now()
current_week = int(str(now.year)[-2:] + ('0' + str(now.isocalendar()[1]))[-2:])
current_day = datetime.today().strftime(r'%Y%m%d')

# Get data coming from the VBA array and clean it
df = pd.read_csv(main_path + "csvFile.csv", encoding = "ISO-8859-1")
df.replace(np.nan, 0, inplace=True)
df.drop_duplicates(keep=False, inplace=True)

# These lines will be used in the loop
df_counted = df[['PO nr.', 'PO line', 'New Prom. Week', 'New Prom. Day', 'Comment']][df["New Prom. Week"].astype(int) >= current_week].reset_index(drop=True)

# Make sure that G01 is open as expected
if bot_control(image_1, 0.7) != None:
    bot.click(), time.sleep(1)
    bot.press("enter"), time.sleep(1)
    bot.press("ctrl"), time.sleep(1)
else:
    print("Please open Rex in the G01 screen before executing the script.")
    print("Press 'End' and try again. This script ends here.")

    # Allow the use to read the problem
    while keyboard.is_pressed('end') != True:
        time.sleep(2)
    exit()

# Avoid mistakes in the first loop
previous_rex_po = 0

# Variables to count which lines are updated
count_lines = 0
total_lines = len(df_counted)

# If one line fails, this variable is set to False
change_screen = True

for po_number in df_counted['PO nr.'].items():
    # Extract the data to be updated per order lines
    rex_po = int(po_number[1])
    rex_line = str(int(df_counted['PO line'][po_number[0]])) + '   '
    new_prom_week = df_counted['New Prom. Week'][po_number[0]]
    new_prom_day = df_counted['New Prom. Day'][po_number[0]]

    # Skip the lines which are empty
    if rex_po == 0 or rex_line == '0   ' or new_prom_week == 0 or new_prom_day == 0:
        break

    if df_counted['Comment'][po_number[0]] != 0:
        information = df_counted['Comment'][po_number[0]]
    else:
        information = 'Bowser updated.'

    # Extra information in the line should be erased
    information = information + (' ' * 66)
    information = information[:57] + '-' + current_day

    # The day requested needs to be a valid one
    if new_prom_day > 5:
        new_prom_day = 5

    if int(current_week) < int(new_prom_week):

        # If this is the first PO line, we need to enter G01
        if count_lines == 0:
            bot.typewrite(str(rex_po)), time.sleep(1)
            bot.press("enter", presses=3, interval=1), time.sleep(1)

            # If the PO is not open, skip the update
            if bot.locateCenterOnScreen(image_2, confidence=0.7) != None:
                update_line(rex_line, new_prom_week, new_prom_day, information, wait=1)
                count_lines += 1
            else:
                bot.press("ctrl"), time.sleep(1)
                bot.hotkey("shift", "tab"), time.sleep(1)

        # If previous PO is the same as the new one, then continue in the same PO
        elif rex_po == previous_rex_po:

            # If the PO is not open, skip the update
            if bot.locateCenterOnScreen(image_2, confidence=0.7) != None:
                update_line(rex_line, new_prom_week, new_prom_day, information, wait=1)
                change_screen = True
                count_lines += 1

        # If previous PO is dfferent to the new one, then exit that PO
        elif rex_po != previous_rex_po:

            if change_screen:
                bot.press("F3"), time.sleep(2)
                bot.press("enter"), time.sleep(1)
                bot.press("ctrl"), time.sleep(1)

            # End the loop, if Bowser is not in the right place
            counter = 0
            while bot.locateCenterOnScreen(image_1, confidence=0.7) == None:
                # Wait 2 loops, if we done find the image break
                counter += 1
                if counter == 2: exit()
                time.sleep(1)

            bot.typewrite(str(rex_po)), time.sleep(1)
            bot.press("enter", presses=3, interval=1), time.sleep(1)

            # If the PO is not open, skip the update
            if bot.locateCenterOnScreen(image_2, confidence=0.7) != None:
                update_line(rex_line, new_prom_week, new_prom_day, information, wait=1)
                change_screen = True
                count_lines += 1
            else:
                bot.press("enter"), time.sleep(1)
                bot.press("ctrl"), time.sleep(1)
                change_screen = False
                count_lines += 1

    previous_rex_po = rex_po

    # Follow up advance of the script
    print(count_lines, total_lines, rex_po, rex_line)

# Finish the loop
bot.press("F3")
