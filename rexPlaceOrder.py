"""
This script uses PyAutoGui to control the mouse and keyboard and place manual orders in the ERP system

Author: Mauricio Montilla
"""

from datetime import datetime
import time

try:
    import pandas as pd
    import numpy as np
    import pyautogui as bot
    import keyboard

except:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'pandas'])
    pip(['install', 'numpy'])
    pip(['install', 'pyautogui'])
    pip(['install', 'keyboard'])

    import pandas as pd
    import numpy as np
    import pyautogui as bot
    import keyboard

# Remove warnings
pd.options.mode.chained_assignment = None
time.sleep(1)

# Locates the center of the image that it gets
def bot_control(image, conf, grayscale=True):
    cords = None
    counter = 0

    while cords == None:
        cords = bot.locateCenterOnScreen(
            image, grayscale=grayscale, confidence=conf)

        # Wait five seconds, if we done find the image break
        counter += 1
        if counter == 5: break

    if cords != None:
        bot.moveTo(cords, duration=2), time.sleep(1)
    else:
        print("Mario cannot find " + image)

    return cords

# Updates the promised weeks accordingly
def create_order_line(pn, quantity, week, day, wait):
    bot.press("ctrl"), time.sleep(wait)
    bot.typewrite(pn[-18:]), time.sleep(wait)
    bot.press("tab"), time.sleep(wait)
    bot.typewrite(quantity[-8:]), time.sleep(wait)
    bot.press("tab"), time.sleep(wait)
    bot.typewrite(week[-4:]), time.sleep(wait)
    bot.press("tab"), time.sleep(wait)
    bot.typewrite(day[-1:]), time.sleep(wait)
    bot.press("enter", presses=2), time.sleep(wait)
    bot.press("ctrl"), time.sleep(wait)

# Files locations
main_path = '\\\\SW03000.se.hvwan.net\\Dept\\P\\PSA\\Inventory Management\\Python\\'
image_1 = main_path + "images\\G01_SE\\8_inside_G01.png"
image_2 = main_path + "images\\G01_SE\\9_inside_A01.png"

# Current week to filter the dataframes
now = datetime.now()
current_week = int(str(now.year)[-2:] + ('0' + str(now.isocalendar()[1]))[-2:])

# Get data coming from the VBA array
df = pd.read_csv(main_path + "csvFile.csv", encoding="ISO-8859-1")
df = df.replace(np.nan, 0)

# Make sure that G01 is open as expected
if bot_control(image_1, 0.7) != None:
    bot.click(), time.sleep(1)
    bot.press("enter"), time.sleep(1)
    bot.press("ctrl"), time.sleep(1)

# if it is A01
else:
    print("Please open Rex in the G01 screen before executing the script.")
    print("Press 'End' and try again. This script ends here.")

    # Allow the use to read the problem
    while keyboard.is_pressed('end') != True:
        time.sleep(2)
    exit()

for line in range(len(df['Pn'])):

    # Prepare the pn
    pn = int(df['Pn'].iloc[line])
    # Prepare the quantity
    quantity = int(df['Qty.'].iloc[line])

    # Prepare the week requested, the week requested needs to be a valid one
    week_req = int(df['Week Req.'].iloc[line])

    if (current_week + 1) >= week_req:
        week_req = 0
        df['Day Req.'].iloc[line] = 0

    # Prepare the day requested, the day requested needs to be a valid one
    day_req = int(df['Day Req.'].iloc[line])

    if day_req > 5:
        day_req = 5

    # Fit the variables to Rex size
    pn = '000000000' + str(pn)
    quantity = '00000000' + str(quantity)
    week_req = '000000000' + str(week_req)
    day_req = '000000000' + str(day_req)

    # Include only lines which are vlid
    if int(pn) > 0 and int(quantity) > 0:
        create_order_line(pn, quantity, week_req, day_req, 1)

    # End the loop, if Bowser is not in the right place
    counter = 0

    while bot.locateCenterOnScreen(image_1, confidence=0.7) == None:
        # Wait 5 seconds, if we done find the image break
        counter += 1
        if counter ==5: exit()

        time.sleep(1)

    # Follow up advance of the script
    print(line, len(df['Pn']), int(pn), int(quantity))

# Finish the loop
bot.press("tab", presses=6), time.sleep(1)
bot.typewrite("1")
