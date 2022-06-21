"""
This script 
"""

import time
import re
import warnings

# Avoid warnings
warnings.simplefilter("ignore")

try:
    import pyperclip
    import keyboard

except ImportError:
    # In case that an external package is missing, the script will install them
    from pip._internal import main as pip
    pip(['install', 'pyperclip'])
    pip(['install', 'keyboard'])

    import pyperclip
    import keyboard

# Create the desired pattern
partNumRegex1 = re.compile(r'\d{3} \d{2} \d{2}-\d{2}')
partNumRegex2 = re.compile(r'\d{9}')
partNumRegex3 = re.compile(r'\d{10}')

# Loop as long as the '+' is not pressed, or the application is closed
while keyboard.is_pressed('-') != True:

    try:
        pn = 'Nothing found'

        # Search for the pattern in the clipboard data
        mo1 = partNumRegex1.search(pyperclip.paste())
        mo2 = partNumRegex2.search(pyperclip.paste())
        mo3 = partNumRegex3.search(pyperclip.paste())

        # if the pattern is found, copy it into the clipboard
        if mo3 != None:
            po = str(mo3.group())
            pyperclip.copy(po)

            print(po)

        elif mo2 != None:
            pn = str(mo2.group()) + '     '
            pyperclip.copy(pn)

            print(pn)

        elif mo1 != None:
            pn = str(mo1.group()).replace(' ', '').replace('-', '') + '     '
            pyperclip.copy(pn)

            print(pn)

    except:
        pass

    # Wait one second between iteration
    time.sleep(1)
