import pyautogui
import pyperclip
import time
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from pathlib import Path


# ---------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# Folder where the Excel file and screenshot will be saved
output_folder = Path.cwd()

# Current date and time
now = datetime.now()

current_date = now.strftime("%Y-%m-%d")
current_datetime = now.strftime("%Y-%m-%d %H:%M:%S")

excel_filename = f"daily_report_{current_date}.xlsx"
screenshot_filename = f"daily_report_{current_date}.png"

excel_path = output_folder / excel_filename
screenshot_path = output_folder / screenshot_filename


# ---------------------------------------------------------
# STEP 1 - OPEN CHROME
# ---------------------------------------------------------

print("Step 1: Opening Chrome...")

pyautogui.hotkey("win", "r")
time.sleep(1)

pyautogui.write("chrome", interval=0.1)
pyautogui.press("enter")

# Give Chrome enough time to start
time.sleep(4)


# ---------------------------------------------------------
# STEP 2 - OPEN WEATHER WEBSITE
# ---------------------------------------------------------

print("Step 2: Opening AccuWeather...")

pyautogui.hotkey("ctrl", "l")

pyautogui.write(
    "https://www.accuweather.com/",
    interval=0.03
)

pyautogui.press("enter")

# Wait for website to load
time.sleep(7)


# ---------------------------------------------------------
# STEP 3 - COPY PAGE CONTENT
# ---------------------------------------------------------

print("Step 3: Copying webpage information...")

# Make sure Chrome has focus
pyautogui.click(800, 400)

time.sleep(1)

# Select all page content
pyautogui.hotkey("ctrl", "a")

time.sleep(1)

# Copy
pyautogui.hotkey("ctrl", "c")

time.sleep(2)


# ---------------------------------------------------------
# STEP 4 - GET COPIED DATA
# ---------------------------------------------------------

print("Step 4: Reading copied information...")

page_text = pyperclip.paste()

print("\nCopied webpage text:")
print("--------------------------------")
print(page_text[:1000])
print("--------------------------------")


# ---------------------------------------------------------
# STEP 5 - EXTRACT TEMPERATURE
# ---------------------------------------------------------

print("Step 5: Searching for temperature...")

temperature = None

# Look for values such as:
# 32°
# 32°F
# 32 C
# 32°C

temperature_pattern = r"\b(-?\d{1,3})\s*°?\s*(F|C|°F|°C)?\b"

matches = re.findall(
    temperature_pattern,
    page_text,
    re.IGNORECASE
)

if matches:
    # Take the first reasonable temperature value
    for value, unit in matches:

        number = int(value)

        if -50 <= number <= 60:

            if unit:
                temperature = f"{number}°{unit.replace('°', '').upper()}"
            else:
                temperature = f"{number}°"

            break


# If temperature was not detected
if temperature is None:
    temperature = "Temperature not detected"


print("Fetched data:", temperature)


# ---------------------------------------------------------
# STEP 6 - CREATE COMMENT
# ---------------------------------------------------------

if "not detected" in temperature.lower():

    comment = "Weather information could not be detected automatically."

else:

    # Convert temperature to Celsius if possible
    try:

        number = int(re.search(r"-?\d+", temperature).group())

        if "F" in temperature.upper():

            celsius = (number - 32) * 5 / 9

        else:

            celsius = number

        if 20 <= celsius <= 30:

            comment = "Good for outdoor activities."

        elif celsius > 30:

            comment = "Quite warm; stay hydrated."

        else:

            comment = "Cool weather; suitable for outdoor activities."

    except:

        comment = "Weather information retrieved successfully."


print("Comment:", comment)


# ---------------------------------------------------------
# STEP 7 - CREATE EXCEL FILE
# ---------------------------------------------------------

print("Step 7: Creating Excel report...")

workbook = Workbook()

worksheet = workbook.active
worksheet.title = "Daily Report"


# Headers
worksheet["A1"] = "Date & Time"
worksheet["B1"] = "Fetched Data"
worksheet["C1"] = "Comment"


# Header formatting
for cell in worksheet[1]:

    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")


# Data row
worksheet["A2"] = current_datetime
worksheet["B2"] = temperature
worksheet["C2"] = comment


# Column widths
worksheet.column_dimensions["A"].width = 22
worksheet.column_dimensions["B"].width = 20
worksheet.column_dimensions["C"].width = 55


# Save Excel file
workbook.save(excel_path)

print(f"Excel file saved: {excel_path}")


# ---------------------------------------------------------
# STEP 8 - OPEN EXCEL
# ---------------------------------------------------------

print("Step 8: Opening Excel...")

pyautogui.hotkey("win", "r")

time.sleep(1)

pyautogui.write(
    str(excel_path),
    interval=0.02
)

pyautogui.press("enter")

# Wait for Excel to open
time.sleep(6)


# ---------------------------------------------------------
# STEP 9 - TAKE SCREENSHOT
# ---------------------------------------------------------

print("Step 9: Taking screenshot...")

# Give Excel time to display the workbook
time.sleep(2)

# Make sure Excel is visible
pyautogui.hotkey("alt", "tab")

time.sleep(1)

# Take screenshot of the entire screen
screenshot = pyautogui.screenshot()

screenshot.save(screenshot_path)

print(f"Screenshot saved: {screenshot_path}")


# ---------------------------------------------------------
# COMPLETE
# ---------------------------------------------------------

print("\n========================================")
print("       DAILY REPORT COMPLETED")
print("========================================")

print(f"Excel file : {excel_path}")
print(f"Screenshot : {screenshot_path}")
print(f"Date/Time  : {current_datetime}")
print(f"Data       : {temperature}")
print(f"Comment    : {comment}")

print("========================================")