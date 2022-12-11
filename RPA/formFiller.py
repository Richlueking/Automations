#! python3
# formFiller.py - Automatically fills in the form.  https://autbor.com/form
import os
import openpyxl as xl
import pyautogui as auto
import time
import sys

wb = xl.load_workbook('formfillerData.xlsx')
print(os.getcwd())
print(type(wb))

sheet = wb['Sheet1']  # Get a sheet from the workbook.
print(sheet)

anotherSheet = wb.active  # Get the active sheet.
print(anotherSheet)
print('\n')

typeSpeed_fast = 0.10
typeSpeed_Medium = 0.20
typeSpeed_slow = 0.50

auto.hotkey('win', 'add')
time.sleep(3)
auto.hotkey('win', 'm')


def incMag():
    auto.hotkey('win', 'add')


def decMag():
    auto.hotkey('win', '-')


def window():
    auto.click(x=7, y=767)
    auto.click(x=7, y=767)


# opening the form
form = 'https://docs.google.com/forms/d/e/1FAIpQLSdBdtwwQsumIn3cVyH5WKReTIpwNOWDIjxZiz1WFz9USti5Mw/viewform'

window()
incMag()
time.sleep(3)
auto.write('chrome', typeSpeed_fast)
decMag()
time.sleep(3)
auto.keyDown('enter')
auto.keyUp('enter')
time.sleep(3)

auto.click(x=151, y=63)
incMag()
auto.write(form, typeSpeed_fast)
decMag()
auto.keyDown('enter')
auto.keyUp('enter')

# first converted the sheets data into a list of dictionary
rows = sheet.rows
headers = [cell.value for cell in next(rows)]  # get the value from each row
# print(headers)

all_rows = []
for row in rows:
    data = {}
    for title, cell in zip(headers, row):
        data[title] = cell.value
    all_rows.append(data)

# Fill out the Name field.
print('>>> 2-SECOND PAUSE TO LET USER PRESS CTRL-C <<<')
for info in all_rows:
    time.sleep(2)
    auto.write('\t')
    auto.write('\t')
    auto.write('\t')
    auto.write('\t')
    auto.write('\t')
    time.sleep(2)
    auto.press('delete')

    auto.click(x=65, y=481, duration=0.90)
    incMag()
    auto.write(info['Name-of-Account-Holder'] + '\t', typeSpeed_fast)
    decMag()

    auto.press('delete')
    auto.click(x=67, y=632, duration=0.90)
    incMag()
    auto.write(info['Branch-Sort-Code'] + '\t', typeSpeed_fast)
    decMag()
    auto.click(x=628, y=161, duration=0.90)
    auto.scroll(-50)  # scroll down 10 "clicks"
    time.sleep(2)

    auto.press('delete')
    time.sleep(1)
    auto.click(x=67, y=173, duration=0.90)
    incMag()
    auto.write(info['postal-address'] + '\t', typeSpeed_fast)
    decMag()
    time.sleep(1)

    auto.click(x=66, y=346, duration=0.90)
    auto.press('delete')
    time.sleep(1)
    incMag()
    auto.write(info['Service-User-Number'] + '\t', typeSpeed_fast)
    decMag()
    time.sleep(1)

    auto.click(x=68, y=495, duration=0.90)
    auto.press('delete')
    time.sleep(1)
    incMag()
    auto.write(info['Reference'] + '\t', typeSpeed_fast)
    decMag()

    # Click Submit.
    def enterKey():
        auto.keyDown('enter')
        auto.keyUp('enter')
    enterKey()
    time.sleep(3)

    # Click the Submit another response link.
    auto.write('\t')
    auto.write('\t')
    enterKey()
else:
    auto.confirm(title='Confirm', text='Task Completed')
    time.sleep(3)
    enterKey()
