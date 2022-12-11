import openpyxl as xl
import pyautogui as auto
import time
import sys

time.sleep(2)

email = 'richlueking302@gmail.com'
subject = 'Listing of all employee'
fileName = 'employees report'

typeSpeed_fast = 0.10
typeSpeed_Medium = 0.20
typeSpeed_slow = 0.50

auto.hotkey('win', 'add')
time.sleep(3)
auto.hotkey('win', 'm')

time.sleep(3)


def incMag():
    auto.hotkey('win', 'add')


def decMag():
    auto.hotkey('win', '-')


def window():
    auto.click(x=7, y=767)
    auto.click(x=7, y=767)


def enterKey():
    auto.keyDown('enter')
    auto.keyUp('enter')


# opening sql management  studio
window()
incMag()
time.sleep(2)
auto.write('ssms', typeSpeed_Medium)
decMag()
enterKey()

# opening database in sql studio
time.sleep(40)
auto.click(x=611, y=418, duration=0.90)
window()
incMag()
time.sleep(3)
auto.write('AutoQuerying', typeSpeed_Medium)
decMag()
enterKey()

# Activating the database for querying
time.sleep(5)
auto.click(x=97, y=90, duration=0.90)
incMag()
time.sleep(5)
auto.write('AdventureWorks', typeSpeed_Medium)
decMag()
enterKey()

# querying the data
auto.click(x=264, y=140, duration=0.90)
incMag()
time.sleep(3)
auto.write('select * from HumanResources.Employee', typeSpeed_Medium)
decMag()
auto.click(x=229, y=138, duration=0.90)
auto.click(x=283, y=88, duration=0.90)

# sending to excel
auto.click(x=191, y=471, duration=0.99)
time.sleep(3)
incMag()
auto.hotkey('ctrl', 'shiftleft', 'c')
decMag()

# opening and sending to excel
window()
incMag()
auto.write('excel', typeSpeed_Medium)
decMag()
time.sleep(2)
enterKey()
time.sleep(6)
enterKey()
time.sleep(3)
enterKey()

time.sleep(5)
enterKey()
time.sleep(10)
enterKey()
time.sleep(4)
auto.hotkey('ctrl', 'v')
time.sleep(2)
enterKey()

# saving spreadsheet
auto.hotkey('ctrl', 's')
time.sleep(3)
auto.click(x=519, y=487, duration=0.99)
time.sleep(3)
auto.click(x=63, y=160, duration=0.99)

time.sleep(3)
auto.click(x=199, y=342, duration=0.99)
incMag()
auto.press('backspace')
auto.press('backspace')
auto.press('backspace')
auto.press('backspace')
auto.press('backspace')

time.sleep(1)
auto.write('employees report', typeSpeed_Medium)
decMag()
time.sleep(2)
enterKey()

# emailing spreadsheet
window()
incMag()
time.sleep(2)
auto.write('mail', typeSpeed_Medium)
decMag()
enterKey()
time.sleep(4)

auto.click(x=25, y=101, duration=0.90)
time.sleep(3)


auto.click(x=983, y=193, duration=0.90)
incMag()
auto.write(email, typeSpeed_Medium)
decMag()
time.sleep(2)


auto.write('\t')
auto.write('\t')
auto.write('\t')
auto.write('\t')
auto.click(x=959, y=232, duration=0.90)
incMag()
auto.write(subject, typeSpeed_Medium)
decMag()
enterKey()

auto.click(x=1078, y=56, duration=0.90)
auto.click(x=980, y=94, duration=0.90)
time.sleep(3)
auto.click(x=209, y=417, duration=0.90)
incMag()
auto.write(fileName, typeSpeed_Medium)
decMag()
enterKey()
time.sleep(3)
auto.click(x=1268, y=49, duration=0.90)
time.sleep(2)
auto.confirm(title='Confirm', text='Task Completed')
time.sleep(3)
enterKey()

# next: fix excel delay, add alert and assign texts to variable
