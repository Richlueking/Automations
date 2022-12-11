import webbrowser
import bs4
import requests
import pyautogui as auto
import time
import sys

typeSpeed_fast = 0.30
typeSpeed_Medium = 0.60
typeSpeed_slow = 0.90

auto.hotkey('win', 'add')
time.sleep(3)
auto.hotkey('win', 'm')

writeTo = 'word'


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


window()
incMag()
auto.write(writeTo, typeSpeed_Medium)
decMag()
time.sleep(2)
enterKey()
time.sleep(6)
enterKey()
time.sleep(3)
enterKey()

url1 = 'https://knoema.com/data/liberia+consumption+gasoline'

url2 = 'https://knoema.com/data/nigeria+consumption''+gasoline+million-metric-tons-of-oil-equivalent'
url3 = 'https://knoema.com/data/ghana+consumption+gasoline+million-metric-tons-of-oil-equivalent'

auto.click(x=34, y=221, duration=0.90)
incMag()

auto.write('World - Consumption - Gasoline')
enterKey()
enterKey()

res = requests.get(url1)
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, features="html.parser")
header = soup.select('h1')
t = header[0].getText()
tInfo = soup.select('.text')
info = tInfo[0].getText()
fullInfo = t, info
for text in fullInfo:
    auto.write(text)

res = requests.get(url2)
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, features="html.parser")
header = soup.select('h1')
t = header[0].getText()
tInfo = soup.select('.text')
info = tInfo[0].getText()
fullInfo = t, info
for text in fullInfo:
    auto.write(text)

res = requests.get(url3)
res.raise_for_status()
soup = bs4.BeautifulSoup(res.text, features="html.parser")
header = soup.select('h1')
t = header[0].getText()
tInfo = soup.select('.text')
info = tInfo[0].getText()
fullInfo = t, info
for text in fullInfo:
    auto.write(text)
    decMag()


'''styling the  report'''

# styling the heading
auto.click(x=17, y=268, duration=0.90)
incMag()
tm = 0
while tm < 30:
    tm += 1
    auto.hotkey('shift', 'right')
    continue
decMag()

time.sleep(3)
auto.hotkey('ctrl', 'shift', 'p')
time.sleep(3)
auto.write('30', typeSpeed_fast)
enterKey()

time.sleep(3)
auto.hotkey('ctrl', 'I')
time.sleep(2)
auto.hotkey('ctrl', 'U')
time.sleep(3)


# styling liberia
auto.click(x=341, y=386, duration=0.99)
enterKey()
auto.keyDown('up')
auto.keyUp('up')

incMag()
tm = 0
while tm < 42:
    tm += 1
    auto.hotkey('shift', 'right')
    continue
decMag()

time.sleep(3)
auto.hotkey('ctrl', 'shift', 'p')
time.sleep(3)
auto.write('20', typeSpeed_fast)
enterKey()
time.sleep(3)
auto.hotkey('ctrl', 'I')
time.sleep(2)
auto.hotkey('ctrl', 'B')


# styling Nigeria
auto.click(x=341, y=484, duration=0.99)
enterKey()
auto.keyDown('up')
auto.keyUp('up')

incMag()
tm = 0
while tm < 40:
    tm += 1
    auto.hotkey('shift', 'right')
    continue
decMag()

time.sleep(3)
auto.hotkey('ctrl', 'shift', 'p')
time.sleep(3)
auto.write('20', typeSpeed_fast)
enterKey()
time.sleep(3)
auto.hotkey('ctrl', 'I')
time.sleep(2)
auto.hotkey('ctrl', 'B')


# styling Ghana
auto.click(x=340, y=584, duration=0.99)
enterKey()
auto.keyDown('up')
auto.keyUp('up')

incMag()
tm = 0
while tm < 40:
    tm += 1
    auto.hotkey('shift', 'right')
    continue
decMag()

time.sleep(3)
auto.hotkey('ctrl', 'shift', 'p')
time.sleep(3)
auto.write('20', typeSpeed_fast)
enterKey()
time.sleep(3)
auto.hotkey('ctrl', 'I')
time.sleep(2)
auto.hotkey('ctrl', 'B')

auto.click(x=615, y=16, duration=0.99)



