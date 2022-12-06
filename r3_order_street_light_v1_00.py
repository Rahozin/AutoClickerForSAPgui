"""Making an order"""

from time import sleep
from datetime import datetime, timedelta
import sys
import pyautogui
import keyboard
import pyperclip
import arrow
import win32api
import os


def allert():
    """Allert to check customer`s card"""
    message = 'Ти перевірив картку дебітора?'
    missing = pyautogui.confirm(
        text=message, title='Увага!', buttons=['Так', 'Ні'])
    if missing == 'Ні':
        sys.exit()


def check_layout():
    """Exit if layout isn't ENG"""
    layout = win32api.GetKeyboardLayout()
    if layout != 67699721:
        MESSAGE3 = 'Зміни розкладку клавіатури на "ENG"\n і перезапусти програму.'
        missing3 = pyautogui.alert(
            text=MESSAGE3, title='Неправильна розкладка клавіатури!', button='Ok')
        if missing3 == 'Ok':
            sys.exit()


def quick_click(clicks, pic_name):
    """Clicking exactly when pic is appear + interapt"""
    if keyboard.is_pressed('ctrl') is False:
        time_start = datetime.now()
        time_finish = time_start + timedelta(seconds=10)
        go_click = True
        while go_click:
            target = pyautogui.locateOnScreen(
                PICS_PATH + pic_name, confidence=0.8)
            if target:
                if clicks == 1:
                    pyautogui.click(target)
                    print('click "' + pic_name + '" *1')
                if clicks == 2:
                    pyautogui.doubleClick(target)
                    print('click "' + pic_name + '" *2')
                break
            if time_start >= time_finish:
                message = 'Не знаходжу кнопку: ' + pic_name + \
                    '.\n' + 'Можна зробити крок вручну і продовжити'
                missing = pyautogui.confirm(
                    text=message, title='Щось пішло не так!', buttons=['Продовжити', 'Зупинити'])
                if missing == 'Зупинити':
                    sys.exit()
                else:
                    go_click = False
            time_start = datetime.now()
            sleep(0.2)
    else:
        pyautogui.alert('Програма була зупинена!')
        sys.exit()


def click(pic_name):
    """1 click using quick_click function"""
    clicks = 1
    quick_click(clicks, pic_name)


def click2(pic_name):
    """2 clicks using quick_click function"""
    clicks = 2
    quick_click(clicks, pic_name)


def write(string):
    """Simplyfying common functions"""
    keyboard.write(string)
    print('written "' + string + '"')


def key(string):
    """Simplyfying common functions"""
    if keyboard.is_pressed('ctrl') is False:
        pyautogui.press(string)
        print('key "' + string + '"')
    else:
        pyautogui.alert('Програма була зупинена!')
        sys.exit()


def hotkey(string1, string2):
    """Simplyfying common functions + interapt"""
    if keyboard.is_pressed('ctrl') is False:
        pyautogui.hotkey(string1, string2)
        print('hotkey "' + string1 + " + " + string2 + '"')
    else:
        pyautogui.alert('Програма була зупинена!')
        sys.exit()


def press_and_release(string):
    """Simplyfying common functions + interapt"""
    if keyboard.is_pressed('ctrl') is False:
        keyboard.press_and_release(string)
        print('hotkey "' + string + '"')
    else:
        pyautogui.alert('Програма була зупинена!')
        sys.exit()


def login():
    """loging in"""
    os.system('"C:/Program Files (x86)/SAP/FrontEnd/SAPgui/SAPgui.exe" /M/r3-poedb.pl.energy.gov.ua/S/3600/G/PUBLIC')
    # click2('r3.png')
    click('login.png')
    write('vtv3')
    click('password.png')
    write('1234')
    key('enter')
    while pyautogui.locateOnScreen(PICS_PATH + 'new_order.png') is None:
        if pyautogui.locateOnScreen(PICS_PATH + 'multiple_registration.png'):
            click('multiple_registration.png')
            key('enter')


def create_order():
    """creating an order"""
    click2('new_order.png')
    click('sector.png')
    key('backspace')
    write(SECTOR)
    key('enter')


def fill_order():
    """filling an order"""
    click('EDRPOU.png')
    write(EDRPOU)
    key('enter')
    click('choose_order.png')
    key('enter')
    click('material.png')
    sleep(SLEEPTIME*2)

    output = ''
    output = output + MATERIAL + '\t' + PILLARS + \
        '\t'*4 + str(50150) + '\t' + '\n'
    pyperclip.copy(output)
    press_and_release('ctrl + v')
    pyautogui.press("enter", presses=2, interval=1)
    print('key "enter" * 2')
    sleep(SLEEPTIME*2)


def fill_work_type():
    """filling lines work type"""
    click2('work_type1.png')
    sleep(SLEEPTIME*2.5)
    hotkey("shift", "f5")
    click('work_type2.png')
    write(WORK_TYPE)
    hotkey("shift", "f7")
    sleep(SLEEPTIME)


def paying_date():
    """filling lines names"""
    pyautogui.click(pyautogui.locateOnScreen(PICS_PATH + 'paying_date1.png',
                    confidence=0.8), clicks=10, interval=0.1)
    click('paying_date2.png')
    sleep(SLEEPTIME*2.5)
    click('paying_date3.png')
    key("tab")
    sleep(SLEEPTIME)
    write(REDLINE_DATE)
    key("tab")
    sleep(SLEEPTIME)
    write(MONTH1)
    key("tab")
    sleep(SLEEPTIME)
    write(year)
    key("tab")
    sleep(SLEEPTIME)
    write(MONTH2)
    key("tab")
    sleep(SLEEPTIME)
    write(year)
    sleep(SLEEPTIME)


def additional_information():
    """additional information adn save the order"""
    hotkey("shift", "f4")
    pyautogui.click(pyautogui.locateOnScreen(
        PICS_PATH + 'add_inf1.png', confidence=0.8), clicks=10, interval=0.1)
    click('add_inf2.png')

    sleep(SLEEPTIME)
    # write(OBJECT_NAME)
    pyperclip.copy(OBJECT_NAME)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    sleep(SLEEPTIME)
    key('tab')
    sleep(SLEEPTIME)
    # write(OBJECT_ADDRESS)
    pyperclip.copy(OBJECT_ADDRESS)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')

    click('add_inf3.png')
    pyautogui.dragRel(0, 350, 1)
    print('drag down 350')

    click('add_inf4.png')
    key('tab')
    sleep(SLEEPTIME)
    write(START_DATE)
    key('tab')
    sleep(SLEEPTIME)
    # write(OWNER_PERSON1)
    pyperclip.copy(OWNER_PERSON1)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    # write(OWNER_PERSON2)
    pyperclip.copy(OWNER_PERSON2)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    # write(CUSTOMER_PERSON1)
    pyperclip.copy(CUSTOMER_PERSON1)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    # write(CUSTOMER_PERSON2)
    pyperclip.copy(CUSTOMER_PERSON2)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    key('tab')
    sleep(SLEEPTIME)
    # write(CUSTOMER_POSITION)
    pyperclip.copy(CUSTOMER_POSITION)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    # write(CUSTOMER_NAME)
    pyperclip.copy(CUSTOMER_NAME)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    key('tab')
    sleep(SLEEPTIME)
    key('tab')
    sleep(SLEEPTIME)
    write(MONTHS_NUMBER)
    key('tab')
    sleep(SLEEPTIME)
    key('tab')
    sleep(SLEEPTIME)
    write(EMAIL_POE)
    key('tab')
    sleep(SLEEPTIME)
    write(CUSTOVER_EMAIL)
    sleep(SLEEPTIME)

    click('save_order1.png')
    click('save_order2.png')
    click('save_order3.png')
    click('save_order4.png')


def save_order_number():
    """saving order number"""
    click2('order_number1.png')
    click('order_number2.png')
    hotkey("ctrl", "a")
    sleep(SLEEPTIME)
    hotkey("ctrl", "c")
    number = pyperclip.paste()
    print('get to clipboard "' + str(number) + '"')
    with open('street_light_orders.txt', "a+", encoding="UTF-8") as f:
        f.write(f'{date}_{CUSTOMER_ORGANIZATION}_{number}\n')
        f.close()
    hotkey('shift', 'f3')
    return number


def make_contract(number):
    """making contract"""
    click2('make_contract1.png')
    sleep(SLEEPTIME*5)
    click('make_contract2.png')
    sleep(SLEEPTIME*5)
    # write(VARIANT)
    pyperclip.copy(VARIANT)
    sleep(SLEEPTIME)
    press_and_release('ctrl + v')
    sleep(SLEEPTIME*5)
    click('make_contract3.png')
    sleep(SLEEPTIME*5)
    write(number)
    sleep(SLEEPTIME*5)
    key('f8')
    sleep(SLEEPTIME*5)
    key('enter')
    sleep(SLEEPTIME*5)
    key('enter')
    sleep(SLEEPTIME*5)


def save_contract(number):
    """saving contract"""
    click('save_doc1.png')
    sleep(SLEEPTIME*5)
    click('save_doc2.png')
    sleep(SLEEPTIME*5)
    write('PDF!')
    key('enter')
    sleep(SLEEPTIME*10)
    pyautogui.hotkey('ctrl', 'shift', 's')
    print('hotkey "ctrl + shift + s"')
    sleep(SLEEPTIME*5)
    # write(r'Y:\shared\ВРСП\Вуличне освітлення\2022_Договори' + '\\' +
    #       date + '_' + number + '_Договір ВО_' + CUSTOMER_ORGANIZATION + '_' + PILLARS + ' опор')
    write(
        rf'Y:\shared\ВРСП\Вуличне освітлення_2020-2022\2022_Договори\{date}_{number}_Договір ВО_{CUSTOMER_ORGANIZATION}_{PILLARS} опор')
    sleep(SLEEPTIME*10)
    click('save_doc3.png')
    MESSAGE4 = 'Проект договору лежить в:\n "Y:\shared\ВРСП\Вуличне освітлення\\2022_Договори".'
    pyautogui.alert(text=MESSAGE4, title='Програма завершена!', button='Ok')


if __name__ == '__main__':

    allert()
    check_layout()

    pyautogui.FAILSAFE = True

    SECTOR = "4"
    WORK_TYPE = "04"
    MATERIAL = "27001903"
    SLEEPTIME = 0.2
    EMAIL_POE = "vpr30@poe.pl.ua"
    PICS_PATH = 'pics/street_light/'

    date = arrow.now().format('YY-MM-DD')
    year = arrow.now().format('YYYY')

    with open("street_light_input.txt", "r", encoding="UTF-8") as file:
        EDRPOU = file.readline().replace('\n', '')
        CUSTOMER_ORGANIZATION = file.readline().replace('\n', '')
        PILLARS = file.readline().replace('\n', '')
        REDLINE_DATE = file.readline().replace('\n', '')
        MONTH1 = file.readline().replace('\n', '')
        MONTH2 = file.readline().replace('\n', '')
        OBJECT_NAME = file.readline().replace('\n', '')
        OBJECT_ADDRESS = file.readline().replace('\n', '')
        START_DATE = file.readline().replace('\n', '')
        OWNER_PERSON1 = file.readline().replace('\n', '')
        OWNER_PERSON2 = file.readline().replace('\n', '')
        CUSTOMER_PERSON1 = file.readline().replace('\n', '')
        CUSTOMER_PERSON2 = file.readline().replace('\n', '')
        CUSTOMER_POSITION = file.readline().replace('\n', '')
        CUSTOMER_NAME = file.readline().replace('\n', '')
        MONTHS_NUMBER = file.readline().replace('\n', '')
        CUSTOVER_EMAIL = file.readline().replace('\n', '')
        VARIANT = file.readline().replace('\n', '')
        file.close()

    login()
    create_order()
    fill_order()
    fill_work_type()
    paying_date()
    additional_information()
    order_number = save_order_number()
    make_contract(order_number)
    save_contract(order_number)
