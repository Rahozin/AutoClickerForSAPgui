"""Making bill and aps"""

from time import sleep
from datetime import datetime, timedelta
import sys
import pyautogui
import keyboard
import pyperclip
import arrow
import win32api


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
    click2('r3.png')
    click('login.png')
    write('vtv3')
    click('password.png')
    write('1234')
    key('enter')
    while pyautogui.locateOnScreen(PICS_PATH + 'new_order.png') is None:
        if pyautogui.locateOnScreen(PICS_PATH + 'multiple_registration.png'):
            click('multiple_registration.png')
            key('enter')


def fill_contract():
    """filling contract number and date"""
    click2('fill_contract1.png')
    sleep(SLEEPTIME*2)
    pyperclip.copy(order_number)
    press_and_release('ctrl + v')
    sleep(SLEEPTIME*2)
    click('fill_contract2.png')
    click('fill_contract3.png')
    sleep(SLEEPTIME*2)
    click('fill_contract4.png')
    click('fill_contract5.png')
    click('fill_contract6.png')
    click('fill_contract7.png')
    hotkey("ctrl", "a")
    write(contract_number)
    key("tab")
    write(contract_date)
    click('fill_contract8.png')
    click('fill_contract9.png')


def make_bill():
    """making bill"""
    f_type = "рахунок"
    click2('make_bill1.png')
    sleep(SLEEPTIME*5)
    write(order_number)
    pyautogui.scroll(-5000)
    print('scroll down 5000')
    click('make_bill2.png')
    click('make_bill3.png')
    key('f8')
    sleep(SLEEPTIME*5)
    return f_type


def make_aps():
    """making aps"""
    f_type = "АНП"
    click2('make_aps1.png')
    sleep(SLEEPTIME*5)
    write(order_number)
    pyautogui.scroll(-5000)
    print('scroll down 5000')
    click('make_aps2.png')
    click('make_aps3.png')
    key('tab')
    hotkey('ctrl', 'a')
    # write(SIGN1)
    pyperclip.copy(SIGN1)
    press_and_release('ctrl + v')
    key('tab')
    hotkey('ctrl', 'a')
    write(SIGN2)
    key('f8')
    return f_type


def go_main_manu():
    """return to main manu"""
    while pyautogui.locateOnScreen(PICS_PATH + 'save_file_checker2.png', confidence=0.9):
        pyautogui.hotkey('alt', 'f4')
        print('hotkey "alt + f4"')
        sleep(SLEEPTIME*5)
    # click('go_main_menu1.png')
    # click('go_main_menu2.png')
    # click('fill_contract1.png')
    click('new_session_after_bill.png')


def save_file(f_type):
    """saving file"""
    click('save_file1.png')
    sleep(SLEEPTIME*5)
    click('save_file2.png')
    sleep(SLEEPTIME*10)
    write('PDF!')
    key('enter')
    sleep(SLEEPTIME*5)
    while pyautogui.locateOnScreen(PICS_PATH + 'save_file_checker1.png', confidence=0.9) is None:
        pyautogui.hotkey('ctrl', 'shift', 's')
        print('hotkey "ctrl + shift + s"')
        sleep(SLEEPTIME*5)
    write(r'Y:\shared\ВРСП\Договора на ТУ\_Рахунки і АНП' + '\\' +
          date + '_' + contract_number + '_від_' + contract_date + '_' + f_type)
    sleep(SLEEPTIME*5)
    click('save_file3.png')
    if f_type == "АНП":
        MESSAGE4 = 'Рахунок і АНП лежать в:\n "Y:/shared/ВРСП/Договора на ТУ/_Рахунки і АНП".'
        pyautogui.alert(
            text=MESSAGE4, title='Програма завершена!', button='Ok')


if __name__ == '__main__':

    pyautogui.FAILSAFE = True

    SLEEPTIME = 0.2
    PICS_PATH = 'pics/bill&aps/'
    PICS_PATH_CHECK = 'pics/check/'
    SIGN1 = 'Директор сервісного центру Гопка О.М.'
    SIGN2 = '.'

    date = arrow.now().format('YYMMDD')

    check_layout()

    # order_number = '552517'
    order_number = pyautogui.prompt("Номер замовлення: ")
    if order_number is None:
        sys.exit()

    # contract_number = '93'
    contract_number = pyautogui.prompt("Номер договора: ")

    # contract_date = '05.01.2022'
    contract_date = pyautogui.prompt("Дата договора у форматі DD.MM.YYYY: ")

    login()
    fill_contract()
    file_type = make_bill()
    save_file(file_type)
    go_main_manu()
    file_type = make_aps()
    save_file(file_type)
