# import autoclicker
import pyautogui
# import autoclicker
from time import sleep


def lang_en():
    """layout to en"""
    while pyautogui.locateOnScreen(PICS_PATH + 'lang_en.png',
                                   region=(1800, 1030, 120, 50), confidence=0.8) is None:
        pyautogui.hotkey("ctrl", "shift")
        sleep(SLEEPTIME)


SLEEPTIME = 0.2
PICS_PATH = 'pics/'

lang_en()
exec(open('autoclicker.py', encoding="UTF-8").read())
# autoclicker
# sys.run(autoclicker.py)
