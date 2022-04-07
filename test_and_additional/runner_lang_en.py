import os
import pyautogui
from time import sleep


def lang_en_manualy():
    """Layout to en manualy"""
    pyautogui.keyDown('win')
    pyautogui.press('space')
    sleep(0.2)
    pyautogui.click(x=1770, y=540)
    pyautogui.keyUp('win')
    sleep(0.2)


lang_en_manualy()

os.system('"versions\\R3_order_spec_last.exe"')
