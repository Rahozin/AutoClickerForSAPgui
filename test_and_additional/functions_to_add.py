def lang_en():
    """Layout to en"""
    while pyautogui.locateOnScreen(PICS_PATH + 'lang_en.png',
                                   region=(1800, 1030, 120, 50), confidence=0.8) is None:
        pyautogui.hotkey("ctrl", "shift")
        sleep(SLEEPTIME)


def lang_en_manualy():
    """Layout to en manualy"""
    pyautogui.keyDown('win')
    key('space')
    sleep(SLEEPTIME)
    pyautogui.click(x=1770, y=540)
    pyautogui.keyUp('win')
    sleep(SLEEPTIME)


def check_done(pic_name, i):
    """Checking if event was done"""

    time_start = datetime.now()
    time_finish = time_start + timedelta(seconds=5)
    check = True
    while check:
        target = pyautogui.locateOnScreen(
            PICS_PATH_CHECK + pic_name, confidence=0.95)
        if target:
            break
        if time_start >= time_finish:
            print('Була пропущена лінія.' + str(i+1) + 'Повторне внесення')
            i = i - 1
            check = False
        time_start = datetime.now()
        sleep(SLEEPTIME)
    return i
