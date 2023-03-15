def on_press(key):
    if key == keyboard.Key.esc:
        quit()
        return False  # stop listener
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys

    if k in ["g", "G"]:
        print("PROCEED TO REVERSE PO")
        return False


import pyautogui
import ctypes
import os
from pynput import keyboard

""" for i in range(10):
      pyautogui.moveTo(100, 100, duration=0.25)
      pyautogui.moveTo(200, 100, duration=0.25)
      pyautogui.moveTo(200, 200, duration=0.25)
      pyautogui.moveTo(100, 200, duration=0.25) """
# ctypes.windll.user32.SetCursorPos(315,455)
# pyautogui.click()#select all

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1.5

# after entering PO number:
ctypes.windll.user32.SetCursorPos(350, 165)  # find
pyautogui.click()

# click select all
ctypes.windll.user32.SetCursorPos(320, 460)  # select all
pyautogui.click()

# wait for untick bottom total and confirmation
listener = keyboard.Listener(on_press=on_press)
listener.start()  # start to listen on a separate thread
listener.join()

# click row-reverse receipt
ctypes.windll.user32.SetCursorPos(455, 165)
pyautogui.click()
ctypes.windll.user32.SetCursorPos(465, 462)
pyautogui.click()

# click close, and confirm
ctypes.windll.user32.SetCursorPos(410, 165)
pyautogui.click()
ctypes.windll.user32.SetCursorPos(450, 350)
pyautogui.click()

# click "open-receipt-reversal inquiry" to begin inputting new PO
ctypes.windll.user32.SetCursorPos(150, 720)
pyautogui.click()


# wait for signs to continue

# pyautogui.write('i3614795',interval=0.3)
# pyautogui.hotkey('tab')

# pyautogui.click()#goto end
