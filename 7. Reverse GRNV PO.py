import pyautogui
import ctypes
from pynput import keyboard


# make sure JDE is at Purchasing--Purchasing Corrections--Open Receipt-Reversal Inquiry inferface

# put "OPen receipt _Reversal Inquiry " barely above taskbar

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1.0


def on_press(key):
    if key == keyboard.Key.esc:
        print("quit script ")
        quit()
        return False  # stop listener
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys
    if k in ["g", "G"]:
        print("Proceed to Reverse")
        return False
    if k in ["right", "left", "up", "down", "enter"]:
        return False


def main():
    ctypes.windll.user32.SetCursorPos(
        100, 1024
    )  # open another instance of Reversal inquiry
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(610, 240)  # enter PO
    pyautogui.click()
    pyautogui.dragTo(310, 240, 0.5, button="left")  # drag so it erases *
    # Listen in to allow entering PO
    listener = keyboard.Listener(on_press=on_press)
    listener.start()  # start to listen on a separate thread
    listener.join()

    ctypes.windll.user32.SetCursorPos(360, 170)  # click find
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(320, 500)  # click first line in grid
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(320, 530)  # move mouse to next line
    # stop and allow user to select grid
    # press g to proceed reversing
    listener = keyboard.Listener(on_press=on_press)
    listener.start()  # start to listen on a separate thread
    listener.join()  # remove if main thread is polling self.keys
    #
    ctypes.windll.user32.SetCursorPos(455, 170)  # select Row
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(470, 465)  # select Reverse Receipt
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(400, 170)  # click close
    pyautogui.click()
    ctypes.windll.user32.SetCursorPos(455, 350)  # click OK to confirm
    pyautogui.click()


while True:  # repeat until the try statement succeeds   # or "a+", whatever you need
    try:
        main()
    except IOError:
        quit()
