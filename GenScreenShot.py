import pyautogui as pyg
import time
import os

path = os.path.dirname(os.path.realpath(__file__))
fn="Screen.png"

def genscreen(filename="Script.py"):
    #Run CMD
    pyg.keyDown('win')
    pyg.press('r')
    pyg.keyUp('win')
    time.sleep(0.5)
    pyg.typewrite('cmd')
    pyg.press('enter')
    print("Got Shell !")
    time.sleep(0.5)
    #Change Directory
    pyg.typewrite(path[:2])
    pyg.press('enter')
    print("Accessing Directory ...")
    time.sleep(0.5)
    pyg.typewrite('cd '+path)
    pyg.press('enter')
    print("Running It Now ...")
    time.sleep(0.5)
    #Execute Python Script
    pyg.typewrite('python '+filename)
    pyg.press('enter')
    print("Screening ...")
    time.sleep(1)
    #Take Screenshot
    pyg.screenshot(fn)
    print("Catching A Breath ...")
    time.sleep(1)
    print("Exiting ...")
    pyg.typewrite('exit')
    pyg.press('enter')
    time.sleep(0.5)
    return fn
