import ReadPySource as rps
import docx
from docx.shared import Inches
import pyautogui as pyg
import time

print("Preparing ...")
time.sleep(0.5)
aim="WAP to calculate area of a circle."
code=rps.read()

pyg.keyDown('win')
pyg.press('r')
pyg.keyUp('win')
time.sleep(0.5)
pyg.typewrite('cmd')
pyg.press('enter')
print("Got Shell !")
time.sleep(0.5)
pyg.typewrite('D:')
pyg.press('enter')
print("Accessing Directory ...")
time.sleep(0.5)
pyg.typewrite('cd D:\\Python\\ProPy')
pyg.press('enter')
print("Running It Now ...")
time.sleep(0.5)
pyg.typewrite('python Script.py')
pyg.press('enter')

print("Screening ...")
time.sleep(1)
pyg.screenshot("out.png")

print("Taking 5 ...")
time.sleep(5)

print("Exiting ...")
pyg.typewrite('exit')
pyg.press('enter')
time.sleep(0.5)

print("Creating Document ...")

doc=docx.Document()

p=doc.add_paragraph("")
p.add_run("Aim : ").bold=True
p.add_run(aim)
p=doc.add_paragraph("")
p.add_run("Code :").bold=True
for c in code:
    doc.add_paragraph(c)
p=doc.add_paragraph("")
p.add_run("Output :").bold=True
doc.add_picture("out.png",width=Inches(5.5))
print("Saving Document ...")
doc.save('done.docx')
