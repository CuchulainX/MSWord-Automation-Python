from time import sleep
import win32com.client as win32
from Tkinter import *
master = Tk()
import os
import tkMessageBox
from tkFileDialog   import askopenfilename
#Import your feature file for which yyou want to automate
from CountTableCol9 import CountTableColumn9
from CountTableCol12 import CountTableColumn12
path_to_script = os.path.dirname(os.path.abspath(__file__))

def Col9():
    PVFileName= askopenfilename() 
    CountTableColumn9(PVFileName)
    sleep(5)
    tkMessageBox.showinfo(title="MS Word document Checker Completed", message="Report.txt is generated")
def Col12():
    PVFileName= askopenfilename() 
    CountTableColumn12(PVFileName)
    sleep(5)
    tkMessageBox.showinfo(title="MS Word document Checker Completed", message="Report.txt is generated")
errmsg = 'Error!'
Button(text='Select Word document for 9 column Check', command=Col9).pack(fill=X)
Button(text='Select Word document for 12 column Check', command=Col12).pack(fill=X)
mainloop()
