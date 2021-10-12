#!/usr/bin/env python3

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='File Conversion Tool', bg = 'lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)

def getExcel ():
    global read_file
    global data
    import_file_path = filedialog.askopenfilename()
#     read_file = pd.read_excel (import_file_path)
    read_file = openpyxl.load_workbook(import_file_path)
    sheet = read_file.active
    data = sheet.rows

browseButton_Excel = tk.Button(text="      Import Excel File     ", command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_Excel)

def convertToCSV ():
    global read_file
    global data

    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
    csv = open(export_file_path, "w+")
    count=0
    for row in data:
       # print(row)
       if(count>=3):
              csv.write('\n')
              l = list(row)
              for i in range(len(l)):
                     eachWord=str(l[i].value)
                     if "," in eachWord:
                            eachWord='"'+eachWord+'"'
                     if i == len(l) - 1:
                            csv.write(eachWord)
                     else:
                            csv.write(eachWord + ',')
       count=count+1    
    csv.close()

saveAsButton_CSV = tk.Button(text='Convert Excel to CSV', command=convertToCSV, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=saveAsButton_CSV)

def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
     
exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=exitButton)

root.mainloop()