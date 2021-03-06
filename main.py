import time
from openpyxl import Workbook, load_workbook
import cv2
import pandas as pd
from pyzbar.pyzbar import decode
import tkinter
from tkinter import messagebox

root = tkinter.Tk()
root.withdraw()
cap = cv2.VideoCapture(0)

file = 'fio.xlsx'
xl = pd.ExcelFile(file)
active_sheet = xl.sheet_names[0]
wordbook = load_workbook(filename=file)
sheet = wordbook[active_sheet]
df = pd.read_excel(file, sheet_name=active_sheet)
saved = df.values.tolist()
saved = [item[0] for item in saved]

while True:
    success, frame = cap.read()
    for code in decode(frame):
        print(code.data.decode('utf-8'))
        print(saved)
        if str(code.data.decode('utf-8')) not in saved:
            new_fio = code.data.decode('utf-8')
            sheet.append([new_fio])
            wordbook.save(filename=file)
            df = pd.read_excel(file, sheet_name=active_sheet)
            saved = df.values.tolist()
            saved = [item[0] for item in saved]
            time.sleep(1)
        else:
            messagebox.showinfo("Title", "Пользователь " + str(code.data.decode('utf-8')) + " уже существует!")
    cv2.imshow('Testing-code-scan', frame)
    cv2.waitKey(10)
# cap.release()

