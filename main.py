import time
from tkinter.filedialog import askopenfile

from openpyxl import Workbook, load_workbook
import cv2
import pandas as pd
from pyzbar.pyzbar import decode
import tkinter as tk
from tkinter import messagebox, CENTER
import imutils
from PIL import Image, ImageTk
import PIL
from datetime import date

# root = tk.Tk()
# root.title('Регистрация участников')
# root.geometry("800x600")
# root.withdraw()
# cap = cv2.VideoCapture(0)
#
# lbimage = tk.Label()
# lbimage.pack(side="top", anchor="ne")
# lbText = tk.Label()
# lbText.place(anchor=CENTER, relx=0.5, rely=0.5)

file = 'test.xlsx'
xl = pd.ExcelFile(file)
sheets = xl.sheet_names
active_sheet = xl.sheet_names[0]
wordbook = load_workbook(filename=file)
sheet = wordbook[active_sheet]
df = pd.read_excel(file, sheet_name=active_sheet)
saved = df.values.tolist()
print(active_sheet)
print(saved)
# print(sheets)

# today = date.today()
# d1 = today.strftime("%d/%m/%Y")
d1 = "08-04-2022"
print("d1 =", d1)
if d1 not in sheets:
    wordbook.create_sheet(d1)
    wordbook[d1]['A1'] = 'ФИО'
    wordbook.save(file)

# def capture():
#     try:
#         # file = 'fio.xlsx'
#         file = 'test.xlsx'
#         xl = pd.ExcelFile(file)
#         active_sheet = xl.sheet_names[0]
#         wordbook = load_workbook(filename=file)
#         sheet = wordbook[active_sheet]
#         df = pd.read_excel(file, sheet_name=active_sheet)
#         saved = df.values.tolist()
#         saved = [item[0] for item in saved]
#
#         success, frame = cap.read()
#         if success:
#             frame = imutils.resize(frame, width=320)
#             frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
#             img = PIL.ImageTk.PhotoImage(image=Image.fromarray(frame))
#             lbimage.imgtk = img
#             lbimage.configure(image=img)
#             for code in decode(frame):
#                 # print(saved)
#                 if int(code.data.decode('utf-8')) not in saved:
#                     new_fio = int(code.data.decode('utf-8'))
#                     sheet.append([new_fio])
#                     wordbook.save(filename=file)
#                     df = pd.read_excel(file, sheet_name=active_sheet)
#                     saved = df.values.tolist()
#                     saved = [item[0] for item in saved]
#                 else:
#                     lbText.config(text="Последний пользователь " + str(saved[-1]), font=("roboto", 30))
#                     print("Последний пользователь " + code.data.decode('utf-8'))
#
#             root.after(1, capture)
#     except Exception as exc:
#         print(exc)
#
#
# root.after(1000, capture)
# root.mainloop()
