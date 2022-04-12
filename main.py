import time
import traceback
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

root = tk.Tk()
root.title('Регистрация участников')
root.geometry("800x600")
# root.withdraw()
cap = cv2.VideoCapture(0)

lbimage = tk.Label()
lbimage.pack(side="top", anchor="ne")
lbText = tk.Label()
lbText.place(anchor=CENTER, relx=0.5, rely=0.5)


# file = 'test.xlsx'
# xl = pd.ExcelFile(file)
# sheets = xl.sheet_names
# active_sheet = xl.sheet_names[0]
# wordbook = load_workbook(filename=file)
# sheet = wordbook[active_sheet]
# df = pd.read_excel(file, sheet_name=active_sheet)
# saved = df.values.tolist()
# print(active_sheet)
# print(saved)
#
# today = date.today()
# d1 = today.strftime("%d/%m/%Y").replace('/','-')
# # d1 = "08-04-2022"
# print("d1 =", d1)
# if d1 not in sheets:
#     wordbook.create_sheet(d1)
#     wordbook[d1]['A1'] = 'ФИО'
#     wordbook[d1]['B1'] = 'Посещение'
#     wordbook.save(file)

def capture():
    try:
        # file = 'fio.xlsx'
        file = 'test.xlsx'
        xl = pd.ExcelFile(file)
        sheets = xl.sheet_names
        fio_sheet = xl.sheet_names[0]
        wordbook = load_workbook(filename=file)
        # cur_sheet = wordbook[fio_sheet]
        df = pd.read_excel(file, sheet_name=fio_sheet)
        fio_list = df.values.tolist()
        # fio_list = [item[0] for item in fio_list]
        fio_map = {item[0]: [item[0], item[1], item[2], item[3]] for item in fio_list}

        success, frame = cap.read()
        if success:
            frame = imutils.resize(frame, width=320)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = PIL.ImageTk.PhotoImage(image=Image.fromarray(frame))
            lbimage.imgtk = img
            lbimage.configure(image=img)
            today = date.today()
            today = today.strftime("%d/%m/%Y").replace('/', '-')
            for code in decode(frame):

                if today not in sheets:
                    wordbook.create_sheet(today)
                    wordbook[today]['A1'] = 'Номер'
                    wordbook[today]['B1'] = 'ФИО'
                    wordbook[today]['C1'] = 'Класс'
                    wordbook[today]['D1'] = 'Посещение'
                    wordbook.save(file)

                df = pd.read_excel(file, sheet_name=today)
                saved_visitors = df.values.tolist()
                saved_visitors = [int(item[0]) for item in saved_visitors]

                if int(code.data.decode('utf-8')) not in saved_visitors:
                    user_id = int(code.data.decode('utf-8'))
                    cur_sheet = wordbook[today]
                    cur_sheet.append(fio_map[user_id])
                    wordbook.save(filename=file)
                    # df = pd.read_excel(file, sheet_name=fio_sheet)
                    # saved_visitors = df.values.tolist()
                    # saved_visitors = [item[0] for item in saved_visitors]
                else:
                    lbText.config(text="Последний пользователь " + str(saved_visitors[-1]), font=("roboto", 30))
                    print("Последний пользователь " + code.data.decode('utf-8'))

            root.after(1, capture)
    except Exception as exc:
        print(exc)
        print(traceback.format_exc())


root.after(1000, capture)
root.mainloop()
