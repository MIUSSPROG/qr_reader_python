import time
from openpyxl import Workbook, load_workbook
import cv2
import pandas as pd
from pyzbar.pyzbar import decode
import tkinter as tk
from tkinter import messagebox, CENTER
import imutils
from PIL import Image, ImageTk
import PIL

root = tk.Tk()
root.geometry("800x600")
# root.withdraw()
cap = cv2.VideoCapture(0)

lbimage = tk.Label()
lbimage.pack(side="top", anchor="ne")
# been = False
lbText = tk.Label()
lbText.place(anchor=CENTER, relx=0.5, rely=0.5)


def capture():
    try:
        file = 'fio.xlsx'
        xl = pd.ExcelFile(file)
        active_sheet = xl.sheet_names[0]
        wordbook = load_workbook(filename=file)
        sheet = wordbook[active_sheet]
        df = pd.read_excel(file, sheet_name=active_sheet)
        saved = df.values.tolist()
        saved = [item[0] for item in saved]

        success, frame = cap.read()
        if success:
            frame = imutils.resize(frame, width=320)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = PIL.ImageTk.PhotoImage(image=Image.fromarray(frame))
            lbimage.imgtk = img
            lbimage.configure(image=img)
            for code in decode(frame):
                # print(saved)
                if str(code.data.decode('utf-8')) not in saved:
                    new_fio = code.data.decode('utf-8')
                    sheet.append([new_fio])
                    wordbook.save(filename=file)
                    df = pd.read_excel(file, sheet_name=active_sheet)
                    saved = df.values.tolist()
                    saved = [item[0] for item in saved]
                else:
                    # lbText(text="Последний пользователь " + str(saved[-1]), font=("roboto", 30))
                    # lbText.text = "Последний пользователь " + str(saved[-1])
                    # lbText.font = ("roboto", 30)
                    lbText.config(text="Последний пользователь " + str(saved[-1]), font=("roboto", 30))
                    print("Последний пользователь " + code.data.decode('utf-8'))

            root.after(1, capture)
    except Exception as exc:
        print(exc)


root.after(1000, capture)
root.mainloop()

# while True:
#     success, frame = cap.read()
#     frame = imutils.resize(frame, width=320)
#     # frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
#     img = PIL.ImageTk.PhotoImage(image=Image.fromarray(frame))
#     lbimage.imgtk = img
#     lbimage.configure(image=img)
# for code in decode(frame):
#     print(code.data.decode('utf-8'))
#     print(saved)
#     if str(code.data.decode('utf-8')) not in saved:
#         new_fio = code.data.decode('utf-8')
#         sheet.append([new_fio])
#         wordbook.save(filename=file)
#         df = pd.read_excel(file, sheet_name=active_sheet)
#         saved = df.values.tolist()
#         saved = [item[0] for item in saved]
#         time.sleep(1)
#     else:
#         messagebox.showinfo("Title", "Пользователь " + str(code.data.decode('utf-8')) + " уже существует!")
# cv2.imshow('Testing-code-scan', frame)
# cv2.waitKey(10)

# root.after(1000, capture)
# root.mainloop()
# cap.release()
