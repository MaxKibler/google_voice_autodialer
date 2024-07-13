import pandas as pd
import pyautogui as auto
from tkinter import *
import keyboard
import time

# path to xlsx with phone numbers goes here
sheet = pd.read_excel('')

# make sure column with phone numbers is named Phone1
phone_number = sheet['Phone1']
phone_list_int = sheet['Phone1'].tolist()
phone_list_0 = [str(num) for num in phone_list_int]
phone_list = [num.replace('.0', '') for num in phone_list_0]

# These are the xy coordinates for button locations on the google voice window as well as the tkinter window that opens.
# If you need to adjust these values run xy_tool.py to find coordinates.
call_box_right = 1735, 442
call_box_left = 1489, 445
call_box_middle = 1604, 441
call_button = 1839, 437
hang_up = 1691, 978
screen_center = 997, 517
screen_bottom_right = 1919, 1079
screen_top_left = 0, 0
screen_bottom_left = 0, 1079

answer_count = 0
answer_count_list = []
bad_num_count = 0


def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry('{}x{}+{}+{}'.format(width, height, x, y))


def create_window():
    global root
    root = Tk()
    root.title("Info")
    root.bind('<q>', close_window)
    label = Label(root, text=f"{phone_number[index]}", font=("Helvetica", 14))
    root.focus_force()
    root.after(100, check_and_center_window)
    label.pack()
    if keyboard.is_pressed("v"):
        print("null")
    center_window(root)
    root.mainloop()


def check_and_center_window():
    if root.winfo_exists():
        center_window(root)
    else:
        root.after(100, check_and_center_window)


def close_window(event):
    root.destroy()


for index, number in enumerate(phone_list):
    auto.moveTo(call_box_middle)
    auto.click()
    auto.press('del', presses=20)
    auto.moveTo(call_box_middle)
    auto.click()
    time.sleep(.25)
    auto.write(number)
    time.sleep(.5)
    auto.moveTo(call_button)
    auto.click()
    create_window()
    auto.moveTo(hang_up)
    auto.click()
    time.sleep(2)

