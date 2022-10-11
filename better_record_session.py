import threading

import PySimpleGUI as sg
import time
import wmi
import win32gui, win32con
import os
import re
import sys
import time
import json
from datetime import datetime as date
from pynput.keyboard import Key, Controller

import pendulum
import pendulum
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pythoncom

import json
from pynput.keyboard import Key, Controller

import glob
import os

import os

import jdatetime

from tinydb import TinyDB, Query


class DB:
    db = None

    def __init__(self):
        self.db = TinyDB('db.json')
        # return db

    def records_to_list(self):
        datas = self.db.all()
        for a in iter(self.db):
            # jdatetime.
            print(a)
        data_list = []
        for data in datas:
            # file_name, duration, link
            data_list.append([data['uuid'],data['file_name'], data['duration'], data['link']])

        return data_list


def change_video_name(dars="dars"):
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    video_folder = desktop + '\\recorded\\'
    print(desktop)

    list_of_files = glob.glob(video_folder + '*')  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)
    filename = str(dars) + "_" + jdatetime.datetime.now().strftime("%d %b %Y %H-%M") + ".mp4"
    os.rename(latest_file, video_folder + filename)
    # print(jdatetime.datetime.now().strftime("%d %b %Y %H:%M"))


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def get_crd():
    path_to_json = resource_path("./crd.json")

    with open(path_to_json, "r") as handler:
        info = json.load(handler)
    return info
    # users = info["users"]
    # passwords = info["passwords"]


def get_url(string):
    # findall() has been used
    # with valid conditions for urls in string
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    url = re.findall(regex, string)
    return [x[0] for x in url]


# time.sleep(10)

def maximize_adobe_connect():
    hwnd = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)


# import wmi library


def force_close_adobe_connect():
    pythoncom.CoInitialize()
    # pythoncom.

    ti = 0

    name = 'connect.exe'

    f = wmi.WMI()

    for process in f.Win32_Process():

        if process.name == name:
            process.Terminate()
            print("closed succesfully")

            ti += 1

    if ti == 0:
        print("Process not found!!!")


# force_close_adobe_connect()

def which_class():
    for classroom in get_crd()['class']:
        today = date.today().strftime("%A")

        now = date.now()

        hour_from = pendulum.parse(classroom['hour_from'])
        hour_to = pendulum.parse(classroom['hour_to'])
        hour_now = pendulum.parse(now.strftime("%H:%M:%S"))

        if today == classroom['day']:
            if hour_from.timestamp() < hour_now.timestamp() < hour_to.timestamp():
                print("kelas hast ")
                print(classroom['name'])
                return classroom

    print("There is no Class")
    return False
    # sys.exit()


def is_class_finished():
    for classroom in get_crd()['class']:
        today = date.today().strftime("%A")

        now = date.now()

        hour_from = pendulum.parse(classroom['hour_from'])
        hour_to = pendulum.parse(classroom['hour_to'])
        hour_now = pendulum.parse(now.strftime("%H:%M:%S"))

        if today == classroom['day']:
            if hour_from.timestamp() < hour_now.timestamp() < hour_to.timestamp():
                print(classroom['name'])
                print("kelas dar haale ejras")
                return False
            else:
                return True

    # print("There is no Class")
    # sys.exit()
    return True


def go_to_class(link, duration, filename):
    selected_class_link = link
    # if

    username = get_crd()['username']
    password = get_crd()['password']
    # print(username,password)

    s = Service(resource_path('./driver/chromedriver.exe'))
    driver = webdriver.Chrome(service=s)

    # driver = webdriver.Chrome("chromedriver.exe")
    while True:
        try:
            driver.get("https://vc4011.kntu.ac.ir/login/index.php")
            break
        except Exception as e:
            print(e)
    username_field = driver.find_element(By.ID, "username")
    username_field.send_keys(username)
    password_field = driver.find_element(By.ID, "password")
    password_field.send_keys(password)

    login_button = driver.find_element(By.ID, "loginbtn")
    login_button.click()
    print(selected_class_link)
    driver.get(selected_class_link)
    # driver.get("https://vc4002.kntu.ac.ir/mod/adobeconnect/view.php?id=3393")
    # fanavari
    # driver.get("https://vc4002.kntu.ac.ir/mod/adobeconnect/view.php?id=2931")

    # class_link = driver.find_element(By.XPATH,
    #                                  "/html/body/div[1]/div[2]/div/div[3]/div/section/div/div[1]/form/div[3]/div/input")
    # print(class_link)
    #
    # adobe_connect_link = class_link.get_attribute('onclick')
    # print(adobe_connect_link)
    # print(get_url(adobe_connect_link))
    # driver.get(get_url(adobe_connect_link)[0])

    # keyboard = Controller()
    # keyboard.press(Key.left)
    # keyboard.release(Key.left)
    # keyboard.press(Key.enter)
    # keyboard.release(Key.enter)
    time.sleep(5)
    keyboard = Controller()
    keyboard.press(Key.left)
    keyboard.release(Key.left)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

    time.sleep(5)
    maximize_adobe_connect()
    keyboard.press(Key.esc)
    keyboard.release(Key.esc)
    time.sleep(3)
    print("hot keys for start recording ")

    keyboard.press(Key.f8)
    keyboard.release(Key.f8)
    # time.sleep(15)

    time.sleep(5)
    with keyboard.pressed(Key.ctrl):
        keyboard.press("\\")
        keyboard.release('\\')

    time.sleep(duration)
    keyboard.press(Key.f8)
    keyboard.release(Key.f8)
    print("close adobe connect!")
    force_close_adobe_connect()
    print("close chrome")

    driver.close()

    print("kelas tamoom shod")
    change_video_name(filename)
    with keyboard.pressed(Key.ctrl):
        keyboard.press("\\")
        keyboard.release('\\')
    return

    # while True:
    #     if is_class_finished():
    #         print("hot keys for stop recording ")
    #
    #
    #         # break
    #     else:
    #         print("wait for finishing class ...")
    #         time.sleep(10)
    #
    #         # return driver


# course_name="felan dars"

# Define the window's contents
# def UI():
#     layout = [
#         [sg.Text("link video", key="course_name")],
#         [sg.Input(key='LINK')],
#         [sg.Text("file name", key="course_name")],
#         [sg.Input(key='FILE_NAME')],
#         [sg.Text("duration in Hour and Minutes", key="course_name")],
#         [sg.Text("H",), sg.Input(key='HOUR', size=(20, 20),default_text=0), sg.Text("M"), sg.Input(key='MINUTE', default_text=0,size=(20, 20))],
#
#         # [sg.Text(size=(40,1), key='-OUTPUT-')],
#         [sg.Button('Record Class')],
#         [sg.Button('Quit')]
#     ]
#
#     # Create the window
#     window = sg.Window('kntu recorder', layout, finalize=True)
#     window.minimize()
#
#     # Display and interact with the Window using an Event Loop
#     while True:
#         print("1")
#
#         event, values = window.read()
#
#         # See if user wants to quit or window was closed
#         if event == sg.WINDOW_CLOSED or event == 'Quit':
#             break
#         # Output a message to the window
#         # if event == "Record Class":
#         #     print(values['LINK'])
#         #     link = values['LINK']
#         #     file_name = values['FILE_NAME']
#         #     duration = int(values['HOUR']) * 3600 + int(values['MINUTE']) * 60
#         #
#         #     print(link)
#         #     print(file_name)
#         #     print(duration)
#         #     print("stop kon")
#         #     window.close()
#         #     go_to_class(link=link,filename=file_name,duration=duration)
#
#         # window['-OUTPUT-'].update('Hello ' + values['-INPUT-'] + "! Thanks for trying PySimpleGUI")
#
#     # Finish up by removing from the screen
#     window.close()

def start_recording_classes(record_list, window):
    for record in record_list:
        print(record)
        go_to_class(link=str(record[2]), filename=record[0], duration=record[1])
        window['_table_'].Update(row_colors=[[record_list.index(record), 'green']])
        time.sleep(3)
    # print('Starting thread - will sleep for {} seconds'.format(seconds))
    # time.sleep(seconds)                  # sleep for a while
    # window.write_event_value('-THREAD-', '** DONE **')  # put a message into queue for GUI


def UI():
    db_manager = DB()

    record_list = db_manager.records_to_list()
    data_headings = ['id','file_name', 'duration', 'link']
    layout = [
        [sg.Text("link video", key="course_name")],
        [sg.Input(key='LINK')],
        [sg.Text("file name", key="course_name")],
        [sg.Input(key='FILE_NAME')],
        [sg.Text("duration in Hour and Minutes", key="course_name")],
        [sg.Text("H", ), sg.Input(key='HOUR', size=(20, 20), default_text=0), sg.Text("M"),
         sg.Input(key='MINUTE', default_text=0, size=(20, 20))],

        # [sg.Text(size=(40,1), key='-OUTPUT-')],
        [sg.Table(values=record_list, headings=data_headings,
                  max_col_width=35,
                  col_widths=[5, 8, 35],
                  auto_size_columns=True,
                  justification='left',
                  enable_events=True,
                  num_rows=6, key='_table_',
                  expand_x=True,
                  expand_y=True,
                  )

         ],
        [sg.Button('Add Class'), sg.Button('Record Class')],
        [sg.Button('Quit')]
    ]

    # Create the window
    window = sg.Window('kntu recorder', layout, finalize=True,resizable=True)

    # Display and interact with the Window using an Event Loop

    # record_list = []
    while True:
        print("1")

        event, values = window.read()

        # See if user wants to quit or window was closed
        if event == sg.WINDOW_CLOSED or event == 'Quit':
            break
        # Output a message to the window

        if event == "Add Class":
            print(values['LINK'])
            link = values['LINK']
            file_name = values['FILE_NAME']
            duration = int(values['HOUR']) * 3600 + int(values['MINUTE']) * 60
            db_manager.db.insert({"uuid":jdatetime.datetime.now().strftime("%Y%m%d%H%M%S"),"link": link, "file_name": file_name, "duration": duration})

            print(link)
            print(file_name)
            print(duration)
            print("stop kon")
            # record_list.append([file_name, duration, link])
            record_list = db_manager.records_to_list()

            window['_table_'].Update(values=record_list)

        if event == "Record Class":
            window.minimize()

            threading.Thread(target=start_recording_classes, args=(record_list, window,), daemon=True).start()

        if event == "_table_":
            data_selected = [record_list[row] for row in values[event]]
            print(data_selected)


    # Finish up by removing from the screen
    window.close()


UI()
