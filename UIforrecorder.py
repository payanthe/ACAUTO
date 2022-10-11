import threading
import subprocess
import os.path
import PySimpleGUI as sg
import time
import wmi
import win32gui, win32con
import os
import re
import sys
import time
from tinydb import Query
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

from tinydb import TinyDB, Query, where
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler()
    ]
)


def change_video_name(dars="dars"):
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    video_folder = desktop + '\\recorded\\'
    logging.info(desktop)

    list_of_files = glob.glob(video_folder + '*')  # * means all if need specific format then *.csv
    latest_file = max(list_of_files, key=os.path.getctime)
    # logging.info(latest_file)
    filename = str(dars) + "_" + jdatetime.datetime.now().strftime("%d %b %Y %H-%M") + ".mp4"
    os.rename(latest_file, video_folder + filename)
    return (video_folder + filename)
    # logging.info(jdatetime.datetime.now().strftime("%d %b %Y %H:%M"))


def force_close_adobe_connect():
    pythoncom.CoInitialize()

    ti = 0

    name = 'connect.exe'

    f = wmi.WMI()

    for process in f.Win32_Process():

        if process.name == name:
            process.Terminate()
            logging.info("closed succesfully")

            ti += 1

    if ti == 0:
        logging.info("Process not found!!!")


def get_url(string):
    # findall() has been used
    # with valid conditions for urls in string
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    url = re.findall(regex, string)
    return [x[0] for x in url]


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


class DB:
    db = None

    def __init__(self):
        self.db = TinyDB('db.json')
        # return db

    def update_db(self):
        self.db = TinyDB('db.json')

    def records_to_list(self):
        self.update_db()
        datas = self.db.all()
        # for a in iter(self.db):
        #     # jdatetime.
        #     logging.info(a)
        data_list = []
        for data in datas:
            # file_name, duration, link
            data_list.append([data['uuid'], data['date'], data['name'], data['status'],data['start'],data['end']])

        return data_list

    @staticmethod
    def read_crd():
        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.dirname(__file__)
            return os.path.join(base_path, relative_path)

            # users = info["users"]
            # passwords = info["passwords"]

        path_to_json = resource_path("./crd.json")

        with open(path_to_json, "r") as handler:
            info = json.load(handler)
        return info

    def get_class_list(self):
        class_list = []
        for uni_class in self.read_crd()['class']:
            # class_headings = ['name', 'day', 'time']
            class_list.append(
                [uni_class['name'], uni_class['day'], uni_class['hour_from'] + " -- " + uni_class['hour_to']])

        return class_list

    def which_class(self):
        self.update_db()
        for classroom in self.read_crd()['class']:

            today = date.now().strftime("%A")

            now = date.now()

            hour_from = pendulum.parse(classroom['hour_from'])
            hour_to = pendulum.parse(classroom['hour_to'])
            hour_now = pendulum.parse(now.strftime("%H:%M:%S"))

            if today == classroom['day'] and hour_from.timestamp() < hour_now.timestamp() < hour_to.timestamp():
                logging.info("kelas hast ")
                logging.info(classroom['name'])
                self.update_db()
                logging.info(self.db.all())
                results = self.db.search(
                    Query().fragment({'name': classroom['name'], "date": jdatetime.datetime.now().strftime("%Y-%m-%d"),
                                      "status": "finished"}))
                # logging.info("inja resulte")
                # logging.info(results)
                if len(results):
                    return False
                return classroom

        # logging.info("There is no Class")
        return False
        # sys.exit()

    def which_class_to_list(self):
        now_class = self.which_class()

        if now_class is not False:
            return [now_class['name'], now_class['day'], now_class['hour_from'] + " -- " + now_class['hour_to']]
        return False

    def is_class_finished(self, class_name=None):
        self.update_db()
        nowclass = self.which_class()
        if nowclass is not False:
            results = self.db.search(
                Query().fragment({'name': nowclass['name'], "date": jdatetime.datetime.now().strftime("%Y-%m-%d")}))
            if results:
                for result in results:
                    if result['status'] == "finishit":
                        logging.info("CLASS FINISHED BY STOP IT")
                        return True
        for classroom in self.read_crd()['class']:
            today = date.today().strftime("%A")

            now = date.now()

            hour_from = pendulum.parse(classroom['hour_from'])
            hour_to = pendulum.parse(classroom['hour_to'])
            hour_now = pendulum.parse(now.strftime("%H:%M:%S"))

            if today == classroom['day']:
                if hour_from.timestamp() < hour_now.timestamp() < hour_to.timestamp():
                    if class_name is not None:
                        if class_name != classroom['name']:
                            # logging.info("choon classromm nam yeki nabood")
                            # logging.info(class_name)
                            # logging.info(classroom['name'])
                            return True
                    # logging.info(classroom['name'])
                    # logging.info("kelas dar haale ejras")
                    return False
                else:
                    if class_name == classroom['name']:
                        return True
                    # if class_name is not None:
                    #     if class_name != classroom['name']:
                    #         logging.info("choon classromm nam yeki nabood")
                    #         logging.info(class_name)
                    #         logging.info(classroom['name'])
                    #         return True
                    # logging.info("inja shodesh")
                    # return True

        # logging.info("There is no Class")
        # sys.exit()
        return False

    def add_recording_class_to_db(self):
        myclass = self.which_class()
        if myclass is not False:
            self.db.insert(
                {"uuid": jdatetime.datetime.now().strftime("%Y%m%d%H%M%S"), "status": "recording",
                 "name": myclass['name'], "file_name": "",
                 "date": jdatetime.datetime.now().strftime("%Y-%m-%d"),
                 "start": jdatetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S"), "end": ""})

    def change_status_to_finished(self, class_name,file_name):
        self.update_db()
        # db.update(delete('key1'), User.name == 'John')
        # myclass= Query()
        self.db.update({"status": "finished","file_name":file_name,"end":jdatetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")},
                       (where('name') == class_name) & (where('date') == jdatetime.datetime.now().strftime("%Y-%m-%d")))

    def change_status_to_finish_it(self):
        self.update_db()
        myclass = self.which_class()
        if myclass is not False:
            self.db.update({"status": "finishit"}, (where('name') == myclass['name']) & (
                    where('date') == jdatetime.datetime.now().strftime("%Y-%m-%d")))
    def get_class_record_by_uuid(self,uuid):
        self.update_db()
        datas=self.db.search(Query().uuid==uuid)
        if datas:
            return datas[0]
        else:
            return None

def UI():
    db_manager = DB()

    record_list = db_manager.records_to_list()
    class_list = db_manager.get_class_list()
    data_headings = ['id', 'date', 'name', 'status','start','end']
    class_headings = ['name', 'day', 'time']
    layout = [
        [sg.Text("status"), sg.Text("not runing", key="status")],
        [sg.Table(values=class_list, headings=class_headings,
                  max_col_width=35,
                  col_widths=[5, 8, 35],
                  auto_size_columns=True,
                  justification='left',
                  enable_events=True,
                  num_rows=6, key='_classtable_',
                  expand_x=True,
                  expand_y=True,
                  )

         ],

        # [sg.Input(key='LINK')],
        # [sg.Text("file name", key="course_name")],
        # [sg.Input(key='FILE_NAME')],
        # [sg.Text("duration in Hour and Minutes", key="course_name")],
        # [sg.Text("H", ), sg.Input(key='HOUR', size=(20, 20), default_text=0), sg.Text("M"),
        #  sg.Input(key='MINUTE', default_text=0, size=(20, 20))],

        # [sg.Text(size=(40,1), key='-OUTPUT-')],
        [sg.Table(values=record_list, headings=data_headings,
                  max_col_width=35,
                  col_widths=[5, 8, 35],
                  auto_size_columns=True,
                  # justification='left',
                  justification='center',
                  enable_events=True,
                  num_rows=6, key='_table_',
                  expand_x=True,
                  expand_y=True,

                  )

         ],
        [
            # sg.Button('Add Class'), sg.Button('Record Class'),
         sg.Button('Stop Class')],
        [sg.Button('Quit')]
    ]

    # Create the window
    window = sg.Window('kntu recorder', layout, finalize=True, resizable=True)

    # Display and interact with the Window using an Event Loop

    # record_list = []
    def greenify_selected_class(window):
        db_manager = DB()
        now_class = db_manager.which_class_to_list()
        class_list = db_manager.get_class_list()

        window['_classtable_'].Update(class_list)

        if now_class is not False:
            window['_classtable_'].Update(class_list)

            # logging.info(window['_classtable_'])
            window['_classtable_'].Update(row_colors=[[class_list.index(now_class), 'green']])
        time.sleep(10)
        # for record in record_list:
        #     logging.info(record)
        #     # go_to_class(link=str(record[2]), filename=record[0], duration=record[1])
        #     window['_table_'].Update(row_colors=[[record_list.index(record), 'green']])
        #     # time.sleep(3)

    def thread(window):
        time.sleep(5)
        greenify_selected_class(window)
        # db_manager = DB()
        # logging.info(db_manager.which_class_to_list())

        # logging.info("salam")
        # time.sleep(10)
        # logging.info("hala")
        # time.sleep(10)
        # logging.info("ok")

    def go_to_class(window):

        db_manager = DB()
        db_manager.read_crd()
        my_which_class = db_manager.which_class()
        greenify_selected_class(window)
        if my_which_class == False:
            logging.info("hich kelasi nist!")
            window['status'].Update("Waiting For class " + jdatetime.datetime.now().strftime("%m/%d %H:%M:%S"))

            return
        selected_class_link = my_which_class['link']
        # if

        username = db_manager.read_crd()['username']
        password = db_manager.read_crd()['password']
        # logging.info(username,password)

        s = Service(resource_path('./driver/chromedriver.exe'))
        driver = webdriver.Chrome(service=s)

        # driver = webdriver.Chrome("chromedriver.exe")
        driver.get("https://vc4011.kntu.ac.ir/login/index.php")
        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys(username)
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys(password)

        login_button = driver.find_element(By.ID, "loginbtn")
        login_button.click()

        driver.get(selected_class_link)
        # driver.get("https://vc4002.kntu.ac.ir/mod/adobeconnect/view.php?id=3393")
        # fanavari
        # driver.get("https://vc4002.kntu.ac.ir/mod/adobeconnect/view.php?id=2931")
        class_link = driver.find_element(By.XPATH,
                                         "/html/body/div[1]/div[2]/div/div[3]/div/section/div/div[1]/form/div[3]/div/input")
        logging.info(class_link)

        adobe_connect_link = class_link.get_attribute('onclick')
        logging.info(adobe_connect_link)
        logging.info(get_url(adobe_connect_link))
        driver.get(get_url(adobe_connect_link)[0])

        # keyboard = Controller()
        # keyboard.press(Key.left)
        # keyboard.release(Key.left)
        # keyboard.press(Key.enter)
        # keyboard.release(Key.enter)

        # time.sleep(5)
        keyboard = Controller()
        keyboard.press(Key.left)
        keyboard.release(Key.left)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)

        time.sleep(5)
        keyboard.press(Key.esc)
        keyboard.release(Key.esc)
        # time.sleep(15)
        time.sleep(3)
        logging.info("hot keys for start recording ")

        db_manager.add_recording_class_to_db()

        keyboard.press(Key.f8)
        keyboard.release(Key.f8)
        time.sleep(2)
        window['status'].Update("Recording class " + jdatetime.datetime.now().strftime("%m/%d %H:%M:%S"))
        with keyboard.pressed(Key.ctrl):
            keyboard.press("\\")
            keyboard.release('\\')

        while True:
            logging.info(my_which_class['name'])
            if db_manager.is_class_finished(my_which_class['name']):
                logging.info("hot keys for stop recording ")
                # keyboard.keyDown('shift')

                keyboard.press(Key.f8)
                time.sleep(1)
                keyboard.release(Key.f8)
                logging.info("close adobe connect!")
                force_close_adobe_connect()
                logging.info("close chrome")

                driver.close()

                logging.info("kelas tamoom shod")
                file_name=change_video_name(my_which_class['name'])
                db_manager.change_status_to_finished(my_which_class['name'],file_name)
                with keyboard.pressed(Key.ctrl):
                    keyboard.press("\\")
                    keyboard.release('\\')
                # time.sleep(10)
                return
                # break
            else:
                logging.info("wait for finishing class ...")
                window['status'].Update(
                    "Wait For Finishing class " + jdatetime.datetime.now().strftime("%m/%d %H:%M:%S"))

                time.sleep(10)

                # return driver

    def start_recording(window):
        logging.info('Start Recording')

        window['status'].Update("Running")
        while True:
            go_to_class(window)
            # record_list = db_manager.records_to_list()

            window['_table_'].Update(values=db_manager.records_to_list())
            time.sleep(30)

    threading.Thread(target=start_recording, args=(window,), daemon=True).start()

    while True:

        logging.info("1")

        event, values = window.read()

        # See if user wants to quit or window was closed
        if event == sg.WINDOW_CLOSED or event == 'Quit':
            break
        # Output a message to the window

        if event == "Add Class":
            logging.info(values['LINK'])
            link = values['LINK']
            file_name = values['FILE_NAME']
            duration = int(values['HOUR']) * 3600 + int(values['MINUTE']) * 60
            db_manager.db.insert(
                {"uuid": jdatetime.datetime.now().strftime("%Y%m%d%H%M%S"), "link": link, "file_name": file_name,
                 "duration": duration})

            logging.info(link)
            logging.info(file_name)
            logging.info(duration)
            logging.info("stop kon")
            # record_list.append([file_name, duration, link])
            record_list = db_manager.records_to_list()

            window['_table_'].Update(values=record_list)

        if event == "Record Class":
            window.minimize()
        if event == "Stop Class":
            db_manager.update_db()
            db_manager.change_status_to_finish_it()
            # logging.info("stop class")

            # db_manager.db.insert(
            #     {"uuid": jdatetime.datetime.now().strftime("%Y%m%d%H%M%S"), "link": link, "file_name": file_name,"class"
            #      "duration": duration,"status":"finished"})
            # todo:close adobe connect
            # add class to db to not record again

            # threading.Thread(target=start_recording_classes, args=(record_list, window,), daemon=True).start()

        if event == "_table_":
            data_selected = [record_list[row] for row in values[event]]
            print(data_selected[0][0])
            selected_class_uuid=data_selected[0][0]
            # print()
            class_record=db_manager.get_class_record_by_uuid(selected_class_uuid)
            # if class_record:
            #     class_record=[]
            if class_record is not None and os.path.exists(class_record['file_name']):
                print("vojood dare")
                file_name=class_record['file_name']
                # subprocess.Popen(r'explorer /select,"C:\path\of\folder\file"')
                subprocess.Popen(f'explorer /select,"{file_name}"')
            # logging.info(2/0)
            # logging.info(data_selected)

    # Finish up by removing from the screen
    window.close()


UI()
