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

import json
from pynput.keyboard import Key, Controller

import glob
import os

import os

import jdatetime


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


def go_to_class():
    my_which_class = which_class()

    if my_which_class == False:
        print("hich kelasi nist!")
        return
    selected_class_link = my_which_class['link']
    # if

    username = get_crd()['username']
    password = get_crd()['password']
    # print(username,password)

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
    print(class_link)

    adobe_connect_link = class_link.get_attribute('onclick')
    print(adobe_connect_link)
    print(get_url(adobe_connect_link))
    driver.get(get_url(adobe_connect_link)[0])

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
    keyboard.press(Key.esc)
    keyboard.release(Key.esc)
    time.sleep(15)
    print("hot keys for start recording ")
    keyboard.press(Key.f8)
    keyboard.release(Key.f8)
    time.sleep(5)
    with keyboard.pressed(Key.ctrl):
        keyboard.press("\\")
        keyboard.release('\\')

    while True:
        if is_class_finished():
            print("hot keys for stop recording ")

            keyboard.press(Key.f8)
            keyboard.release(Key.f8)
            print("close adobe connect!")
            force_close_adobe_connect()
            print("close chrome")

            driver.close()

            print("kelas tamoom shod")
            change_video_name(my_which_class['name'])
            with keyboard.pressed(Key.ctrl):
                keyboard.press("\\")
                keyboard.release('\\')
            return
            # break
        else:
            print("wait for finishing class ...")
            time.sleep(10)

            # return driver


while True:
    # driver=go_to_class()
    print("check for class ...")
    go_to_class()
    print("wait for class ...")
    time.sleep(60)
