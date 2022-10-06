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

import json
from pynput.keyboard import Key, Controller

import glob
import os

import os

import jdatetime

def start_recording_classes(record_list, window):

    for record in record_list:
        window['_table_'].Update(row_colors=[[record_list.index(record), 'green']])
        time.sleep(3)
    # print('Starting thread - will sleep for {} seconds'.format(seconds))
    # time.sleep(seconds)                  # sleep for a while
    # window.write_event_value('-THREAD-', '** DONE **')  # put a message into queue for GUI
def UI():
    data_headings=['file_name','duration','link']
    layout = [
        [sg.Text("link video", key="course_name")],
        [sg.Input(key='LINK')],
        [sg.Text("file name", key="course_name")],
        [sg.Input(key='FILE_NAME')],
        [sg.Text("duration in Hour and Minutes", key="course_name")],
        [sg.Text("H",), sg.Input(key='HOUR', size=(20, 20),default_text=0), sg.Text("M"), sg.Input(key='MINUTE', default_text=0,size=(20, 20))],

        # [sg.Text(size=(40,1), key='-OUTPUT-')],
                 [sg.Table(values=[], headings=data_headings,
                           max_col_width=65,
                           col_widths=[5, 8, 35],
                           auto_size_columns=True,
                           justification='left',
                           enable_events=True,
                           num_rows=6, key='_table_',
                           )

                  ],
        [sg.Button('Add Class'),sg.Button('Record Class')],
        [sg.Button('Quit')]
    ]

    # Create the window
    window = sg.Window('kntu recorder', layout, finalize=True)

    # Display and interact with the Window using an Event Loop
    record_list=[]
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

            print(link)
            print(file_name)
            print(duration)
            print("stop kon")
            record_list.append([file_name, duration, link])
            window['_table_'].Update(values=record_list)

        if event == "Record Class":
            window.minimize()

            threading.Thread(target=start_recording_classes, args=(record_list, window,), daemon=True).start()




    # Finish up by removing from the screen
    window.close()


UI()