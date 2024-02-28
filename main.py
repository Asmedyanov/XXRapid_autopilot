# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings
import pyautogui
import time
import wmi
import pywinauto
from pywinauto.application import Application
import pandas as pd
from tkinter import filedialog

import os

# Press the green button in the gutter to run the script.1
if __name__ == '__main__':
    # os.chdir('D:/programs/4spec')
    file_name = filedialog.askopenfilename(initialdir='D:/Butterfly_MAGEN_2024')
    camera_data = pd.read_excel(file_name, index_col=0)

    my_wmi = wmi.WMI()
    my_id = 0
    for process in my_wmi.Win32_Process():
        if process.Name == '4spec.exe':
            forspec = Application().connect(process=process.ProcessId)
            break

    camera_control = forspec['USB Camera Control   Version 2.30,    KHSDLL v.0495']

    for i in camera_data.index:

        camera_control['Start OptionComboBox'].select(f'{i}')  # select camera
        camera_control.Multiple.click()
        camera_control['-Trig'].click()
        camera_control.UnitEdit.set_text(f'{camera_data["unit[ns]"][i]}')
        camera_control.SequenceEdit.set_text(
            f'{camera_data["t1[ns]"][i]},{camera_data["dt1[ns]"][i]},{camera_data["t2[ns]"][i]},{camera_data["dt2[ns]"][i]}')
        camera_control['MCP Gain VoltageEdit'].set_text(f'{camera_data["MCP[V]"][i]}')  # MCP
        camera_control['FixedEdit'].set_text(f'{camera_data["CCD[dB]"][i]}')  # CCD
        try:
            camera_control['Send It'].click()
        except:
            pass

    '''camera_control['Start OptionComboBox'].select('2')  # select camera
    camera_control.Multiple.click()
    camera_control['-Trig'].click()
    camera_control.UnitEdit.set_text('1')
    camera_control.SequenceEdit.set_text('400,10,800,12')
    camera_control['MCP Gain VoltageEdit'].set_text('700')  # MCP
    camera_control['FixedEdit'].set_text('6')  # CCD
    try:
        camera_control['Send It'].click()
    except:
        pass

    camera_control['Start OptionComboBox'].select('3')  # select camera
    camera_control.Multiple.click()
    camera_control['-Trig'].click()
    camera_control.UnitEdit.set_text('1')
    camera_control.SequenceEdit.set_text('600,10,800,12')
    camera_control['MCP Gain VoltageEdit'].set_text('700')  # MCP
    camera_control['FixedEdit'].set_text('6')  # CCD
    try:
        camera_control['Send It'].click()
    except:
        pass

    camera_control['Start OptionComboBox'].select('4')  # select camera
    camera_control.Multiple.click()
    camera_control['-Trig'].click()
    camera_control.UnitEdit.set_text('1')
    camera_control.SequenceEdit.set_text('800,10,800,12')
    camera_control['MCP Gain VoltageEdit'].set_text('700')  # MCP
    camera_control['FixedEdit'].set_text('6')  # CCD
    try:
        camera_control['Send It'].click()
    except:
        pass'''
