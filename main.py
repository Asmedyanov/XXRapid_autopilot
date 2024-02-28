# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings
import pyautogui
import time
import wmi
import pywinauto
from pywinauto.application import Application
import os

# Press the green button in the gutter to run the script.1
if __name__ == '__main__':
    # os.chdir('D:/programs/4spec')
    my_wmi = wmi.WMI()
    my_id = 0
    for process in my_wmi.Win32_Process():
        if process.Name == '4spec.exe':
            forspec = Application().connect(process=process.ProcessId)
            break

    # print(forspec.winows())



    camera_control = forspec['USB Camera Control   Version 2.30,    KHSDLL v.0495']

    camera_control['Start OptionComboBox'].select('1')  # select camera
    camera_control.Multiple.click()
    camera_control['-Trig'].click()
    camera_control.UnitEdit.set_text('1')
    camera_control.SequenceEdit.set_text('200,10,800,12')
    camera_control['MCP Gain VoltageEdit'].set_text('700')  # MCP
    camera_control['FixedEdit'].set_text('6')  # CCD
    try:
        camera_control['Send It'].click()
    except:
        pass
    #camera_control.print_control_identifiers()

    camera_control['Start OptionComboBox'].select('2') #select camera
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
        pass


    # my_wmi = wmi.WMI()

    # pyautogui.PAUSE = 10
    '''pyautogui.hotkey('win', '1')
    start_time = time.time()
    current_time = start_time
    process_name_list = []
    for process in my_wmi.Win32_Process():
        process_name_list.append(process.Name)
    while not ('4spec.exe' in process_name_list):
        process_name_list = []
        for process in my_wmi.Win32_Process():
            process_name_list.append(process.Name)
    print('4spec is open')
    pyautogui.hotkey('win', '1')'''

    # time.sleep(20)
    # pyautogui.keyDown('alt')
    # pyautogui.press('tab')
    # time.sleep(20)
    # pyautogui.press('left')
    # pyautogui.keyUp('alt')
    # pyautogui.hotkey('enter')
    # for i in range(10):
    # pyautogui.hotkey('tab')

    # pyautogui.hotkey('alt','tab')
    # pyautogui.hotkey('alt','tab')
    # time.sleep(10)

    # pyautogui.hotkey('alt', 'tab')
    # pyautogui.hotkey('alt','tab')
    # print(f'Hi, {2}')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
