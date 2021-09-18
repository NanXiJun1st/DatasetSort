import os
from os.path import join, getsize
from tkinter.constants import DISABLED
import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import WINDOW_CLOSED
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook

sg.theme('DarkAmber')

def GUI(*args):
    layout = [
        [sg.Text('需计算文件夹:',size=(15,1)),sg.InputText(key='-FolderPath-',enable_events=True),sg.FolderBrowse(initial_folder='C:/Users/VertecxGd/Desktop/')],
        [sg.Text('数据统计表:',size=(15,1)),sg.InputText(key='-FilePath-',enable_events=True),sg.FilesBrowse(file_types = (('Excel Files','*.xlsx'),),initial_folder='C:/Users/VertecxGd/Desktop/')],
        [sg.Submit(),sg.Cancel()]
        ]

    window = sg.Window('GetFolderSize',layout,location=(0,0))

    while True:
        event, values = window.read()
        if event in (WINDOW_CLOSED,'Cancel'):
            break
        if event == '-FolderPath-':
            for dir in os.listdir(values['-FolderPath-']):
                sg.Print(dir)
        if event == 'Submit':
            GetFolderSize(values['-FolderPath-'],values['-FilePath-'])
    window.close()

def GetFolderSize(dir,filePath):
    if len(filePath)>0 & os.path.exists(filePath):
        wb = load_workbook(filePath)
        sheet = wb.active

#根据Excel表中记录的文件夹名称数据与选中的路径数据获取文件夹路径
    num = 1
    for values in sheet['A']:
        num += 1
        sg.Print(sheet['A'+str(num)].value)
        if sheet['A'+str(num)].value:
            dirList = dir + '/' + sheet['A'+str(num)].value

            #计算文件夹大小并写入Excel表对应的列中
            size = 0
            for root, dirs, files in os.walk(dirList):
                size += sum([getsize(join(root, name)) for name in files])
            sheet['M' + str(num)].value = size/1024/1024

    wb.save(filePath)


GUI()