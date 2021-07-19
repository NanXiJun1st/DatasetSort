import PySimpleGUI as sg
from openpyxl import Workbook
import os
import numpy as np

from openpyxl.reader.excel import load_workbook

sg.theme('DarkAmber')
def ChooseFolder(*args):
    layout = [[sg.Text('IWSDataset')], 
              [sg.Text('根目录',size=(15,1),key='-Folder-'),sg.InputText(),sg.FolderBrowse()],
              [sg.Text('选择文件',size=(15,1),key='-file-'),sg.InputText(),sg.FilesBrowse()],
              [sg.Combo(['脑膜瘤','胶质瘤'],size=(10,1))],
            #   [sg.popup_scrolled('Log',size=(80,None),key='-log-')],
              [sg.Submit(),sg.Cancel()]
             ]

    window = sg.Window('IWS DATASET',layout,location=(0,0))

    while True:
        event,values=window.read()
        if event in (sg.WIN_CLOSED,'Cancel'):
            break
        # print(f'You choose filenames{values[0]} and {values[1]}')
        return values[1]
    window.close()

def ReadXlsxFile(*args):
    wb = Workbook()
    wb = load_workbook("./数据收集对应关系总表_20210628.xlsx")
    sheet = wb.active
    print(sheet["A1"].value)


ChooseFolder()
ReadXlsxFile()