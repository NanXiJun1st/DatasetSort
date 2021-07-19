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
              [sg.Text('病种类型',size=(15,1)),sg.Combo(['脑膜瘤','胶质瘤'],default_value = '脑膜瘤', size=(10,1),key='-DiseazeName-')],
              [sg.Text('输出信息',size=(15,1)),sg.Combo(['数量','所在行','文件夹名称'],default_value='数量',size=(10,1),key='-LogData-')],
              [sg.Submit(),sg.Cancel()],
              [sg.Text('Log')]
             ]

    window = sg.Window('IWS DATASET',layout,location=(0,0))

    while True:
        event,values=window.read()
        if event in (sg.WIN_CLOSED,'Cancel'):
            break
        # print(f'You choose filenames{values[0]} and {values[1]}')
        ReadXlsxFile(values[1],values['-DiseazeName-'],values['-LogData-'])
    window.close()

def ReadXlsxFile(path,diseazeName,logData):
    wb = Workbook()
    wb = load_workbook(path)
    sheet = wb.active
    diseazeNameList = sheet['G']
    num = 0
    for i in diseazeNameList:
        if i.value == diseazeName:
            num+=1
            if logData == '数量':
                sg.Print(num)
            if logData =='所在行':
                sg.Print(i.row)
            if logData == '文件夹名称':
                sg.Print(sheet['A'+str(i.row)].value)

ChooseFolder()