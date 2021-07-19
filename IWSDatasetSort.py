from typing import Text
import PySimpleGUI as sg
from openpyxl import Workbook
import os
import numpy as np
import re

from openpyxl.reader.excel import load_workbook

sg.theme('DarkAmber')
def ChooseFolder(*args):
    frame_layout = [ [sg.Listbox(values=['胶质瘤85例已标注','脑膜瘤103例已标注','脑膜瘤110例未标注'],size=(30,7))]]
    layout = [[sg.Text('IWSDataset')], 
              [sg.Text('根目录',size=(15,1),key='-Folder-'),sg.InputText(),sg.FolderBrowse()],
              [sg.Text('选择文件',size=(15,1),key='-file-'),sg.InputText(),sg.FilesBrowse()],
              [sg.Text('所有病例',size=(15,1)),sg.Combo(['未标注','已标注'],default_value='未标注', size=(10,1),key=('-UntaggedDiseazeName-'))],
              [sg.Text('病种类型',size=(15,1)),sg.Combo(['脑膜瘤','胶质'],default_value = '脑膜瘤', size=(10,1),key='-TaggedDiseazeName-'),
              sg.Text('因存在“胶质细胞瘤”的命名,所以检索时以“胶质”作为检索项')],
              [sg.Text('输出信息',size=(15,1)),sg.Combo(['数量','所在行','文件夹名称'],default_value='数量',size=(10,1),key='-LogData-')],
              [sg.Frame('已提交数据',frame_layout,font='Any 12',title_color='blue')],
              [sg.Submit(),sg.Cancel()]
             ]

    

    window = sg.Window('IWS DATASET',layout,location=(0,0))

    while True:
        event,values=window.read()
        if event in (sg.WIN_CLOSED,'Cancel'):
            break
        # print(f'You choose filenames{values[0]} and {values[1]}')
        ReadXlsxFile(values[1],values['-UntaggedDiseazeName-'], values['-TaggedDiseazeName-'],values['-LogData-'])
    window.close()

def ReadXlsxFile(path,UntaggedDiseazeName,TaggedDiseazeName,logData):
    wb = Workbook()
    wb = load_workbook(path)
    sheet = wb.active
    if UntaggedDiseazeName == '未标注':
        diseazeNameList = sheet['E']
    else:
        diseazeNameList = sheet['G']

    num = 0
    for i in diseazeNameList:
        if str(i.value).find(TaggedDiseazeName) >= 0 :
            num+=1
            if logData == '数量':
                sg.Print(num)
            if logData =='所在行':
                sg.Print(i.row)
            if logData == '文件夹名称':
                sg.Print(sheet['A'+str(i.row)].value)

ChooseFolder()