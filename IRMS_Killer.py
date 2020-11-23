import csv
import urllib.parse
import requests
from lxml import etree
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from openpyxl import load_workbook

def File_Process():
    global List_7013
    List_7013 = []
    with open('豪德置业.csv') as file_csv:
        reader_obj = csv.reader(file_csv)
        List_7013 = list(reader_obj)

    for row_data in range(len(List_7013)):
        del List_7013[row_data][0]
        del List_7013[row_data][2:4]
        del List_7013[row_data][3]
        del List_7013[row_data][4:6]
        del List_7013[row_data][6]
        del List_7013[row_data][8:12]
        del List_7013[row_data][9]
        del List_7013[row_data][11:39]

    WB_obj = load_workbook('豪德置业.xlsx')
    WS_obj = WB_obj['BOX_Topology']

    global List_Box_Name
    List_Box_Name = []
    for row_num in range(1,201):
        for column_num in range(1,201):
            v = WS_obj.cell(row_num, column_num)
            if v.value == 'empty':
                continue
            if v.value == None:
                break
                # continue
            List_Box_Name.append(list([v.value, row_num, column_num]))

    WS_obj = WB_obj['Info']
    global Longitude_Start
    global Latitude_Start
    global Horizontal_Density
    global Vertical_Density
    Longitude_Start = WS_obj['B1'].value
    Latitude_Start = WS_obj['B2'].value
    Horizontal_Density = WS_obj['B3'].value
    Vertical_Density = WS_obj['B4'].value

    global List_Box_Type
    List_Box_Type = []
    cell_range = WS_obj['H2': 'J3']
    for row_data in cell_range:
        List_Temp_1 = []
        for cell in row_data:
            List_Temp_1.append(cell.value)
        List_Box_Type.append(List_Temp_1)

    WS_obj = WB_obj['Template']
    global List_Template
    List_Template = []
    cell_range = WS_obj['A2': 'R11']
    for row_data in cell_range:
        List_Temp_1 = []
        for cell in row_data:
            List_Temp_1.append(cell.value)
        List_Template.append(List_Temp_1)

def Filling_Box_Info():
    global List_Box_Info
    List_Box_Info = []
    for box_num in range(len(List_Box_Name)):
        List_Box_Info.append(dict({'Box_Name':List_Box_Name[box_num][0]}))
        List_Box_Info[box_num]['Longitude'] = Longitude_Start + (List_Box_Name[box_num][2] - 1) * Horizontal_Density
        List_Box_Info[box_num]['latitude'] = Latitude_Start - (List_Box_Name[box_num][1] - 1) * Vertical_Density
        if (List_Box_Info[box_num]['Box_Name'].find('GJ') != -1) or (List_Box_Info[box_num]['Box_Name'].find('gj') != -1):
            List_Box_Info[box_num]['Box_Type'] = 'guangjiaojiexiang'
        else:
            List_Box_Info[box_num]['Box_Type'] = 'guangfenxianxiang'

if __name__ == '__main__':
    File_Process()
    Filling_Box_Info()
    # print(List_Box_Name)
    # print(List_Box_Info)
    print(List_7013[0])
