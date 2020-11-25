import csv
from urllib import parse
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
        for row_data in range(len(List_Box_Type)):
            if List_Box_Info[box_num]['Box_Type'] == List_Box_Type[row_data][0]:
                List_Box_Info[box_num]['Box_Type_ID'] = List_Box_Type[row_data][1]
                List_Box_Info[box_num]['ResPoint_Type_ID'] = List_Box_Type[row_data][2]
        for row_data in range(len(List_Template)):
            if List_7013[1][1] == List_Template[row_data][2]:
                List_Box_Info[box_num]['Box_City_ID'] = List_Template[row_data][1]
                List_Box_Info[box_num]['Box_County_ID'] = List_Template[row_data][3]
    Swimming_Pool(Query_Box_ID_ResPont_ID_Alias, List_Box_Info)

def Query_Box_ID_ResPont_ID_Alias(Para_List_Box_Info):
    URL_Box_Query = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Para_List_Box_Info['Box_Type']+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_List_Box_Info['Box_Name']+'%\'" returnfields="INT_ID,ZH_LABEL,STRUCTURE_ID,ALIAS"/></request>'
    Form_Info_Encoded = "xml="+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Respond_Body = requests.post(URL_Box_Query, data=Form_Info_Encoded, headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = etree.HTML(Respond_Body)
    List_Response_Key = Respond_Body.xpath("//fv/@k")
    List_Response_Value = Respond_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    for box_num in range(len(List_Box_Info)):
        if Dic_Response['ZH_LABEL'] == List_Box_Info[box_num]['Box_Name']:
            List_Box_Info[box_num]['Box_ID'] = Dic_Response['INT_ID']
            List_Box_Info[box_num]['ResPoint_ID'] = Dic_Response['STRUCTURE_ID']
            List_Box_Info[box_num]['Alias'] = Dic_Response['ALIAS']

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Push_Box(Para_List_Box_Info):
    URL_Push_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/SynchroController/synchroData?sid='+Para_List_Box_Info['Box_ID']+'&sType='+Para_List_Box_Info['Box_Type']+'&longi='+str(Para_List_Box_Info['Longitude'])+'&lati='+str(Para_List_Box_Info['latitude'])
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded"}
    Respond_Body = requests.post(URL_Push_Box, headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = etree.HTML(Respond_Body)
    Push_Result = Respond_Body.xpath('//message/text()')
    print('P1-{}-{}'.format(Para_List_Box_Info['Box_Name'], Push_Result[0]))

if __name__ == '__main__':
    File_Process()
    Filling_Box_Info()
    Swimming_Pool(Push_Box, List_Box_Info)


