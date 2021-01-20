import csv
import math
import requests
import copy
import json
import datetime
from urllib import parse
from lxml import etree
from concurrent.futures import ThreadPoolExecutor
from openpyxl import load_workbook
from selenium import webdriver
from time import sleep
from collections import Counter

'''
7013CSV是基础,除P1外都需要

一级箱子DL_2FS_Count,可以考虑放弃使用&
光路申请(含工单建立)
提交工单(理想情况)
光路设计
工单号写入(含工单建立)
'''

File_Name = ['豪德置业']

P1_Push_Box                    = False
P2_Generate_Support_Segment    = False
P3_Generate_Cable_Segment      = False
P4_Cable_Lay                   = False
P5_Generate_ODM_and_Tray       = False
P6_Termination_and_Direct_Melt = False
P7_Generate_Optical_Circut     = True

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Generate_Local_Data(Para_File_Name):
    '''读取本地数据,在不查询的前提下填充几个List'''

    '''读取并整理7013表,生成List_7013'''
    global List_7013
    List_7013 = []
    with open(Para_File_Name+'.csv') as file_csv:
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
    Task_Name_ID_List = List_7013[1][5].split('-')

    '''读取并整理Sheet_Template,生成List_Template'''
    WB_obj = load_workbook(Para_File_Name+'.xlsx')
    WS_obj = WB_obj['Template']
    List_Template = []
    cell_range = WS_obj['A2': 'S11']
    for row_data in cell_range:
        List_Temp_1 = []
        for cell in row_data:
            List_Temp_1.append(cell.value)
        List_Template.append(List_Temp_1)
    for list_num in List_Template:
        if List_7013[1][1] == list_num[0]:
            List_Template_Selected = copy.deepcopy(list_num)

    '''读取并整理Sheet_Info,生成List_Box_Type和其他参数'''
    WS_obj = WB_obj['Info']
    Longitude_Start = WS_obj['B1'].value
    Latitude_Start = WS_obj['B2'].value
    Horizontal_Density = WS_obj['B3'].value
    Vertical_Density = WS_obj['B4'].value
    List_Box_Type = []
    cell_range = WS_obj['H2': 'J3']
    for row_data in cell_range:
        List_Temp_1 = []
        for cell in row_data:
            List_Temp_1.append(cell.value)
        List_Box_Type.append(List_Temp_1)

    '''读取并整理Sheet_Box_Topology,生成List_Box_Data'''
    WS_obj = WB_obj['BOX_Topology']
    List_Box_Name = []
    for row_num in range(1,201):
        for column_num in range(1,201):
            v = WS_obj.cell(row_num, column_num)
            if v.value == 'empty':
                continue
            if v.value == None:
                break
            List_Box_Name.append(list([v.value, row_num, column_num]))
    global List_Box_Data
    List_Box_Data = []
    for box_num in range(len(List_Box_Name)):
        List_Box_Data.append(dict({'Box_Name':List_Box_Name[box_num][0]}))
        List_Box_Data[box_num]['Longitude'] = Longitude_Start + (List_Box_Name[box_num][2] - 1) * Horizontal_Density
        List_Box_Data[box_num]['Latitude'] = Latitude_Start - (List_Box_Name[box_num][1] - 1) * Vertical_Density
        if (List_Box_Data[box_num]['Box_Name'].find('GJ') != -1) or (List_Box_Data[box_num]['Box_Name'].find('gj') != -1):
            List_Box_Data[box_num]['Box_Type'] = 'guangjiaojiexiang'
        else:
            List_Box_Data[box_num]['Box_Type'] = 'guangfenxianxiang'
        for row_data in range(len(List_Box_Type)):
            if List_Box_Data[box_num]['Box_Type'] == List_Box_Type[row_data][0]:
                List_Box_Data[box_num]['Box_Type_ID'] = List_Box_Type[row_data][1]
                List_Box_Data[box_num]['ResPoint_Type_ID'] = List_Box_Type[row_data][2]
        List_Box_Data[box_num]['City_ID'] = List_Template_Selected[16]
        List_Box_Data[box_num]['County_ID'] = List_Template_Selected[1]

    '''读取并整理Sheet_OCS_List,生成List_CS_Data'''
    WS_obj = WB_obj['OCS_List']
    global List_CS_Data
    List_CS_Data = []
    for row_num in range(1,501):
        OCS_A_Box_Name = WS_obj.cell(row_num, 1).value
        OCS_Z_Box_Name = WS_obj.cell(row_num, 2).value
        OCS_Width = WS_obj.cell(row_num, 3).value
        if OCS_A_Box_Name == None:
            break
        List_CS_Data.append(dict({'A_Box_Name': OCS_A_Box_Name}))
        List_CS_Data[row_num - 1]['Z_Box_Name'] = OCS_Z_Box_Name
        List_CS_Data[row_num - 1]['Width'] = OCS_Width
    
    #Length_Prepare
    Horizontal_Metre = 111.11 * 1000 * math.cos(Latitude_Start * math.pi / 180)
    Vertical_Metre = 111.11 * 1000
    
    for dic_num_in_osc in List_CS_Data:
        for dic_num_in_box in List_Box_Data:
            if dic_num_in_osc['A_Box_Name'] == dic_num_in_box['Box_Name']:
                dic_num_in_osc['A_Box_Type_ID'] = dic_num_in_box['Box_Type_ID']
                dic_num_in_osc['A_Box_Type'] = dic_num_in_box['Box_Type']
                dic_num_in_osc['A_ResPoint_Type_ID'] = dic_num_in_box['ResPoint_Type_ID']
                dic_num_in_osc['A_Longitude'] = dic_num_in_box['Longitude']
                dic_num_in_osc['A_Latitude'] = dic_num_in_box['Latitude']
            if dic_num_in_osc['Z_Box_Name'] == dic_num_in_box['Box_Name']:
                dic_num_in_osc['Z_Box_Type_ID'] = dic_num_in_box['Box_Type_ID']
                dic_num_in_osc['Z_Box_Type'] = dic_num_in_box['Box_Type']
                dic_num_in_osc['Z_ResPoint_Type_ID'] = dic_num_in_box['ResPoint_Type_ID']
                dic_num_in_osc['Z_Longitude'] = dic_num_in_box['Longitude']
                dic_num_in_osc['Z_Latitude'] = dic_num_in_box['Latitude']
        dic_num_in_osc['Length'] = round(math.sqrt(((dic_num_in_osc['A_Longitude'] - dic_num_in_osc['Z_Longitude']) * Horizontal_Metre) ** 2 + ((dic_num_in_osc['A_Latitude'] - dic_num_in_osc['Z_Latitude'])* Vertical_Metre) ** 2))
        dic_num_in_osc['Business_Level'] = 8
        dic_num_in_osc['Life_Cycle'] = 8
        dic_num_in_osc['Owner_Type'] = 1
        dic_num_in_osc['Owner_Name'] = 0
        dic_num_in_osc['Field_Type'] = '市城区域'
        dic_num_in_osc['City_ID'] = List_Template_Selected[16]
        dic_num_in_osc['County_ID'] = List_Template_Selected[1]
        dic_num_in_osc['DQS_Project_ID'] = List_Template_Selected[3]
        dic_num_in_osc['DQS_ID'] = List_Template_Selected[5]
        dic_num_in_osc['DQS_County_ID'] = List_Template_Selected[7]
        dic_num_in_osc['DQS_Maintainer_ID'] = List_Template_Selected[9]
        dic_num_in_osc['Task_Name_ID'] = Task_Name_ID_List[0]

    '''生成List_OC_Data'''
    global List_OC_Data
    List_OC_Data = []
    for each_7013_line in List_7013: 
        if each_7013_line[10] == '二级':
            List_OC_Data.append(dict({
                'Z_POS_Name': each_7013_line[11],
                'Z_Box_Name':  each_7013_line[3],
                'A_POS_Name': each_7013_line[8],
                'City_ID': List_Template_Selected[16],
                'County_ID': List_Template_Selected[1],
                'Business_Name': Task_Name_ID_List[1],
                'Project_Code': each_7013_line[4],
                'Task_Name': each_7013_line[5],
                'Task_Name_ID': Task_Name_ID_List[0],
                'DQS_Project': List_Template_Selected[11],
                'DQS_Project_ID': List_Template_Selected[3],
                'DQS': List_Template_Selected[12],
                'DQS_ID': List_Template_Selected[5],
                'DQS_County': List_Template_Selected[13],
                'DQS_County_ID': List_Template_Selected[7],
                'DQS_Maintainer': List_Template_Selected[8],
                'DQS_Maintainer_ID': List_Template_Selected[9],
                'AEquType': 10002,
                'ZEquType': 10002,
                'SXBussType': '家客业务',
                'AJoinName': '网元成端设备',
                'ZJoinName': '网元成端设备',
                'FiberCount': 1,
                'BussType': 207, # 设备类型 207 PON/402 直联光路
                'AppType': 1005, # 应用类型  1005 业务光路/1006 尾纤光路
                'ServiceLevel': 101, # 101 家客接入/111 驻地网
                }))
    for each_oc_data in List_OC_Data:
        for each_7013_line in List_7013:
            if each_oc_data['A_POS_Name'] == each_7013_line[11]: # 中文名称 each_7013_line[2] # 资管中文名称
                each_oc_data['A_Box_Name'] = each_7013_line[3]
        if each_oc_data['A_Box_Name'] == each_oc_data['Z_Box_Name']:
            each_oc_data['BussType'] = 402
            each_oc_data['AppType'] = 1006
        for each_box_data in List_Box_Data:
            if each_oc_data['A_Box_Name'] == each_box_data['Box_Name']:
                each_oc_data['A_Box_Type'] = each_box_data['Box_Type']
                each_oc_data['A_Box_Type_ID'] = each_box_data['Box_Type_ID']
                each_oc_data['A_ResPoint_Type_ID'] = each_box_data['ResPoint_Type_ID']
            if each_oc_data['Z_Box_Name'] == each_box_data['Box_Name']:
                each_oc_data['Z_Box_Type'] = each_box_data['Box_Type']
                each_oc_data['Z_Box_Type_ID'] = each_box_data['Box_Type_ID']
                each_oc_data['Z_ResPoint_Type_ID'] = each_box_data['ResPoint_Type_ID']

def Generate_Topology():
    global CS_Topology
    CS_Topology = []
    Line_Num = 0
    Segment_Num = 0
    for cable_seg_num in range(len(List_CS_Data)):
        if cable_seg_num == 0:
            CS_Topology.append([List_CS_Data[cable_seg_num]['A_Box_Name']])
            CS_Topology[cable_seg_num].append(List_CS_Data[cable_seg_num]['Z_Box_Name'])
            Segment_Num += 1
            continue
        if List_CS_Data[cable_seg_num]['A_Box_Name'] == CS_Topology[Line_Num][Segment_Num]:
            CS_Topology[Line_Num].append(List_CS_Data[cable_seg_num]['Z_Box_Name'])
            Segment_Num += 1
            continue
        else:
            CS_Topology.append([List_CS_Data[cable_seg_num]['A_Box_Name']])
            Line_Num += 1
            Segment_Num = 1
            CS_Topology[Line_Num].append(List_CS_Data[cable_seg_num]['Z_Box_Name'])

def Generate_FS_Data():
    #filled by 0
    for box_info in List_Box_Data:
        box_info['1FS_Count'] = 0
        box_info['2FS_Count'] = 0
        box_info['DL_2FS_Count'] = 0
        box_info['ODM_Rows'] = 0
        box_info['Tray_Count'] = 0
    #1fs_count&2fs_count
    for row in List_7013:
        for box_info in List_Box_Data:
            if row[3] == box_info['Box_Name']:
                if row[10] == '一级':
                    box_info['1FS_Count'] += 1
                elif row[10] == '二级':
                    box_info['2FS_Count'] += 1
    #dl_2fs_count
    for cable_num in range(len(CS_Topology) - 1, -1, -1):
        for cable_seg_num in range(len(CS_Topology[cable_num]) - 1, -1, -1):
            if cable_seg_num == len(CS_Topology[cable_num]) - 1:#Tail
                for box_info in List_Box_Data: 
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = 0
            elif cable_seg_num == 0:#Head
                for box_info in List_Box_Data:
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = 0
            else:#Middle
                for box_info in List_Box_Data:
                    if CS_Topology[cable_num][cable_seg_num + 1] == box_info['Box_Name']:
                        DL_2FS_Count_temp = box_info['2FS_Count'] + box_info['DL_2FS_Count']
                for box_info in List_Box_Data:
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = DL_2FS_Count_temp
    #dl_2fs_count 1fs_box_correction
    for box_info in List_Box_Data:
        if box_info['1FS_Count'] != 0:
            DL_2FS_Count_temp = []
            for cable_num in CS_Topology:
                if box_info['Box_Name'] == cable_num[0]:
                    for  box_info_2 in List_Box_Data:
                        if cable_num[1] == box_info_2['Box_Name']:
                            DL_2FS_Count_temp.append(str(box_info_2['DL_2FS_Count'] + box_info_2['2FS_Count']))
                            box_info['DL_2FS_Count'] = '&'.join(DL_2FS_Count_temp)
    #ODM_Rows$Tray_Count
    for box_info in List_Box_Data:
        if box_info['1FS_Count'] == 0:
            for cable_seg_data in List_CS_Data:
                if box_info['Box_Name'] == cable_seg_data['Z_Box_Name']:
                    box_info['ODM_Rows'] = box_info['Tray_Count'] = math.ceil(cable_seg_data['Width'] / 12)
        elif box_info['1FS_Count'] != 0:
            for cable_seg_data in List_CS_Data:
                if box_info['Box_Name'] == cable_seg_data['A_Box_Name']:
                    box_info['ODM_Rows'] = box_info['Tray_Count'] = math.ceil(cable_seg_data['Width'] / 12) + box_info['ODM_Rows']

def Generate_Termination_and_Direct_Melt_Data():
    for box_info in List_Box_Data:
        if box_info['1FS_Count'] == 0:
            for cable_seg_data in List_CS_Data:
                if box_info['Box_Name'] == cable_seg_data['Z_Box_Name']:
                    if cable_seg_data['Width'] >= (box_info['2FS_Count'] * 3 + box_info['DL_2FS_Count'] * 3):
                        box_info['BackUp_Fiber_Count'] = 2
                    elif (cable_seg_data['Width'] < (box_info['2FS_Count'] * 3 + box_info['DL_2FS_Count'] * 3)) and (cable_seg_data['Width'] >= (box_info['2FS_Count'] * 2 + box_info['DL_2FS_Count'] * 2)):
                        box_info['BackUp_Fiber_Count'] = 1
                    else:
                        box_info['BackUp_Fiber_Count'] = 0
            box_info['Termination_Start'] = 1
            box_info['Termination_Count'] = box_info['2FS_Count'] * (box_info['BackUp_Fiber_Count'] + 1)
            box_info['Direct_Melt_Start'] = box_info['2FS_Count'] * (box_info['BackUp_Fiber_Count'] + 1) + 1
            box_info['Direct_Melt_Count'] = box_info['DL_2FS_Count'] * (box_info['BackUp_Fiber_Count'] + 1)
            if box_info['Direct_Melt_Count'] == 0: #尾箱没有直熔数据
                box_info['Direct_Melt_Start'] = 0
        elif box_info['1FS_Count'] != 0:

            List_Width = []
            for cable_seg_data in List_CS_Data:
                if box_info['Box_Name'] == cable_seg_data['A_Box_Name']:
                    List_Width.append(cable_seg_data['Width'])

            List_DL_2FS = box_info['DL_2FS_Count'].split('&')
            box_info['BackUp_Fiber_Count'] = []
            box_info['Termination_Start'] = []
            box_info['Termination_Count'] = []
            box_info['Direct_Melt_Start'] = []
            box_info['Direct_Melt_Count'] = []

            for each_DL_2FS, each_Width in zip(List_DL_2FS, List_Width):
                if each_Width >= (box_info['2FS_Count'] * 3 + int(each_DL_2FS) * 3):
                    box_info['BackUp_Fiber_Count'].append(2)
                elif (each_Width < (box_info['2FS_Count'] * 3 + int(each_DL_2FS) * 3)) and (each_Width >= (box_info['2FS_Count'] * 2 + int(each_DL_2FS) * 2)):
                    box_info['BackUp_Fiber_Count'].append(1)
                else:
                    box_info['BackUp_Fiber_Count'].append(0)
            
            for each_BackUp_Fiber_Count, each_DL_2FS in zip(box_info['BackUp_Fiber_Count'], List_DL_2FS):
                box_info['Termination_Start'].append(1)
                box_info['Termination_Count'].append(int(each_DL_2FS) * (int(each_BackUp_Fiber_Count) + 1))
                box_info['Direct_Melt_Start'] = '0'
                box_info['Direct_Melt_Count'] = '0'
            
            box_info['BackUp_Fiber_Count'] = [str(i) for i in box_info['BackUp_Fiber_Count']]
            box_info['Termination_Start'] = [str(i) for i in box_info['Termination_Start']]
            box_info['Termination_Count'] = [str(i) for i in box_info['Termination_Count']]
            box_info['Direct_Melt_Start'] = [str(i) for i in box_info['Direct_Melt_Start']]
            box_info['Direct_Melt_Count'] = [str(i) for i in box_info['Direct_Melt_Count']]

            box_info['BackUp_Fiber_Count'] = '&'.join(box_info['BackUp_Fiber_Count'])
            box_info['Termination_Start'] = '&'.join(box_info['Termination_Start'])
            box_info['Termination_Count'] = '&'.join(box_info['Termination_Count'])
            box_info['Direct_Melt_Start'] = '&'.join(box_info['Direct_Melt_Start'])
            box_info['Direct_Melt_Count'] = '&'.join(box_info['Direct_Melt_Count'])

def Generate_OC_POS_Data_and_OC_Name():
    List_OC_Name = []
    for each_oc_data in List_OC_Data:
        for each_box_data in List_Box_Data:
            if each_oc_data['A_Box_Name'] == each_box_data['Box_Name']:
                for keykey, valuevalue in each_box_data['POS'].items():
                    if each_oc_data['A_POS_Name'] == keykey:
                        each_oc_data['A_POS_ID'] = valuevalue
            if each_oc_data['Z_Box_Name'] == each_box_data['Box_Name']:
                for keykey, valuevalue in each_box_data['POS'].items():
                    if each_oc_data['Z_POS_Name'] == keykey:
                        each_oc_data['Z_POS_ID'] = valuevalue

        each_oc_data['OC_Name'] = each_oc_data['A_Box_Name'] + '资源点' + '-' + each_oc_data['Z_Box_Name'] + '资源点' + 'F0001'
        List_OC_Name.append(each_oc_data['OC_Name'])

    Counter_OC_Name = Counter()
    for oc_name in List_OC_Name:
        Counter_OC_Name[oc_name] += 1

    for keykey, valuevalue in Counter_OC_Name.items():
        for increase_num in range(valuevalue):
            for each_oc_data in List_OC_Data:
                if each_oc_data['OC_Name'] == keykey:
                    each_oc_data['OC_Name'] = each_oc_data['OC_Name'][:(len(each_oc_data['OC_Name']) - 4)] + '{:04d}'.format(int(each_oc_data['OC_Name'][(len(each_oc_data['OC_Name']) - 4):]) + int(increase_num))
                    break


def Query_Project_Code_ID():
    URL_Query_Project_Code_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/datatrans/expall?model=guangfenxianxiang&fname=PROJECTCODE&p1='+List_7013[1][4]
    Response_Body = requests.get(URL_Query_Project_Code_ID)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//@name")
    List_Response_Value = Response_Body.xpath("//@value")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    Dic_Project_Code_ID = copy.deepcopy(Dic_Response)
    for box_num in List_Box_Data:
        box_num['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]
    for ocs_num in List_CS_Data:
        ocs_num['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]
    for oc_data in List_OC_Data:
        oc_data['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]

def Query_Box_ID_ResPoint_ID_Alias(Para_List_Box_Data):
    URL_Query_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Para_List_Box_Data['Box_Type']+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_List_Box_Data['Box_Name']+'%\'" returnfields="INT_ID,ZH_LABEL,STRUCTURE_ID,ALIAS"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Box, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_Response_Value_tv = Response_Body.xpath("//fv/@tv")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    for box_num in range(len(List_Box_Data)):
        if Dic_Response['ZH_LABEL'] == List_Box_Data[box_num]['Box_Name']:
            List_Box_Data[box_num]['Box_ID'] = Dic_Response['INT_ID']
            List_Box_Data[box_num]['ResPoint_ID'] = Dic_Response['STRUCTURE_ID']
            List_Box_Data[box_num]['Alias'] = Dic_Response['ALIAS']
            List_Box_Data[box_num]['ResPoint_Name'] = List_Response_Value_tv[0]
    for ocs_num in List_CS_Data:
        for box_num in List_Box_Data:
            if ocs_num['A_Box_Name'] == box_num['Box_Name']:
                ocs_num['A_Box_ID'] = box_num['Box_ID']
                ocs_num['A_ResPoint_ID'] = box_num['ResPoint_ID']
        for box_num in List_Box_Data:
            if ocs_num['Z_Box_Name'] == box_num['Box_Name']:
                ocs_num['Z_Box_ID'] = box_num['Box_ID']
                ocs_num['Z_ResPoint_ID'] = box_num['ResPoint_ID']
    for each_oc_data in List_OC_Data:
        for each_box_data in List_Box_Data:
            if each_oc_data['A_Box_Name'] == each_box_data['Box_Name']:
                each_oc_data['A_Box_ID'] = each_box_data['Box_ID']
                each_oc_data['A_ResPoint_ID'] = each_box_data['ResPoint_ID']
                each_oc_data['A_ResPoint_Name'] = each_box_data['ResPoint_Name']
            if each_oc_data['Z_Box_Name'] == each_box_data['Box_Name']:
                each_oc_data['Z_Box_ID'] = each_box_data['Box_ID']
                each_oc_data['Z_ResPoint_ID'] = each_box_data['ResPoint_ID']
                each_oc_data['Z_ResPoint_Name'] = each_box_data['ResPoint_Name']

def Query_Support_Sys_and_Cable_Sys():
    Task_Name_ID_List = List_7013[1][5].split('-')
    Support_Sys_Name = List_7013[1][0]+List_7013[1][1]+Task_Name_ID_List[1]+'区内引上系统'
    Cable_Sys_Name = List_7013[1][0]+List_7013[1][1]+Task_Name_ID_List[1]+'接入层架空光缆'

    #Support_System
    URL_Query_SS_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGetPage'
    Form_Info = '<request><query mc="xitong" where="1=1 AND LINESEG_TYPE=\'9108\' AND ZH_LABEL LIKE \'%'+Support_Sys_Name+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'pageNum=1&'+'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_SS_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_SS_Count = Response_Body.xpath('//@counts')
    if List_SS_Count[0] == '0':
        #Generate Support_Sys
        URL_Add_Support_Sys = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalSave'
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Data[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Data[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = str(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Support_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Support_Sys_ID = Response_Body.xpath('//@newid')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_CS_Data:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('引上系统ID-{}'.format(Support_Sys_ID))

    elif List_SS_Count[0] == '1':
        #Get Support_Sys_ID
        Support_Sys_ID = Response_Body.xpath('//@int_id')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_CS_Data:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('引上系统ID-{}'.format(Support_Sys_ID))

    #Cable_System
    URL_Query_CS_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGetPage'
    Form_Info = '<request><query mc="guanglanxitong" where="1=1 AND ZH_LABEL LIKE \'%'+Cable_Sys_Name+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'pageNum=1&'+'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_CS_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_CS_Count = Response_Body.xpath('//@counts')
    if List_CS_Count[0] == '0':
        #Generate Cable_Sys
        URL_Add_Cable_Sys = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalSave'
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Data[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Data[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info = '<response mode="add"><mc type="guanglanxitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Data[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Data[0]['County_ID'])+'"/><fv k="ZH_LABEL" v="'+Cable_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/><fv k="OPT_TYPE" v="GYTA-12"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = str(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Cable_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Cable_Sys_ID = Response_Body.xpath('//@newid')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_CS_Data:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('光缆系统ID-{}'.format(Cable_Sys_ID))

    elif List_CS_Count[0] == '1':
        #Get Cable_Sys_ID
        Cable_Sys_ID = Response_Body.xpath('//@int_id')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_CS_Data:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('光缆系统ID-{}'.format(Cable_Sys_ID))

def Query_JSESSIONIRMS_and_route():
    Browser_Obj = webdriver.Ie()
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_was/uploadResource/public/login/login.html')
    Browser_Obj.find_element_by_id('username').send_keys('tyyangwei')
    Browser_Obj.find_element_by_id('password').send_keys('tyyw789...')
    Browser_Obj.find_element_by_class_name('lwb_login').click()
    sleep(5)
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_wcm/middelwebpage/app_recoder_log.jsp?app_flg=zhwlzygl_ywzl')
    sleep(5)
    Dic_Cookie_JSESSIONIRMS = Browser_Obj.get_cookie('JSESSIONIRMS')
    Dic_Cookie_route = Browser_Obj.get_cookie('route')
    JSESSIONIRMS_Value = Dic_Cookie_JSESSIONIRMS['value']
    route_Value = Dic_Cookie_route['value']
    Browser_Obj.quit()
    global Jsessionirms_v, route_v
    Jsessionirms_v = JSESSIONIRMS_Value
    route_v = route_Value

def Query_Support_Seg_ID_Cable_Seg_ID(Para_List_CS_Data):
    List_CS_Support_Seg_Name_ID_Cable_Name_ID = {}
    URL_Query_Support_Seg_ID_Cable_Seg_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'

    Form_Info = '<request><query mc="yinshangduan" where="1=1 AND ZH_LABEL LIKE \'%'+str(Para_List_CS_Data['A_Box_Name'])+'资源点-'+str(Para_List_CS_Data['Z_Box_Name'])+'资源点'+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Support_Seg_ID_Cable_Seg_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_Name'] = List_Response_Value[1]
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_ID'] = List_Response_Value[0]

    Form_Info = '<request><query mc="guanglanduan" where="1=1 AND ZH_LABEL LIKE \'%'+str(Para_List_CS_Data['A_Box_Name'])+'资源点-'+str(Para_List_CS_Data['Z_Box_Name'])+'资源点'+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Support_Seg_ID_Cable_Seg_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_Name'] = List_Response_Value[1]
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_ID'] = List_Response_Value[0]
    for cable_seg in List_CS_Data:
        if Para_List_CS_Data['A_Box_Name'] == cable_seg['A_Box_Name']:
            cable_seg['Support_Seg_Name'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_Name']
            cable_seg['Support_Seg_ID'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_ID']
            cable_seg['Cable_Seg_Name'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_Name']
            cable_seg['Cable_Seg_ID'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_ID']

def Query_ODM_ID_and_Terminarl_IDs(Para_List_Box_Data):
    URL_Query_ODM_ID_and_Terminarl_IDs = 'http://10.209.199.74:8120/nxapp/room/queryShelfAndModuleData.ilf'
    if Para_List_Box_Data['Box_Type'] == 'guangfenxianxiang':
        Para_List_Box_Data['Box_Type_Short'] = 'gfxx'
    elif Para_List_Box_Data['Box_Type'] == 'guangjiaojiexiang':
        Para_List_Box_Data['Box_Type_Short'] = 'gjjx'
    Form_Info = '<params><param key="type" value="'+Para_List_Box_Data['Box_Type_Short']+'"/><param key="rack_id" value="'+Para_List_Box_Data['Box_ID']+'"/></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&model=odm&" +  "lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_ODM_ID_and_Terminarl_IDs, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = parse.unquote(bytes(Response_Body.text, encoding="utf-8"))
    Response_Body = etree.HTML(Response_Body)
    List_Terminal_IDs = Response_Body.xpath('//@id')
    List_Terminal_IDs.pop(0)
    Para_List_Box_Data['ODM_ID'] = List_Terminal_IDs.pop(0)
    Para_List_Box_Data['Terminal_IDs'] = '&'.join(List_Terminal_IDs)

def Query_CS_Fiber_IDs(Para_List_CS_Data):
    URL_Query_CS_Fiber_IDs = 'http://10.209.199.74:8120/igisserver_osl/rest/EquipEditModule1/getEquipModuleTerminals'
    Form_Info = '<params><param key="equ_type" value="fiberseg"/><param key="equ_id" value="'+Para_List_CS_Data['Cable_Seg_ID']+'"/></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_CS_Fiber_IDs, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_CS_Fiber_IDs = Response_Body.xpath('//@id')
    List_CS_Fiber_IDs.pop(0)
    Para_List_CS_Data['CS_Fiber_IDs'] = '&'.join(List_CS_Fiber_IDs)

def Query_POS_ID(Para_List_Box_Data):
    URL_Query_POS_ID = 'http://10.209.199.72:7112/irms/opticOpenApplyAction!queryEqu.ilf'
    Form_Info_Encoded = 'equType=POS&countyId=' + str(Para_List_Box_Data['County_ID']) + '&siteId=' + str(Para_List_Box_Data['ResPoint_ID']) + '&sitetype=' + str(Para_List_Box_Data['ResPoint_Type_ID']) + '&sitename=' + parse.quote_plus(Para_List_Box_Data['ResPoint_Name']) + '&cityId=' + str(Para_List_Box_Data['City_ID'])
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host':'10.209.199.72:7112', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_POS_ID, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': Jsessionirms_v, 'route': route_v})
    # Response_Body = requests.post(URL_Query_POS_ID, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': 'Cb15gGLSqLh2fbGDx9j8TVKfphF25YmQnN7h6JVyxnL7lFGg7Nvn!418505528', 'route': '5fb592aa37b5606b0629ebaa738ace15'})
    Response_Body = Response_Body.text
    Response_Body = Response_Body.replace('success:true','"success":true')
    Response_Body = Response_Body.replace('totalProperty','"totalProperty"')
    Response_Body = Response_Body.replace('data','"data"')
    Response_Body = Response_Body.replace('int_id','"int_id"')
    Response_Body = Response_Body.replace('endEquipType','"endEquipType"')
    Response_Body = Response_Body.replace('trans_site_type','"trans_site_type"')
    Response_Body = Response_Body.replace('trans_site_id','"trans_site_id"')
    Response_Body = Response_Body.replace('trans_site_name','"trans_site_name"')
    Response_Body = Response_Body.replace('zh_label','"zh_label"')
    Response_Body = Response_Body.replace('\'','\"')
    Response_Body = json.loads(Response_Body)
    List_POS_Name = []
    List_POS_ID = []
    for each_POS_data in Response_Body['data']:
        List_POS_Name.append(each_POS_data['zh_label'])
        List_POS_ID.append(each_POS_data['int_id'])
    Para_List_Box_Data['POS'] = dict(zip(List_POS_Name, List_POS_ID))

def Query_Work_Sheet_ID():
    Work_Sheet_Count = math.ceil(len(List_OC_Data) / 40)
    for work_sheet_num in range(Work_Sheet_Count):
        Work_Sheet_Name = '关于' + List_OC_Data[0]['Business_Name'] + '二级光路申请yry' + '0' + str(work_sheet_num)
        print(Work_Sheet_Name)
        Work_Sheet_Times = []
        Work_Sheet_Times.append(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        Work_Sheet_Times.append((datetime.datetime.now() + datetime.timedelta(days=2)).strftime('%Y-%m-%d %H:%M:%S'))
        Work_Sheet_Times.append((datetime.datetime.now() + datetime.timedelta(days=4)).strftime('%Y-%m-%d %H:%M:%S'))
        Work_Sheet_Times.append((datetime.datetime.now() + datetime.timedelta(days=6)).strftime('%Y-%m-%d %H:%M:%S'))
        URL_Query_Work_Sheet_Exist = 'http://10.209.199.72:7112/irms/tasklistAction!waitedTaskAJAX.ilf'
        Form_Info_Encoded = 'processinstname=' + parse.quote_plus(Work_Sheet_Name)
        Request_Lenth = str(len(Form_Info_Encoded))
        Request_Header = {'Host':'10.209.199.72:7112', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Query_Work_Sheet_Exist, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': Jsessionirms_v, 'route': route_v})
        # Response_Body = requests.post(URL_Query_Work_Sheet_Exist, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': 'xfGJgLTFcNG4Hbhx12vg3pT1FMtdDLqVfgpt8R10H1vFnM63GFxh!-48630710', 'route': 'b7724c4ffeaf48382b8d4d099b73de2f'})
        Response_Body = Response_Body.text
        Response_Body = json.loads(Response_Body)
        Exist_Work_Sheet = Response_Body['totalCount']
        if Exist_Work_Sheet == 0: # Create
            URL_Create_New_Work_Sheet_Step_1 = 'http://10.209.199.72:7112/irms/opticalSchedulingAction!init.ilf'
            Response_Body = requests.get(URL_Create_New_Work_Sheet_Step_1, cookies={'JSESSIONIRMS': Jsessionirms_v, 'route': route_v})
            # Response_Body = requests.get(URL_Create_New_Work_Sheet_Step_1, cookies={'JSESSIONIRMS': 'xfGJgLTFcNG4Hbhx12vg3pT1FMtdDLqVfgpt8R10H1vFnM63GFxh!-48630710', 'route': 'b7724c4ffeaf48382b8d4d099b73de2f'})
            Response_Body = bytes(Response_Body.text, encoding="utf-8")
            Response_Body = etree.HTML(Response_Body)
            List_Work_Sheet_Info = Response_Body.xpath('//input/@value')
            URL_Create_New_Work_Sheet_Step_2 = 'http://10.209.199.72:7112/irms/opticalSchedulingAction!submit.ilf?needSh=false'
            Form_Info_Encoded = 'ownerId=' + List_Work_Sheet_Info[1] + '&deptId=' + List_Work_Sheet_Info[2] + '&flowId=' + List_Work_Sheet_Info[3] + '&companyName=' + parse.quote_plus(List_Work_Sheet_Info[7]) + '&companyId=' + List_Work_Sheet_Info[8] + '&workitemId=' + List_Work_Sheet_Info[10] + '&activeName=' + List_Work_Sheet_Info[14] + '&formNo=' + List_Work_Sheet_Info[15] + '&title=' + parse.quote_plus(Work_Sheet_Name) + '&startFlag=' + List_Work_Sheet_Info[19] + '&ownerName=' + parse.quote_plus(List_Work_Sheet_Info[21]) + '&cellPhone=' +List_Work_Sheet_Info[22] + '&deptName=' + parse.quote_plus(List_Work_Sheet_Info[23]) + '&startTime=' + parse.quote_plus(Work_Sheet_Times[0]) + '&acceptTime=' + parse.quote_plus(Work_Sheet_Times[1]) + '&replyTime=' + parse.quote_plus(Work_Sheet_Times[2]) + '&requestTime=' + parse.quote_plus(Work_Sheet_Times[3]) + '&urgentDegree=' + parse.quote_plus('一般') + '&schedulingReason='
            Request_Lenth = str(len(Form_Info_Encoded))
            Request_Header = {'Host':'10.209.199.72:7112', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth}
            Response_Body = requests.post(URL_Create_New_Work_Sheet_Step_2, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': Jsessionirms_v, 'route': route_v})
            # Response_Body = requests.post(URL_Query_POS_ID, data=Form_Info_Encoded, headers=Request_Header, cookies={'JSESSIONIRMS': 'Cb15gGLSqLh2fbGDx9j8TVKfphF25YmQnN7h6JVyxnL7lFGg7Nvn!418505528', 'route': '5fb592aa37b5606b0629ebaa738ace15'})
            
            Repete_Start = (int(work_sheet_num) * 40)
            Repete_End = (int(work_sheet_num + 1) * 40)
            for oc_num in range(Repete_Start,Repete_End):
                List_OC_Data[oc_num]['Pro_ID'] = List_Work_Sheet_Info[3]
                if (oc_num + 1) == len(List_OC_Data):
                    break

        else:
            Repete_Start = (int(work_sheet_num) * 40)
            Repete_End = (int(work_sheet_num + 1) * 40)
            for oc_num in range(Repete_Start,Repete_End):
                List_OC_Data[oc_num]['Pro_ID'] = Response_Body['root'][0]['FLOW_ID']
                if (oc_num + 1) == len(List_OC_Data):
                    break


def Execute_Push_Box(Para_List_Box_Data):
    URL_Push_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/SynchroController/synchroData?sid=' + str(Para_List_Box_Data['Box_ID']) + '&sType=' + str(Para_List_Box_Data['Box_Type']) + '&longi='+str(Para_List_Box_Data['Longitude']) + '&lati=' + str(Para_List_Box_Data['Latitude'])
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded"}
    Response_Body = requests.post(URL_Push_Box, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    Push_Result = Response_Body.xpath('//message/text()')
    print('P1-{}-{}'.format(Para_List_Box_Data['Box_Name'], Push_Result[0]))

def Execute_Generate_Support_Segment():
    URL_Generate_Support_Segment = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesAdd'
    Form_Info_Head = '<xmldata mode="PipeLineAddMode"><mc type="yinshangduan">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for ocs_num in List_CS_Data:
        Form_Info_Body = '<mo group="1" ax="'+str(ocs_num['A_Longitude'])+'" ay="'+str(ocs_num['A_Latitude'])+'" zx="'+str(ocs_num['Z_Longitude'])+'" zy="'+str(ocs_num['Z_Latitude'])+'"><fv k="CITY_ID" v="'+str(ocs_num['City_ID'])+'"/><fv k="COUNTY_ID" v="'+str(ocs_num['County_ID'])+'"/><fv k="RELATED_SYSTEM" v="'+str(ocs_num['Support_Sys_ID'])+'"/><fv k="A_OBJECT_ID" v="'+str(ocs_num['A_ResPoint_ID'])+'"/><fv k="ZH_LABEL" v="'+str(ocs_num['A_Box_Name'])+'资源点-'+str(ocs_num['Z_Box_Name'])+'资源点引上段'+'"/><fv k="MAINTAINOR" v="'+str(ocs_num['DQS_Maintainer_ID'])+'"/><fv k="M_LENGTH" v="'+str(ocs_num['Length'])+'"/><fv k="STATUS" v="'+str(ocs_num['Life_Cycle'])+'"/><fv k="QUALITOR_COUNTY" v="'+str(ocs_num['DQS_County_ID'])+'"/><fv k="QUALITOR" v="'+str(ocs_num['DQS_ID'])+'"/><fv k="QUALITOR_PROJECT" v="'+str(ocs_num['DQS_Project_ID'])+'"/><fv k="OWNERSHIP" v="'+str(ocs_num['Owner_Type'])+'"/><fv k="RES_OWNER" v="'+str(ocs_num['Owner_Name'])+'"/><fv k="A_OBJECT_TYPE" v="'+str(ocs_num['A_ResPoint_Type_ID'])+'"/><fv k="INT_ID" v="new-27991311"/><fv k="TASK_NAME" v="'+str(ocs_num['Task_Name_ID'])+'"/><fv k="PROJECT_CODE" v="'+str(ocs_num['Project_Code_ID'])+'"/><fv k="RESOURCE_LOCATION" v="'+str(ocs_num['Field_Type'])+'"/><fv k="Z_OBJECT_TYPE" v="'+str(ocs_num['Z_ResPoint_Type_ID'])+'"/><fv k="Z_OBJECT_ID" v="'+str(ocs_num['Z_ResPoint_ID'])+'"/><fv k="SYSTEM_LEVEL" v="'+str(ocs_num['Business_Level'])+'"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Support_Segment, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Generate_Support_Segment_State = Response_Body.xpath('//@loaded')
    print('P2-引上段建立-{}'.format(List_Generate_Support_Segment_State[0]))

def Execute_Generate_Cable_Segment():
    URL_Generate_Cable_Segment = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesAdd?coreNamingRules=0'
    Form_Info_Head = '<xmldata mode="FibersegAddMode"><mc type="guanglanduan">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for ocs_num in List_CS_Data:
        Form_Info_Body = '<mo group="1" ax="'+str(ocs_num['A_Longitude'])+'" ay="'+str(ocs_num['A_Latitude'])+'" zx="'+str(ocs_num['Z_Longitude'])+'" zy="'+str(ocs_num['Z_Latitude'])+'"><fv k="QUALITOR_PROJECT" v="'+str(ocs_num['DQS_Project_ID'])+'"/><fv k="RES_OWNER" v="'+str(ocs_num['Owner_Name'])+'"/><fv k="RELATED_SYSTEM" v="'+str(ocs_num['Cable_Sys_ID'])+'"/><fv k="QUALITOR" v="'+str(ocs_num['DQS_ID'])+'"/><fv k="ZH_LABEL" v="'+str(ocs_num['A_Box_Name'])+'资源点-'+str(ocs_num['Z_Box_Name'])+'资源点'+'"/><fv k="STATUS" v="'+str(ocs_num['Life_Cycle'])+'"/><fv k="FIBER_TYPE" v="2"/><fv k="INT_ID" v="new-27991311"/><fv k="A_OBJECT_TYPE" v="'+str(ocs_num['A_ResPoint_Type_ID'])+'"/><fv k="Z_OBJECT_ID" v="'+str(ocs_num['Z_ResPoint_ID'])+'"/><fv k="A_OBJECT_ID" v="'+str(ocs_num['A_ResPoint_ID'])+'"/><fv k="Z_OBJECT_TYPE" v="'+str(ocs_num['Z_ResPoint_Type_ID'])+'"/><fv k="M_LENGTH" v="'+str(ocs_num['Length'])+'"/><fv k="CITY_ID" v="'+str(ocs_num['City_ID'])+'"/><fv k="FIBER_NUM" v="'+str(ocs_num['Width'])+'"/><fv k="MAINTAINOR" v="'+str(ocs_num['DQS_Maintainer_ID'])+'"/><fv k="WIRE_SEG_TYPE" v="GYTA-'+str(ocs_num['Width'])+'"/><fv k="SERVICE_LEVEL" v="14"/><fv k="COUNTY_ID" v="'+str(ocs_num['County_ID'])+'"/><fv k="PROJECTCODE" v="'+str(ocs_num['Project_Code_ID'])+'"/><fv k="TASK_NAME" v="'+str(ocs_num['Task_Name_ID'])+'"/><fv k="OWNERSHIP" v="'+str(ocs_num['Owner_Type'])+'"/><fv k="QUALITOR_COUNTY" v="'+str(ocs_num['DQS_County_ID'])+'"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Cable_Segment, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Generate_Cable_Segment_State = Response_Body.xpath('//@loaded')
    print('P3-光缆段建立-{}'.format(List_Generate_Cable_Segment_State[0]))

def Execute_Cable_Lay(Para_List_CS_Data):
    URL_Cable_Lay = 'http://10.209.199.74:8120/igisserver_osl/rest/optCabLayInspur/saveFiberSegM1'
    Form_Info = '<xmldata><fiberseg id="'+str(Para_List_CS_Data['Cable_Seg_ID'])+'" aid="'+str(Para_List_CS_Data['A_ResPoint_ID'])+'" zid="'+str(Para_List_CS_Data['Z_ResPoint_ID'])+'"/><cablays><cablay id="'+str(Para_List_CS_Data['Support_Seg_ID'])+'" type="yinshangduan" name="'+str(Para_List_CS_Data['Support_Seg_Name'])+'" aid="'+str(Para_List_CS_Data['A_ResPoint_ID'])+'" zid="'+str(Para_List_CS_Data['Z_ResPoint_ID'])+'"/></cablays></xmldata>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Cable_Lay, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Cable_Lay_State = Response_Body.xpath('//@msg')
    print('P4-敷设-{}-{}'.format(List_Cable_Lay_State[0], Para_List_CS_Data['Cable_Seg_Name']))

def Execute_Generate_ODM(Para_List_Box_Data):
    URL_Generate_ODM = 'http://10.209.199.74:8120/nxapp/room/editODMData.ilf'
    Form_Info = '<params><odm id="" rowflag="+" rownum="1" colflag="+" colnum="12"><attribute module_rowno="1" rowno="'+str(Para_List_Box_Data['ODM_Rows'])+'" columnno="12" terminal_arr="0" maintain_county="'+str(Para_List_Box_Data['County_ID'])+'" maintain_city="'+str(Para_List_Box_Data['City_ID'])+'" structure_id="'+str(Para_List_Box_Data['ResPoint_ID'])+'" structure_type="'+str(Para_List_Box_Data['ResPoint_Type_ID'])+'" related_rack="'+str(Para_List_Box_Data['Box_ID'])+'" related_type="'+str(Para_List_Box_Data['Box_Type_ID'])+'" status="8" model="odm"/></odm></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&model=odm&" +  "lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_ODM, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = parse.unquote(bytes(Response_Body.text, encoding="utf-8"))
    Response_Body = etree.HTML(Response_Body)
    List_ODM_ID = Response_Body.xpath('//@int_id')
    Para_List_Box_Data['ODM_ID'] = List_ODM_ID[0]
    print('P5-ODM-{}-in-{}'.format(List_ODM_ID[0], Para_List_Box_Data['Box_Name']))

def Execute_Generate_Tray(Para_List_Box_Data):
    URL_Generate_Tray = "http://10.209.199.74:8120/nxapp/room/editTray_sx.ilf"
    Form_Info = '<params model="tray"><obj related_rack="'+str(Para_List_Box_Data['Box_ID'])+'" related_type="'+str(Para_List_Box_Data['Box_Type_ID'])+'" structure_id="'+str(Para_List_Box_Data['ResPoint_ID'])+'" structure_type="'+str(Para_List_Box_Data['ResPoint_Type_ID'])+'" deviceshelf_id="'+str(Para_List_Box_Data['ODM_ID'])+'" tray_no="1" tray_num="'+str(Para_List_Box_Data['Tray_Count'])+'" row_count="1" col_count="12" int_id=""/></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&model=odm&" +  "lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Tray, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = parse.unquote(bytes(Response_Body.text, encoding="utf-8"))
    Response_Body = etree.HTML(Response_Body)
    List_Terminal_IDs = Response_Body.xpath('//terminal//@int_id')
    print('P5-Terminal_IDs_Count-{}-{}'.format(len(List_Terminal_IDs), Para_List_Box_Data['Box_Name']))
    Para_List_Box_Data['Terminal_IDs'] = '&'.join(List_Terminal_IDs)

def Execute_Termination(Para_List_Box_Data):

    if Para_List_Box_Data['1FS_Count'] == 0:

        List_Terminal_IDs = Para_List_Box_Data['Terminal_IDs'].split('&')

        if Para_List_Box_Data['Box_Type'] == 'guangfenxianxiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gfxx'
        elif Para_List_Box_Data['Box_Type'] == 'guangjiaojiexiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gjjx'

        for cable_seg_data in List_CS_Data:
            if Para_List_Box_Data['Box_Name'] == cable_seg_data['Z_Box_Name']:
                List_CS_Fiber_IDs = cable_seg_data['CS_Fiber_IDs'].split('&')

        termination_datas = []
        for termination_num in range(Para_List_Box_Data['Termination_Count']):
            termination_data = '<param fiber_id="'+str(List_CS_Fiber_IDs[termination_num])+'" z_equ_id="'+str(Para_List_Box_Data['Box_ID'])+'" z_equ_type="'+str(Para_List_Box_Data['Box_Type_Short'])+'" z_port_id="'+str(List_Terminal_IDs[termination_num])+'" a_equ_id="" a_port_id="" room_id="'+str(Para_List_Box_Data['ResPoint_ID'])+'"/>'
            termination_datas.append(termination_data)

        Form_Info_Body = ''.join(termination_datas)

    elif Para_List_Box_Data['1FS_Count'] != 0:

        List_CS_Fiber_IDs_Joined = []
        for cable_seg_data in List_CS_Data:
            if Para_List_Box_Data['Box_Name'] == cable_seg_data['A_Box_Name']:
                List_CS_Fiber_IDs_Joined.append(cable_seg_data['CS_Fiber_IDs'])
        List_CS_Fiber_IDs_Separated = [i.split('&') for i in List_CS_Fiber_IDs_Joined]

        List_Termination_Count = Para_List_Box_Data['Termination_Count'].split('&')

        List_Terminal_IDs = Para_List_Box_Data['Terminal_IDs'].split('&')

        if Para_List_Box_Data['Box_Type'] == 'guangfenxianxiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gfxx'
        elif Para_List_Box_Data['Box_Type'] == 'guangjiaojiexiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gjjx'

        Termination_Pointer = 0
        termination_datas = []
        for each_Termination_Count, each_CS_Fiber_IDs in zip(List_Termination_Count, List_CS_Fiber_IDs_Separated):

            if Termination_Pointer % 12 != 0:
                for i in range(11):
                    i += 0
                    Termination_Pointer += 1
                    if Termination_Pointer % 12 == 0:
                        break

            for termination_num in range(int(each_Termination_Count)):
                termination_data = '<param fiber_id="'+str(each_CS_Fiber_IDs[termination_num])+'" a_equ_id="'+str(Para_List_Box_Data['Box_ID'])+'" a_equ_type="'+Para_List_Box_Data['Box_Type_Short']+'" a_port_id="'+str(List_Terminal_IDs[Termination_Pointer])+'" z_equ_id="" z_port_id="" room_id="'+Para_List_Box_Data['ResPoint_ID']+'"/>'
                Termination_Pointer += 1
                termination_datas.append(termination_data)

            Form_Info_Body = ''.join(termination_datas)

    URL_Termination = 'http://10.209.199.74:8120/igisserver_osl/rest/jumperandfiber/saveFiber'
    Form_Info_Head = '<params>'
    Form_Info_Butt = '</params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = 'params=' + parse.quote_plus(Form_Info_Head + Form_Info_Body + Form_Info_Butt) + '&lifeparams=' + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Termination, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Termination_State = Response_Body.xpath('//@msg')
    print('P6-{}-{}'.format(Para_List_Box_Data['Box_Name'], List_Termination_State[0]))

def Execute_Direct_Melt(Para_List_Box_Data):
    if Para_List_Box_Data['1FS_Count'] == 0:

        for cable_seg_data in List_CS_Data:
            if Para_List_Box_Data['Box_Name'] == cable_seg_data['Z_Box_Name']:
                List_UL_CS_Fiber_IDs = cable_seg_data['CS_Fiber_IDs'].split('&')
            if Para_List_Box_Data['Box_Name'] == cable_seg_data['A_Box_Name']:
                List_DL_CS_Fiber_IDs = cable_seg_data['CS_Fiber_IDs'].split('&')
            
        if Para_List_Box_Data['Box_Type'] == 'guangfenxianxiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gfxx'
        elif Para_List_Box_Data['Box_Type'] == 'guangjiaojiexiang':
            Para_List_Box_Data['Box_Type_Short'] = 'gjjx'

        URL_Direct_Melt = 'http://10.209.199.74:8120/igisserver_osl/rest/fibercorekiss/fiberKiss'

        Form_Info_Head = '<params respoint_id="'+Para_List_Box_Data['ResPoint_ID']+'" respoint_type="'+str(Para_List_Box_Data['ResPoint_Type_ID'])+'" parent_rack_id="null" parent_rack_type="null">'
        each_direct_melt_datas = []
        for each_direct_melt_data_num in range(Para_List_Box_Data['Direct_Melt_Count']):
            each_direct_melt_data = '<param a_fiber_id="'+str(List_UL_CS_Fiber_IDs[each_direct_melt_data_num + Para_List_Box_Data['Direct_Melt_Start'] - 1])+'" z_fiber_id="'+str(List_DL_CS_Fiber_IDs[each_direct_melt_data_num])+'"/>'
            each_direct_melt_datas.append(each_direct_melt_data)
        Form_Info_Body = ''.join(each_direct_melt_datas)
        Form_Info_Butt = '</params>'
        Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
        Form_Info_Encoded = 'params=' + parse.quote_plus(Form_Info_Head + Form_Info_Body + Form_Info_Butt) + '&lifeparams=' + parse.quote_plus(Form_Info_Tail)
        Request_Lenth = str(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Direct_Melt, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        List_Direct_Melt_State = Response_Body.xpath('//@msg')
        print('P6-{}-{}'.format(Para_List_Box_Data['Box_Name'], List_Direct_Melt_State[0]))

    elif Para_List_Box_Data['1FS_Count'] != 0:
        print('P6-{}-是一级分纤箱,不涉及直熔'.format(Para_List_Box_Data['Box_Name']))


def Main_Process(Para_File_Name):
    Generate_Local_Data(Para_File_Name)
    if (P1_Push_Box or 
        P2_Generate_Support_Segment or 
        P3_Generate_Cable_Segment or 
        P4_Cable_Lay or 
        P5_Generate_ODM_and_Tray or 
        P6_Termination_and_Direct_Melt or 
        P7_Generate_Optical_Circut):
        print('查询Box/ResPoint开始')
        Swimming_Pool(Query_Box_ID_ResPoint_ID_Alias, List_Box_Data)
        print('查询Box/ResPoint结束')
    if (P2_Generate_Support_Segment or 
        P3_Generate_Cable_Segment or 
        P4_Cable_Lay or 
        P5_Generate_ODM_and_Tray or 
        P6_Termination_and_Direct_Melt):
        print('查询Support_Sys_ID/Cable_Sys_ID')
        Query_Support_Sys_and_Cable_Sys()
    if (P2_Generate_Support_Segment or 
        P3_Generate_Cable_Segment or 
        P7_Generate_Optical_Circut):
        print('查询Project_Code_ID')
        Query_Project_Code_ID()
    if (P4_Cable_Lay or 
        P6_Termination_and_Direct_Melt):
        print('查询Support_Seg_ID/Cable_Seg_ID开始')
        Swimming_Pool(Query_Support_Seg_ID_Cable_Seg_ID, List_CS_Data)
        print('查询Support_Seg_ID/Cable_Seg_ID结束')
    if (P5_Generate_ODM_and_Tray or 
        P6_Termination_and_Direct_Melt):
        Generate_Topology()
        Generate_FS_Data()
        if not ('ODM_ID' in List_Box_Data[0]):
            print('查询ODM_ID开始')
            Swimming_Pool(Query_ODM_ID_and_Terminarl_IDs, List_Box_Data)
            print('查询ODM_ID结束')
    if P6_Termination_and_Direct_Melt:
        print('查询Cable_Fiber_ID开始')
        Swimming_Pool(Query_CS_Fiber_IDs, List_CS_Data)
        Generate_Termination_and_Direct_Melt_Data()
        print('查询Cable_Fiber_ID结束')
    if P7_Generate_Optical_Circut:
        print('查询POS_ID开始')
        Query_JSESSIONIRMS_and_route()
        Swimming_Pool(Query_POS_ID, List_Box_Data)
        Generate_OC_POS_Data_and_OC_Name()
        print('查询POS_ID结束')


    if P1_Push_Box:
        print('P1-开始')
        Swimming_Pool(Execute_Push_Box, List_Box_Data)
        print('P1-结束')
    if P2_Generate_Support_Segment:
        print('P2-开始')
        Execute_Generate_Support_Segment()
        print('P2-结束')
    if P3_Generate_Cable_Segment:
        print('P3-开始')
        Execute_Generate_Cable_Segment()
        print('P3-结束')
    if P4_Cable_Lay:
        print('P4-开始')
        Swimming_Pool(Execute_Cable_Lay, List_CS_Data)
        print('P4-结束')
    if P5_Generate_ODM_and_Tray:
        print('P5-开始')
        Swimming_Pool(Execute_Generate_ODM, List_Box_Data)
        Swimming_Pool(Execute_Generate_Tray, List_Box_Data)
        print('P5-结束')
    if P6_Termination_and_Direct_Melt:
        print('P6-开始')
        Swimming_Pool(Execute_Termination, List_Box_Data)
        Swimming_Pool(Execute_Direct_Melt, List_Box_Data)
        print('P6-结束')
    if P7_Generate_Optical_Circut:
        print('P7-开始')
        print('P7-结束')


if __name__ == '__main__':
    # for each_File_Name in File_Name:
    #     Main_Process(each_File_Name)

    # for each_oc_data in List_OC_Data:
    #     print(each_oc_data['OC_Name'])

    # print(sorted(List_CS_Data[10].items(), key = lambda item:item[0]))

    Generate_Local_Data('豪德置业')
    Query_JSESSIONIRMS_and_route()
    Query_Work_Sheet_ID()


    test001 = ['', '115501', '166', '1611142394539', '', '', '', '山西省太原市', '324', '', '0', '', '', '', 'glsq', 'SX-020-210120-195', '', '', '光路调度-210120-157', 'true', '', '杨伟', '13603516690', '工程建设部', 'Wed Jan 20 19:33:13 CST 2021', 'Fri Jan 22 19:33:13 CST 2021', 'Sun Jan 24 19:33:13 CST 2021', '', '附件上传', '提交']