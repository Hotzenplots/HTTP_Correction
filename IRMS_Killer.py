import csv
import math
import requests
import copy
from urllib import parse
from lxml import etree
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from time import sleep

File_Name = '豪德置业'

P1_Push_Box                 = False
P2_Generate_Support_Segment = False
P3_Generate_Cable_Segment   = False
P4_Cable_Lay                = False
P5_Generate_ODM_and_Tray    = True

def File_Process_and_Generate_Basic_Data():
    '''读取本地数据,在不查询的前提下填充几个List'''

    '''读取并整理7013表,生成List_7013'''
    global List_7013
    List_7013 = []
    with open(File_Name+'.csv') as file_csv:
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
    WB_obj = load_workbook(File_Name+'.xlsx')
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

    '''读取并整理Sheet_Box_Topology,生成List_Box_Info'''
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
    global List_Box_Info
    List_Box_Info = []
    for box_num in range(len(List_Box_Name)):
        List_Box_Info.append(dict({'Box_Name':List_Box_Name[box_num][0]}))
        List_Box_Info[box_num]['Longitude'] = Longitude_Start + (List_Box_Name[box_num][2] - 1) * Horizontal_Density
        List_Box_Info[box_num]['Latitude'] = Latitude_Start - (List_Box_Name[box_num][1] - 1) * Vertical_Density
        if (List_Box_Info[box_num]['Box_Name'].find('GJ') != -1) or (List_Box_Info[box_num]['Box_Name'].find('gj') != -1):
            List_Box_Info[box_num]['Box_Type'] = 'guangjiaojiexiang'
        else:
            List_Box_Info[box_num]['Box_Type'] = 'guangfenxianxiang'
        for row_data in range(len(List_Box_Type)):
            if List_Box_Info[box_num]['Box_Type'] == List_Box_Type[row_data][0]:
                List_Box_Info[box_num]['Box_Type_ID'] = List_Box_Type[row_data][1]
                List_Box_Info[box_num]['ResPoint_Type_ID'] = List_Box_Type[row_data][2]
        List_Box_Info[box_num]['City_ID'] = List_Template_Selected[1]
        List_Box_Info[box_num]['County_ID'] = List_Template_Selected[16]

    '''读取并整理Sheet_OCS_List,生成List_CS_Info'''
    WS_obj = WB_obj['OCS_List']
    global List_CS_Info
    List_CS_Info = []
    for row_num in range(1,501):
        OCS_A_Box_Name = WS_obj.cell(row_num, 1).value
        OCS_Z_Box_Name = WS_obj.cell(row_num, 2).value
        OCS_Width = WS_obj.cell(row_num, 3).value
        if OCS_A_Box_Name == None:
            break
        List_CS_Info.append(dict({'A_Box_Name': OCS_A_Box_Name}))
        List_CS_Info[row_num - 1]['Z_Box_Name'] = OCS_Z_Box_Name
        List_CS_Info[row_num - 1]['Width'] = OCS_Width
    
    #Length_Prepare
    Horizontal_Metre = 111.11 * 1000 * math.cos(Latitude_Start * math.pi / 180)
    Vertical_Metre = 111.11 * 1000
    
    for dic_num_in_osc in List_CS_Info:
        for dic_num_in_box in List_Box_Info:
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

def Query_Project_Code_ID():
    URL_Query_Project_Code_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/datatrans/expall?model=guangfenxianxiang&fname=PROJECTCODE&p1='+List_7013[1][4]
    Response_Body = requests.get(URL_Query_Project_Code_ID)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//@name")
    List_Response_Value = Response_Body.xpath("//@value")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    Dic_Project_Code_ID = copy.deepcopy(Dic_Response)
    for box_num in List_Box_Info:
        box_num['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]
    for ocs_num in List_CS_Info:
        ocs_num['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]

def Query_Box_ID_ResPont_ID_Alias(Para_List_Box_Info):
    URL_Query_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Para_List_Box_Info['Box_Type']+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_List_Box_Info['Box_Name']+'%\'" returnfields="INT_ID,ZH_LABEL,STRUCTURE_ID,ALIAS"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Box, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    for box_num in range(len(List_Box_Info)):
        if Dic_Response['ZH_LABEL'] == List_Box_Info[box_num]['Box_Name']:
            List_Box_Info[box_num]['Box_ID'] = Dic_Response['INT_ID']
            List_Box_Info[box_num]['ResPoint_ID'] = Dic_Response['STRUCTURE_ID']
            List_Box_Info[box_num]['Alias'] = Dic_Response['ALIAS']
    for ocs_num in List_CS_Info:
        for box_num in List_Box_Info:
            if ocs_num['A_Box_Name'] == box_num['Box_Name']:
                ocs_num['A_Box_ID'] = box_num['Box_ID']
                ocs_num['A_ResPoint_ID'] = box_num['ResPoint_ID']
        for box_num in List_Box_Info:
            if ocs_num['Z_Box_Name'] == box_num['Box_Name']:
                ocs_num['Z_Box_ID'] = box_num['Box_ID']
                ocs_num['Z_ResPoint_ID'] = box_num['ResPoint_ID']

def Prepare_Support_Sys_and_Cable_Sys():
    Task_Name_ID_List = List_7013[1][5].split('-')
    Support_Sys_Name = List_7013[1][0]+List_7013[1][1]+Task_Name_ID_List[1]+'区内引上系统'
    Cable_Sys_Name = List_7013[1][0]+List_7013[1][1]+Task_Name_ID_List[1]+'接入层架空光缆'

    #Support_System
    URL_Query_SS_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGetPage'
    Form_Info = '<request><query mc="xitong" where="1=1 AND LINESEG_TYPE=\'9108\' AND ZH_LABEL LIKE \'%'+Support_Sys_Name+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'pageNum=1&'+'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_SS_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_SS_Count = Response_Body.xpath('//@counts')
    if List_SS_Count[0] == '0':
        #Generate Support_Sys
        URL_Add_Support_Sys = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalSave'
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Info[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = chr(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Support_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Support_Sys_ID = Response_Body.xpath('//@newid')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_CS_Info:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('引上系统ID-{}'.format(Support_Sys_ID))

    elif List_SS_Count[0] == '1':
        #Get Support_Sys_ID
        Support_Sys_ID = Response_Body.xpath('//@int_id')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_CS_Info:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('引上系统ID-{}'.format(Support_Sys_ID))

    #Cable_System
    URL_Query_CS_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGetPage'
    Form_Info = '<request><query mc="guanglanxitong" where="1=1 AND ZH_LABEL LIKE \'%'+Cable_Sys_Name+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'pageNum=1&'+'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Query_CS_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_CS_Count = Response_Body.xpath('//@counts')
    if List_CS_Count[0] == '0':
        #Generate Cable_Sys
        URL_Add_Cable_Sys = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalSave'
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Info[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info = '<response mode="add"><mc type="guanglanxitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_CS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_CS_Info[0]['County_ID'])+'"/><fv k="ZH_LABEL" v="'+Cable_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/><fv k="OPT_TYPE" v="GYTA-12"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = chr(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Cable_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Cable_Sys_ID = Response_Body.xpath('//@newid')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_CS_Info:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('光缆系统ID-{}'.format(Cable_Sys_ID))

    elif List_CS_Count[0] == '1':
        #Get Cable_Sys_ID
        Cable_Sys_ID = Response_Body.xpath('//@int_id')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_CS_Info:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('光缆系统ID-{}'.format(Cable_Sys_ID))

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Get_JSESSIONIRMS_and_route():
    Browser_Obj = webdriver.Ie()
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_was/uploadResource/public/login/login.html')
    Browser_Obj.find_element_by_id('username').send_keys('tyyangwei')
    Browser_Obj.find_element_by_id('password').send_keys('tyyw159...')
    Browser_Obj.find_element_by_class_name('lwb_login').click()
    sleep(2)
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_wcm/middelwebpage/app_recoder_log.jsp?app_flg=zhwlzygl_ywzl')
    sleep(2)
    Dic_Cookie_JSESSIONIRMS = Browser_Obj.get_cookie('JSESSIONIRMS')
    Dic_Cookie_route = Browser_Obj.get_cookie('route')
    JSESSIONIRMS_Value = Dic_Cookie_JSESSIONIRMS['value']
    route_Value = Dic_Cookie_route['value']
    Browser_Obj.quit()
    return [JSESSIONIRMS_Value, route_Value]

def Push_Box(Para_List_Box_Info):
    URL_Push_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/SynchroController/synchroData?sid=' + str(Para_List_Box_Info['Box_ID']) + '&sType=' + str(Para_List_Box_Info['Box_Type']) + '&longi='+str(Para_List_Box_Info['Longitude']) + '&lati=' + str(Para_List_Box_Info['Latitude'])
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded"}
    Response_Body = requests.post(URL_Push_Box, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    Push_Result = Response_Body.xpath('//message/text()')
    print('P1-{}-{}'.format(Para_List_Box_Info['Box_Name'], Push_Result[0]))

def Generate_Support_Segment():
    URL_Generate_Support_Segment = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesAdd'
    Form_Info_Head = '<xmldata mode="PipeLineAddMode"><mc type="yinshangduan">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for ocs_num in List_CS_Info:
        Form_Info_Body = '<mo group="1" ax="'+str(ocs_num['A_Longitude'])+'" ay="'+str(ocs_num['A_Latitude'])+'" zx="'+str(ocs_num['Z_Longitude'])+'" zy="'+str(ocs_num['Z_Latitude'])+'"><fv k="CITY_ID" v="'+str(ocs_num['City_ID'])+'"/><fv k="COUNTY_ID" v="'+str(ocs_num['County_ID'])+'"/><fv k="RELATED_SYSTEM" v="'+str(ocs_num['Support_Sys_ID'])+'"/><fv k="A_OBJECT_ID" v="'+str(ocs_num['A_ResPoint_ID'])+'"/><fv k="ZH_LABEL" v="'+str(ocs_num['A_Box_Name'])+'资源点-'+str(ocs_num['Z_Box_Name'])+'资源点引上段'+'"/><fv k="MAINTAINOR" v="'+str(ocs_num['DQS_Maintainer_ID'])+'"/><fv k="M_LENGTH" v="'+str(ocs_num['Length'])+'"/><fv k="STATUS" v="'+str(ocs_num['Life_Cycle'])+'"/><fv k="QUALITOR_COUNTY" v="'+str(ocs_num['DQS_County_ID'])+'"/><fv k="QUALITOR" v="'+str(ocs_num['DQS_ID'])+'"/><fv k="QUALITOR_PROJECT" v="'+str(ocs_num['DQS_Project_ID'])+'"/><fv k="OWNERSHIP" v="'+str(ocs_num['Owner_Type'])+'"/><fv k="RES_OWNER" v="'+str(ocs_num['Owner_Name'])+'"/><fv k="A_OBJECT_TYPE" v="'+str(ocs_num['A_ResPoint_Type_ID'])+'"/><fv k="INT_ID" v="new-27991311"/><fv k="TASK_NAME" v="'+str(ocs_num['Task_Name_ID'])+'"/><fv k="PROJECT_CODE" v="'+str(ocs_num['Project_Code_ID'])+'"/><fv k="RESOURCE_LOCATION" v="'+str(ocs_num['Field_Type'])+'"/><fv k="Z_OBJECT_TYPE" v="'+str(ocs_num['Z_ResPoint_Type_ID'])+'"/><fv k="Z_OBJECT_ID" v="'+str(ocs_num['Z_ResPoint_ID'])+'"/><fv k="SYSTEM_LEVEL" v="'+str(ocs_num['Business_Level'])+'"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Support_Segment, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Generate_Support_Segment_State = Response_Body.xpath('//@loaded')
    print('P2-引上段建立-{}'.format(List_Generate_Support_Segment_State[0]))

def Generate_Cable_Segment():
    URL_Generate_Cable_Segment = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesAdd?coreNamingRules=0'
    Form_Info_Head = '<xmldata mode="FibersegAddMode"><mc type="guanglanduan">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for ocs_num in List_CS_Info:
        Form_Info_Body = '<mo group="1" ax="'+str(ocs_num['A_Longitude'])+'" ay="'+str(ocs_num['A_Latitude'])+'" zx="'+str(ocs_num['Z_Longitude'])+'" zy="'+str(ocs_num['Z_Latitude'])+'"><fv k="QUALITOR_PROJECT" v="'+str(ocs_num['DQS_Project_ID'])+'"/><fv k="RES_OWNER" v="'+str(ocs_num['Owner_Name'])+'"/><fv k="RELATED_SYSTEM" v="'+str(ocs_num['Cable_Sys_ID'])+'"/><fv k="QUALITOR" v="'+str(ocs_num['DQS_ID'])+'"/><fv k="ZH_LABEL" v="'+str(ocs_num['A_Box_Name'])+'资源点-'+str(ocs_num['Z_Box_Name'])+'资源点'+'"/><fv k="STATUS" v="'+str(ocs_num['Life_Cycle'])+'"/><fv k="FIBER_TYPE" v="2"/><fv k="INT_ID" v="new-27991311"/><fv k="A_OBJECT_TYPE" v="'+str(ocs_num['A_ResPoint_Type_ID'])+'"/><fv k="Z_OBJECT_ID" v="'+str(ocs_num['Z_ResPoint_ID'])+'"/><fv k="A_OBJECT_ID" v="'+str(ocs_num['A_ResPoint_ID'])+'"/><fv k="Z_OBJECT_TYPE" v="'+str(ocs_num['Z_ResPoint_Type_ID'])+'"/><fv k="M_LENGTH" v="'+str(ocs_num['Length'])+'"/><fv k="CITY_ID" v="'+str(ocs_num['City_ID'])+'"/><fv k="FIBER_NUM" v="'+str(ocs_num['Width'])+'"/><fv k="MAINTAINOR" v="'+str(ocs_num['DQS_Maintainer_ID'])+'"/><fv k="WIRE_SEG_TYPE" v="GYTA-'+str(ocs_num['Width'])+'"/><fv k="SERVICE_LEVEL" v="14"/><fv k="COUNTY_ID" v="'+str(ocs_num['County_ID'])+'"/><fv k="PROJECTCODE" v="'+str(ocs_num['Project_Code_ID'])+'"/><fv k="TASK_NAME" v="'+str(ocs_num['Task_Name_ID'])+'"/><fv k="OWNERSHIP" v="'+str(ocs_num['Owner_Type'])+'"/><fv k="QUALITOR_COUNTY" v="'+str(ocs_num['DQS_County_ID'])+'"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Cable_Segment, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Generate_Cable_Segment_State = Response_Body.xpath('//@loaded')
    print('P3-光缆段建立-{}'.format(List_Generate_Cable_Segment_State[0]))

def Query_Support_Seg_ID_Cable_Seg_ID(Para_List_CS_Info):
    List_CS_Support_Seg_Name_ID_Cable_Name_ID = {}
    URL_Query_Support_Seg_ID_Cable_Seg_ID = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'

    Form_Info = '<request><query mc="yinshangduan" where="1=1 AND ZH_LABEL LIKE \'%'+str(Para_List_CS_Info['A_Box_Name'])+'资源点-'+str(Para_List_CS_Info['Z_Box_Name'])+'资源点'+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Support_Seg_ID_Cable_Seg_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_Name'] = List_Response_Value[1]
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_ID'] = List_Response_Value[0]

    Form_Info = '<request><query mc="guanglanduan" where="1=1 AND ZH_LABEL LIKE \'%'+str(Para_List_CS_Info['A_Box_Name'])+'资源点-'+str(Para_List_CS_Info['Z_Box_Name'])+'资源点'+'%\'" returnfields="INT_ID,ZH_LABEL"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Support_Seg_ID_Cable_Seg_ID, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_Name'] = List_Response_Value[1]
    List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_ID'] = List_Response_Value[0]
    for cable_seg in List_CS_Info:
        if Para_List_CS_Info['A_Box_Name'] == cable_seg['A_Box_Name']:
            cable_seg['Support_Seg_Name'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_Name']
            cable_seg['Support_Seg_ID'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Support_Seg_ID']
            cable_seg['Cable_Seg_Name'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_Name']
            cable_seg['Cable_Seg_ID'] = List_CS_Support_Seg_Name_ID_Cable_Name_ID['Cable_Seg_ID']

def Cable_Lay(Para_List_CS_Info):
    URL_Cable_Lay = 'http://10.209.199.74:8120/igisserver_osl/rest/optCabLayInspur/saveFiberSegM1'
    Form_Info = '<xmldata><fiberseg id="'+str(Para_List_CS_Info['Cable_Seg_ID'])+'" aid="'+str(Para_List_CS_Info['A_ResPoint_ID'])+'" zid="'+str(Para_List_CS_Info['Z_ResPoint_ID'])+'"/><cablays><cablay id="'+str(Para_List_CS_Info['Support_Seg_ID'])+'" type="yinshangduan" name="'+str(Para_List_CS_Info['Support_Seg_Name'])+'" aid="'+str(Para_List_CS_Info['A_ResPoint_ID'])+'" zid="'+str(Para_List_CS_Info['Z_ResPoint_ID'])+'"/></cablays></xmldata>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Cable_Lay, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Cable_Lay_State = Response_Body.xpath('//@msg')
    print('P4-敷设-{}-{}'.format(List_Cable_Lay_State[0], Para_List_CS_Info['Cable_Seg_Name']))

def Generate_Topology():
    global CS_Topology
    CS_Topology = []
    Line_Num = 0
    Segment_Num = 0
    for cable_seg_num in range(len(List_CS_Info)):
        if cable_seg_num == 0:
            CS_Topology.append([List_CS_Info[cable_seg_num]['A_Box_Name']])
            CS_Topology[cable_seg_num].append(List_CS_Info[cable_seg_num]['Z_Box_Name'])
            Segment_Num += 1
            continue
        if List_CS_Info[cable_seg_num]['A_Box_Name'] == CS_Topology[Line_Num][Segment_Num]:
            CS_Topology[Line_Num].append(List_CS_Info[cable_seg_num]['Z_Box_Name'])
            Segment_Num += 1
            continue
        else:
            CS_Topology.append([List_CS_Info[cable_seg_num]['A_Box_Name']])
            Line_Num += 1
            Segment_Num = 1
            CS_Topology[Line_Num].append(List_CS_Info[cable_seg_num]['Z_Box_Name'])

def Generate_FS_Info():
    #filled by 0
    for box_info in List_Box_Info:
        box_info['1FS_Count'] = 0
        box_info['2FS_Count'] = 0
        box_info['DL_2FS_Count'] = 0
        box_info['ODM_Rows'] = 0
        box_info['Tray_Count'] = 0
    #1fs_count&2fs_count
    for row in List_7013:
        for box_info in List_Box_Info:
            if row[3] == box_info['Box_Name']:
                if row[10] == '一级':
                    box_info['1FS_Count'] += 1
                elif row[10] == '二级':
                    box_info['2FS_Count'] += 1
    #dl_2fs_count
    for cable_num in range(len(CS_Topology) - 1, -1, -1):
        for cable_seg_num in range(len(CS_Topology[cable_num]) - 1, -1, -1):
            if cable_seg_num == len(CS_Topology[cable_num]) - 1:#Tail
                for box_info in List_Box_Info: 
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = 0
            elif cable_seg_num == 0:#Head
                for box_info in List_Box_Info:
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = 0
            else:#Middle
                for box_info in List_Box_Info:
                    if CS_Topology[cable_num][cable_seg_num + 1] == box_info['Box_Name']:
                        DL_2FS_Count_temp = box_info['2FS_Count'] + box_info['DL_2FS_Count']
                for box_info in List_Box_Info:
                    if CS_Topology[cable_num][cable_seg_num] == box_info['Box_Name']:
                        box_info['DL_2FS_Count'] = DL_2FS_Count_temp
    #dl_2fs_count 1fs_box_correction
    for box_info in List_Box_Info:
        if box_info['1FS_Count'] != 0:
            DL_2FS_Count_temp = []
            for cable_num in CS_Topology:
                if box_info['Box_Name'] == cable_num[0]:
                    for  box_info_2 in List_Box_Info:
                        if cable_num[1] == box_info_2['Box_Name']:
                            DL_2FS_Count_temp.append(str(box_info_2['DL_2FS_Count'] + box_info_2['2FS_Count']))
                            box_info['DL_2FS_Count'] = '&'.join(DL_2FS_Count_temp)
    #ODM_Rows$Tray_Count
    for box_info in List_Box_Info:
        if box_info['1FS_Count'] == 0:
            for cable_seg_info in List_CS_Info:
                if box_info['Box_Name'] == cable_seg_info['Z_Box_Name']:
                    box_info['ODM_Rows'] = box_info['Tray_Count'] = math.ceil(cable_seg_info['Width'] / 12)
        elif box_info['1FS_Count'] != 0:
            for cable_seg_info in List_CS_Info:
                if box_info['Box_Name'] == cable_seg_info['A_Box_Name']:
                    box_info['ODM_Rows'] = box_info['Tray_Count'] = math.ceil(cable_seg_info['Width'] / 12) + box_info['ODM_Rows']

def Generate_ODM(Para_List_Box_Info):
    URL_Generate_ODM = 'http://10.209.199.74:8120/nxapp/room/editODMData.ilf'
    Form_Info = '<params><odm id="" rowflag="+" rownum="1" colflag="+" colnum="12"><attribute module_rowno="1" rowno="'+str(Para_List_Box_Info['ODM_Rows'])+'" columnno="12" terminal_arr="0" maintain_county="'+str(Para_List_Box_Info['County_ID'])+'" maintain_city="'+str(Para_List_Box_Info['City_ID'])+'" structure_id="'+str(Para_List_Box_Info['ResPoint_ID'])+'" structure_type="'+str(Para_List_Box_Info['ResPoint_Type_ID'])+'" related_rack="'+str(Para_List_Box_Info['Box_ID'])+'" related_type="'+str(Para_List_Box_Info['Box_Type_ID'])+'" status="8" model="odm"/></odm></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&model=odm&" +  "lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_ODM, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = parse.unquote(bytes(Response_Body.text, encoding="utf-8"))
    Response_Body = etree.HTML(Response_Body)
    List_ODM_ID = Response_Body.xpath('//@int_id')
    Para_List_Box_Info['ODM_ID'] = List_ODM_ID[0]
    print('P5-ODM-{}-in-{}'.format(List_ODM_ID[0], Para_List_Box_Info['Box_Name']))

def Generate_Tray(Para_List_Box_Info):
    URL_Generate_Tray = "http://10.209.199.74:8120/nxapp/room/editTray_sx.ilf"
    Form_Info = '<params model="tray"><obj related_rack="'+str(Para_List_Box_Info['Box_ID'])+'" related_type="'+str(Para_List_Box_Info['Box_Type_ID'])+'" structure_id="'+str(Para_List_Box_Info['ResPoint_ID'])+'" structure_type="'+str(Para_List_Box_Info['ResPoint_Type_ID'])+'" deviceshelf_id="'+str(Para_List_Box_Info['ODM_ID'])+'" tray_no="1" tray_num="'+str(Para_List_Box_Info['Tray_Count'])+'" row_count="1" col_count="12" int_id=""/></params>'
    Form_Info_Tail = '<params><param key="pro_task_id" value=""/><param key="status" value="8"/><param key="photo" value="null"/><param key="isvirtual" value="0"/><param key="virtualtype" value=""/></params>'
    Form_Info_Encoded = "params=" + parse.quote_plus(Form_Info) + "&model=odm&" +  "lifeparams=" + parse.quote_plus(Form_Info_Tail)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Generate_Tray, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = parse.unquote(bytes(Response_Body.text, encoding="utf-8"))
    Response_Body = etree.HTML(Response_Body)
    List_Terminal_IDs = Response_Body.xpath('//terminal//@int_id')
    print('P5-Terminal_IDs_Count-{}'.format(len(List_Terminal_IDs)))
    Para_List_Box_Info['Terminal_IDs'] = '&'.join(List_Terminal_IDs)

if __name__ == '__main__':

    File_Process_and_Generate_Basic_Data()
    Swimming_Pool(Query_Box_ID_ResPont_ID_Alias, List_Box_Info)
    if P2_Generate_Support_Segment or P3_Generate_Cable_Segment or P4_Cable_Lay:
        Prepare_Support_Sys_and_Cable_Sys()
        Query_Project_Code_ID()
    if P5_Generate_ODM_and_Tray:
        Generate_Topology()
        Generate_FS_Info()

    if P1_Push_Box:
        print('P1-开始')
        Swimming_Pool(Push_Box, List_Box_Info)
        print('P1-结束')
    if P2_Generate_Support_Segment:
        print('P2-开始')
        Generate_Support_Segment()
        print('P2-结束')
    if P3_Generate_Cable_Segment:
        print('P3-开始')
        Generate_Cable_Segment()
        print('P3-结束')
    if P4_Cable_Lay:
        print('P4-开始')
        Swimming_Pool(Query_Support_Seg_ID_Cable_Seg_ID, List_CS_Info)
        Swimming_Pool(Cable_Lay, List_CS_Info)
        print('P4-结束')
    if P5_Generate_ODM_and_Tray:
        print('P5-开始')
        # Swimming_Pool(Generate_ODM, List_Box_Info)
        Swimming_Pool(Generate_Tray, List_Box_Info)
        print('P5-结束')


        # print(sorted(List_CS_Info[10].items(), key = lambda item:item[0]))
        # print(List_CS_Info[69])
        # print(List_Box_Info[20])
