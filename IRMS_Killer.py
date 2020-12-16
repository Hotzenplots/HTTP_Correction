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

P1_Push_Box                     = False
P2_Draw_Support_Sys_and_Segment = True

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
    cell_range = WS_obj['A2': 'R11']
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
        for row_data in range(len(List_Template)):
            if List_7013[1][1] == List_Template[row_data][2]:
                List_Box_Info[box_num]['Box_City_ID'] = List_Template[row_data][1]
                List_Box_Info[box_num]['Box_County_ID'] = List_Template[row_data][3]

    '''读取并整理Sheet_OCS_List,生成List_OCS_Info'''
    WS_obj = WB_obj['OCS_List']
    global List_OCS_Info
    List_OCS_Info = []
    for row_num in range(1,501):
        OCS_A_Box_Name = WS_obj.cell(row_num, 1).value
        OCS_Z_Box_Name = WS_obj.cell(row_num, 2).value
        OCS_Width = WS_obj.cell(row_num, 3).value
        if OCS_A_Box_Name == None:
            break
        List_OCS_Info.append(dict({'A_Box_Name': OCS_A_Box_Name}))
        List_OCS_Info[row_num - 1]['Z_Box_Name'] = OCS_Z_Box_Name
        List_OCS_Info[row_num - 1]['Width'] = OCS_Width
    
    #Length_Prepare
    Horizontal_Metre = 111.11 * 1000 * math.cos(Latitude_Start * math.pi / 180)
    Vertical_Metre = 111.11 * 1000
    
    for dic_num_in_osc in List_OCS_Info:
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
    List_Respond_Key = Response_Body.xpath("//@name")
    List_Respond_Value = Response_Body.xpath("//@value")
    Dic_Respond = dict(zip(List_Respond_Key,List_Respond_Value))
    Dic_Project_Code_ID = copy.deepcopy(Dic_Respond)
    for box_num in List_Box_Info:
        box_num['Project_Code_ID'] = Dic_Project_Code_ID[List_7013[1][4]]
    for ocs_num in List_OCS_Info:
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
    List_Respond_Key = Response_Body.xpath("//fv/@k")
    List_Respond_Value = Response_Body.xpath("//fv/@v")
    Dic_Respond = dict(zip(List_Respond_Key,List_Respond_Value))
    for box_num in range(len(List_Box_Info)):
        if Dic_Respond['ZH_LABEL'] == List_Box_Info[box_num]['Box_Name']:
            List_Box_Info[box_num]['Box_ID'] = Dic_Respond['INT_ID']
            List_Box_Info[box_num]['ResPoint_ID'] = Dic_Respond['STRUCTURE_ID']
            List_Box_Info[box_num]['Alias'] = Dic_Respond['ALIAS']
    for ocs_num in List_OCS_Info:
        for box_num in List_Box_Info:
            if ocs_num['A_Box_Name'] == box_num['Box_Name']:
                ocs_num['A_Box_ID'] = box_num['Box_ID']
                ocs_num['A_ResPoint_ID'] = box_num['ResPoint_ID']
        for box_num in List_Box_Info:
            if ocs_num['Z_Box_Name'] == box_num['Box_Name']:
                ocs_num['Z_Box_ID'] = box_num['Box_ID']
                ocs_num['Z_ResPoint_ID'] = box_num['ResPoint_ID']

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Push_Box(Para_List_Box_Info):
    URL_Push_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/SynchroController/synchroData?sid=' + str(Para_List_Box_Info['Box_ID']) + '&sType=' + str(Para_List_Box_Info['Box_Type']) + '&longi='+str(Para_List_Box_Info['Longitude']) + '&lati=' + str(Para_List_Box_Info['Latitude'])
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded"}
    Response_Body = requests.post(URL_Push_Box, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    Push_Result = Response_Body.xpath('//message/text()')
    print('P1-{}-{}'.format(Para_List_Box_Info['Box_Name'], Push_Result[0]))

def Add_Support_Segment():
    URL_Add_Support_Segment = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesAdd'
    Form_Info_Head = '<xmldata mode="PipeLineAddMode"><mc type="yinshangduan">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for ocs_num in List_OCS_Info:
        Form_Info_Body = '<mo group="1" ax="'+str(ocs_num['A_Longitude'])+'" ay="'+str(ocs_num['A_Latitude'])+'" zx="'+str(ocs_num['Z_Longitude'])+'" zy="'+str(ocs_num['Z_Latitude'])+'"><fv k="CITY_ID" v="'+str(ocs_num['City_ID'])+'"/><fv k="COUNTY_ID" v="'+str(ocs_num['County_ID'])+'"/><fv k="RELATED_SYSTEM" v="'+str(ocs_num['Support_Sys_ID'])+'"/><fv k="A_OBJECT_ID" v="'+str(ocs_num['A_ResPoint_ID'])+'"/><fv k="ZH_LABEL" v="'+str(ocs_num['A_Box_Name'])+'资源点-'+str(ocs_num['Z_Box_Name'])+'资源点引上段'+'"/><fv k="MAINTAINOR" v="'+str(ocs_num['DQS_Maintainer_ID'])+'"/><fv k="M_LENGTH" v="'+str(ocs_num['Length'])+'"/><fv k="STATUS" v="'+str(ocs_num['Life_Cycle'])+'"/><fv k="QUALITOR_COUNTY" v="'+str(ocs_num['DQS_County_ID'])+'"/><fv k="QUALITOR" v="'+str(ocs_num['DQS_ID'])+'"/><fv k="QUALITOR_PROJECT" v="'+str(ocs_num['DQS_Project_ID'])+'"/><fv k="OWNERSHIP" v="'+str(ocs_num['Owner_Type'])+'"/><fv k="RES_OWNER" v="'+str(ocs_num['Owner_Name'])+'"/><fv k="A_OBJECT_TYPE" v="'+str(ocs_num['A_ResPoint_Type_ID'])+'"/><fv k="INT_ID" v="new-27991311"/><fv k="TASK_NAME" v="'+str(ocs_num['Task_Name_ID'])+'"/><fv k="PROJECT_CODE" v="'+str(ocs_num['Project_Code_ID'])+'"/><fv k="RESOURCE_LOCATION" v="'+str(ocs_num['Field_Type'])+'"/><fv k="Z_OBJECT_TYPE" v="'+str(ocs_num['Z_ResPoint_Type_ID'])+'"/><fv k="Z_OBJECT_ID" v="'+str(ocs_num['Z_ResPoint_ID'])+'"/><fv k="SYSTEM_LEVEL" v="'+str(ocs_num['Business_Level'])+'"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    # Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info_Body = List_Form_Info_Body[52]+List_Form_Info_Body[53]+List_Form_Info_Body[54]
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = chr(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Add_Support_Segment, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Add_Support_Segment_State = Response_Body.xpath('//@loaded')
    print('P2-引上段建立-{}'.format(List_Add_Support_Segment_State[0]))
    # print(Form_Info_Body)

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
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_OCS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_OCS_Info[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = chr(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Support_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Support_Sys_ID = Response_Body.xpath('//@newid')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_OCS_Info:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('P2-引上系统ID-{}'.format(Support_Sys_ID))

    elif List_SS_Count[0] == '1':
        #Get Support_Sys_ID
        Support_Sys_ID = Response_Body.xpath('//@int_id')
        Support_Sys_ID = Support_Sys_ID[0]
        for ocs_num in List_OCS_Info:
            ocs_num['Support_Sys_ID'] = Support_Sys_ID
        print('P2-引上系统ID-{}'.format(Support_Sys_ID))

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
        Form_Info = '<response mode="add"><mc type="xitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_OCS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_OCS_Info[0]['County_ID'])+'"/><fv k="LINESEG_TYPE" v="9108"/><fv k="ZH_LABEL" v="'+Support_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/></mo></mc></response>'
        Form_Info = '<response mode="add"><mc type="guanglanxitong"><mo group="1"><fv k="SYSTEM_LEVEL" v="8"/><fv k="CITY_ID" v="'+str(List_OCS_Info[0]['City_ID'])+'"/><fv k="STATUS" v="8"/><fv k="COUNTY_ID" v="'+str(List_OCS_Info[0]['County_ID'])+'"/><fv k="ZH_LABEL" v="'+Cable_Sys_Name+'"/><fv k="INT_ID" v="new-27991311"/><fv k="OPT_TYPE" v="GYTA-12"/></mo></mc></response>'
        Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
        Request_Lenth = chr(len(Form_Info_Encoded))
        Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
        Response_Body = requests.post(URL_Add_Cable_Sys, data=Form_Info_Encoded, headers=Request_Header)
        Response_Body = bytes(Response_Body.text, encoding="utf-8")
        Response_Body = etree.HTML(Response_Body)
        Cable_Sys_ID = Response_Body.xpath('//@newid')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_OCS_Info:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('P2-光缆系统ID-{}'.format(Cable_Sys_ID))

    elif List_CS_Count[0] == '1':
        #Get Cable_Sys_ID
        Cable_Sys_ID = Response_Body.xpath('//@int_id')
        Cable_Sys_ID = Cable_Sys_ID[0]
        for ocs_num in List_OCS_Info:
            ocs_num['Cable_Sys_ID'] = Cable_Sys_ID
        print('P2-光缆系统ID-{}'.format(Cable_Sys_ID))

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

if __name__ == '__main__':
    File_Process_and_Generate_Basic_Data()
    if P1_Push_Box:
        print('P1-开始')
        Swimming_Pool(Query_Box_ID_ResPont_ID_Alias, List_Box_Info)
        Swimming_Pool(Push_Box, List_Box_Info)
        print('P1-结束')
    if P2_Draw_Support_Sys_and_Segment:
        print('P2-开始')
        if not P1_Push_Box:
            Swimming_Pool(Query_Box_ID_ResPont_ID_Alias, List_Box_Info)
        Prepare_Support_Sys_and_Cable_Sys()
        Query_Project_Code_ID()
        Add_Support_Segment()
        print('P2-结束')
        # print(sorted(List_OCS_Info[10].items(), key = lambda item:item[0]))
        # print(List_OCS_Info[69])
        # print(List_Box_Info[20])
