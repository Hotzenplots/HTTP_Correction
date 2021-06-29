import concurrent.futures
import openpyxl
import urllib
import requests
import lxml
import json
import os
import copy

File_Name = ['平舆小黑黑']

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with concurrent.futures.ThreadPoolExecutor(max_workers=35) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Generate_Local_Data(Para_File_Name):

    WB_obj = openpyxl.load_workbook(Para_File_Name+'.xlsx') 

    WS_obj = WB_obj['Info']

    global Force_Query,P1_Push_Box

    Force_Query                 = WS_obj['B8'].value
    P1_Push_Box                 = WS_obj['E3'].value
    P2_Generate_Support_Segment = WS_obj['E4'].value
    
    if P1_Push_Box:

        WS_obj = WB_obj['Info']
        Longitude_Start     = WS_obj['B2'].value
        Latitude_Start      = WS_obj['B3'].value
        Horizontal_Density  = WS_obj['B4'].value
        Vertical_Density    = WS_obj['B5'].value
        Anchor_Point_Buttom = WS_obj['B6'].value
        Anchor_Point_Right  = WS_obj['B7'].value
        WS_obj = WB_obj['BOX_Topology']

        '''相对坐标'''
        List_Box_Name = []
        for row_num in range(1,201):
            for column_num in range(1,201):
                v = WS_obj.cell(row_num, column_num)
                if v.value == 'empty':
                    continue
                if v.value == None:
                    break
                List_Box_Name.append(list([v.value, row_num, column_num]))

        '''经纬度'''
        global List_Box_Data
        List_Box_Data = []
        for box_num in range(len(List_Box_Name)):
            List_Box_Data.append(dict({'Box_Name':List_Box_Name[box_num][0]}))
            List_Box_Data[box_num]['Longitude'] = Longitude_Start + (List_Box_Name[box_num][2] - 1) * Horizontal_Density
            List_Box_Data[box_num]['Latitude'] = Latitude_Start - (List_Box_Name[box_num][1] - 1) * Vertical_Density

        '''处理锚点'''
        if Anchor_Point_Buttom:
            Move_Up = List_Box_Name[-1][1]
            for each_box_data in List_Box_Data:
                each_box_data['Latitude'] = each_box_data['Latitude'] + Vertical_Density * (Move_Up - 1)
        if Anchor_Point_Right:
            Move_Left = List_Box_Name[-1][2]
            for each_box_data in List_Box_Data:
                each_box_data['Longitude'] = each_box_data['Longitude'] - Horizontal_Density * (Move_Left - 1)

        '''箱体类型'''
        List_Box_Type = [['guangfenxianxiang', 9204, 9115], ['guangjiaojiexiang', 9203, 9115]]
        for box_num in range(len(List_Box_Name)):
            if (List_Box_Data[box_num]['Box_Name'].find('GJ') != -1) or (List_Box_Data[box_num]['Box_Name'].find('gj') != -1):
                List_Box_Data[box_num]['Box_Type'] = 'guangjiaojiexiang'
            else:
                List_Box_Data[box_num]['Box_Type'] = 'guangfenxianxiang'
            for row_data in range(len(List_Box_Type)):
                if List_Box_Data[box_num]['Box_Type'] == List_Box_Type[row_data][0]:
                    List_Box_Data[box_num]['Box_Type_ID'] = List_Box_Type[row_data][1]
                    List_Box_Data[box_num]['ResPoint_Type_ID'] = List_Box_Type[row_data][2]

def SaSaSa_Save(Para_File_Name):

    List_Sorted_Box_Data = []
    for each_box_data in List_Box_Data:
        each_sorted_box_data = dict(sorted(each_box_data.items(), key = lambda item:item[0]))
        List_Sorted_Box_Data.append(each_sorted_box_data)
    JS_List_Box_Data = json.dumps(List_Sorted_Box_Data,ensure_ascii=False)
    with open(Para_File_Name+'Box.json', 'w', encoding='utf-8') as File_Box:
        File_Box.write(JS_List_Box_Data)

def LoLoLo_Load(Para_File_Name):

    if Force_Query:
        pass
    else:
        if os.path.isfile(Para_File_Name+'Box.json'):
            with open(Para_File_Name+'Box.json', 'r', encoding='utf-8') as File_Box:
                global List_Box_Data
                JS_Stream = File_Box.read()
                if len(JS_Stream) != 0:
                    List_Box_Data_Saved = json.loads(JS_Stream)
                    List_Box_Data = copy.deepcopy(List_Box_Data_Saved)

def Query_Box_Info(Para_List_Box_Data):
    URL_Query_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Para_List_Box_Data['Box_Type']+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_List_Box_Data['Box_Name']+'%\'" returnfields="INT_ID,ZH_LABEL,STRUCTURE_ID,ALIAS"/></request>'
    Form_Info_Encoded = 'xml='+urllib.parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Box, data=Form_Info_Encoded, headers=Request_Header)
    
    '''处理502错误,递归'''
    if Response_Body.status_code == 502: 
        Query_Box_Info(Para_List_Box_Data)

    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = lxml.etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    List_Response_Value_tv = Response_Body.xpath("//fv/@tv")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))

    if len(Dic_Response) == 0:
        print('<<<'+Para_List_Box_Data['Box_Name']+'>>>'+'箱体名字错误')
        exit()

    if P1_Push_Box:
        for box_num in range(len(List_Box_Data)):
            if Dic_Response['ZH_LABEL'] == List_Box_Data[box_num]['Box_Name']:
                List_Box_Data[box_num]['Box_ID'] = int(Dic_Response['INT_ID'])
                List_Box_Data[box_num]['ResPoint_ID'] = int(Dic_Response['STRUCTURE_ID'])
                List_Box_Data[box_num]['Alias'] = Dic_Response['ALIAS']
                List_Box_Data[box_num]['ResPoint_Name'] = List_Response_Value_tv[0]

def Execute_Push_Box():
    URL_Push_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=move'
    Form_Info_Head = '<xmldata mode="SinglePointEditMode"><mc type="ziyuandian">'
    Form_Info_Tail = '</mc></xmldata>'
    List_Form_Info_Body = []
    for each_box_data in List_Box_Data:
        Form_Info_Body = '<mo group="1"><fv k="LATITUDE" v="' + str(each_box_data['Latitude']) + '"/><fv k="INT_ID" v="' + str(each_box_data['ResPoint_ID']) + '"/><fv k="ZH_LABEL" v="' + each_box_data['ResPoint_Name'] + '"/><fv k="LONGITUDE" v="' + str(each_box_data['Longitude']) + '"/></mo>'
        List_Form_Info_Body.append(Form_Info_Body)
    Form_Info_Body = ''.join(List_Form_Info_Body)
    Form_Info = Form_Info_Head + Form_Info_Body + Form_Info_Tail
    Form_Info_Encoded = 'xml=' + urllib.parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {'Host': '10.209.199.74:8120','Content-Type': 'application/x-www-form-urlencoded','Content-Length': Request_Lenth}
    Response_Body = requests.post(URL_Push_Box, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = lxml.etree.HTML(Response_Body)
    Push_Result = Response_Body.xpath('//@loaded')
    print('P1-推箱子-{}'.format(Push_Result[0]))

def Main_Process(Para_File_Name):

    Generate_Local_Data(Para_File_Name)

    if P1_Push_Box:

        LoLoLo_Load(Para_File_Name)

        if Force_Query or ('Box_ID' not in List_Box_Data[0]):
            print('查询Box/ResPoint开始')
            Swimming_Pool(Query_Box_Info, List_Box_Data)
            print('查询Box/ResPoint结束')

        SaSaSa_Save(Para_File_Name)

        Execute_Push_Box()

if __name__ == '__main__':
    for each_File_Name in File_Name:
        Main_Process(each_File_Name)

