import csv
import requests
import re
from urllib import parse
from lxml import etree
from concurrent.futures import ThreadPoolExecutor

File_Name = '豪德置业光缆'

# Add_Redundant = True
Add_Redundant = False
# Clear_Redundant = True
Clear_Redundant = False

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Prepare_Data(Para_File_Name):
    with open(Para_File_Name+'.csv') as file_csv:
        reader_obj = csv.reader(file_csv)
        List_Cable_Data = list(reader_obj)
        global List_Cable_Name
        List_Cable_Name = []
        for each_cable_data in List_Cable_Data:
            if each_cable_data[2] == '光缆段名称':
                continue
            List_Cable_Name.append(each_cable_data[2])

def Cable_Add_Redundant(Para_Cable_Name):
    URL_Query_Cable_Data = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="guanglanduan" where="1=1 AND ZH_LABEL LIKE \'%'+Para_Cable_Name+'%\'" returnfields="INT_ID,ZH_LABEL,M_LENGTH"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Cable_Data, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Return_Cable_Data = Response_Body.xpath("//@v")

    URL_Query_Cable_Redundant = 'http://10.209.199.74:8120/igisserver_osl/rest/pipelineSection/getSupportByCab?int_id=' + str(List_Return_Cable_Data[0]) + '&queryType=glpl'
    Response_Body = requests.get(URL_Query_Cable_Redundant)
    Response_Body = Response_Body.text
    Response_Body = Response_Body.replace('cabcir_length=""', 'cabcir_length="30"')
    Response_Body = Response_Body.replace('isChange=""', 'isChange="0"')

    URL_Add_Cable_Redundant = 'http://10.209.199.74:8120/igisserver_osl/rest/pipelineSection/changeCabCir?queryType=glpl'
    Form_Info = Response_Body
    Form_Info_Encoded = 'dataXML='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Add_Cable_Redundant, data=Form_Info_Encoded, headers=Request_Header)

    URL_Query_Cable_Data = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info_Encoded = "xml="+parse.quote_plus('''<request><query mc="guanglanduan" ids="" where="1=1 AND ZH_LABEL LIKE '%'''+Para_Cable_Name+'''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,CITY_ID,COUNTY_ID,STATUS,PRO_TASK_ID,CJ_TASK_ID,A_EQUIP_TYPE,CJ_STATUS,A_EQUIP_ID,A_ROOM_ID,A_OBJECT_TYPE,A_OBJECT_ID,Z_EQUIP_TYPE,Z_EQUIP_ID,Z_ROOM_ID,Z_OBJECT_TYPE,Z_OBJECT_ID,SERVICE_LEVEL,C_LENGTH,M_LENGTH,FIBER_NUM,RELATED_SYSTEM,WIRE_SEG_TYPE,DIRECTION,IS_ALTER,FIBER_TYPE,DIA,GT_VERSION,VENDOR,OPTI_CABLE_TYPE,MAINT_DEP,MAINT_MODE,LAY_TYPE,RELATED_IS_AREA,IS_WRONG,MAINT_STATE,AREA_LEVEL,WRONG_INFO,SYS_VERSION,OWNERSHIP,PHONE_NO,TIME_STAMP,RES_OWNER,PRODUCE_DATE,PROJECTCODE,BUSINESS,TASK_NAME,ASSENT_NO,SERVICER,RUWANG_DATE,STUFF,BUILDER,TUIWANG_DATE,QUALITOR_PROJECT,PROJECT_ID,QUALITOR,QUALITOR_COUNTY,MAINTAINOR,QRCODE,SEG_NO,REMARK,INDEX_IN_BRANCH,IS_COMPLETE_LAY,MAINTAIN_CITY,MAINTAIN_COUNTY,CREATOR,CREAT_TIME,MODIFIER,MODIFY_TIME,STATEFLAG,BUILD_DATE,PROJECT_NAME,USER_NAME,PURPOSE,PROJECT,YJTZ,JSLY,JSYXJ,SQJSNF"/></request>''')
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Cable_Data, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))

    URL_Renew_Cable = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update&coreNamingRules=0"
    Form_Info_Encoded = "xml="+parse.quote_plus('<xmldata mode="FibersegEditMode"><mc type="guanglanduan"><mo group="1"><fv k="INT_ID" v="'+Dic_Response["INT_ID"]+'"/><fv k="FIBER_TYPE" v="2"/><fv k="COUNTY_ID" v="'+Dic_Response["COUNTY_ID"]+'"/><fv k="STATUS" v="8"/><fv k="OWNERSHIP" v="1"/><fv k="SYS_VERSION" v="外线v3.0.0-核心模型v4.0"/><fv k="PROJECTCODE" v="'+Dic_Response["PROJECTCODE"]+'"/><fv k="A_ROOM_ID" v="'+Dic_Response["A_ROOM_ID"]+'"/><fv k="Z_ROOM_ID" v="'+Dic_Response["Z_ROOM_ID"]+'"/><fv k="ASSENT_NO" v="'+Dic_Response["ASSENT_NO"]+'"/><fv k="QUALITOR_PROJECT" v="'+Dic_Response["QUALITOR_PROJECT"]+'"/><fv k="TASK_NAME" v="'+Dic_Response["TASK_NAME"]+'"/><fv k="IS_ALTER" v="0"/><fv k="OPTI_CABLE_TYPE" v="1"/><fv k="ZH_LABEL" v="'+Dic_Response["ZH_LABEL"]+'"/><fv k="LAY_TYPE" v="6"/><fv k="CITY_ID" v="'+Dic_Response["CITY_ID"]+'"/><fv k="RELATED_SYSTEM" v="'+Dic_Response["RELATED_SYSTEM"]+'"/><fv k="IS_COMPLETE_LAY" v="1"/><fv k="A_OBJECT_ID" v="'+Dic_Response["A_OBJECT_ID"]+'"/><fv k="M_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_OBJECT_TYPE" v="'+Dic_Response["A_OBJECT_TYPE"]+'"/><fv k="FIBER_NUM" v="'+Dic_Response["FIBER_NUM"]+'"/><fv k="Z_OBJECT_ID" v="'+Dic_Response["Z_OBJECT_ID"]+'"/><fv k="QUALITOR_COUNTY" v="'+Dic_Response["QUALITOR_COUNTY"]+'"/><fv k="Z_OBJECT_TYPE" v="'+Dic_Response["Z_OBJECT_TYPE"]+'"/><fv k="MAINTAINOR" v="'+Dic_Response["MAINTAINOR"]+'"/><fv k="Z_EQUIP_TYPE" v="'+Dic_Response["Z_EQUIP_TYPE"]+'"/><fv k="C_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_EQUIP_ID" v="'+Dic_Response["A_EQUIP_ID"]+'"/><fv k="A_EQUIP_TYPE" v="'+Dic_Response["A_EQUIP_TYPE"]+'"/><fv k="QRCODE" v="'+Dic_Response["QRCODE"]+'"/><fv k="Z_EQUIP_ID" v="'+Dic_Response["Z_EQUIP_ID"]+'"/><fv k="QUALITOR" v="'+Dic_Response["QUALITOR"]+'"/><fv k="WIRE_SEG_TYPE" v="'+Dic_Response["WIRE_SEG_TYPE"]+'"/><fv k="SERVICE_LEVEL" v="'+Dic_Response["SERVICE_LEVEL"]+'"/><fv k="RES_OWNER" v="0"/><fv k="VENDOR" v="烽火"/></mo></mc></xmldata>')
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Renew_Cable, data=Form_Info_Encoded, headers=Request_Header)

def Cable_Clear_Redundant(Para_Cable_Name):
    URL_Query_Cable_Data = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="guanglanduan" where="1=1 AND ZH_LABEL LIKE \'%'+Para_Cable_Name+'%\'" returnfields="INT_ID,ZH_LABEL,M_LENGTH"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Cable_Data, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Return_Cable_Data = Response_Body.xpath("//@v")

    URL_Query_Cable_Redundant = 'http://10.209.199.74:8120/igisserver_osl/rest/pipelineSection/getSupportByCab?int_id=' + str(List_Return_Cable_Data[0]) + '&queryType=glpl'
    Response_Body = requests.get(URL_Query_Cable_Redundant)
    Response_Body = Response_Body.text
    Response_Body = re.sub(r'cabcir_length="[\d]*"','cabcir_length=""', Response_Body)
    Response_Body = Response_Body.replace('isChange=""', 'isChange="0"')

    URL_Clear_Cable_Redundant = 'http://10.209.199.74:8120/igisserver_osl/rest/pipelineSection/changeCabCir?queryType=glpl'
    Form_Info = Response_Body
    Form_Info_Encoded = 'dataXML='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Clear_Cable_Redundant, data=Form_Info_Encoded, headers=Request_Header)

    URL_Query_Cable_Data = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info_Encoded = "xml="+parse.quote_plus('''<request><query mc="guanglanduan" ids="" where="1=1 AND ZH_LABEL LIKE '%'''+Para_Cable_Name+'''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,CITY_ID,COUNTY_ID,STATUS,PRO_TASK_ID,CJ_TASK_ID,A_EQUIP_TYPE,CJ_STATUS,A_EQUIP_ID,A_ROOM_ID,A_OBJECT_TYPE,A_OBJECT_ID,Z_EQUIP_TYPE,Z_EQUIP_ID,Z_ROOM_ID,Z_OBJECT_TYPE,Z_OBJECT_ID,SERVICE_LEVEL,C_LENGTH,M_LENGTH,FIBER_NUM,RELATED_SYSTEM,WIRE_SEG_TYPE,DIRECTION,IS_ALTER,FIBER_TYPE,DIA,GT_VERSION,VENDOR,OPTI_CABLE_TYPE,MAINT_DEP,MAINT_MODE,LAY_TYPE,RELATED_IS_AREA,IS_WRONG,MAINT_STATE,AREA_LEVEL,WRONG_INFO,SYS_VERSION,OWNERSHIP,PHONE_NO,TIME_STAMP,RES_OWNER,PRODUCE_DATE,PROJECTCODE,BUSINESS,TASK_NAME,ASSENT_NO,SERVICER,RUWANG_DATE,STUFF,BUILDER,TUIWANG_DATE,QUALITOR_PROJECT,PROJECT_ID,QUALITOR,QUALITOR_COUNTY,MAINTAINOR,QRCODE,SEG_NO,REMARK,INDEX_IN_BRANCH,IS_COMPLETE_LAY,MAINTAIN_CITY,MAINTAIN_COUNTY,CREATOR,CREAT_TIME,MODIFIER,MODIFY_TIME,STATEFLAG,BUILD_DATE,PROJECT_NAME,USER_NAME,PURPOSE,PROJECT,YJTZ,JSLY,JSYXJ,SQJSNF"/></request>''')
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Cable_Data, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))

    URL_Renew_Cable = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update&coreNamingRules=0"
    Form_Info_Encoded = "xml="+parse.quote_plus('<xmldata mode="FibersegEditMode"><mc type="guanglanduan"><mo group="1"><fv k="INT_ID" v="'+Dic_Response["INT_ID"]+'"/><fv k="FIBER_TYPE" v="2"/><fv k="COUNTY_ID" v="'+Dic_Response["COUNTY_ID"]+'"/><fv k="STATUS" v="8"/><fv k="OWNERSHIP" v="1"/><fv k="SYS_VERSION" v="外线v3.0.0-核心模型v4.0"/><fv k="PROJECTCODE" v="'+Dic_Response["PROJECTCODE"]+'"/><fv k="A_ROOM_ID" v="'+Dic_Response["A_ROOM_ID"]+'"/><fv k="Z_ROOM_ID" v="'+Dic_Response["Z_ROOM_ID"]+'"/><fv k="ASSENT_NO" v="'+Dic_Response["ASSENT_NO"]+'"/><fv k="QUALITOR_PROJECT" v="'+Dic_Response["QUALITOR_PROJECT"]+'"/><fv k="TASK_NAME" v="'+Dic_Response["TASK_NAME"]+'"/><fv k="IS_ALTER" v="0"/><fv k="OPTI_CABLE_TYPE" v="1"/><fv k="ZH_LABEL" v="'+Dic_Response["ZH_LABEL"]+'"/><fv k="LAY_TYPE" v="6"/><fv k="CITY_ID" v="'+Dic_Response["CITY_ID"]+'"/><fv k="RELATED_SYSTEM" v="'+Dic_Response["RELATED_SYSTEM"]+'"/><fv k="IS_COMPLETE_LAY" v="1"/><fv k="A_OBJECT_ID" v="'+Dic_Response["A_OBJECT_ID"]+'"/><fv k="M_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_OBJECT_TYPE" v="'+Dic_Response["A_OBJECT_TYPE"]+'"/><fv k="FIBER_NUM" v="'+Dic_Response["FIBER_NUM"]+'"/><fv k="Z_OBJECT_ID" v="'+Dic_Response["Z_OBJECT_ID"]+'"/><fv k="QUALITOR_COUNTY" v="'+Dic_Response["QUALITOR_COUNTY"]+'"/><fv k="Z_OBJECT_TYPE" v="'+Dic_Response["Z_OBJECT_TYPE"]+'"/><fv k="MAINTAINOR" v="'+Dic_Response["MAINTAINOR"]+'"/><fv k="Z_EQUIP_TYPE" v="'+Dic_Response["Z_EQUIP_TYPE"]+'"/><fv k="C_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_EQUIP_ID" v="'+Dic_Response["A_EQUIP_ID"]+'"/><fv k="A_EQUIP_TYPE" v="'+Dic_Response["A_EQUIP_TYPE"]+'"/><fv k="QRCODE" v="'+Dic_Response["QRCODE"]+'"/><fv k="Z_EQUIP_ID" v="'+Dic_Response["Z_EQUIP_ID"]+'"/><fv k="QUALITOR" v="'+Dic_Response["QUALITOR"]+'"/><fv k="WIRE_SEG_TYPE" v="'+Dic_Response["WIRE_SEG_TYPE"]+'"/><fv k="SERVICE_LEVEL" v="'+Dic_Response["SERVICE_LEVEL"]+'"/><fv k="RES_OWNER" v="0"/><fv k="VENDOR" v="烽火"/></mo></mc></xmldata>')
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Renew_Cable, data=Form_Info_Encoded, headers=Request_Header)

if Add_Redundant:
    Prepare_Data(File_Name)
    Swimming_Pool(Cable_Add_Redundant,List_Cable_Name)

if Clear_Redundant:
    Prepare_Data(File_Name)
    Swimming_Pool(Cable_Clear_Redundant,List_Cable_Name)
