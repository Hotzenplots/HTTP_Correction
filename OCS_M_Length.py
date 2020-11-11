# -*- coding: UTF-8 -*-
import csv
import urllib.parse
import requests
from lxml import etree

with open("8017OpticalCable.csv") as file_csv:
    reader_obj = csv.reader(file_csv)
    CSV_List = list(reader_obj)
    file_csv.close()
    
Renew_Count = 0
# for RL1 in list(range(len(CSV_List))):
for RL1 in list(range(2)):
    if RL1 == 0:
        continue
    if CSV_List[RL1][20] != CSV_List[RL1][21]:
        Dic_Response = []
        URL_Query_OpticalCable = "http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet"
        XML_Info_Encoded = "xml="+urllib.parse.quote('''<request><query mc="guanglanduan" ids="" where="1=1 AND ZH_LABEL LIKE '%'''+CSV_List[RL1][2]+'''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,CITY_ID,COUNTY_ID,STATUS,PRO_TASK_ID,CJ_TASK_ID,A_EQUIP_TYPE,CJ_STATUS,A_EQUIP_ID,A_ROOM_ID,A_OBJECT_TYPE,A_OBJECT_ID,Z_EQUIP_TYPE,Z_EQUIP_ID,Z_ROOM_ID,Z_OBJECT_TYPE,Z_OBJECT_ID,SERVICE_LEVEL,C_LENGTH,M_LENGTH,FIBER_NUM,RELATED_SYSTEM,WIRE_SEG_TYPE,DIRECTION,IS_ALTER,FIBER_TYPE,DIA,GT_VERSION,VENDOR,OPTI_CABLE_TYPE,MAINT_DEP,MAINT_MODE,LAY_TYPE,RELATED_IS_AREA,IS_WRONG,MAINT_STATE,AREA_LEVEL,WRONG_INFO,SYS_VERSION,OWNERSHIP,PHONE_NO,TIME_STAMP,RES_OWNER,PRODUCE_DATE,PROJECTCODE,BUSINESS,TASK_NAME,ASSENT_NO,SERVICER,RUWANG_DATE,STUFF,BUILDER,TUIWANG_DATE,QUALITOR_PROJECT,PROJECT_ID,QUALITOR,QUALITOR_COUNTY,MAINTAINOR,QRCODE,SEG_NO,REMARK,INDEX_IN_BRANCH,IS_COMPLETE_LAY,MAINTAIN_CITY,MAINTAIN_COUNTY,CREATOR,CREAT_TIME,MODIFIER,MODIFY_TIME,STATEFLAG,BUILD_DATE,PROJECT_NAME,USER_NAME,PURPOSE,PROJECT,YJTZ,JSLY,JSYXJ,SQJSNF"/></request>''')
        XML_Info_Encoded = XML_Info_Encoded.replace("/","%2f")
        Request_Lenth = chr(len(XML_Info_Encoded))
        HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
        Respond_Body = requests.post(URL_Query_OpticalCable,data=XML_Info_Encoded,headers=HTTP_Header)
        Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
        Respond_Body = etree.HTML(Respond_Body)
        List_Response_Key = Respond_Body.xpath("//fv/@k")
        List_Response_Value = Respond_Body.xpath("//fv/@v")
        Dic_Response = dict(zip(List_Response_Key,List_Response_Value))

        URL_Renew_OpticalCable = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update&coreNamingRules=0"
        XML_Info_Encoded = "xml="+urllib.parse.quote('<xmldata mode="FibersegEditMode"><mc type="guanglanduan"><mo group="1"><fv k="INT_ID" v="'+Dic_Response["INT_ID"]+'"/><fv k="FIBER_TYPE" v="2"/><fv k="COUNTY_ID" v="'+Dic_Response["COUNTY_ID"]+'"/><fv k="STATUS" v="8"/><fv k="OWNERSHIP" v="1"/><fv k="SYS_VERSION" v="外线v3.0.0-核心模型v4.0"/><fv k="PROJECTCODE" v="'+Dic_Response["PROJECTCODE"]+'"/><fv k="A_ROOM_ID" v="'+Dic_Response["A_ROOM_ID"]+'"/><fv k="Z_ROOM_ID" v="'+Dic_Response["Z_ROOM_ID"]+'"/><fv k="ASSENT_NO" v="'+Dic_Response["ASSENT_NO"]+'"/><fv k="QUALITOR_PROJECT" v="'+Dic_Response["QUALITOR_PROJECT"]+'"/><fv k="TASK_NAME" v="'+Dic_Response["TASK_NAME"]+'"/><fv k="IS_ALTER" v="0"/><fv k="OPTI_CABLE_TYPE" v="1"/><fv k="ZH_LABEL" v="'+Dic_Response["ZH_LABEL"]+'"/><fv k="LAY_TYPE" v="6"/><fv k="CITY_ID" v="'+Dic_Response["CITY_ID"]+'"/><fv k="RELATED_SYSTEM" v="'+Dic_Response["RELATED_SYSTEM"]+'"/><fv k="IS_COMPLETE_LAY" v="1"/><fv k="A_OBJECT_ID" v="'+Dic_Response["A_OBJECT_ID"]+'"/><fv k="M_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_OBJECT_TYPE" v="'+Dic_Response["A_OBJECT_TYPE"]+'"/><fv k="FIBER_NUM" v="'+Dic_Response["FIBER_NUM"]+'"/><fv k="Z_OBJECT_ID" v="'+Dic_Response["Z_OBJECT_ID"]+'"/><fv k="QUALITOR_COUNTY" v="'+Dic_Response["QUALITOR_COUNTY"]+'"/><fv k="Z_OBJECT_TYPE" v="'+Dic_Response["Z_OBJECT_TYPE"]+'"/><fv k="MAINTAINOR" v="'+Dic_Response["MAINTAINOR"]+'"/><fv k="Z_EQUIP_TYPE" v="'+Dic_Response["Z_EQUIP_TYPE"]+'"/><fv k="C_LENGTH" v="'+Dic_Response["C_LENGTH"]+'"/><fv k="A_EQUIP_ID" v="'+Dic_Response["A_EQUIP_ID"]+'"/><fv k="A_EQUIP_TYPE" v="'+Dic_Response["A_EQUIP_TYPE"]+'"/><fv k="QRCODE" v="'+Dic_Response["QRCODE"]+'"/><fv k="Z_EQUIP_ID" v="'+Dic_Response["Z_EQUIP_ID"]+'"/><fv k="QUALITOR" v="'+Dic_Response["QUALITOR"]+'"/><fv k="WIRE_SEG_TYPE" v="'+Dic_Response["WIRE_SEG_TYPE"]+'"/><fv k="SERVICE_LEVEL" v="'+Dic_Response["SERVICE_LEVEL"]+'"/><fv k="RES_OWNER" v="0"/><fv k="VENDOR" v="烽火"/></mo></mc></xmldata>')
        XML_Info_Encoded = XML_Info_Encoded.replace("/","%2f")
        Request_Lenth = chr(len(XML_Info_Encoded))
        HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
        Respond_Body = requests.post(URL_Renew_OpticalCable,data=XML_Info_Encoded,headers=HTTP_Header)
        Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
        Respond_Body = etree.HTML(Respond_Body)
        Renew_State = Respond_Body.xpath("//@loaded")
        Renew_Count = Renew_Count + 1
        print("%d" % Renew_Count + " " + Dic_Response["ZH_LABEL"] + " " + Renew_State[0])