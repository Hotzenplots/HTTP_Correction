# -*- coding: UTF-8 -*-
import csv
import urllib.parse
import requests
from lxml import etree

def Renew_SFB():
    with open("SFB8002.csv") as file_csv:
        reader_obj = csv.reader(file_csv)
        List_CSV = list(reader_obj)
        file_csv.close()

    Renew_Count = 0
    for RL1 in list(range(len(List_CSV))):
        if RL1 == 0:
            continue
        List_CSV[RL1][2] = List_CSV[RL1][2].replace("\n","")
        Dic_Response = []
        URL_Query_SFB = "http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet"
        XML_Info_Encoded = "xml=" + urllib.parse.quote('''<request><query mc="guangfenxianxiang" where="1=1 AND ZH_LABEL LIKE '%'''+List_CSV[RL1][2]+'''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,LONGITUDE,LATITUDE,CITY_ID,COUNTY_ID,STATUS,MAINTAIN_CITY,PRO_TASK_ID,MAINTAIN_COUNTY,CJ_TASK_ID,CJ_STATUS,STRUCTURE_TYPE,RELATED_IS_AREA,STRUCTURE_ID,SERVICE_LEVEL,ROOM_ID,CREATOR,IS_USAGE_STATE,CREAT_TIME,DESIGN_CAPACITY,FIBERDP_NO,MODIFIER,MODIFY_TIME,FREE_CAPACITY,IS_WRONG,GT_VERSION,WRONG_INFO,CREATTIME,SYS_VERSION,INSTALL_CAPACITY,TIME_STAMP,MAINT_DEP,STATEFLAG,MAINT_MODE,PROJECT_ID,MAINT_STATE,LABEL_DEV,LAND_HEIGHT,LOCATION,MOD_NUM,OWNERSHIP,MODEL,PRESERVER,PRESERVER_ADDR,PRESERVER_PHONE,RES_OWNER,SERVICER,TERMINAL_NUM,TIER_COL_COUNT,TIER_ROW_COUNT,USED_CAPACITY,IS_SF_POINT,USER1NAME,SF_POINT_LEV,REMARK,BUILD_DATE,RUWANG_DATE,JIAOWEI_DATE,RELATED_MG_AREA,BUS_AREAR_ID,BUILD_AREA_ID,DEVICE_VENDOR,DETAIL_ADDRESS,FACE_COUNT,EACH_SIDE_ROWS_NUM,EACH_SIDE_COL_NUM,FIBER_COUNT,FIBER_COUNT_FREE,QUALITOR,QUALITOR_COUNTY,QUALITOR_PROJECT,ZICHAN_BELONG,INBOX_NE_TYPE,UP_OPTI_TRANS_BOX,MAINTAINOR,SBXH,FULL_ADDR,TOWN,IS_5G,ROAD_ID,XQRL,JSYXJ,VILLAGE_ID,FLOOR_ID,SQJSNF,GJJB,UNIT_ID,LAYER_ID,ROOM_NUN_ID,SPEC_TYPE,PROJECTCODE,TASK_NAME,ROWNO,COLUMNNO,MAINTAIN_DEPARTMENT,PURPOSE"/></request>''')
        XML_Info_Encoded = XML_Info_Encoded.replace("/","%2f")
        Request_Lenth = chr(len(XML_Info_Encoded))
        HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
        Respond_Body = requests.post(URL_Query_SFB,data=XML_Info_Encoded,headers=HTTP_Header)
        Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
        Respond_Body = etree.HTML(Respond_Body)
        List_Response_Key = Respond_Body.xpath("//fv/@k")
        List_Response_Value = Respond_Body.xpath("//fv/@v")
        Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
        URL_Renew_SFB = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update"
        XML_Info_Encoded = "xml=" + urllib.parse.quote('<xmldata mode="OptoeleEditMode"><mc type="ziyuandian"><mo group="1"><fv k="INT_ID" v="'+Dic_Response["INT_ID"]+'"/><fv k="LONGITUDE" v="'+Dic_Response["LONGITUDE"]+'"/><fv k="LATITUDE" v="'+Dic_Response["LATITUDE"]+'"/></mo></mc><mc type="guangfenxianxiang"><mo group="1"><fv k="SYS_VERSION" v="外线v3.0.0-核心模型v4.0"/><fv k="TIER_COL_COUNT" v="0"/><fv k="QUALITOR_COUNTY" v="'+Dic_Response["QUALITOR_COUNTY"]+'"/><fv k="QUALITOR_PROJECT" v="'+Dic_Response["QUALITOR_PROJECT"]+'"/><fv k="MAINTAINOR" v="'+Dic_Response["MAINTAINOR"]+'"/><fv k="STATUS" v="8"/><fv k="USED_CAPACITY" v="0"/><fv k="ZICHAN_BELONG" v="1"/><fv k="STRUCTURE_TYPE" v="'+Dic_Response["STRUCTURE_TYPE"]+'"/><fv k="UP_OPTI_TRANS_BOX" v="0"/><fv k="ROWNO" v="0"/><fv k="STRUCTURE_ID" v="'+Dic_Response["STRUCTURE_ID"]+'"/><fv k="OWNERSHIP" v="1"/><fv k="MOD_NUM" v="0"/><fv k="QUALITOR" v="'+Dic_Response["QUALITOR"]+'"/><fv k="FIBER_COUNT_FREE" v="0"/><fv k="EACH_SIDE_ROWS_NUM" v="0"/><fv k="ROAD_ID" v="'+Dic_Response["ROAD_ID"]+'"/><fv k="CREATTIME" v="2018-11-21"/><fv k="LONGITUDE" v="'+Dic_Response["LONGITUDE"]+'"/><fv k="VILLAGE_ID" v="'+Dic_Response["VILLAGE_ID"]+'"/><fv k="INBOX_NE_TYPE" v="1"/><fv k="TOWN" v="'+Dic_Response["TOWN"]+'"/><fv k="FIBER_COUNT" v="0"/><fv k="LATITUDE" v="'+Dic_Response["LATITUDE"]+'"/><fv k="TERMINAL_NUM" v="0"/><fv k="RUWANG_DATE" v=""/><fv k="SPEC_TYPE" v="3"/><fv k="CITY_ID" v="445835190"/><fv k="PURPOSE" v="0"/><fv k="COUNTY_ID" v="'+Dic_Response["COUNTY_ID"]+'"/><fv k="TIER_ROW_COUNT" v="0"/><fv k="ZH_LABEL" v="'+Dic_Response["ZH_LABEL"]+'"/><fv k="RES_OWNER" v="0"/><fv k="INT_ID" v="'+Dic_Response["INT_ID"]+'"/><fv k="PROJECTCODE" v="2929"/><fv k="TASK_NAME" v="1111111"/><fv k="EACH_SIDE_COL_NUM" v="0"/></mo></mc></xmldata>')
        XML_Info_Encoded = XML_Info_Encoded.replace("/","%2f")
        Request_Lenth = chr(len(XML_Info_Encoded))
        HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
        Respond_Body = requests.post(URL_Renew_SFB,data=XML_Info_Encoded,headers=HTTP_Header)
        Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
        Respond_Body = etree.HTML(Respond_Body)
        Renew_State = Respond_Body.xpath("//@loaded")
        Renew_Count += 1
        print("%d" % Renew_Count + " " + Dic_Response["ZH_LABEL"] + " " + Renew_State[0])

if __name__ == '__main__':
    Renew_SFB()

