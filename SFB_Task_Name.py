# -*- coding: UTF-8 -*-
import csv
import urllib.parse
import requests
import concurrent.futures
import lxml.etree

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Open_File():
    with open("SFB_Task_Name.csv") as file_csv:
        global List_CSV
        List_CSV = []
        reader_obj = csv.reader(file_csv)
        List_CSV = list(reader_obj)
        List_CSV.pop(0)

def Renew_SFB(Para_List_CSV):
    Dic_Response = []
    URL_Query_SFB = "http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet"
    XML_Info_Encoded = "xml=" + urllib.parse.quote_plus('''<request><query mc="guangfenxianxiang" where="1=1 AND ZH_LABEL LIKE '%''' + Para_List_CSV[0] + '''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,LONGITUDE,LATITUDE,CITY_ID,COUNTY_ID,STATUS,MAINTAIN_CITY,PRO_TASK_ID,MAINTAIN_COUNTY,CJ_TASK_ID,CJ_STATUS,STRUCTURE_TYPE,RELATED_IS_AREA,STRUCTURE_ID,SERVICE_LEVEL,ROOM_ID,CREATOR,IS_USAGE_STATE,CREAT_TIME,DESIGN_CAPACITY,FIBERDP_NO,MODIFIER,MODIFY_TIME,FREE_CAPACITY,IS_WRONG,GT_VERSION,WRONG_INFO,CREATTIME,SYS_VERSION,INSTALL_CAPACITY,TIME_STAMP,MAINT_DEP,STATEFLAG,MAINT_MODE,PROJECT_ID,MAINT_STATE,LABEL_DEV,LAND_HEIGHT,LOCATION,MOD_NUM,OWNERSHIP,MODEL,PRESERVER,PRESERVER_ADDR,PRESERVER_PHONE,RES_OWNER,SERVICER,TERMINAL_NUM,TIER_COL_COUNT,TIER_ROW_COUNT,USED_CAPACITY,IS_SF_POINT,USER1NAME,SF_POINT_LEV,REMARK,BUILD_DATE,RUWANG_DATE,JIAOWEI_DATE,RELATED_MG_AREA,BUS_AREAR_ID,BUILD_AREA_ID,DEVICE_VENDOR,DETAIL_ADDRESS,FACE_COUNT,EACH_SIDE_ROWS_NUM,EACH_SIDE_COL_NUM,FIBER_COUNT,FIBER_COUNT_FREE,QUALITOR,QUALITOR_COUNTY,QUALITOR_PROJECT,ZICHAN_BELONG,INBOX_NE_TYPE,UP_OPTI_TRANS_BOX,MAINTAINOR,SBXH,FULL_ADDR,TOWN,IS_5G,ROAD_ID,XQRL,JSYXJ,VILLAGE_ID,FLOOR_ID,SQJSNF,GJJB,UNIT_ID,LAYER_ID,ROOM_NUN_ID,SPEC_TYPE,PROJECTCODE,TASK_NAME,ROWNO,COLUMNNO,MAINTAIN_DEPARTMENT,PURPOSE"/></request>''')
    Request_Lenth = chr(len(XML_Info_Encoded))
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Respond_Body = requests.post(URL_Query_SFB, data=XML_Info_Encoded, headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = lxml.etree.HTML(Respond_Body)
    List_Response_Key = Respond_Body.xpath("//fv/@k")
    List_Response_Value = Respond_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))

    temp_list = Para_List_CSV[1].split('-')

    URL_Renew_SFB = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update"
    XML_Info_Encoded = "xml=" + urllib.parse.quote_plus('<xmldata mode="OptoeleEditMode"><mc type="guangfenxianxiang"><mo group="1"><fv k="PRESERVER" v="' + Dic_Response["PRESERVER"] + '"/><fv k="STATUS" v="' + Dic_Response["STATUS"] + '"/><fv k="SYS_VERSION" v="' + Dic_Response["SYS_VERSION"] + '"/><fv k="GJJB" v="' + Dic_Response["GJJB"] + '"/><fv k="SQJSNF" v="' + Dic_Response["SQJSNF"] + '"/><fv k="RUWANG_DATE" v="' + Dic_Response["RUWANG_DATE"] + '"/><fv k="SF_POINT_LEV" v="' + Dic_Response["SF_POINT_LEV"] + '"/><fv k="COUNTY_ID" v="' + Dic_Response["COUNTY_ID"] + '"/><fv k="FIBER_COUNT_FREE" v="' + Dic_Response["FIBER_COUNT_FREE"] + '"/><fv k="MOD_NUM" v="' + Dic_Response["MOD_NUM"] + '"/><fv k="RELATED_MG_AREA" v="' + Dic_Response["RELATED_MG_AREA"] + '"/><fv k="MODEL" v="' + Dic_Response["MODEL"] + '"/><fv k="INSTALL_CAPACITY" v="' + Dic_Response["INSTALL_CAPACITY"] + '"/><fv k="USED_CAPACITY" v="' + Dic_Response["USED_CAPACITY"] + '"/><fv k="STRUCTURE_ID" v="' + Dic_Response["STRUCTURE_ID"] + '"/><fv k="STRUCTURE_TYPE" v="' + Dic_Response["STRUCTURE_TYPE"] + '"/><fv k="FREE_CAPACITY" v="' + Dic_Response["FREE_CAPACITY"] + '"/><fv k="CJ_TASK_ID" v="' + Dic_Response["CJ_TASK_ID"] + '"/><fv k="IS_WRONG" v="' + Dic_Response["IS_WRONG"] + '"/><fv k="LAND_HEIGHT" v="' + Dic_Response["LAND_HEIGHT"] + '"/><fv k="IS_5G" v="' + Dic_Response["IS_5G"] + '"/><fv k="JIAOWEI_DATE" v="' + Dic_Response["JIAOWEI_DATE"] + '"/><fv k="ALIAS" v="' + Dic_Response["ALIAS"] + '"/><fv k="TOWN" v="' + Dic_Response["TOWN"] + '"/><fv k="DESIGN_CAPACITY" v="' + Dic_Response["DESIGN_CAPACITY"] + '"/><fv k="MAINTAIN_CITY" v="' + Dic_Response["MAINTAIN_CITY"] + '"/><fv k="SERVICE_LEVEL" v="' + Dic_Response["SERVICE_LEVEL"] + '"/><fv k="MAINTAIN_COUNTY" v="' + Dic_Response["MAINTAIN_COUNTY"] + '"/><fv k="INBOX_NE_TYPE" v="' + Dic_Response["INBOX_NE_TYPE"] + '"/><fv k="ROAD_ID" v="' + Dic_Response["ROAD_ID"] + '"/><fv k="FIBERDP_NO" v="' + Dic_Response["FIBERDP_NO"] + '"/><fv k="LABEL_DEV" v="' + Dic_Response["LABEL_DEV"] + '"/><fv k="JSYXJ" v="' + Dic_Response["JSYXJ"] + '"/><fv k="IS_USAGE_STATE" v="' + Dic_Response["IS_USAGE_STATE"] + '"/><fv k="FULL_ADDR" v="' + Dic_Response["FULL_ADDR"] + '"/><fv k="RES_OWNER" v="' + Dic_Response["RES_OWNER"] + '"/><fv k="CJ_STATUS" v="' + Dic_Response["CJ_STATUS"] + '"/><fv k="PROJECTCODE" v="' + Dic_Response["PROJECTCODE"] + '"/><fv k="TASK_NAME" v="' + temp_list[0] + '"/><fv k="VILLAGE_ID" v="' + Dic_Response["VILLAGE_ID"] + '"/><fv k="MAINT_MODE" v="' + Dic_Response["MAINT_MODE"] + '"/><fv k="BUILD_DATE" v="' + Dic_Response["BUILD_DATE"] + '"/><fv k="QUALITOR" v="' + Dic_Response["QUALITOR"] + '"/><fv k="XQRL" v="' + Dic_Response["XQRL"] + '"/><fv k="MAINT_STATE" v="' + Dic_Response["MAINT_STATE"] + '"/><fv k="QUALITOR_COUNTY" v="' + Dic_Response["QUALITOR_COUNTY"] + '"/><fv k="QUALITOR_PROJECT" v="' + Dic_Response["QUALITOR_PROJECT"] + '"/><fv k="MAINTAINOR" v="' + Dic_Response["MAINTAINOR"] + '"/><fv k="USER1NAME" v="' + Dic_Response["USER1NAME"] + '"/><fv k="FLOOR_ID" v="' + Dic_Response["FLOOR_ID"] + '"/><fv k="REMARK" v="' + Dic_Response["REMARK"] + '"/><fv k="UNIT_ID" v="' + Dic_Response["UNIT_ID"] + '"/><fv k="WRONG_INFO" v="' + Dic_Response["WRONG_INFO"] + '"/><fv k="MAINT_DEP" v="' + Dic_Response["MAINT_DEP"] + '"/><fv k="TIER_ROW_COUNT" v="' + Dic_Response["TIER_ROW_COUNT"] + '"/><fv k="PRO_TASK_ID" v="' + Dic_Response["PRO_TASK_ID"] + '"/><fv k="UP_OPTI_TRANS_BOX" v="' + Dic_Response["UP_OPTI_TRANS_BOX"] + '"/><fv k="ROOM_ID" v="' + Dic_Response["ROOM_ID"] + '"/><fv k="MAINTAIN_DEPARTMENT" v="' + Dic_Response["MAINTAIN_DEPARTMENT"] + '"/><fv k="PROJECT_ID" v="' + Dic_Response["PROJECT_ID"] + '"/><fv k="PURPOSE" v="' + Dic_Response["PURPOSE"] + '"/><fv k="ZICHAN_BELONG" v="' + Dic_Response["ZICHAN_BELONG"] + '"/><fv k="ROOM_NUN_ID" v="' + Dic_Response["ROOM_NUN_ID"] + '"/><fv k="LAYER_ID" v="' + Dic_Response["LAYER_ID"] + '"/><fv k="INT_ID" v="' + Dic_Response["INT_ID"] + '"/><fv k="CITY_ID" v="' + Dic_Response["CITY_ID"] + '"/><fv k="TIER_COL_COUNT" v="' + Dic_Response["TIER_COL_COUNT"] + '"/><fv k="OWNERSHIP" v="' + Dic_Response["OWNERSHIP"] + '"/><fv k="LONGITUDE" v="' + Dic_Response["LONGITUDE"] + '"/><fv k="LATITUDE" v="' + Dic_Response["LATITUDE"] + '"/><fv k="ROWNO" v="' + Dic_Response["ROWNO"] + '"/><fv k="TERMINAL_NUM" v="' + Dic_Response["TERMINAL_NUM"] + '"/><fv k="COLUMNNO" v="' + Dic_Response["COLUMNNO"] + '"/><fv k="EACH_SIDE_COL_NUM" v="' + Dic_Response["EACH_SIDE_COL_NUM"] + '"/><fv k="EACH_SIDE_ROWS_NUM" v="' + Dic_Response["EACH_SIDE_ROWS_NUM"] + '"/><fv k="ZH_LABEL" v="' + Dic_Response["ZH_LABEL"] + '"/><fv k="LOCATION" v="' + Dic_Response["LOCATION"] + '"/><fv k="DETAIL_ADDRESS" v="' + Dic_Response["DETAIL_ADDRESS"] + '"/><fv k="RELATED_IS_AREA" v="' + Dic_Response["RELATED_IS_AREA"] + '"/><fv k="FIBER_COUNT" v="' + Dic_Response["FIBER_COUNT"] + '"/><fv k="DEVICE_VENDOR" v="' + Dic_Response["DEVICE_VENDOR"] + '"/><fv k="PRESERVER_PHONE" v="' + Dic_Response["PRESERVER_PHONE"] + '"/><fv k="SPEC_TYPE" v="' + Dic_Response["SPEC_TYPE"] + '"/><fv k="CREATTIME" v="' + Dic_Response["CREATTIME"] + '"/><fv k="SERVICER" v="' + Dic_Response["SERVICER"] + '"/><fv k="SBXH" v="' + Dic_Response["SBXH"] + '"/><fv k="BUILD_AREA_ID" v="' + Dic_Response["BUILD_AREA_ID"] + '"/><fv k="PRESERVER_ADDR" v="' + Dic_Response["PRESERVER_ADDR"] + '"/><fv k="FACE_COUNT" v="' + Dic_Response["FACE_COUNT"] + '"/><fv k="GT_VERSION" v="' + Dic_Response["GT_VERSION"] + '"/><fv k="IS_SF_POINT" v="' + Dic_Response["IS_SF_POINT"] + '"/><fv k="BUS_AREAR_ID" v="' + Dic_Response["BUS_AREAR_ID"] + '"/></mo></mc></xmldata>')
    Request_Lenth = chr(len(XML_Info_Encoded))
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Respond_Body = requests.post(URL_Renew_SFB, data=XML_Info_Encoded, headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = lxml.etree.HTML(Respond_Body)
    Renew_State = Respond_Body.xpath("//@loaded")
    print(Dic_Response["ZH_LABEL"] + " " + Renew_State[0])

if __name__ == '__main__':
    Open_File()
    Swimming_Pool(Renew_SFB, List_CSV)
