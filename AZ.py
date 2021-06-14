import requests
import lxml.etree
import urllib
import csv
import json
import concurrent.futures


def File():

    with open("AZ.csv") as file_csv:
        global List_CSV
        List_CSV = []
        reader_obj = csv.reader(file_csv)
        List_CSV = list(reader_obj)
        List_CSV.pop(0)

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Execute_Update_1(Para_List_CSV):

    temp1 = Para_List_CSV[1].split('-')
    temp2 = temp1[0].replace('资源点','')
    print(temp2)

    Dic_Response2 = []
    URL_Query_SFB = "http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet"
    XML_Info_Encoded = "xml=" + urllib.parse.quote_plus('''<request><query mc="guangfenxianxiang" where="1=1 AND ZH_LABEL LIKE '%''' +temp2 + '''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,LONGITUDE,LATITUDE,CITY_ID,COUNTY_ID,STATUS,MAINTAIN_CITY,PRO_TASK_ID,MAINTAIN_COUNTY,CJ_TASK_ID,CJ_STATUS,STRUCTURE_TYPE,RELATED_IS_AREA,STRUCTURE_ID,SERVICE_LEVEL,ROOM_ID,CREATOR,IS_USAGE_STATE,CREAT_TIME,DESIGN_CAPACITY,FIBERDP_NO,MODIFIER,MODIFY_TIME,FREE_CAPACITY,IS_WRONG,GT_VERSION,WRONG_INFO,CREATTIME,SYS_VERSION,INSTALL_CAPACITY,TIME_STAMP,MAINT_DEP,STATEFLAG,MAINT_MODE,PROJECT_ID,MAINT_STATE,LABEL_DEV,LAND_HEIGHT,LOCATION,MOD_NUM,OWNERSHIP,MODEL,PRESERVER,PRESERVER_ADDR,PRESERVER_PHONE,RES_OWNER,SERVICER,TERMINAL_NUM,TIER_COL_COUNT,TIER_ROW_COUNT,USED_CAPACITY,IS_SF_POINT,USER1NAME,SF_POINT_LEV,REMARK,BUILD_DATE,RUWANG_DATE,JIAOWEI_DATE,RELATED_MG_AREA,BUS_AREAR_ID,BUILD_AREA_ID,DEVICE_VENDOR,DETAIL_ADDRESS,FACE_COUNT,EACH_SIDE_ROWS_NUM,EACH_SIDE_COL_NUM,FIBER_COUNT,FIBER_COUNT_FREE,QUALITOR,QUALITOR_COUNTY,QUALITOR_PROJECT,ZICHAN_BELONG,INBOX_NE_TYPE,UP_OPTI_TRANS_BOX,MAINTAINOR,SBXH,FULL_ADDR,TOWN,IS_5G,ROAD_ID,XQRL,JSYXJ,VILLAGE_ID,FLOOR_ID,SQJSNF,GJJB,UNIT_ID,LAYER_ID,ROOM_NUN_ID,SPEC_TYPE,PROJECTCODE,TASK_NAME,ROWNO,COLUMNNO,MAINTAIN_DEPARTMENT,PURPOSE"/></request>''')
    Request_Lenth = chr(len(XML_Info_Encoded))
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Respond_Body = requests.post(URL_Query_SFB, data=XML_Info_Encoded, headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = lxml.etree.HTML(Respond_Body)
    List_Response_Key = Respond_Body.xpath("//fv/@k")
    List_Response_Value = Respond_Body.xpath("//fv/@v")
    Dic_Response2 = dict(zip(List_Response_Key,List_Response_Value))
    


    URL_Query_OpticalCable = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    XML_Info_Encoded = 'xml=' + urllib.parse.quote_plus('''<request><query mc="guanglanduan" ids="" where="1=1 AND ZH_LABEL LIKE '%''' + Para_List_CSV[1] + '''%'" returnfields="INT_ID,ZH_LABEL,ALIAS,CITY_ID,COUNTY_ID,STATUS,PRO_TASK_ID,CJ_TASK_ID,A_EQUIP_TYPE,CJ_STATUS,A_EQUIP_ID,A_ROOM_ID,A_OBJECT_TYPE,A_OBJECT_ID,Z_EQUIP_TYPE,Z_EQUIP_ID,Z_ROOM_ID,Z_OBJECT_TYPE,Z_OBJECT_ID,SERVICE_LEVEL,C_LENGTH,M_LENGTH,FIBER_NUM,RELATED_SYSTEM,WIRE_SEG_TYPE,DIRECTION,IS_ALTER,FIBER_TYPE,DIA,GT_VERSION,VENDOR,OPTI_CABLE_TYPE,MAINT_DEP,MAINT_MODE,LAY_TYPE,RELATED_IS_AREA,IS_WRONG,MAINT_STATE,AREA_LEVEL,WRONG_INFO,SYS_VERSION,OWNERSHIP,PHONE_NO,TIME_STAMP,RES_OWNER,PRODUCE_DATE,PROJECTCODE,BUSINESS,TASK_NAME,ASSENT_NO,SERVICER,RUWANG_DATE,STUFF,BUILDER,TUIWANG_DATE,QUALITOR_PROJECT,PROJECT_ID,QUALITOR,QUALITOR_COUNTY,MAINTAINOR,QRCODE,SEG_NO,REMARK,INDEX_IN_BRANCH,IS_COMPLETE_LAY,MAINTAIN_CITY,MAINTAIN_COUNTY,CREATOR,CREAT_TIME,MODIFIER,MODIFY_TIME,STATEFLAG,BUILD_DATE,PROJECT_NAME,USER_NAME,PURPOSE,PROJECT,YJTZ,JSLY,JSYXJ,SQJSNF"/></request>''')
    HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded"}
    Respond_Body = requests.post(URL_Query_OpticalCable,data=XML_Info_Encoded,headers=HTTP_Header)
    Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
    Respond_Body = lxml.etree.HTML(Respond_Body) 
    List_Response_Key = Respond_Body.xpath("//fv/@k")
    List_Response_Value = Respond_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    if (Dic_Response['A_EQUIP_ID'] == '') or (Dic_Response['A_EQUIP_TYPE'] == '') or (Dic_Response['A_ROOM_ID'] == ''):
        URL_Renew_OpticalCable = "http://10.209.199.74:8120/igisserver_osl/rest/ResourceController/resourcesUpdate?isUpdate=update&coreNamingRules=0"
        XML_Info_Encoded = "xml=" + urllib.parse.quote_plus('<xmldata mode="FibersegEditMode"><mc type="guanglanduan"><mo group="1"><fv k="A_OBJECT_TYPE" v="' + str(Dic_Response["A_OBJECT_TYPE"]) + '"/><fv k="AREA_LEVEL" v="' + str(Dic_Response["AREA_LEVEL"]) + '"/><fv k="STATUS" v="' + str(Dic_Response["STATUS"]) + '"/><fv k="Z_OBJECT_TYPE" v="' + str(Dic_Response["Z_OBJECT_TYPE"]) + '"/><fv k="SQJSNF" v="' + str(Dic_Response["SQJSNF"]) + '"/><fv k="Z_EQUIP_TYPE" v="' + str(Dic_Response["Z_EQUIP_TYPE"]) + '"/><fv k="Z_OBJECT_ID" v="' + str(Dic_Response["Z_OBJECT_ID"]) + '"/><fv k="A_EQUIP_ID" v="' + str(Dic_Response2['INT_ID']) + '"/><fv k="RUWANG_DATE" v="' + str(Dic_Response["RUWANG_DATE"]) + '"/><fv k="A_EQUIP_TYPE" v="' + str(9204) + '"/><fv k="YJTZ" v="' + str(Dic_Response["YJTZ"]) + '"/><fv k="QUALITOR" v="' + str(Dic_Response["QUALITOR"]) + '"/><fv k="Z_EQUIP_ID" v="' + str(Dic_Response["Z_EQUIP_ID"]) + '"/><fv k="PROJECT_NAME" v="' + str(Dic_Response["PROJECT_NAME"]) + '"/><fv k="JSLY" v="' + str(Dic_Response["JSLY"]) + '"/><fv k="LAY_TYPE" v="' + str(Dic_Response["LAY_TYPE"]) + '"/><fv k="ASSENT_NO" v="' + str(Dic_Response["ASSENT_NO"]) + '"/><fv k="MAINTAIN_COUNTY" v="' + str(Dic_Response["MAINTAIN_COUNTY"]) + '"/><fv k="PROJECT_ID" v="' + str(Dic_Response["PROJECT_ID"]) + '"/><fv k="C_LENGTH" v="' + str(Dic_Response["C_LENGTH"]) + '"/><fv k="MAINTAIN_CITY" v="' + str(Dic_Response["MAINTAIN_CITY"]) + '"/><fv k="MAINT_DEP" v="' + str(Dic_Response["MAINT_DEP"]) + '"/><fv k="SYS_VERSION" v="' + str(Dic_Response["SYS_VERSION"]) + '"/><fv k="USER_NAME" v="' + str(Dic_Response["USER_NAME"]) + '"/><fv k="ALIAS" v="' + str(Dic_Response["ALIAS"]) + '"/><fv k="CJ_STATUS" v="' + str(Dic_Response["CJ_STATUS"]) + '"/><fv k="BUILDER" v="' + str(Dic_Response["BUILDER"]) + '"/><fv k="MAINT_STATE" v="' + str(Dic_Response["MAINT_STATE"]) + '"/><fv k="A_ROOM_ID" v="' + str(Dic_Response2['STRUCTURE_ID']) + '"/><fv k="STUFF" v="' + str(Dic_Response["STUFF"]) + '"/><fv k="JSYXJ" v="' + str(Dic_Response["JSYXJ"]) + '"/><fv k="FIBER_TYPE" v="' + str(Dic_Response["FIBER_TYPE"]) + '"/><fv k="MAINT_MODE" v="' + str(Dic_Response["MAINT_MODE"]) + '"/><fv k="Z_ROOM_ID" v="' + str(Dic_Response["Z_ROOM_ID"]) + '"/><fv k="BUILD_DATE" v="' + str(Dic_Response["BUILD_DATE"]) + '"/><fv k="WIRE_SEG_TYPE" v="' + str(Dic_Response["WIRE_SEG_TYPE"]) + '"/><fv k="QUALITOR_PROJECT" v="' + str(Dic_Response["QUALITOR_PROJECT"]) + '"/><fv k="REMARK" v="' + str(Dic_Response["REMARK"]) + '"/><fv k="INT_ID" v="' + str(Dic_Response["INT_ID"]) + '"/><fv k="COUNTY_ID" v="' + str(Dic_Response["COUNTY_ID"]) + '"/><fv k="TASK_NAME" v="' + str(Dic_Response["TASK_NAME"]) + '"/><fv k="IS_ALTER" v="' + str(Dic_Response["IS_ALTER"]) + '"/><fv k="RES_OWNER" v="' + str(Dic_Response["RES_OWNER"]) + '"/><fv k="GT_VERSION" v="' + str(Dic_Response["GT_VERSION"]) + '"/><fv k="PRO_TASK_ID" v="' + str(Dic_Response["PRO_TASK_ID"]) + '"/><fv k="CJ_TASK_ID" v="' + str(Dic_Response["CJ_TASK_ID"]) + '"/><fv k="QUALITOR_COUNTY" v="' + str(Dic_Response["QUALITOR_COUNTY"]) + '"/><fv k="DIRECTION" v="' + str(Dic_Response["DIRECTION"]) + '"/><fv k="PRODUCE_DATE" v="' + str(Dic_Response["PRODUCE_DATE"]) + '"/><fv k="SERVICER" v="' + str(Dic_Response["SERVICER"]) + '"/><fv k="FIBER_NUM" v="' + str(Dic_Response["FIBER_NUM"]) + '"/><fv k="OPTI_CABLE_TYPE" v="' + str(Dic_Response["OPTI_CABLE_TYPE"]) + '"/><fv k="PROJECTCODE" v="' + str(Dic_Response["PROJECTCODE"]) + '"/><fv k="VENDOR" v="' + str(Dic_Response["VENDOR"]) + '"/><fv k="IS_WRONG" v="' + str(Dic_Response["IS_WRONG"]) + '"/><fv k="PHONE_NO" v="' + str(Dic_Response["PHONE_NO"]) + '"/><fv k="RELATED_IS_AREA" v="' + str(Dic_Response["RELATED_IS_AREA"]) + '"/><fv k="WRONG_INFO" v="' + str(Dic_Response["WRONG_INFO"]) + '"/><fv k="DIA" v="' + str(Dic_Response["DIA"]) + '"/><fv k="SERVICE_LEVEL" v="' + str(Dic_Response["SERVICE_LEVEL"]) + '"/><fv k="BUSINESS" v="' + str(Dic_Response["BUSINESS"]) + '"/><fv k="TUIWANG_DATE" v="' + str(Dic_Response["TUIWANG_DATE"]) + '"/><fv k="IS_COMPLETE_LAY" v="' + str(Dic_Response["IS_COMPLETE_LAY"]) + '"/><fv k="QRCODE" v="' + str(Dic_Response["QRCODE"]) + '"/><fv k="ZH_LABEL" v="' + str(Dic_Response["ZH_LABEL"]) + '"/><fv k="M_LENGTH" v="' + str(Dic_Response["M_LENGTH"]) + '"/><fv k="OWNERSHIP" v="' + str(Dic_Response["OWNERSHIP"]) + '"/><fv k="RELATED_SYSTEM" v="' + str(Dic_Response["RELATED_SYSTEM"]) + '"/><fv k="PURPOSE" v="' + str(Dic_Response["PURPOSE"]) + '"/><fv k="INDEX_IN_BRANCH" v="' + str(Dic_Response["INDEX_IN_BRANCH"]) + '"/><fv k="SEG_NO" v="' + str(Dic_Response["SEG_NO"]) + '"/><fv k="CITY_ID" v="' + str(Dic_Response["CITY_ID"]) + '"/><fv k="A_OBJECT_ID" v="' + str(Dic_Response["A_OBJECT_ID"]) + '"/><fv k="MAINTAINOR" v="' + str(Dic_Response["MAINTAINOR"]) + '"/><fv k="PROJECT" v="' + str(Dic_Response["PROJECT"]) + '"/></mo></mc></xmldata>')
        XML_Info_Encoded = XML_Info_Encoded.replace("/","%2f")
        Request_Lenth = chr(len(XML_Info_Encoded))
        HTTP_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
        Respond_Body = requests.post(URL_Renew_OpticalCable,data=XML_Info_Encoded,headers=HTTP_Header)
        Respond_Body = bytes(Respond_Body.text, encoding="utf-8")
        Respond_Body = lxml.etree.HTML(Respond_Body)
        Renew_State = Respond_Body.xpath("//@loaded")
        print(Renew_State[0],Para_List_CSV[1])

if __name__ == '__main__':
    File()
    # for csv in List_CSV:
    #     Execute_Update_1(csv)
    Swimming_Pool(Execute_Update_1,List_CSV)
