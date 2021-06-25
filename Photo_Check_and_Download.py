import concurrent.futures
import csv
import urllib.parse
import requests
import lxml.etree

List_Box_Info = []

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with concurrent.futures.ThreadPoolExecutor(max_workers=35) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

def Query_tsid_and_route():
    Query_Session = requests.session()
    Response_Body = Query_Session.post('http://10.231.251.132:7113/rmw/index.jsp', data={'foreign': 'true', 'useraccount': 'baiyunpeng'})
    cookies = requests.utils.dict_from_cookiejar(Response_Body.cookies)
    global route_v, tsid_v
    route_v = cookies['route']
    tsid_v = cookies['tsid']

def Prepare_Data():
    with open('光分纤箱的配置查询结果.csv') as file_csv:
        global List_CSV
        List_CSV = []
        reader_obj = csv.reader(file_csv)
        List_CSV = list(reader_obj)
        List_CSV.pop(0)

def Query_Box_Info(Para_Box_Name):

    if (Para_Box_Name[3].find('GJ') != -1) or (Para_Box_Name[3].find('gj') != -1):
        Box_Type = 'guangjiaojiexiang'
    else:
        Box_Type = 'guangfenxianxiang'

    URL_Query_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Box_Type+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_Box_Name[3]+'%\'" returnfields="INT_ID,ZH_LABEL,LONGITUDE,LATITUDE"/></request>'
    Form_Info_Encoded = 'xml='+urllib.parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Box, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = lxml.etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    if len(Dic_Response) == 0:
        print(Para_Box_Name[3],'名称有误')
        exit()
    List_Box_Info.append(Dic_Response)

def Query_Photo(Para_Box_Info):

    if (Para_Box_Info['ZH_LABEL'].find('GJ') != -1) or (Para_Box_Info['ZH_LABEL'].find('gj') != -1):
        resClassName = 'OPTI_TRAN_BOX'
    else:
        resClassName = 'OPTI_SFB'

    URL_Query_Photo = 'http://10.231.251.132:7113/rmw/datamanage/resmaintain/resMaintainAction!preUploadPhoto.action?objId='+Para_Box_Info['INT_ID']+'&resClassName='+resClassName
    Response_Body = requests.get(URL_Query_Photo, cookies={'tsid': tsid_v, 'route': route_v})
    Response_Body = lxml.etree.HTML(Response_Body.text)
    List_Photo =  Response_Body.xpath("//img/@src")
    List_Photo.pop(0)
    Para_Box_Info['Photo_Count'] = len(List_Photo)
    Para_Box_Info['Photo_IDs'] = [i[len(i)-6:] for i in List_Photo]

def Download_Photo(Para_Box_Info):

    for photo_id in Para_Box_Info['Photo_IDs']:
        URL_Download = 'http://10.231.251.132:7113/rmw/datamanage/resmaintain/resMaintainAction!sxDownloadImg.action?id='+str(photo_id)
        Response_Body = requests.get(URL_Download, cookies={'tsid': tsid_v, 'route': route_v})
        Obj_Img = Response_Body.content
        with open( Para_Box_Info['ZH_LABEL'] + '__' + str(photo_id) + '.jpg','wb' ) as file_img:
            file_img.write(Obj_Img)

if __name__ == '__main__':
    Prepare_Data()
    # for each_data in List_CSV:
    #     print(each_data[3])
    Query_tsid_and_route()
    Swimming_Pool(Query_Box_Info, List_CSV)
    Swimming_Pool(Query_Photo, List_Box_Info)
    Swimming_Pool(Download_Photo, List_Box_Info)
