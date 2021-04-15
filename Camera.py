from concurrent.futures import ThreadPoolExecutor
from lxml import etree
from urllib import parse
import csv
import requests
import os
import math
import pyexiv2

def Query_tsid_and_route():
    Query_Session = requests.session()
    Response_Body = Query_Session.post('http://10.231.251.132:7113/rmw/index.jsp', data={'foreign': 'true', 'useraccount': 'baiyunpeng'})
    cookies = requests.utils.dict_from_cookiejar(Response_Body.cookies)
    global route_v, tsid_v
    route_v = cookies['route']
    tsid_v = cookies['tsid']
    print('route: ' + route_v)
    print('tsid: ' + tsid_v)

def Query_Box_Info(Para_Box_Name):

    if (Para_Box_Name.find('GJ') != -1) or (Para_Box_Name.find('gj') != -1):
        Box_Type = 'guangjiaojiexiang'
    else:
        Box_Type = 'guangfenxianxiang'

    URL_Query_Box = 'http://10.209.199.74:8120/igisserver_osl/rest/generalSaveOrGet/generalGet'
    Form_Info = '<request><query mc="'+Box_Type+'" ids="" where="1=1 AND ZH_LABEL LIKE \'%'+Para_Box_Name+'%\'" returnfields="INT_ID,ZH_LABEL,LONGITUDE,LATITUDE"/></request>'
    Form_Info_Encoded = 'xml='+parse.quote_plus(Form_Info)
    Request_Lenth = str(len(Form_Info_Encoded))
    Request_Header = {"Host": "10.209.199.74:8120","Content-Type": "application/x-www-form-urlencoded","Content-Length": Request_Lenth}
    Response_Body = requests.post(URL_Query_Box, data=Form_Info_Encoded, headers=Request_Header)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Response_Key = Response_Body.xpath("//fv/@k")
    List_Response_Value = Response_Body.xpath("//fv/@v")
    Dic_Response = dict(zip(List_Response_Key,List_Response_Value))
    List_Box_Info.append(Dic_Response)

def Modify_Phote_Coordinate():
    for box_info in List_Box_Info:
        for photo_name in box_info['Box_Photos']:
            Obj_img = pyexiv2.Image(photo_name, 'gb2312')

            Exif_Data = Obj_img.read_exif()
            Origin_Longitude = Exif_Data['Exif.GPSInfo.GPSLongitude']
            Origin_Latitude = Exif_Data['Exif.GPSInfo.GPSLatitude']
            Origin_Longitude_DD = DMS2DD(Origin_Longitude)
            Origin_Latitude_DD = DMS2DD(Origin_Latitude)

            Dic_New_Coordinate = {'Exif.GPSInfo.GPSLongitude': DD2DMS(float(box_info['LONGITUDE'])), 'Exif.GPSInfo.GPSLatitude': DD2DMS(float(box_info['LATITUDE']))}
            Obj_img.modify_exif(Dic_New_Coordinate)
            Exif_Data = Obj_img.read_exif()
            New_Longitude = Exif_Data['Exif.GPSInfo.GPSLongitude']
            New_Latitude = Exif_Data['Exif.GPSInfo.GPSLatitude']
            New_Longitude_DD = DMS2DD(New_Longitude)
            New_Latitude_DD = DMS2DD(New_Latitude)

            Obj_img.close()

            print(photo_name,r'原坐标/新坐标:',str(round(Origin_Longitude_DD,4)) + r'/' + str(round(New_Longitude_DD,4)),'|||',str(round(Origin_Latitude_DD,4)) + r'/'+str(round(New_Latitude_DD,4)))

def DD2DMS(para_DD):
    List_Temp_1 = math.modf(para_DD)
    DMS_Degree = str(int(List_Temp_1[1])) + '/1'
    DMS_Minute = str(int(List_Temp_1[0] * 60)) + '/1'
    DMS_Second = str(int(round(((List_Temp_1[0] * 60) - (int(List_Temp_1[0] * 60))) * 60 * 1000, 0))) + '/1000'
    Str_DMS = DMS_Degree + ' ' + DMS_Minute + ' ' + DMS_Second
    return Str_DMS

def DMS2DD(para_DMS):
    List_Temp_1 = para_DMS.split(' ')
    List_Temp_2 =[]
    for sub_string in List_Temp_1:
        List_Temp_2.append(eval(sub_string))
    Dig_DD = List_Temp_2[0] + List_Temp_2[1] / 60 + List_Temp_2[2] / 3600
    return Dig_DD

def Upload_Photo(Para_Box_Info):

    if (Para_Box_Info['ZH_LABEL'].find('GJ') != -1) or (Para_Box_Info['ZH_LABEL'].find('gj') != -1):
        resClassName = 'OPTI_TRAN_BOX'
    else:
        resClassName = 'OPTI_SFB'

    for photo_name in Para_Box_Info['Box_Photos']:
        URL_Upload = 'http://10.231.251.132:7113/rmw/datamanage/resmaintain/resMaintainAction!uploadPhoto.action?objId='+Para_Box_Info['INT_ID']+'&resClassName='+resClassName+'&extensionName='+photo_name[(len(photo_name)-3):len(photo_name)]
        files = {'filename':(photo_name, open(photo_name, 'rb'),'image/pjpeg')}
        Response_Body = requests.post(URL_Upload, files=files, cookies={'tsid': tsid_v, 'route': route_v})
        print(photo_name, Response_Body.status_code)

def Processing_Local_File():
    global List_Box_Info, List_Box_Name,List_Folder_File_List
    List_Box_Name = List_Box_Info = List_Folder_File_List =[]
    with open('Camera.csv') as file_csv:
        reader_obj = csv.reader(file_csv)
        List_Box_Name = list(reader_obj)
        List_Box_Name.pop(0)
    List_Box_Name = [str(i) for j in List_Box_Name for i in j]
    List_Folder_File_List = os.listdir()
    List_Folder_File_List.remove('Camera.py')
    List_Folder_File_List.remove('Camera.csv')

def Find_Photos():
    for box_info in List_Box_Info:
        List_Box_Photos = []
        for photo_name in List_Folder_File_List:
            if (photo_name.find(box_info['ZH_LABEL']) != -1):
                List_Box_Photos.append(photo_name)
        box_info['Box_Photos'] = List_Box_Photos

def Clear_Photo(Para_Box_Info):

    if (Para_Box_Info['ZH_LABEL'].find('GJ') != -1) or (Para_Box_Info['ZH_LABEL'].find('gj') != -1):
        resClassName = 'OPTI_TRAN_BOX'
    else:
        resClassName = 'OPTI_SFB'

    URL_Query_Photo = 'http://10.231.251.132:7113/rmw/datamanage/resmaintain/resMaintainAction!preUploadPhoto.action?objId='+Para_Box_Info['INT_ID']+'&resClassName='+resClassName
    Response_Body = requests.get(URL_Query_Photo, cookies={'tsid': tsid_v, 'route': route_v})
    Response_Body = etree.HTML(Response_Body.text)
    List_Photo =  Response_Body.xpath("//img/@src")
    List_Photo.pop(0)
    if len(List_Photo) == 0:
        print(Para_Box_Info['ZH_LABEL']+'无照片')
    elif len(List_Photo) != 0:
        for photo_address in List_Photo:
            URL_Delete_Photo = 'http://10.231.251.132:7113/rmw/datamanage/resmaintain/resMaintainAction!detelePhoto.action?photoId='+photo_address[len(photo_address) - 6:len(photo_address)]+'&objId='+Para_Box_Info['INT_ID']+'&resClassName='+resClassName
            Response_Body = requests.get(URL_Delete_Photo, cookies={'tsid': tsid_v, 'route': route_v})
        print(Para_Box_Info['ZH_LABEL']+'照片已清空')

def Swimming_Pool(Para_Functional_Function,Para_Some_Iterable_Obj):
    with ThreadPoolExecutor(max_workers=10) as Pool_Executor:
        Pool_Executor.map(Para_Functional_Function,(Para_Some_Iterable_Obj))

if __name__ == '__main__':
    Processing_Local_File()
    Swimming_Pool(Query_Box_Info, List_Box_Name)
    Find_Photos()
    Modify_Phote_Coordinate()
    Query_tsid_and_route()
    # Swimming_Pool(Clear_Photo, List_Box_Info)
    # Swimming_Pool(Upload_Photo, List_Box_Info)
    for box_info in List_Box_Info:
        Clear_Photo(box_info)
        Upload_Photo(box_info)