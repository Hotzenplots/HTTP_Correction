# -*- coding: utf-8 -*-
import csv
import pyexiv2
import math
'''
    Beta1 20210306 Initial Release
    Beta2 20210412 更新编码,支持中文文件
'''

def DMS2DD(para_DMS):
    List_Temp_1 = para_DMS.split(' ')
    List_Temp_2 =[]
    for sub_string in List_Temp_1:
        List_Temp_2.append(eval(sub_string))
    Dig_DD = List_Temp_2[0] + List_Temp_2[1] / 60 + List_Temp_2[2] / 3600
    return Dig_DD

def DD2DMS(para_DD):
    List_Temp_1 = math.modf(para_DD)
    DMS_Degree = str(int(List_Temp_1[1])) + '/1'
    DMS_Minute = str(int(List_Temp_1[0] * 60)) + '/1'
    DMS_Second = str(int(round(((List_Temp_1[0] * 60) - (int(List_Temp_1[0] * 60))) * 60 * 1000, 0))) + '/1000'
    Str_DMS = DMS_Degree + ' ' + DMS_Minute + ' ' + DMS_Second
    return Str_DMS

def Modify_Coordinate():
    with open('GPS.csv') as file_csv:
        reader_obj = csv.reader(file_csv)
        List_GPS = list(reader_obj)
        List_GPS.pop(0)

    for gps_info_line in List_GPS:
        Obj_img = pyexiv2.Image(gps_info_line[0], 'gb2312')
        Exif_Data = Obj_img.read_exif()
        Origin_Longitude = Exif_Data['Exif.GPSInfo.GPSLongitude']
        Origin_Latitude = Exif_Data['Exif.GPSInfo.GPSLatitude']
        Origin_Longitude_DD = DMS2DD(Origin_Longitude)
        Origin_Latitude_DD = DMS2DD(Origin_Latitude)
        print(gps_info_line[0],'原坐标为:',round(Origin_Longitude_DD,4),round(Origin_Latitude_DD,4))
        Dic_New_Coordinate = {'Exif.GPSInfo.GPSLongitude': DD2DMS(float(gps_info_line[1])), 'Exif.GPSInfo.GPSLatitude': DD2DMS(float(gps_info_line[2]))}
        Obj_img.modify_exif(Dic_New_Coordinate)
        Exif_Data = Obj_img.read_exif()
        Origin_Longitude = Exif_Data['Exif.GPSInfo.GPSLongitude']
        Origin_Latitude = Exif_Data['Exif.GPSInfo.GPSLatitude']
        Origin_Longitude_DD = DMS2DD(Origin_Longitude)
        Origin_Latitude_DD = DMS2DD(Origin_Latitude)
        Obj_img.close()
        print(gps_info_line[0],'新坐标为:',round(Origin_Longitude_DD,4),round(Origin_Latitude_DD,4))

if __name__ == '__main__':
    Modify_Coordinate()