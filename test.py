from lxml import etree
x1 = '''<xml success="true" msg="" total="4"><path id="439854750;439854826" type="9204-9204" flowId="1620870853980" workid="1186797"><dpath id="439854750;439854826" name="太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001,太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001"><line id="37871D8EEAFF43479DFDF00C50F26FBD" portid="865757919" fiberno="1-1" stat="0" groupno="" name="【太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001-ODM01-01-01】(太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001资源点-太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001资源点光缆段)【太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001-ODM01-01-01】(太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001资源点-太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001资源点光缆段)"/><line id="1A7C50F7988E4790B3C6AF8BB17DC06B" portid="865757922" fiberno="4-4" stat="0" groupno="" name="【太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001-ODM01-01-04】(太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001资源点-太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001资源点光缆段)【太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001-ODM01-01-04】"/><line id="F4711A3AE1C444F6992D60EE046ECDEC" portid="865757922" fiberno="4-4" stat="0" groupno="" name="【太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001-ODM01-01-04】(太原小店区肥皂厂回迁楼3号楼2单元1层竖井内GF0001资源点-太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001资源点光缆段)【太原小店区肥皂厂回迁楼3号楼1单元1层竖井内GF0001-ODM01-01-04】"/></dpath></path></xml>'''
x1 = etree.HTML(x1)
List_Path_IDs = x1.xpath('//line/@id')

# 处理重复通路的系统bug开始
List_FiberNo = x1.xpath('//line/@fiberno')
dic_te = dict(zip(List_FiberNo,List_Path_IDs))
print(List_FiberNo)
print(List_Path_IDs)
print(dic_te)
