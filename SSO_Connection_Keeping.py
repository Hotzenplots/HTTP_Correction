import requests
import datetime
import json
from time import sleep
from urllib import parse
from lxml import etree

JSESSIONIRMS_7112 = 'bBvMgz1VnyZJDJZ6lmmFHPLpd2H2sS1yB9yTyt2034l2p4cJJg3m!1269070846'
route_7112 = 'b7724c4ffeaf48382b8d4d099b73de2f'

tsid_7016 = 'b1775b07cf28bb964cae34a0a8cb'
route_7016 = 'e3e320018ad690f48b373ab456d467cd'

def Keep_7112():
    URL_7112 = 'http://10.209.199.72:7112/irms/login.ilf'
    Session_7112 = requests.session()
    requests.utils.add_dict_to_cookiejar(Session_7112.cookies,{'JSESSIONIRMS':JSESSIONIRMS_7112, 'route': route_7112})
    Form_Info_Encoded = 'start=' + parse.quote_plus('1') + '&rp='+ parse.quote_plus('100')
    Request_Lenth = str(len(Form_Info_Encoded))
    Session_7112.headers.update({'Host':'10.209.199.72:7112', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth})
    Response_Body = Session_7112.post(URL_7112,data=Form_Info_Encoded)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Title = Response_Body.xpath("//title/text()")
    if List_Title[0] == '中国移动综合网络资源管理系统':
        print('7112', '正常连接')
    else:
        print('7112', '登录超时')

    Form_Info_Encoded = 'msg=' + parse.quote_plus('old')
    Response_Body = Session_7112.post(URL_7112,data=Form_Info_Encoded)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Title = Response_Body.xpath("//title/text()")
    if List_Title[0] == 'Electric Operation Maintenance System':
        print('光路调度', '正常连接')
    else:
        print('光路调度', '登录超时')

    URL_GLDD ='http://10.209.199.72:7112/irms/tasklistAction!waitedTaskAJAX.ilf'
    Session_7112.headers.clear()
    Response_Body = Session_7112.get(URL_GLDD)
    Response_Body = Response_Body.text
    Response_Body = json.loads(Response_Body)
    Exist_Work_Sheet = Response_Body['totalCount']
    print('光路工单',Exist_Work_Sheet,'个')

def Keep_7016(para_tsid, para_route):
    URL_7016 = 'http://10.209.199.32:7016/irms/tasklistAction!waitedTaskAJAX.ilf'
    Session_7016 = requests.session()
    requests.utils.add_dict_to_cookiejar(Session_7016.cookies,{'tsid':para_tsid, 'route': para_route})
    Form_Info_Encoded = 'start=' + parse.quote_plus('0') + '&limit='+ parse.quote_plus('15')
    Request_Lenth = str(len(Form_Info_Encoded))
    Session_7016.headers.update({'Host':'10.209.199.32:7016', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth})
    Response_Body = Session_7016.post(URL_7016,data=Form_Info_Encoded)
    Response_Body = bytes(Response_Body.text, encoding="utf-8")
    Response_Body = etree.HTML(Response_Body)
    List_Title = Response_Body.xpath("//title/text()")
    if len(List_Title) == 0:
        print('7016', '正常连接')
    else:
        print('7016', List_Title[0])

while True:
    print(datetime.datetime.now())
    Keep_7112()
    # Keep_7016(tsid_7016,route_7016)
    sleep(600)

