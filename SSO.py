import requests
import datetime
from time import sleep
from urllib import parse

JSESSIONIRMS_7112 = 'GJ12ghgBnv1XNj4GygP88thQhK2b5HvTLTcTypKhmcChzb2462vp!-2126455540'
route_7112 = '5fb592aa37b5606b0629ebaa738ace15'

tsid_7016 = '73780ccc5eafb823e563aee926e8'
route_7016 = 'e3e320018ad690f48b373ab456d467cd'

def Keep_7112(para_jessionirms, para_route):
    URL_7112 = 'http://10.209.199.72:7112/irms/tasklistAction!waitTask4Person.ilf'
    Session_7112 = requests.session()
    requests.utils.add_dict_to_cookiejar(Session_7112.cookies,{'JSESSIONIRMS':para_jessionirms, 'route': para_route})
    Form_Info_Encoded = 'start=' + parse.quote_plus('1') + '&rp='+ parse.quote_plus('100')
    Request_Lenth = str(len(Form_Info_Encoded))
    Session_7112.headers.update({'Host':'10.209.199.72:7112', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth})
    Response_Body = Session_7112.post(URL_7112,data=Form_Info_Encoded)
    print('7112', Response_Body)

def Keep_7016(para_tsid, para_route):
    URL_7016 = 'http://10.209.199.32:7016/irms/tasklistAction!waitedTaskAJAX.ilf'
    Session_7016 = requests.session()
    requests.utils.add_dict_to_cookiejar(Session_7016.cookies,{'tsid':para_tsid, 'route': para_route})
    Form_Info_Encoded = 'start=' + parse.quote_plus('0') + '&limit='+ parse.quote_plus('15')
    Request_Lenth = str(len(Form_Info_Encoded))
    Session_7016.headers.update({'Host':'10.209.199.32:7016', 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Request_Lenth})
    Response_Body = Session_7016.post(URL_7016,data=Form_Info_Encoded)
    print('7016', Response_Body)

while True:
    print(datetime.datetime.now())
    Keep_7112(JSESSIONIRMS_7112,route_7112)
    Keep_7016(tsid_7016,route_7016)
    sleep(300)