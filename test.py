
import requests
from lxml import etree
from Crypto.Cipher import AES
import base64
# test_dic = {'a':31, 'bc':5, 'c':3, 'asd':4, 'aa':74, 'd':0}
# new_dic = sorted(test_dic.items(), key = lambda item:item[0])
# print(str(dict(new_dic)))

# session = requests.session()
# t = session.post('http://portal.sx.cmcc/pkmslogin.form?uid=tyyangwei',data= {'login-form-type':'pwd','username':'tyyangwei','password':'ZvixritCBQIFNwb0jmzL5A=='})
# t = session.post('http://portal.sx.cmcc/sxmcc_wcm/middelwebpage/encryptportallogin/encryptlogin.jsp?appurl=http://10.209.199.72:7112/irms/sso.action')
# t1 = bytes(t.text, encoding="utf-8")
# t1 = etree.HTML(t1)
# l = t1.xpath('//@value')
# t = requests.post('http://10.209.199.72:7112/irms/sso.action',data={'userdt': l[0], 'ipAddress': l[1]})
# cookies = requests.utils.dict_from_cookiejar(t.cookies)
# print(cookies)

# def add_16(par):
#     # python3字符串是unicode编码，需要 encode才可以转换成字节型数据
#     par = par.encode('utf-8')
#     while len(par) % 16 != 0:
#         par += b'\x00'
#     return par
def add_to_16(text):
    while len(text) % 16 != 0:
        text += '\0'
    return str.encode(text)  # 返回bytes
key='sxportaljiamikey'
key_byte = key.encode()
iv ='sxportaljiamiwyl'
iv_byte = iv.encode()
pw = 'tyyw789...'
pw_byte = pw.encode()
pw_byte = add_to_16(pw)
cipher = AES.new(key_byte, AES.MODE_CBC, iv_byte)
encrypted_byte = cipher.encrypt(pw_byte)
encodestrs = base64.b64encode(encrypted_byte)
print(encodestrs.decode())

'ZvixritCBQIFNwb0jmzL5A=='
# a ='yyz123
# b = a.encode('utf-8')
# print(b)
# c = b.decode('utf-8')
# print(c)
# print(key)
