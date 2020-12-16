from selenium import webdriver
from time import sleep

def Get_JSESSIONIRMS_and_route():
    Browser_Obj = webdriver.Ie()
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_was/uploadResource/public/login/login.html')
    Browser_Obj.find_element_by_id('username').send_keys('tyyangwei')
    Browser_Obj.find_element_by_id('password').send_keys('tyyw159...')
    Browser_Obj.find_element_by_class_name('lwb_login').click()
    sleep(2)
    Browser_Obj.get(r'http://portal.sx.cmcc/sxmcc_wcm/middelwebpage/app_recoder_log.jsp?app_flg=zhwlzygl_ywzl')
    sleep(2)
    Dic_Cookie_JSESSIONIRMS = Browser_Obj.get_cookie('JSESSIONIRMS')
    Dic_Cookie_route = Browser_Obj.get_cookie('route')
    JSESSIONIRMS_Value = Dic_Cookie_JSESSIONIRMS['value']
    route_Value = Dic_Cookie_route['value']
    Browser_Obj.quit()
    return [JSESSIONIRMS_Value, route_Value]

if __name__ == '__main__':
    what = Get_JSESSIONIRMS_and_route()
    print(what)
