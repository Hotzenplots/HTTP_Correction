from datetime import datetime
import time
import threading

Start_time = datetime.now()
def run(n):
    print("task", n)
    time.sleep(1)       #此时子线程停1s
    print('3')
    time.sleep(1)
    print('2')
    time.sleep(1)
    print('1')
    print(temp_data)

if __name__ == '__main__':
    for RL1 in range(5):
        temp_data = RL1
        t = threading.Thread(target=run, args=(RL1,))
        t.setDaemon(True)   #把子进程设置为守护线程，必须在start()之前设置
        t.start()
        t.join() # 设置主线程等待子线程结束
        print("end")
    End_Time = datetime.now()
    print("程序运行了" + "%d" % (End_Time-Start_time).seconds + "秒")


List_DQS = [
    ["古交市",445835320,"tyyangwei",246552123,"lurui2",246552126,"liuhongbo",342412680,"tt_jt_zhangyuelong",685584994],
    ["娄烦县",445835319,"tyyangwei",246552123,"lurui2",246552126,"xuyanan",342412681,"tt_2_gongjianfeng",685583882],
    ["阳曲县",445835318,"tyyangwei",246552123,"lurui2",246552126,"liupeixu",483582248,"tt_1_lvmai",685585124],
    ["清徐县",445835317,"tyyangwei",246552123,"lurui2",246552126,"wangdapeng",248,"tt_2_guojianbo",685583261],
    ["晋源区",445835316,"tyyangwei",246552123,"lurui2",246552126,"tyxuyan",342412683,"tt_2_peihao",685587010],
    ["万柏林区",445835315,"tyyangwei",246552123,"lurui2",246552126,"tyxuyan",342412683,"tt_1_anliguo",685583025],
    ["尖草坪区",445835314,"tyyangwei",246552123,"lurui2",246552126,"zhengzhihua",342412684,"tt_2_songyingjiang"],
    ["杏花岭区",445835313,"tyyangwei",246552123,"lurui2",246552126,"zhengzhihua",342412684,"tt_1_muruibin"],
    ["迎泽区",445835312,"tyyangwei",246552123,"lurui2",246552126,"dongjianfeng",483583693,"tt_1_ttwangzhenhua"],
    ["小店区",445835311,"tyyangwei",246552123,"lurui2",246552126,"liuya",483572456,"tt_1_weixiaodong1",685582176]
    ]