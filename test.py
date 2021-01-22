dic = {'a':31, 'bc':5, 'c':3, 'asd':4, 'aa':74, 'd':0}

fff = sorted(dic.items(), key = lambda item:item[0])


print(str(dict(fff)))
