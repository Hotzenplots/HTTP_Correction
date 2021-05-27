import json

dic = { 'andy':{ 'age': [{r'你好': 55}, 00], 'city': 'beijing', 'skill': 'python' }, 
        'william': { 'age': 25, 'city': 'shanghai', 'skill': 'js' } 
        }
js = json.dumps(dic, ensure_ascii=False)
file = open('test2.txt', 'w')
file.write(js)
file.close()

# file = open('test2.txt', 'r') 
# js = file.read()
# dic = json.loads(js)   
# # print(dic[0]['Box_Name']) 
# file.close()
