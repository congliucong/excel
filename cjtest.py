# -*- coding: utf-8 -*-
import requests

url = "http://www.dzzsb.com.cn:7777/cjxs.asp"

# form-data参数要写成如下格式，注意有None
data = {
    "zkzh": (None, 201381100362),
    "bmxh": (None, 13812020013251),
    "jym": (None, '孔丽玄'),
    "button2": (None, '查询2')
}

response = requests.session().post(url, data=data)
response.encoding='utf-8'
print(response.text)