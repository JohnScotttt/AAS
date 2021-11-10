# -*- coding: UTF-8 -*-
from os import remove
import requests as req
import json

Acc=open('ACCOUNT.txt','r')
account=json.loads(Acc.read())
Acc.close()

# accountkey=['client_id','client_secret','ms_token']
redirect_uri = r'https://login.microsoftonline.com/common/oauth2/nativeclient'
#微软refresh_token获取
def getmstoken():
    ms_headers={
               'Content-Type':'application/x-www-form-urlencoded'
               }
    data={
         'grant_type': 'refresh_token',
         'refresh_token': ms_token,
         'client_id':client_id,
         'client_secret':client_secret,
         'redirect_uri':redirect_uri,
         }
    for retry_ in range(4):
        html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=ms_headers)
        if html.status_code < 300:
            print('微软密钥获取成功')
            break
        else:
            if retry_ == 3:
                print('微软密钥获取失败')
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    return refresh_token

#调用 
client_id=account['client_id']
client_secret=account['client_secret']
ms_token=account['ms_token']
account['ms_token']=getmstoken()
#写入
remove('ACCOUNT.txt')
Acc=open('ACCOUNT.txt','w')
Acc.write(json.dumps(account))
Acc.close()
