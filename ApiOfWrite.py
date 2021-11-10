# -*- coding: UTF-8 -*-
import xlsxwriter
import requests as req
import json,sys,random,time


Acc=open('ACCOUNT.txt','r')
account=json.loads(Acc.read())
Acc.close()
redirect_uri=r'https://login.microsoftonline.com/common/oauth2/nativeclient'
log_list=['']
###########################
# config选项说明
# 0：关闭  ， 1：开启
# allstart：是否全api开启调用，关闭默认随机抽取调用。默认0关闭
# rounds: 轮数，即每次启动跑几轮。
# rounds_delay: 是否开启每轮之间的随机延时，后面两参数代表延时的区间。默认0关闭
# api_delay: 是否开启api之间的延时，默认0关闭
# app_delay: 是否开启账号之间的延时，默认0关闭
########################################
config = {
         'allstart': 0,
         'rounds': 1,
         'rounds_delay': [0,0,5],
         'api_delay': [1,0,5],
         'app_delay': [0,0,5],
         }

#微软access_token获取
def getmstoken():
    #try:except?
    headers={
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
        html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
        if html.status_code < 300:
            print('微软密钥获取成功')
            break
        else:
            if retry_ == 3:
                print('微软密钥获取失败\n')
    jsontxt = json.loads(html.text)       
    return jsontxt['access_token']

#延时
def timeDelay(xdelay):
    if config[xdelay][0] == 1:
        time.sleep(random.randint(config[xdelay][1],config[xdelay][2]))
        
def apiReq(method,url,data='QAQ'):
    timeDelay('api_delay')
    global access_token
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    for retry_ in range(4):        
        if method == 'post':
            posttext=req.post(url,headers=headers,data=data)
        elif method == 'put':
            posttext=req.put(url,headers=headers,data=data)
        elif method == 'delete':
            posttext=req.delete(url,headers=headers)
        else :
            posttext=req.get(url,headers=headers)
        if posttext.status_code < 300:
            print('        操作成功')
            break
            #操作成功跳出循环
        else:
            if retry_ == 3:
                print('        操作失败')
    return posttext     
          

#上传文件到onedrive(小于4M)
def uploadFile(filesname,f):
    global log_list
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+r'/'+filesname+r':/content'
    if apiReq('put',url,f).status_code >= 300 :
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
    
        
# 发送邮件到自定义邮箱
def sendEmail(subject,content,Em):
    global log_list
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={
                'message':{
                          'subject': subject,
                          'body': {'contentType': 'Text', 'content': content},
                          'toRecipients': [{'emailAddress': {'address': Em}}],
                          },
                'saveToSentItems': 'true',
                }            
    if apiReq('post',url,json.dumps(mailmessage)).status_code >= 300 :
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
    	
	
#修改excel(这函数分离好像意义不大)
#api-获取itemid: https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')?select=name,id,webUrl
def excelWrite(filesname,sheet):
    global log_list
    try:
        print('    添加工作表')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+r'/'+filesname+r':/workbook/worksheets/add'
        data={
             "name": sheet
             }
        apiReq('post',url,json.dumps(data))
        print('    添加表格')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
        data={
             "address": "A1:D8",
             "hasHeaders": False
             }
        jsontxt=json.loads(apiReq('post',url,json.dumps(data)).text)
        print('    添加行')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
        rowsvalues=[[0]*4]*2
        for v1 in range(0,2):
            for v2 in range(0,4):
                rowsvalues[v1][v2]=random.randint(1,1200)
        data={
             "values": rowsvalues
             }
        apiReq('post',url,json.dumps(data))
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
        return 
    
def taskWrite(taskname):
    global log_list
    try:
        print("    创建任务列表")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
        data={
             "displayName": taskname
             }
        listjson=json.loads(apiReq('post',url,json.dumps(data)).text)
        print("    创建任务")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
        data={
             "title": taskname,
             }
        taskjson=json.loads(apiReq('post',url,json.dumps(data)).text)
        print("    删除任务")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
        apiReq('delete',url)
        print("    删除任务列表")
        url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
        apiReq('delete',url)
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
        return 
    
def teamWrite(channelname):
    global log_list
    #新建team
    try:
        print('    新建team')
        url=r'https://graph.microsoft.com/v1.0/teams'
        data={
             "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
             "displayName": channelname,
             "description": "My Sample Team’s Description"
             }
        apiReq('post',url,json.dumps(data))
        print("    获取team信息")
        url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
        teamlist = json.loads(apiReq('get',url).text)
        for teamcount in range(teamlist['@odata.count']):
            if teamlist['value'][teamcount]['displayName'] == channelname:
                #创建频道
                print("    创建team频道")
                data={
                     "displayName": channelname,
                     "description": "This channel is where we debate all future architecture plans",
                     "membershipType": "standard"
                     }
                url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels'
                jsontxt = json.loads(apiReq('post',url,json.dumps(data)).text)
                url=r'https://graph.microsoft.com/v1.0/teams/'+teamlist['value'][teamcount]['id']+r'/channels/'+jsontxt['id']
                print("    删除team频道")
                apiReq('delete',url)
                #删除teams
                print("    删除team")
                url=r'https://graph.microsoft.com/v1.0/groups/'+teamlist['value'][teamcount]['id']
                apiReq('delete',url)  
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
        return 
        
def onenoteWrite(notename):
    global log_list
    try:
        print('    创建笔记本')
        url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
        data={
             "displayName": notename,
             }
        notetxt = json.loads(apiReq('post',url,json.dumps(data)).text)
        print('    创建笔记本分区')
        url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks/'+notetxt['id']+r'/sections'
        data={
             "displayName": notename,
             }
        apiReq('post',url,json.dumps(data))
        print('    删除笔记本')
        url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
        apiReq('delete',url)
    except:
        print("        操作中断")
        if sys._getframe().f_code.co_name not in log_list:
            log_list=log_list+sys._getframe().f_code.co_name+','
        return
        
#获取access_token
client_id=account['client_id']
client_secret=account['client_secret']
ms_token=account['ms_token']
access_token=getmstoken()
print('')    
#获取天气
email=json.loads(open('EMAIL.txt').read())['email']
city=json.loads(open('EMAIL.txt').read())['city']
if city=='':
    city='Beijing'
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

#实际运行
print('发送邮件 ( 邮箱单独运行，每次运行只发送一次，防止封号 )')
if email!='':
    sendEmail('weather',weather,email)
else:
    print('尚未配置邮箱')
print('')
#其他api
for _ in range(1,config['rounds']+1):
    timeDelay('rounds_delay')  
    print('第 '+str(_)+' 轮\n')        
    timeDelay('app_delay')
    #生成随机名称
    filesname='QAQ'+str(random.randint(1,600))+r'.xlsx'
    #新建随机xlsx文件
    xls = xlsxwriter.Workbook(filesname)
    xlssheet = xls.add_worksheet()
    for s1 in range(0,4):
        for s2 in range(0,4):
            xlssheet.write(s1,s2,str(random.randint(1,600)))
    xls.close()
    xlspath=sys.path[0]+r'/'+filesname
    print('上传文件')
    with open(xlspath,'rb') as f:
        uploadFile(filesname,f)
    choosenum = random.sample(range(1, 5),2)
    if config['allstart'] == 1 or 1 in choosenum:
        print('excel文件操作')
        excelWrite(filesname,'QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 1 or 2 in choosenum:
        print('team操作')
        teamWrite('QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 1 or 3 in choosenum:
        print('task操作')
        taskWrite('QVQ'+str(random.randint(1,600)))
    if config['allstart'] == 1 or 4 in choosenum:
        print('onenote操作')
        onenoteWrite('QVQ'+str(random.randint(1,600)))
    print('-')
