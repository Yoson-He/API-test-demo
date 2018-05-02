import ConfigParser
import os

#通过自带的ConfigParser模块，读取邮件发送的配置文件，以字典的形式返回
def get_conf():
    conf_file=ConfigParser.ConfigParser()
    conf_file.read(os.path.join(os.getcwd(),"conf.ini"))
    conf={}
    conf["sender"]=conf_file.get("email","sender")
    conf["receiver"]=conf_file.get("email","receiver")
    conf["smtpserver"]=conf_file.get("email","smtpserver")
    conf["username"]=conf_file.get("email","username")
    conf["password"]=conf_file.get("email","password")
    return conf

#使用logging模块，用来作为测试日志，记录测试中系统产生的信息
import logging
log_file=os.path.join(os.getcwd(),"log/test.log")
log_format="[%(asctime)s] [%(levelname)s] %(massage)s"
logging.basicConfig(format=log_format,filename=log_file,filemode="txt",level=logging.DEBUG)
console=logging.StreamHandler()
console.setLevel(logging.DEBUG)
formatter=logging.Formatter(log_format)
console.setFomatter(formatter)
logging.getLogger("").addHandler(console)

#读取testcase.excel文件，获取测试数据，调用interfaceTest方法，将结果保存到errorCase列表中
import xlrd,hashlib,json
def runTest(testCaseFile):
    testCaseFile=os.path.join(os.getwcd(),testCaseFile)
    if not os.path.exists(testCaseFile):
        logging.error("测试用例文件不存在！")
        sys.exit()
    test_case=xlrd.open_workbook(testCaseFile)
    table=test_case.sheet_by_index(0)
    errorCase=[] #用于保存接口返回的内容和http状态码
    
    s=None
    for i in range(1,table.nrows):
        if table.cell(i,9).value.replace("\n","").replace("\r","") != "Yes":
            continue
        num= str(int(table.cell(i,0).value)).replace("\n","").replace("\r","")
        api_purpose=table.cell(i,1).value.replace("\n","").replace("\r","")
        api_host=table.cell(i,2).value.replace("\n","").replace("\r","")
        request_method=table.cell(i,4).value.replace("\n","").replace("\r","")
        requset_data_type=table.cell(i,5).value.replace("\n","").replace("\r","")
        request_data=table.cell(i,6).value.replace("\n","").replace("\r","")
        encryption=table.cell(i,7).value.replace("\n","").replace("\r","")
        check_point=table.cell(i,8).value

    if encryption == "MD5":  #如果数据采用MD5加密，则优先将数据加密
        request_data=json.loads(request_data)
        request_data["pwd"]=md5Encode(request_data["pwd"])
    status,resp,s=interfaceTest(num,apipurpose,api_host,request_url,request_data,check_point,request_method,request_data_type,s)
    if status != 200 or check_point not in resp: #如果状态码不是200，或者返回值没有检查点的内容，那么证明接口产生错误，保存错误信息
        errorCase.append((num+" "+api_purpose,str(status),"http://"+api_host+request_url,resp))

#接受runTest的传参，利用request构造http请求
import requests
def interfaceTest(num,api_purpose,api_host,request_method,request_data_tyoe,request_data,check_point,s=None):
    headers={"conttent-type":"application/x-www-form-urlencoded; charset-UTF-8","X-Requested-With":"XMLHTTPRequest","Connetion":"keep-alive","Referer":"http://"+api_host,"User-Agent":""}

    if s==None:
        s=requests.session()
    if request_method=="post":
        if request_url != "/login":
            r=s.post(url="http://"+api_host+request_url,data=json.loads(request_data),headers=headers) #由于此处数据没有经过加密，所以需要把json格式字符串解码转换成python对象
        elif request_url=="/login":
            s=requests.session()
            r=s.post(url="http://"+api_host+request_url,data=request_data,headers=headers) #由于登录密码不能明文传输，采用md5加密，在之前的代码中已经进行json.loads()转换，so此处不需要解码

    else:
        logging.error(num+" "+api_purpose+"HTTP请求方法错误，请确认[Request Method]字段是否正确！")
        s=None
        return 400,resp,s
    status=r.status_code
    resp=r.text
    print(resp)
    if status==200:
        if re.search(check_point,str(r.text)):
            logging.info(num+" "+apipurpose+"成功，"+str(status)+","+str(r.text))
            return status,resp,s
        else:
            logging.error(num+" "+apipurpose+"失败！！！"+str(status)+","+str(r.text))
            return 200,resp,None
        else:
            logging.error(num+" "+apipurpose+"失败！！！"+str(status)+","+str(r.text))
            return status,resp.decode("utf-8"),None

import hashlib
def md5Encode(data):
    hashobj=hashlib.md5()
    hashobj.update(data.encode("utf-8"))
    return hashobj.hexdigest()

def sendMail(text):
    mail_info=get_conf()
    sender=mail_info["sender"]
    receiver=mail_info["receiver"]
    smtpserver=mail_info["smtpserver"]
    username=mail_info["username"]
    password=mail_info["password"]
    subject="[AutomationTest]接口自动化测试报告通知"
    msg=MIMEText(text,"html","utf-8")
    msg["Subject"]=subject
    msg["From"]=sender
    msg["To"]="".join(receiver)
    smtp.connect(smtpserver)
    smpt.login(username,password)
    smpt.sendmail(sender,receiver,msg.as_string())
    smpt.quit()

def main():
    errorTest =runTest("")
    if len(errorTest)>0:
        html="<html><body>接口自动化定期扫描，共有"+str(len(errorTest))+"个异常接口，列表如下："+"""</p><table><tr><th style="with:100px;text-align:left">接口</th>
    <th style="with:50px;text-align:left>状态</th><th style="with:200px;text-align:left">接口地址</th><th style=text-align:left>接口返回值</th>"""

    for test in errorTest:
        html=html+"<tr><td style=text-align:left>"+test[0]+"</td><td style=text-align:left>"+test[1]+"</td><td style=text-align:left>"+test[2]+"</td></tr>"
        sendMail(html)
    if __name__=="__main__":
        main()
