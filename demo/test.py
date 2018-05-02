
import sys
import operator
import json
import requests
import xlrd
import re
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.header import Header

def str_clean(string):
    """清除空格和换行符,把中文逗号换成英文逗号"""
    try:
        string = string.replace(" ", "").replace("\n", "").replace("，", ",")
        return string
    except Exception:
        print("字符处理出错！！！")

def getTestData(file_path):
    """ 读取文件，获取测试数据 """
    try:
        data = xlrd.open_workbook(file_path)

        # 从config表的1行2列获取host
        config_table = data.sheet_by_name("config")
        host = str_clean(config_table.cell(0, 1).value)
        if host == "":
            print("host为空，请先配置host！！！")
            return None

        #从API表获取api信息并保存到api_list
        api_list = []
        api_table = data.sheet_by_name("API")
        api_count = api_table.nrows
        if api_count == 0:
            print("API表无数据！！！")
            return None

        for i in range(1, api_count):
            api_id = api_table.cell_value(i, 0)
            if api_id == "":
                print("api_id不能为空！！！位置：API表第%s行" % i+1)
                return None

            path = str_clean(api_table.cell_value(i, 1))
            if path == "":
                print("path不能为空！！！位置：API表第%s行" % i+1)
                return None

            method = str_clean(api_table.cell_value(i, 2))
            if path == "":
                print("method不能为空！！！位置：API表第%s行" % i+1)
                return None

            headers = str_clean(api_table.cell_value(i, 3))
            api_list.append({"api_id": api_id, "path": path, "method": method, "headers": headers})

        # 从testCase表中获取测试用例信息并保存到testCase_list
        testCase_list = []
        testCase_table = data.sheet_by_name("test_case")
        testCase_count = testCase_table.nrows
        if testCase_count == 0:
            print("testCase表无数据！！！")
            return None

        for i in range(1,testCase_count):
            case_id = testCase_table.cell_value(i, 0)
            if case_id == "":
                print("case_id不能为空！！！位置：testCase表第%s行" % i+1)
                return None

            api_id = testCase_table.cell_value(i, 1)
            if api_id == "":
                print("api_id不能为空！！！位置：testCase表第%s行" % i+1)
                return None

            level = testCase_table.cell_value(i, 2)
            uri_parameter_value = str_clean(testCase_table.cell_value(i, 3))
            body_parameter_value = str_clean(testCase_table.cell_value(i, 4))

            expected_result = str_clean(testCase_table.cell_value(i, 5))
            if expected_result == "":
                print("expected_result不能为空！！！位置：testCase表第%s行" % i+1)
                return None

            testCase_list.append({"case_id": case_id, "api_id": api_id, "level": level, "uri_parameter_value": uri_parameter_value, "body_parameter_value": body_parameter_value, "expected_result": expected_result})
        print("测试数据已读取")
        return host, api_list, testCase_list
    except Exception:
        print("读取文件出错！！！")
        sys.exit()

def pathProcessor(path):
    """替换path中的参数占位符，如：api/LocalTours/{id}/prices/{useDate}?date={date} 替换成 api/LocalTours/%s/prices/%s?date=%s"""
    return re.sub(r"{.*?}", "%s", path, count=0)

def dataHandle(host, api_list, test_case_list):
    """处理获取的测试数据，拼装URL，返回一个可执行的测试用例列表"""
    try:
        # 替换path中的参数占位符
        api_count = len(api_list)
        for i in range(api_count):
            api_list[i]["path"] = pathProcessor(api_list[i]["path"])

        test_case_count = len(test_case_list)
        for i in range(test_case_count):
            uri_parameter_value = test_case_list[i]["uri_parameter_value"].split(",")#获取用例uri参数值，有多个参数时，以逗号分隔，返回一个参数值列表

            #把用例的uri参数值拼接到对应api的path上
            for j in range(api_count):
                if test_case_list[i]["api_id"] == api_list[j]["api_id"]:

                    if api_list[j]["path"].count("%s") != len(uri_parameter_value):
                        print("参数个数与参数值个数不匹配！！！case_id：%s，api_id：%s" % (test_case_list[i]["case_id"], test_case_list[i]["api_id"]))
                        return None

                    path = api_list[j]["path"]%tuple(uri_parameter_value)#把path中的占位符%s替换为具体的参数值
                    url = "http://" + host + "/" + path
                    test_case_list[i]["url"] = url
                    test_case_list[i]["method"] = api_list[j]["method"]
                    test_case_list[i]["headers"] = api_list[j]["headers"]
                    break

                if j == api_count-1:
                    print("用例找不到对应的api！！！，case_id：%s" % test_case_list[i]["case_id"])
                    return None

        print("数据处理完成")
        return test_case_list
    except Exception:
        print("数据处理出错！！！")
        sys.exit()

def runTest(testCase_list):
    """执行测试用例，返回测试结果"""
    try:
        test_result_all = []
        all_count = len(testCase_list)
        pass_count = 0
        fail_count = 0
        for test_case in testCase_list:
            headers = test_case["headers"]
            if headers != "":
                headers = eval(headers)

            if test_case["method"] == "post":
                test_result = requests.request("post", test_case["url"], data=test_case["body_parameter_value"], headers=headers)
            elif test_case["method"] == "get":
                test_result = requests.request("get", test_case["url"], headers=headers)

            status_code = test_result.status_code
            test_result = str_clean(str(test_result.text))
            expected_result = str_clean(str(test_case["expected_result"]))

            if operator.eq(expected_result, test_result):
                status = "PASS"
                pass_count += 1
            else:
                status = "FAIL"
                fail_count += 1
            test_result_all.append({"case_id": test_case["case_id"], "api_id": test_case["api_id"], "status": status, "status_code": status_code, "test_result": test_result, "expected_result": expected_result})
        print("用例已执行完\n本次执行用例总数：%s， 通过%s个，失败%s个" % (all_count, pass_count, fail_count))
        return test_result_all, all_count, pass_count, fail_count
    except Exception:
        print("执行用例出错！！！")
        sys.exit()

def resultProcessor(test_result_all,all_count, pass_count, fail_count):
    #处理测试结果，格式化测试报告
    test_report = "<html><body><p>本次执行用例总数：%s， 通过%s个，失败%s个" % (all_count, pass_count, fail_count)
    test_report += "<p><table border='1px'><tr><th>status</th><th>status_code</th><th>case_id</th><th>api_id</th><th>expected_result</th><th>test_result</th></tr>"
    for each in test_result_all:
        test_report += "<tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td></tr>" % (each["status"], each["status_code"], each["case_id"], each["api_id"], each["expected_result"], each["test_result"])

    test_report += "</table></body></html>"

    return test_report

def sendMail(text):
    try:
        sender = "TestAutomated@qulv.com"  # 发件人邮箱账号
        password = "ABCabc123"  # 发件人邮箱密码
        receivers = ["yongzhen-he@qulv.com", "Hox@qulv.com"]  # 收件人邮箱账号

        msg = MIMEText(text, "html", "utf-8")
        msg["From"] = formataddr(["接口自动化", sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg["To"] = str(receivers)[1:-1]  # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg["Subject"] = "接口测试报告"        # 邮件主题
     
        server = smtplib.SMTP("mail.qulv.com", 25)  # 创建邮件服务器连接
        server.starttls() #TLS加密
        server.login(sender, password)  # 登录邮件服务器，括号中对应的是发件人邮箱账号、邮箱密码
        server.sendmail(sender, receivers, msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 关闭连接
        print("邮件发送成功")
    except smtplib.SMTPException as reason:
        print(reason)
        sys.exit()

if __name__ == "__main__":
    #读取测试数据
    test_data = getTestData(r"testData.xlsx")
    host = test_data[0]
    api_list = test_data[1]
    testCase_list = test_data[2]

    #处理测试数据，形成有效的测试用例
    testCase_list = dataHandle(host, api_list, testCase_list)

    #执行测试用例
    r = runTest(testCase_list)

    #处理测试结果
    test_result_all = r[0]
    all_count = r[1]
    pass_count = r[2]
    fail_count = r[3]
    test_report = resultProcessor(test_result_all, all_count, pass_count, fail_count)

    #发送测试报告邮件
    sendMail(test_report)





