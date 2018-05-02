# -*- coding: utf-8 -*-
# @Time    : 2018/2/6 15:42
# @Author  : Yoson
# @File    : API_Test.py
# @Software: PyCharm
import operator
import re
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import time
import requests
import xlrd
import xlsxwriter

import report


def str_clean(string):
    """
    清除空格和换行符,
    把中文符号换成英文符号
    """
    try:
        string = string.replace(" ", "").replace("\n", "").replace("\r", "")
        string = string.replace("，", ",").replace("“", "\"").replace("”", "\"").replace("‘", "\'").replace("’", "\'")
        string = string.replace("；", ";").replace("：", ":")
        return string
    except Exception:
        print("字符处理出错！！！")


def get_test_data(file_path):
    """
    读取文件，获取测试数据
    file_path: 文件路径（包括文件名，目标文件必须是Excel文件）
    """
    try:
        print("正在获取数据……")
        data = xlrd.open_workbook(file_path)

    # 从config表获取配置信息（host、项目名称、接口版本）
        config_table = data.sheet_by_name("config")
        config = {}
        config["project_name"] = str_clean(config_table.cell(0, 1).value)
        config["api_version"] = str_clean(config_table.cell(1, 1).value)
        config["receivers"] = str_clean(config_table.cell(3, 1).value)
        config["host"] = str_clean(config_table.cell(2, 1).value)
        if config["host"] == "":
            print("host为空，请先配置host！！！")
            return None
        else:
            print('读取host：%s' % config["host"])

    # 从API表获取api信息并保存到api_list
        api_list = []
        api_table = data.sheet_by_name("API")
        api_count = api_table.nrows
        if api_count == 0:
            print("API表无数据！！！")
            return None

        for i in range(1, api_count):
            api_id = api_table.cell_value(i, 0)
            if api_id == "":
                print("api_id不能为空！！！位置：API表第%s行" % str(i+1))
                return None

            path = str_clean(api_table.cell_value(i, 2))
            if path == "":
                print("path不能为空！！！位置：API表第%s行" % str(i+1))
                return None

            method = str_clean(api_table.cell_value(i, 3))
            if method == "":
                print("method不能为空！！！位置：API表第%s行" % str(i+1))
                return None

            if method not in ('post', 'POST', 'get', 'GET'):
                print("method录入有误：API表第%s行" % str(i+1))
                return None

            headers = str_clean(api_table.cell_value(i, 4))
            if headers != "":
                try:
                    headers = eval(headers)
                except Exception:
                    print('headers格式有误！API表第%s行' % str(i+1))
                    return None
            api_list.append({"api_id": api_id, "path": path, "method": method, "headers": headers})

        print('API信息读取完毕！')

    # 从testCase表中获取测试用例信息并保存到testCase_list
        test_case_list = []
        test_case_table = data.sheet_by_name("test_case")
        test_case_count = test_case_table.nrows
        if test_case_count == 0:
            print("test_case表无数据！！！")
            return None

        for i in range(1, test_case_count):
            case_id = test_case_table.cell_value(i, 0)
            if case_id == "":
                print("case_id不能为空！！！位置：test_case表第%s行" % str(i+1))
                return None
            case_name = test_case_table.cell_value(i, 1)
            if case_id == "":
                print("case_id不能为空！！！位置：test_case表第%s行" % str(i + 1))
                return None
            api_id = test_case_table.cell_value(i, 2)
            if api_id == "":
                print("api_id不能为空！！！位置：test_case表第%s行" % str(i+1))
                return None

            uri_parameters_value = str_clean(test_case_table.cell_value(i, 3))
            if uri_parameters_value != "":
                try:
                    uri_parameters_value = eval(uri_parameters_value)
                except Exception:
                    print('uri_parameters_value格式有误！位置：test_case表第%s行' % str(i+1))
                    return None

            query_parameters_value = str_clean(test_case_table.cell_value(i, 4))
            if query_parameters_value != "":
                try:
                    query_parameters_value = eval(query_parameters_value)
                except Exception:
                    print('query_parameters_value格式有误！位置：test_case表第%s行' % str(i + 1))
                    return None

            body_value = str_clean(test_case_table.cell_value(i, 5))
            if body_value != "":
                try:
                    body_value = eval(body_value)
                except Exception:
                    print("body_value格式有误！位置：test_case表第%s行" % str(i+1))
                    return None

            expected_result = str_clean(test_case_table.cell_value(i, 6))
            if expected_result == "":
                print("expected_result不能为空！！！位置：test_case表第%s行" % str(i+1))
                return None

            test_case_list.append({"case_id": case_id, "case_name": case_name, "api_id": api_id, "query_parameters_value": query_parameters_value,
                                   "uri_parameters_value": uri_parameters_value, "body_value": body_value,
                                   "expected_result": expected_result})

        print("用例数据读取完毕！")
        return config, api_list, test_case_list
    except Exception:
        print("读取文件出错！！！")
        return None


def data_handle(host, api_list, test_case_list):
    """
    处理获取的测试数据，拼装URL，返回一个可执行的测试用例列表
    host: host地址
    api_list: 从文件读取的api列表
    test_case_list: 从文件读取的用例列表
    """

    try:
        print("数据处理中……")
        # 将path中的参数占位符替换成%s占位符
        api_count = len(api_list)
        for i in range(api_count):
            api_list[i]["path"] = re.sub(r"{.*?}", "%s", api_list[i]["path"], count=0)

        test_case_count = len(test_case_list)
        for i in range(test_case_count):
            uri_parameters_value = test_case_list[i]["uri_parameters_value"]

            # 把用例的uri参数值拼接到对应api的path上
            for j in range(api_count):
                if test_case_list[i]["api_id"] == api_list[j]["api_id"]:
                    path = api_list[j]["path"]
                    if path.count("%s") != len(uri_parameters_value):
                        print("参数个数与参数值个数不匹配！！！case_id：%s，api_id：%s"
                              % (test_case_list[i]["case_id"], test_case_list[i]["api_id"]))
                        return None

                    if uri_parameters_value != "":
                        path = path % tuple(uri_parameters_value.values())  # 把path中的占位符替换为具体的参数值
                    url = "http://" + host + "/" + path
                    test_case_list[i]["url"] = url
                    test_case_list[i]["method"] = api_list[j]["method"]
                    test_case_list[i]["headers"] = api_list[j]["headers"]
                    break

                if j == api_count - 1:
                    print("用例找不到对应的api！！！，case_id：%s" % test_case_list[i]["case_id"])
                    return None

        print("数据处理完成")
        return test_case_list
    except Exception:
        print("数据处理出错！！！")
        return None


def run_test(test_case_list):

    '''
    批量执行用例
    '''

    try:
        print("开始执行用例……")
        test_result_all = []
        all_count = len(test_case_list)
        pass_count = 0
        for test_case in test_case_list:
            if test_case["method"] in ("post", "POST"):
                respone_body = requests.request("post", test_case["url"], data=test_case["body_value"], headers=test_case["headers"])
            elif test_case["method"] in ("get", "GET"):
                if test_case["query_parameters_value"] != '':
                    respone_body = requests.request("get", test_case["url"], headers=test_case["headers"], params=test_case["query_parameters_value"])
                else:
                    respone_body = requests.request("get", test_case["url"], headers=test_case["headers"])

            expected_result = str_clean(str(test_case["expected_result"]))
            checkout = _actual_result_check(respone_body, expected_result)
            if checkout[0] == 'PASS':
                pass_count += 1
            test_result_all.append({"case_id": test_case["case_id"], "case_name": test_case["case_name"],
                                    "api_id": test_case["api_id"], "method": test_case["method"],
                                    "url": test_case["url"], "body_value": test_case["body_value"],
                                    "expected_result": expected_result, "actual_result": checkout[2],
                                    "respone_body": respone_body.text, "pass_or_fail": checkout[0]})

        print("用例已执行完，本次执行用例总数：%s" % all_count)
        print("通过用例数：%s" % pass_count)
        return all_count, pass_count, test_result_all
    except Exception:
        print("执行用例出错！！！")
        return None


def _actual_result_check(respone_body, expected_result):
    '''
    用例返回值校验
    '''
    try:
        actual_result = []
        pass_or_fail = 'PASS'
        status_code = respone_body.status_code
        expected_result = expected_result.split("test[")
        expected_result = expected_result[1:]
        for each in expected_result:
            if each[-1] == ";":
                each = each[:-1]
            if re.match(r'(.*?)\]:(.*)', each) is not None:
                name = re.match(r'(.*?)\]:(.*)', each).group(1)
                string = re.match(r'(.*?)\]:(.*)', each).group(2)
                check = 'PASS'

                # 状态码等于
                if re.match(r'responseCode==(.*)', string) is not None:
                    if str(status_code) != re.match(r'responseCode==(.*)', string).group(1):
                        check = 'FAIL'
                        pass_or_fail = 'FAIL'

                # 返回值等于
                elif re.match(r'responseBody==(.*)', string) is not None:
                    if str_clean(respone_body.text) != str_clean(re.match(r'responseBody==(.*)', string).group(1)):
                        check = 'FAIL'
                        pass_or_fail = 'FAIL'

                # 返回值包含
                elif re.match(r'responseBody.has\((.*)\)', string) is not None:
                    if re.match(r'responseBody.has\((.*)\)', string).group(1) not in respone_body.text:
                        check = 'FAIL'
                        pass_or_fail = 'FAIL'

                # 返回值不包含
                elif re.match(r'responseBody.without\((.*)\)', string) is not None:
                    if re.match(r'responseBody.without\((.*)\)', string).group(1) in respone_body.text:
                        check = 'FAIL'
                        pass_or_fail = 'FAIL'

                # 检查一个json值
                elif re.match(r'data\[(.*?)]==(.*)', string) is not None:
                    key = re.match(r'data\[(.*?)]==(.*)', string).group(1)
                    value = re.match(r'data\[(.*?)]==(.*)', string).group(2)
                    i = respone_body.json()
                    if isinstance(i, list):
                        i = i[0]

                    if key in i:
                        actual = i[key]
                    elif 'Data' in i:
                        actual = i['Data'][key]
                    elif 'data' in i:
                        actual = i['data'][key]

                    # 比对的值是数字则将预期结果统一转为浮点型（因为读取的预期值默认都是字符型）
                    if type(actual) in (int, float):
                        value = float(value)

                    # 预期结果是ture、false布尔值时，则转化为布尔类型（因为获取的测试数据默认都是字符类型）
                    elif str.lower(value) == 'true':
                        value = (str.lower(value) == 'true')
                    elif str.lower(value) == 'false':
                        value = (str.lower(value) != 'false')

                    elif str.lower(value) == 'null':
                        value = None

                    if actual != value:
                        check = 'FAIL'
                        pass_or_fail = 'FAIL'

                else:
                    return None
                actual_result.append(check+'：'+name)
            else:
                return None

        return pass_or_fail, status_code, actual_result
    except Exception:
        print("用例返回值校验出错！！！")
        return None

def send_mail(receivers):
    '''
    发送邮件
    '''
    try:
        sender = "TestAutomated@qulv.com"  # 发件人邮箱账号
        password = "ABCabc123"  # 发件人邮箱密码
        # receivers = ["yongzhen-he@qulv.com"]  # 收件人邮箱账号

        msg = MIMEMultipart()
        msg["From"] = formataddr(["接口自动化", sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg["To"] = str(receivers)[1:-1]  # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg["Subject"] = "接口测试报告"  # 邮件主题

        msg.attach(MIMEText('接口测试报告，详见附件……', 'plain', 'utf-8'))  # 邮件正文内容

        # 构造附件，传送当前目录下的 test.txt 文件
        att1 = MIMEText(open('report.xlsx', 'rb').read(), 'base64', 'utf-8')
        att1["Content-Type"] = 'application/octet-stream'
        att1["Content-Disposition"] = 'attachment; filename="report.xlsx"'  # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
        msg.attach(att1)
        server = smtplib.SMTP("mail.qulv.com", 25)  # 创建邮件服务器连接
        server.starttls()  # TLS加密
        server.login(sender, password)  # 登录邮件服务器，括号中对应的是发件人邮箱账号、邮箱密码
        server.sendmail(sender, receivers, msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 关闭连接
        print("邮件发送成功")
    except smtplib.SMTPException as reason:
        print(reason)
        sys.exit()


if __name__ == "__main__":

    # 读取测试数据
    test_data = get_test_data(r"testData.xlsx")
    if test_data is None:
        input()
    config = test_data[0]
    api_list = test_data[1]
    test_case_list = test_data[2]

    # 处理测试数据，形成有效的可执行的测试用例
    test_case_list = data_handle(config["host"], api_list, test_case_list)
    if test_case_list is None:
        input()

    # 执行测试用例
    data = run_test(test_case_list)
    if data is None:
        input()
    all_count = data[0]
    pass_count = data[1]
    test_result = data[2]

    # 生成测试报告
    workbook = xlsxwriter.Workbook('report'+str(time.strftime("%Y-%m-%d %H%M%S", time.localtime()))+'.xlsx')
    worksheet = workbook.add_worksheet("测试概况")
    worksheet2 = workbook.add_worksheet("测试详情")
    report.init(workbook, worksheet, config, all_count, pass_count)
    report.test_detail(workbook, worksheet2, test_result)
    workbook.close()

    # 发送测试报告邮件
    if config["receivers"] != "":
        send_mail(config["receivers"].split(";"))

    input()


