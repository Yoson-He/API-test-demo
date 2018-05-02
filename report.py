# -*- coding: utf-8 -*-
# @Time    : 2018/2/7 16:52
# @Author  : Yoson
# @File    : report.py
# @Software: PyCharm
import time
import xlsxwriter


def get_format(wd, option={}):
    return wd.add_format(option)


def get_format_center(wd,num=0):
    '''
    设置居中
    '''
    return wd.add_format({'valign': 'vcenter'})
#'align': 'center',


def set_border_(wd, num=1):
    return wd.add_format({}).set_border(num)


def _write_center(worksheet, cl, data, wd):
    '''
    写数据
    '''
    return worksheet.write(cl, data, get_format_center(wd))


def init(workbook, worksheet, config, all_count, pass_count):

    # 设置列行的宽高
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 35)
    worksheet.set_column("C:C", 15)
    worksheet.set_column("D:D", 20)

    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)
    worksheet.set_row(5, 30)

    # worksheet.set_row(0, 200)

    define_format_H1 = get_format(workbook, {'bold': True, 'font_size': 18})
    define_format_H2 = get_format(workbook, {'bold': True, 'font_size': 14})
    define_format_H1.set_border(1)

    define_format_H2.set_border(1)
    define_format_H1.set_align("center")
    define_format_H2.set_align("center")
    define_format_H2.set_bg_color("blue")
    define_format_H2.set_color("#ffffff")
    # Create a new Chart object.

    worksheet.merge_range('A1:D1', '测试报告总概况', define_format_H1)
    worksheet.merge_range('A2:D2', '测试概括', define_format_H2)

    _write_center(worksheet, "A3", '项目名称', workbook)
    _write_center(worksheet, "A4", '接口版本', workbook)
    _write_center(worksheet, "A5", '脚本语言', workbook)
    _write_center(worksheet, "A6", 'API_HOST', workbook)

    _write_center(worksheet, "B3", config['project_name'], workbook)
    _write_center(worksheet, "B4", config['api_version'], workbook)
    _write_center(worksheet, "B5", "Python", workbook)
    _write_center(worksheet, "B6", config['host'], workbook)

    _write_center(worksheet, "C3", "接口总数", workbook)
    _write_center(worksheet, "C4", "通过总数", workbook)
    _write_center(worksheet, "C5", "失败总数", workbook)
    _write_center(worksheet, "C6", "测试日期", workbook)

    _write_center(worksheet, "D3", all_count, workbook)
    _write_center(worksheet, "D4", pass_count, workbook)
    _write_center(worksheet, "D5", all_count-pass_count, workbook)
    _write_center(worksheet, "D6", time.strftime('%Y-%m-%d', time.localtime(time.time())), workbook)

    pie(workbook, worksheet)


def pie(workbook, worksheet):
    '''生成饼形图'''
    chart1 = workbook.add_chart({'type': 'pie'})
    chart1.add_series({'name': '接口测试统计', 'categories': '=测试概况!$C$4:$C$5', 'values': '=测试概况!$D$4:$D$5'})
    chart1.set_title({'name': '接口测试统计'})
    chart1.set_style(10)
    worksheet.insert_chart('A9', chart1, {'x_offset': 25, 'y_offset': 10})


def test_detail(workbook, worksheet, test_result):

    # 设置列行的宽高
    worksheet.set_column("A:A", 6)
    worksheet.set_column("B:B", 30)
    worksheet.set_column("C:C", 6)
    worksheet.set_column("D:D", 7)
    worksheet.set_column("E:E", 30)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 30)
    worksheet.set_column("H:H", 30)
    worksheet.set_column("I:I", 30)
    worksheet.set_column("J:J", 8)

    '''
    worksheet.set_row(1, 30)
    worksheet.set_row(2, 30)
    worksheet.set_row(3, 30)
    worksheet.set_row(4, 30)
    worksheet.set_row(5, 30)
    worksheet.set_row(6, 30)
    worksheet.set_row(7, 30)
    '''

    worksheet.merge_range('A1:J1', '测试详情', get_format(workbook, {'bold': True, 'font_size': 18 ,'align': 'center','valign': 'vcenter','bg_color': 'blue', 'font_color': '#ffffff'}))
    _write_center(worksheet, "A2", '用例ID', workbook)
    _write_center(worksheet, "B2", '用例名称', workbook)
    _write_center(worksheet, "C2", 'API_ID', workbook)
    _write_center(worksheet, "D2", 'Method', workbook)
    _write_center(worksheet, "E2", 'URL', workbook)
    _write_center(worksheet, "F2", 'Body', workbook)
    _write_center(worksheet, "G2", '预期结果', workbook)
    _write_center(worksheet, "H2", '实际结果', workbook)
    _write_center(worksheet, "I2", '接口返回值', workbook)
    _write_center(worksheet, "J2", '测试结果', workbook)

    temp = 3
    for item in test_result:
        _write_center(worksheet, "A"+str(temp), item["case_id"], workbook)
        _write_center(worksheet, "B"+str(temp), item["case_name"], workbook)
        _write_center(worksheet, "C"+str(temp), item["api_id"], workbook)
        _write_center(worksheet, "D"+str(temp), item["method"], workbook)
        _write_center(worksheet, "E"+str(temp), item["url"], workbook)
        _write_center(worksheet, "F"+str(temp), str(item["body_value"]), workbook)
        _write_center(worksheet, "G"+str(temp), item["expected_result"], workbook)
        _write_center(worksheet, "H"+str(temp), str(item["actual_result"]), workbook)
        _write_center(worksheet, "I"+str(temp), item["respone_body"], workbook)
        if item["pass_or_fail"] == "FAIL":
            format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_color': 'red'})
            worksheet.write("J"+str(temp), 'FAIL', format)
        else:
            format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_color': 'green'})
            worksheet.write("J" + str(temp), 'PASS', format)
        temp = temp + 1

