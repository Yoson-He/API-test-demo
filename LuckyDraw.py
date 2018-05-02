# -*- coding: utf-8 -*-
# @Time    : 2018/4/3 10:02
# @Author  : Yoson
# @File    : LuckyDraw.py
# @Software: PyCharm
import xlrd


def get_participators(file_path):
    # 1、获取抽奖对象数据源
    participators = []
    data = xlrd.open_workbook(file_path)
    table = data.sheet_by_index(0)
    rows = table.nrows
    for i in range(rows):
        participators.append(table.cell_value(i, 0)+" "+table.cell_value(i, 1))
    return participators


def award_setting():
    # 2、设置奖项（奖项名称、个数、奖品）
    award = {}
    award["name"] = input("奖项名称：")
    award["prize"] = input("奖品：")
    award["times"] = input("抽奖次数：")
    award["count"] = int(input("每次抽奖个数："))

    return award

def draw():
    # 3、选择抽取的奖项进行抽奖


#4、抽奖结果展示&存档

if __name__ == "__main__":
    # 1、获取抽奖对象数据源
    participators = get_participators('1.xlsx')
    print(participators)

    # 2、设置奖项（奖项名称、个数、奖品）
    awards = []
    awards.append(award_setting())
    print(awards)

    while 1:
        choice = int(input("请选择：1、抽奖；2、添加奖项；3、修改奖项；4、作废抽奖结果；5、设置参与者\n"))
        if choice == 1:
            i = 0
            for each in awards:
                print(str(i)+"、"+str(awards[i]))
                j = int(input("请选择要抽取的奖项："))
        elif choice == 2:
            pass
        elif choice == 3:
            pass
        elif choice == 4:
            pass
        elif choice == 5:
            pass
        else:
            print("输入有误\n")
