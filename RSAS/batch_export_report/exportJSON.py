# -*- coding: utf-8 -*-
"""
@Time ： 2022/6/27 9:57
@Auth ： AlairLee
@File ：exportJSON.py
@IDE ：PyCharm
"""

import xlrd, re

# Open the Workbook
workbook = xlrd.open_workbook('资产路径')

# Open the worksheet
worksheet = workbook.sheet_by_index(0)

rows = worksheet.nrows
lists = []


def getName(i):
    name = worksheet.cell_value(i, 1) + '-' + worksheet.cell_value(i, 5)
    return name


def creatClass(i):
    jsData = {
        'report_name': worksheet.cell_value(i, 1) + "-" + worksheet.cell_value(i, 5),
        'export_ip': [worksheet.cell_value(i, 8)],
    }
    return jsData


def getIP(i):
    return worksheet.cell_value(i, 8)


def get_tasklist():
    for i in range(1, rows):
        if getName(i) in str(lists):
            for j in lists:
                if j['report_name'] == getName(i):
                    # 存在10.31.20.10时，判断10.31.20.1为false
                    if getIP(i) not in j['export_ip']:
                        j['export_ip'].append(getIP(i))
        else:
            lists.append(creatClass(i))
    return lists

