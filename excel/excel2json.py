#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
处理Excel文件
'''

from xlrd import open_workbook
import re

filePath = "C://Users//MISuser//Desktop//业务提示消息数据.xls"
resultFilePath = "C:/Users/MISuser/Desktop/test.txt"

wb = open_workbook(filePath)
print("表单数量:", wb.nsheets)
print("表单名称:", wb.sheet_names())

sheet = wb.sheet_by_index(1)
print("表单 %s 共 %d 行 %d 列" % (sheet.name, sheet.nrows, sheet.ncols))

# 获取第一行的列名
fieldList = list(map(lambda x: re.compile("'(.*)'").findall(x), list(map(str, list(sheet.row(0))))))


# 替换行的列名
def replacefield(b):
    result = []
    for i in range(len(fieldList)):
        result.append(re.compile('(.*:)').sub("'" + str(fieldList[i][0]) + "':", str(b[i])))
    return result


rowList = []
for r in range(sheet.nrows):
    rowList.append(replacefield(list(map(str, list(sheet.row(r))))))

# 删除第一行数据
rowList.pop(0)

with open(resultFilePath, 'w', encoding='utf-8') as f:
    f.write("[" + str(rowList)[1:-1].replace("[", "{").replace("]", "}").replace('"', '').replace("'", '"') + "]")
