#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
处理Excel文件
'''

from xlrd import open_workbook
import re
import json

filePath = "C://Users//MISuser//Desktop//业务提示消息数据.xls"
resultFilePath = "C:/Users/MISuser/Desktop/test.txt"
sheetIndex = 0

wb = open_workbook(filePath)
print("表单数量:", wb.nsheets)
print("表单名称:", wb.sheet_names())

sheet = wb.sheet_by_index(sheetIndex)
print("表单 %s 共 %d 行 %d 列" % (sheet.name, sheet.nrows, sheet.ncols))

# 获取第一行的列名
fieldList = list(map(lambda x: re.compile("'(.*)'").findall(x), list(map(str, list(sheet.row(0))))))

rowList = []
for r in range(sheet.nrows):
    rowDir = {}
    for i in range(len(fieldList)):
        rowDir[str(fieldList[i][0])] = str(sheet.cell(r, i).value)
    rowList.append(rowDir)

# 删除第一行数据
rowList.pop(0)

with open(resultFilePath, 'w', encoding='utf-8') as f:
    f.write(json.dumps(rowList).encode('utf-8').decode('unicode_escape'))

print(">>>>>>>>>>>>>写入完成<<<<<<<<<<<<<")
