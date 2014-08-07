# -*- coding: utf-8 -*-
import xlrd
import json
import codecs
import sys
import getopt
from types import *


__opts, _ = getopt.getopt(sys.argv[1:], "p:o:") #获取命令行参数
excelPath = 'file.xls'
outPutName = 'gen.json'
for name, value in __opts:
    if name == "-p": #获取命令行参数e
        excelPath = value
    if name == "-o":
        outPutName = value

print(excelPath)
print(outPutName)

def open_excel(file='file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)


def excel_table_byindex(file= 'file.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = []
             for i in range(len(colnames)):
                print row[i]
                if type(row[i]) == FloatType:
                    row[i] = str(int(row[i]))
                if i == 2:
                    buildNo = row[i].split('-')[0]
                    houseNo = row[i].split('-')[1]
                    app.append(buildNo.encode('UTF-8','ignore'))
                    app.append(houseNo.encode('UTF-8','ignore'))
                    continue
                else:
                    app.append(row[i].encode('UTF-8','ignore'))
         list.append(app)
    return list



def excel_table_byname(file= 'file.xls',colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i].encode('UTF-8')
             list.append(app)
    return list

def saveJsonFile(str):

        #拼接json文件的目录地址
        filename = outPutName
        #将商店的字典写入json文件
        f = codecs.open(filename, 'wb')
        f.write(str)


def main():
   tables = excel_table_byindex(excelPath)
   #print(tables)
   str = json.dumps(tables, ensure_ascii=False)
   #print(str)
   saveJsonFile(str)
   #for row in tables:
       #print row

   #tables = excel_table_byname()
   #for row in tables:
       #print row
   print('has generate json file to ' + outPutName)

if __name__=="__main__":
    main()
