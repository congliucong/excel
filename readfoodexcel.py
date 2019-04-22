# -*- coding: utf-8 -*-
import os
import xlrd
import mysqlConnect
from tqdm import tqdm
import math

# 操作excel LCC

"""
第一个参数为 excel路径，第二个参数为表名，第三个参数为导入数据起始行，由于计算机程序从0开始计数，所以取出第一行表头，一般从1开始
第四个参数为导入数据中止行，最后一个参数tuple，可输入多个表列明
如，excel有3列，则可输入('D:\\food', 1, 10, 'name', 'phone', 'address')
这样对应excel第0列插入到name,第1列出入到phone
"""


def read_excel(path, tablename, startrow, endrow, *rowname):
    # 创建表
    create_table(tablename, *rowname)
    for root, dirs, files in os.walk(path):
        paraArr = []
        for fileName in files:
            file_path = root + '\\' + fileName
            sqlArr = open_excel(fileName, file_path)
            paraArr.extend(sqlArr)

        app = mysqlConnect.MysqlUtil()
        # 清空表
        app.truncateTable('TRUNCATE TABLE '+tablename+'')
        sql = 'insert into ' + tablename + '('
        for row in rowname:
            sql = sql + row + ','
        sql = sql[0:-1] + ') values ('
        for row in rowname:
            sql = sql + '%s' + ','
        sql = sql[0:-1] + ')'
        print('插入语句:' + sql)
        # 批量插入
        total = len(paraArr)
        count = math.ceil(total/100)
        for index in range(0, count):
            if index != count -1:
                start = index * 100
                end = (index+1) * 100
            else:
                start = index * 100
                end = total
            # 切片[0:99] 最后一个不存在，所以要[0:100]
            app.batchinsertTable(sql, paraArr[start:end])
            print('执行'+str(start)+'到'+str(end)+'条成功')


'''
创建表
'''


def create_table(tablename, *rowname):
    row = ''
    for name in rowname:
        row = row+''+name+' varchar(100) NULL,'
    sql = 'CREATE TABLE '+tablename+' ( id int(0) NOT NULL AUTO_INCREMENT,'+row+'  PRIMARY KEY (id))   '
    print('建表语句：'+sql)
    app = mysqlConnect.MysqlUtil()
    app.createTable(tablename, sql)


def open_excel(file_name, file_path):
    sqlArr = []  # 空数组
    file_date = file_name[:-5]
    # 打开文件
    workbook = xlrd.open_workbook(r'' + file_path + '')
    # 获取所有sheet
    all_sheet = workbook.sheet_names()
    for sheet_name in tqdm(all_sheet):
        # print(sheet_name)
        # 根据sheet索引或者名称获取sheet内容
        sheet = workbook.sheet_by_name(sheet_name)
        for index, sheet_row in enumerate(sheet.get_rows()):
            if index == 0:
                continue
            rows = sheet_row
            dataarr = []
            for i in range(0, len(rows)):
                # print(rows[i].value)
                if type(rows[i].value) == float:
                    rowname = str(int(rows[i].value))
                elif type(rows[i].value) == int:
                    rowname = str(rows[i].value)
                else:
                    rowname = rows[i].value
                dataarr.append('' + rowname + '')
            dataarr.append('' + sheet_name + '')
            dataarr.append('' + file_date + '')
            print(dataarr)
            sqlArr.append(dataarr)
    return sqlArr


if __name__ == '__main__':
    # print("开始执行！")
    path = 'D:\\dinner'
    # read_excel(path, 'building', 1, 1629, 'ordernum', 'name', 'nickname', 'dept', 'num', 'innercode', 'outercode', 'applytime', 'returntime', 'sign', 'remark', 'idcard', 'usetime', 'remark2')
    read_excel(path, '04_22_dinner', 1, 1629, 'menu', 'count', 'restaurant', 'day')

    # print("结束执行！")

