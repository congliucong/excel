# -*- coding: utf-8 -*-
import os
import xlrd
import mysqlConnect
from tqdm import tqdm


# 操作excel
def read_excel(path):
    for root, dirs, files in os.walk(path):
        paraArr = []
        for fileName in files:
            sqlArr = []
            file_path = root + '\\' + fileName
            sqlArr = open_excel(fileName, file_path)
            paraArr.extend(sqlArr)

        app = mysqlConnect.MysqlUtil()
        app.truncateTable('TRUNCATE TABLE t_menu')
        sql = 'insert into t_menu(menu,count,restaurant,createtime) values (%s, %s, %s, %s)'
        app.batchinsertTable(sql, paraArr)


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
            t_menu = rows[0].value
            t_count = int(rows[1].value)
            dataarr = ['' + t_menu + '', t_count, '' + sheet_name + '', '' + file_date + '']
            sqlArr.append(dataarr)
    return sqlArr


if __name__ == '__main__':
    path = 'D:\\food'
    read_excel(path)
