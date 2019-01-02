# -*- coding: utf-8 -*-
import os
import xlrd
import mysqlConnect
from tqdm import tqdm


# 操作excel LCC
def read_excel(path):
    print("开始执行！")
    for root, dirs, files in os.walk(path):
        paraArr = []
        for fileName in files:
            sqlArr = []
            file_path = root + '\\' + fileName
            sqlArr = open_excel(fileName, file_path)
            paraArr.extend(sqlArr)

        app = mysqlConnect.MysqlUtil()
        app.truncateTable('TRUNCATE TABLE list1')
        # sql = 'insert into list(fromname,fromphone,fromaddress,rename,rephone,recom,readdress,towhere,what,protect,protectprice,submittime,orderno,fromno,price,weight,type,provice) ' \
        #       'values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
        # print(len(paraArr))
        # app.batchinsertTable(sql, paraArr)
    print("执行结束！")
    return 1


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
        # print(sheet.nrows)
    for index in range(0, sheet.nrows):
        if index == 0:
            continue
        # if index == sheet.nrows-1:
        #     continue
        print('-- '+str(index))
        row_data = sheet.row_values(index)
        # id = str(int(row_data[0]))
        if type(row_data[13]) == float:
            orderno = str(int(row_data[13]))
        else:
            orderno = row_data[13]
        if type(row_data[0]) == float:
            fromname = str(int(row_data[0]))
        else:
            fromname = row_data[0]
        sql = 'insert into list1(order_no, name) values ("' + orderno + '", "' + fromname + '");'
        print(sql)
        # sqlArr.append(dataarr)
        # else:
        #     for index, sheet_row in enumerate(sheet.get_rows()):
        #         if index == 0:
        #             continue
        #         rows = sheet_row
        #         t_menu = rows[0].value
        #         t_count = int(rows[1].value)
        #         dataarr = ['' + t_menu + '', t_count, '' + sheet_name + '', '' + file_date + '']
        #         sqlArr.append(dataarr)



    return sqlArr


if __name__ == '__main__':
    path = 'D:\\ems'
    read_excel(path)

