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
        app.truncateTable('TRUNCATE TABLE list')
        sql = 'insert into list(fromname,fromphone,fromaddress,rename,rephone,recom,readdress,towhere,what,protect,protectprice,submittime,orderno,fromno,price,weight,type,provice) ' \
              'values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
        print(len(paraArr))
        app.batchinsertTable(sql, paraArr)
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
        row_data = sheet.row_values(index)
        fromname = row_data[0]
        fromphone = row_data[1]
        fromaddress = row_data[2]
        rename = row_data[3]
        rephone = row_data[4]
        recom = row_data[5]
        readdress = row_data[6]
        towhere = row_data[7]
        what = row_data[8]
        protect = row_data[9]
        protectprice = row_data[10]
        submittime = row_data[11]
        orderno = row_data[12]
        fromno = row_data[13]
        price = row_data[14]
        weight = row_data[15]
        type = row_data[16]
        provice = row_data[17]
        dataarr = ['' + fromname + '', fromphone, '' + fromaddress + '', '' + rename + '', '' + rephone + '', '' + recom + '', '' + readdress + '',
                   '' + towhere + '','' + what + '', '' + protect + '', '' + protectprice + '', '' + submittime + '', '' + orderno + '', '' + fromno + '',
                   '' + price + '', weight, '' + type + '', '' + provice + '']
        sql = 'insert into list(fromname,fromphone,fromaddress,rename,rephone,recom,readdress,towhere,what,protect,protectprice,submittime,orderno,fromno,price,weight,type,provice)' \
              ' values ("'+ fromname + '", "' + fromphone + '", "' + fromaddress + '", "' + rename + '", "' + rephone + '", "' + recom + '", "' + readdress + '", "' + towhere + '", "' + what + '", "' + protect + '", "' + protectprice + '", "' + submittime + '", "' + orderno + '", "' + fromno + '", "' + price + '", "' + weight + '", "' + type + '", "' + provice + '")'
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

