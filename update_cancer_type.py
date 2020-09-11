# -*- coding: utf-8 -*-
"""
Created on Thu Sep 10 18:06:36 2020

@author: dujidan
"""

#导入模块
import xlrd
import pymysql
import sys


input_excel = sys.argv[1]
input_excel_sheet = sys.argv[2]

# Open the workbook and define the worksheet
print ('Reading Excel ...')
book = xlrd.open_workbook(input_excel)
sheet = book.sheet_by_name(input_excel_sheet)
#book = xlrd.open_workbook("input_cancer_type.xlsx")
#sheet = book.sheet_by_name("input")

#---------------------------------------------------------------------------------
print ('Connecting   database ...')
#打开数据库连接
#注意：这里已经假定存在数据库testdb，db指定了连接的数据库，当然这个参数也可以没有
db = pymysql.connect(host='10.168.1.1', port=3306, user='user', passwd='passwd', db='database', charset='utf8')

#使用cursor方法创建一个游标
cursor = db.cursor()

#查询数据库版本
# 使用execute方法执行SQL语句
cursor.execute("select version()")
# 使用 fetchone() 方法获取一条数据
data = cursor.fetchone()
print(" Database Version:%s" % data)



#---------------------------------------------------------------------------------
print ('Loading information ...')


query = "INSERT INTO `cancer_type` (`编号`, `癌种`, `靶向`,`化疗`) VALUES (%s, %s, %s, %s) \
    ON DUPLICATE KEY UPDATE `癌种` = VALUES (`癌种`), `化疗` = VALUES (`化疗`), `靶向` = VALUES(`靶向`);"

try:
    for i in range(1, sheet.nrows):
        arg_1 = sheet.cell(i,0).value
        arg_2 = sheet.cell(i,1).value
        arg_3 = sheet.cell(i,2).value
        arg_4 = sheet.cell(i,3).value
       
        values = (arg_1, arg_2, arg_3, arg_4)
    
        #print (values)
        cursor.execute(query, values)

except:
    print ('Error:\t',values)


# 关闭游标
cursor.close()
# 提交
db.commit()
# 关闭数据库连接
db.close()

# 打印结果
print ("\nDone! \n")
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print ("我刚导入了 " + str(rows) + " * " + str(columns) + " 数据到MySQL!")
