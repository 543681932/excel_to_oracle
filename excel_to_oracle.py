'''
作者 : 劳嘉俊

联系方式 ： 543681932@qq.com

开发日期 : 2018.07.30

脚本版本 : 1.0

开发python 版本 : 3.6.5

运行python 版本 : 3.5 以上

导入模块:
	os
	sys
	yaml -- PyYAML3.13
	xlrd -- xlrd1.1.0
	cx_Oracle -- cx_Oracle-6.4.1

需求说明:
	编写一个配置文件,指定excel文件中的内容录入到指定Oracle数据库中指定的表中

配置文件示例:
EXCEL:
  FILE: 
  SHEET: 第一页
  HEAD: 2
  START: 3
  END: 6
  MAP_LIST:
    - EXCEL_COL: 丁
      DB_COL: col1
    - EXCEL_COL: 乙
      DB_COL: col2
      DATA_TYPE: int
    - EXCEL_COL: 丙
      DB_COL: col3
      DATA_TYPE: date|%Y%m%d
  DB_TABLE: temp_table
  
DB:
  USER: system
  PASSWORD: oracle
  CON_STR: 192.168.243.131:1521/xe
  BEFORE_SQL: truncate table temp_table
  AFTER_SQL: 

配置说明:
EXCEL:
  FILE: excel文件路径
  SHEET: sheet名称
  HEAD: 头部在第几行
  START: 数据开始在第几行
  END: 数据结束在第几行
  MAP_LIST:
    - EXCEL_COL: excel列名
      DB_COL: 数据库表字段名
    - EXCEL_COL: excel列名
      DB_COL: 数据库表字段名
      DATA_TYPE: excel字段类型 (没有填写的时候默认为字符串)
    - EXCEL_COL: excel列名
      DB_COL: excel字段类型
      DATA_TYPE: date|%Y%m%d (如果是excel列 是时间类型,可以按 date|时间格式转换)
  DB_TABLE: 数据库表名
  
DB:
  USER: 数据库用户名
  PASSWORD: 数据库密码
  CON_STR: IP:端口号/SID
  BEFORE_SQL: 导入前执行的SQL
  AFTER_SQL: 导入完成后执行的SQL

运行示例:
	python excel_to_oracle.py 
	然后输入配置文件路径
	或
	python excel_to_oracle.py 配置文件路径
	或
	直接运行 exe 文件
'''


import os
import sys
import yaml
import xlrd
import cx_Oracle
from datetime import datetime,date


# 读取yaml文件函数，使用 utf-8 编码，处理中文出现乱码情况
def get_yaml(file_pwd):
    f = open(file_pwd, 'r')
    cont = f.read()
    try:
    	x = yaml.load(cont)
    except Exception as e:
    	print("yaml格式异常 或 无法读取文件! 退出程序!")
    	f.close()
    	input("输入任何按键，回车结束")
    	exit()
    else:
    	pass
    finally:
    	f.close()
    return x


# 读取excel文件内容，并以sql 的形式返回一个数组
def get_excel(file_pwd, sheet, head, start, end, col_list):
	# 判断excel 文件是否存在，或者配置信息是否正确
	if os.path.exists(file_pwd):
		pass
	else:
		print("excel文件不存在,或路径错误! 退出程序!")
		input("输入任何按键，回车结束")
		exit()
	excel_file = xlrd.open_workbook(file_pwd)
	table = excel_file.sheet_by_name(sheet)
	# 获取头部信息，并匹配 配置文件中的列的所在的第几列
	list_head = table.row_values(head)
	# print(list_head)
	total_col = len(list_head)
	colx_count = 0
	var_list = []
	sql_list = []

	for li_he in list_head:
		for i in col_list:
			if li_he == i['EXCEL_COL']:
				var_list.append([li_he, i.get('DB_COL'), colx_count, i.get('DATA_TYPE', 'string')])
				if i.get('DB_COL') == None:
					print("配置文件中没有填写 DB_COL，错误，退出程序!")
					input("输入任何按键，回车结束")
					exit()
				else:
					pass
			else:
				pass
		colx_count = colx_count + 1
	# print(var_list)

	# 拼接sql
	row_count = start
	var_sql = ''
	var_value = ''
	var_sql_list = []

	while row_count <= end:
		for var in var_list:
			col_db = var[1]
			col_type = var[3]
			if col_type == 'int':
				col_value = int(table.cell_value(row_count, var[2]))
			elif col_type == 'float':
				col_value = float(table.cell_value(row_count, var[2]))
			elif (col_type.split('|')[0] == 'date') and (table.cell(row_count, var[2]).ctype == 3):
				col_tmp = xlrd.xldate_as_tuple(table.cell_value(row_count, var[2]),excel_file.datemode)
				col_value = date(*col_tmp[:3]).strftime(col_type.split('|')[1])
			else:
				col_value = table.cell_value(row_count, var[2])

			var_sql_list.append([col_db,col_value])
		
		for i in range(len(var_sql_list)):
			if i +1 != len(var_sql_list):
				var_sql = var_sql + str(var_sql_list[i][0]) + ','
				var_value = str(var_value) + "'" + str(var_sql_list[i][1]) + "'" + ','
			else:
				var_sql = var_sql + str(var_sql_list[i][0])
				var_value = str(var_value) + "'" + str(var_sql_list[i][1]) + "'"
			
		# print("表列名为: " + var_sql + " 值为: " + var_value)
		sql = 'insert into ' + excel_table + '(' + var_sql + ') values(' + var_value + ')'
		sql_list.append(sql)

		var_sql_list = []
		var_sql = ''
		var_value = ''
		row_count = row_count + 1

	return sql_list


def inert_oracle(username, pswd, constr, sql_text, be_sql, af_sql):
	connect_string = username + '/' + pswd + '@' + constr
	try:
		db = cx_Oracle.connect(connect_string)
	except Exception as e:
		print("Oracle 连接异常! 退出程序!")
		input("输入任何按键，回车结束")
		exit()

	cursor = db.cursor()

	try:
		if be_sql is not None:
			cursor.execute(be_sql)
			db.commit()
		else:
			pass

		for sql in sql_text:
			cursor.execute(sql)
			db.commit()

		if af_sql is not None:
			cursor.execute(af_sql)
			db.commit()
		else:
			pass
	except Exception as e:
		print(e)
		print("执行sql 语句异常！ 退出程序！")
		input("输入任何按键，回车结束")
		exit()
	finally:
		db.close()




print("程序开始!")

# control_file 为控制文件路径
control_file = ''


if len(sys.argv) > 1:
    control_file = sys.argv[1]
else:
    control_file = input("请输入控制文件路径: ")


# 判断文件是否存在
if os.path.exists(control_file):
    pass
else:
    print("配置文件不存在,或路径错误! 退出程序!")
    input("输入任何按键，回车结束")
    exit()

# 获取 yaml 信息
yaml_info = get_yaml(control_file)
print("读取yaml文件成功!")
excel_info = yaml_info['EXCEL']
db_info = yaml_info['DB']

excel_file = excel_info.get('FILE')
excel_sheet = excel_info.get('SHEET')
excel_head = excel_info.get('HEAD') - 1
excel_start = excel_info.get('START') - 1
excel_end = excel_info.get('END') - 1
excel_col = excel_info.get('MAP_LIST')
excel_table = excel_info.get('DB_TABLE')

db_user = db_info.get('USER')
db_password = db_info.get('PASSWORD')
db_con = db_info.get('CON_STR')
db_be_sql = db_info.get('BEFORE_SQL')
db_af_sql = db_info.get('AFTER_SQL')

# print(excel_col)
sql_list = get_excel(excel_file, excel_sheet, excel_head, excel_start, excel_end, excel_col)
inert_oracle(db_user, db_password, db_con, sql_list, db_be_sql, db_af_sql)

print("程序结束")
input("输入任何按键，回车结束")