# excel_to_oracle

程序使用说明

作者: 劳嘉俊
联系方式：543681932@qq.com
开发日期: 2018.7.31
脚本版本: 1.0
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
  FILE: C:\Users\54368\Desktop\test_excel.xlsx
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
	python 脚本名 
	然后输入配置文件路径
	或
	python 脚本名 配置文件路径
  或
  直接运行 exe 文件
