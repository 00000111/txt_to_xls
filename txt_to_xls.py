#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# В дальнейшем будем дорабатывать этот конвертор, а сейчас он итак работает))
#

import sys
import xlwt
import sqlite3

# Создаем класс базы данных
class Database(object):
	def __init__(self, db_name = 'main.db', login_table_name = 'logins', logout_table_name = 'logouts'):
		self.db_name = db_name
		self.logins = login_table_name
		self.logouts = logout_table_name
	
	# Метод для создающий таблицы для событий входа и выхода
	def CreateDB(self):
		conn = sqlite3.connect(self.db_name)
		cur = conn.cursor()

		cur.execute('''CREATE TABLE ? 
					(LI_Id INTEGER NOT NULL UNIQUE, event text, li_date text, li_time text, username text)''', self.logins)
		cur.execute('''CREATE TABLE ?
					(LO_Id INTEGER NOT NULL UNIQUE, event text, lo_date text, lo_lime text, username text)''', self.logouts)
		
		conn.commit()
		conn.close()

	#Метод для записи ниформации о входе
	def WriteLoginInfo(self, li_date, li_time, username):
		conn = sqlite3.connect(self.db_name)
		cur = conn.cursor()
		
		cur.execute("SELECT Li_Id FROM ? ORDER BY Li_Id DESC LIMIT 1", self.logins)
		line_id = cur.fetcone()
		line_id = line_id[0] + 1

		cur.execute("""INSERT INTO ? (LI_Id, event, li_date, li_time, username) 
					VALUES (?,?,?,?,?)""", self.logins (line_id, u'Вход', li_date, li_time, username))

args = sys.argv
source = "ACD-SRV.log"


# проверяем наличие аргументов запуска, если указан аргумент -f или --file переприсваеваем sourse
if len(args) >= 3:
	if (args[1] == '--file' or args[1] == '-f'):
		source = args[2]
	else:
		print "usage: \n txt_to_xls.py [-f] [file]"


splitted_lines = []

# Вычитываем исходный файл

with open(source, "r") as source:
	for line in source.readlines():
		rows = line.split("--")
		splitted_lines.append(rows)

column_list = zip(*splitted_lines)


# И записываем его в документ Excel
workbook = xlwt.Workbook()
output = workbook.add_sheet('Sheet1')

i = 0
for column in column_list:
	for item in range(len(column)):
		value = column[item].strip()
		output.write(item, i, value.decode('cp1251'))
	i += 1

workbook.save("output.xls")

print "Done!"