#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# В дальнейшем будем дорабатывать этот конвертор, а сейчас он итак работает))
#

import sys
import xlwt

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