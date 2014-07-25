#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import xlwt

args = sys.argv
source = "ACD-SRV.log"

if len(args) >= 3:
	if (args[1] == '--file' or args[1] == '-f'):
		source = args[2]
	else:
		print "usage: \n txt_to_xls.py [-f] [file]"


splitted_lines = []

with open(source, "r") as source:
	for line in source.readlines():
		rows = line.split("--")
		splitted_lines.append(rows)

column_list = zip(*splitted_lines)


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