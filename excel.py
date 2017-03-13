# -*- coding: utf-8 -*-
import sys
import xlrd
import logging

reload(sys)
sys.setdefaultencoding('gbk')

# create logger
logger_name = "example"
logger = logging.getLogger(logger_name)
logger.setLevel(logging.DEBUG)

# create file handler
log_path = "./log.log"
fh = logging.FileHandler(log_path)
fh.setLevel(logging.WARN)

# create formatter
fmt = "%(asctime)-15s %(levelname)s %(filename)s %(lineno)d %(process)d %(message)s"
datefmt = "%a %d %b %Y %H:%M:%S"
formatter = logging.Formatter(fmt, datefmt)

# add handler and formatter to logger
fh.setFormatter(formatter)
logger.addHandler(fh)

g_book = None
'''
def readxls():
	book = xlrd.open_workbook("t.xls")
	print("The number of worksheets is {0}".format(book.nsheets))
	print("Worksheet name(s): {0}".format(book.sheet_names()))
	sh = book.sheet_by_index(0)
	print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
	print("Cell a1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
	for rx in range(sh.nrows):
		print(sh.row(rx))

def readxlsx():
	book = xlrd.open_workbook("t.xlsx")
	print("The number of worksheets is {0}".format(book.nsheets))
	print("Worksheet name(s): {0}".format(book.sheet_names()))
	sh = book.sheet_by_index(0)
	print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
	print("Cell a1 is {0}".format(sh.cell_value(rowx=0, colx=0)))
	for rx in range(sh.nrows):
		print(sh.row(rx))
'''	

def OpenXlsOrXlsxFile(xlspath):
	global g_book #todo(liyh) 后面修改为类的方式 
	logger.warn("open xls " + xlspath)
	g_book = xlrd.open_workbook(xlspath)
	return g_book

def CloseXls():
	global g_book
	if None != g_book:
		g_book = None
	return g_book

def GetSheetByIndex(book, sheetIndex):
	index = 0
	logger.warn("GetSheetByIndex " + str(sheetIndex))
	s = book.sheet_by_index(sheetIndex)    
	return s

def GetRowCount(sheetIndex):
	global g_book 
	if None == g_book:
		return 0
	
	sheet = g_book.sheet_by_index(sheetIndex)
	logger.warn("GetRowCount " )
	rowCount = 0
	rowCount = sheet.nrows
	return rowCount
	
def GetColCount(sheetIndex):
	global g_book
	if None == g_book:
		return 0
	
	sheet = g_book.sheet_by_index(sheetIndex)	
	logger.warn("GetColCount " )
	colCount = 0
	colCount = sheet.ncols
	return colCount
	
def GetValue(sheetIndex, row, col):
	global g_book
	if None == g_book:
		return ""
	
	sheet = g_book.sheet_by_index(sheetIndex)	
	logger.warn("GetValue " )
	val = ""
	val = str(sheet.cell_value(rowx=row, colx=col))
	val=val.strip(' \t\n\r')
	return val

'''
book = OpenXlsOrXlsxFile("t.xlsx")
print book
sheet = GetSheetByIndex(book, 0)
print sheet

print "g_book ", g_book
rows = GetRowCount(0) 
print rows
cols = GetColCount(0)
print cols 
value =GetValue(0, 0, 0)
print value
'''