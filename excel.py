# -*- coding: utf-8 -*-
import sys
import xlrd
import traceback
import logging
from logging.handlers import RotatingFileHandler


reload(sys)
sys.setdefaultencoding('gbk')

# create logger
logger_name = "example"
logger = logging.getLogger(logger_name)
logger.setLevel(logging.DEBUG)

# create file handler
log_path = "./excel.log"
fh = RotatingFileHandler(log_path, mode='a', maxBytes=5*1024*1024, 
                                 backupCount=2, encoding=None, delay=0)
fh.setLevel(logging.WARN)

# create formatter
fmt = "%(asctime)-15s %(levelname)s %(filename)s %(lineno)d %(process)d %(message)s"
datefmt = "%a %d %b %Y %H:%M:%S"
formatter = logging.Formatter(fmt, datefmt)

# add handler and formatter to logger
fh.setFormatter(formatter)
logger.addHandler(fh)

def get_function_name():
    return traceback.extract_stack(None, 2)[0][2]
	
g_book = None
def OpenXlsOrXlsxFile(xlspath):
	global g_book #todo(liyh) 后面修改为类的方式 
	try:
		logger.warn(get_function_name() + xlspath)
		g_book = xlrd.open_workbook(xlspath)
	except:
		e = sys.exc_info()[0]
		logger.error(e)
		logger.error(traceback.format_exc())
	finally:
		return g_book	

def CloseXls():
	global g_book
	if None != g_book:
		g_book = None
	return g_book

def GetSheetByIndex(book, sheetIndex):
	index = 0
	logger.warn(get_function_name() + str(sheetIndex))
	s = book.sheet_by_index(sheetIndex)    
	return s

def GetRowCount(sheetIndex):
	global g_book 
	if None == g_book:
		logger.error(get_function_name() + "g_book is None" )
		return 0
	
	sheet = g_book.sheet_by_index(sheetIndex)
	logger.warn(get_function_name() )
	rowCount = 0
	rowCount = sheet.nrows
	return rowCount
	
def GetColCount(sheetIndex):
	global g_book
	if None == g_book:
		logger.error(get_function_name() + "g_book is None" )
		return 0
	
	sheet = g_book.sheet_by_index(sheetIndex)	
	logger.warn(get_function_name() )
	colCount = 0
	colCount = sheet.ncols
	return colCount
	
def GetValue(sheetIndex, row, col):
	global g_book
	if None == g_book:
		logger.error(get_function_name() + "g_book is None" )
		return ""
	
	val = ""
	try:
		sheet = g_book.sheet_by_index(sheetIndex)	
		logger.warn(get_function_name() )
		
		cty = sheet.cell_type(row, col)
		if xlrd.XL_CELL_EMPTY == cty:
			return val
		elif xlrd.XL_CELL_DATE == cty:
			return 
		elif xlrd.XL_CELL_ERROR == cty:
			return val
		else:	
			val = str(sheet.cell_value(rowx=row, colx=col))
			
		val = str(sheet.cell_value(rowx=row, colx=col))
		val=val.strip(' \t\n\r')
	except:
		e = sys.exc_info()[0]
		logger.error(e)
		logger.error(traceback.format_exc())
	finally:
		return val
	
def test():
	#book = OpenXlsOrXlsxFile("C:/Users/Administrator/Desktop/temp.xls")
	book = OpenXlsOrXlsxFile("C:/Users/Administrator/Desktop/template2.xls")
	print book
	sheet = GetSheetByIndex(book, 0)
	print sheet.nrows  
	print sheet.ncols
	print sheet

	print "g_book ", g_book
	rows = GetRowCount(0) 
	print rows
	cols = GetColCount(0)
	print cols 
	value =GetValue(0, 0, 0)
	print value
	
	for r in range(sheet.nrows):
		for c in range(sheet.ncols):
			print str(sheet.cell_value(r,c))
	
test()
