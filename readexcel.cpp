#define MS_NO_COREDLL
#include "Python.h"



void  openExcelWithPython()
{  
    Py_Initialize();    // 初始化

    PyRun_SimpleString("import sys");
    PyRun_SimpleString("sys.path.append('C:/Python27/Lib/site-packages')");
    PyObject * pModule = NULL;
    PyObject * pFuncOpenXlsOrXlsxFile = NULL;
    PyObject * pFuncGetRowCount = NULL;
    PyObject * pFuncGetColCount = NULL;
    PyObject * pFuncGetValue = NULL;

    PyObject * pExcelBook = NULL;
    PyObject * pCellValue = NULL;
    
    //QString templatePath = qApp->applicationDirPath();
    //PyObject *sys_path = PySys_GetObject("path");
    //PyList_Append(sys_path, PyString_FromString(templatePath.toLocal8Bit()));

    pModule = PyImport_ImportModule("excel");
    if (pModule)
    {
        pFuncOpenXlsOrXlsxFile = PyObject_GetAttrString(pModule, "OpenXlsOrXlsxFile");
        pFuncGetRowCount = PyObject_GetAttrString(pModule, "GetRowCount");
        pFuncGetColCount = PyObject_GetAttrString(pModule, "GetColCount");
        pFuncGetValue = PyObject_GetAttrString(pModule, "GetValue");

        PyObject *pArgs = PyTuple_New(1);
        PyTuple_SetItem(pArgs, 0, Py_BuildValue("s", "t.xlsx"));//0—序号 i表示创建int型变量    
        pExcelBook = PyEval_CallObject(pFuncOpenXlsOrXlsxFile, pArgs);

        if (NULL != pExcelBook)
        {

            int rows = 0;
            int cols = 0;
            int sheetIndex = 0;
            PyObject *pArgssheet = PyTuple_New(1);
            PyTuple_SetItem(pArgssheet, 0, Py_BuildValue("i", sheetIndex));

            PyObject * pRow = PyEval_CallObject(pFuncGetRowCount, pArgssheet);
            PyObject * pCol = PyEval_CallObject(pFuncGetColCount, pArgssheet);

            PyArg_Parse(pRow, "i", &rows);
            PyArg_Parse(pCol, "i", &cols);
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < rows; j++)
                {
                    PyObject *pArgsValue = PyTuple_New(3);
                    PyTuple_SetItem(pArgsValue, 0, Py_BuildValue("i", sheetIndex));
                    PyTuple_SetItem(pArgsValue, 1, Py_BuildValue("i", i));
                    PyTuple_SetItem(pArgsValue, 2, Py_BuildValue("i", j));

                    char *str;
                    pCellValue = PyEval_CallObject(pFuncGetValue, pArgsValue);

                    if (pCellValue)
                    {
                        PyArg_Parse(pCellValue, "s", &str);
                        std::cout << str;

                        QTextCodec *utf8 = QTextCodec::codecForName("UTF-8");
                        QTextCodec *gbk = QTextCodec::codecForName("GB18030");
                        QString str1 = "";
                         str1 = utf8->toUnicode(str); qDebug() << __FUNCTION__ << str1;
                         str1 = gbk->toUnicode(str); qDebug() << __FUNCTION__ << str1;
  
                         QString str2 = "一"; qDebug() << __FUNCTION__ << str2;
   /*                      str1 = ; qDebug() << __FUNCTION__ << str1;
                         str1 = ; qDebug() << __FUNCTION__ << str1;*/

                        qDebug() << __FUNCTION__ << str;
                        qDebug() << __FUNCTION__ << utf8->toUnicode(str);
                    }
                                    
                }
            }
        }
        else
        {
            //打开excel失败
        }
    }
   

    /* char *str;
     PyArg_Parse(result, "s", &str);
     qDebug()<<__FUNCTION__ << str;
     */

    //char *cstr;
    //PyArg_Parse(pstr, "s", &cstr);


    //todo(liyh)释放以c端申请的python对象。//pyobject
    //if (pRet)
    //{
    //    Py_DECREF(pRet);
    //}

    Py_Finalize();      // 释放资源
}