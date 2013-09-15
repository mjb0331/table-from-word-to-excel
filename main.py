'''
Created on 2013年9月15日

@author: lvan
'''

import excel

import win32com
from win32com.client import Dispatch, constants

w = win32com.client.Dispatch('Word.Application')
# 或者使用下面的方法，使用启动独立的进程：
# w = win32com.client.DispatchEx('Word.Application')

# 后台运行，不显示，不警告
w.Visible = 0
w.DisplayAlerts = 0

# 打开新的文件
doc = w.Documents.Open( FileName = filenamein )
# worddoc = w.Documents.Add() # 创建新的文档

# 插入文字
myRange = doc.Range(0,0)
myRange.InsertBefore('Hello from Python!')

# 使用样式
wordSel = myRange.Select()
wordSel.Style = constants.wdStyleHeading1

# 正文文字替换
w.Selection.Find.ClearFormatting()
w.Selection.Find.Replacement.ClearFormatting()
w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)

# 页眉文字替换
w.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
w.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(OldStr, False, False, False, False, False, True, 1, False, NewStr, 2)

# 表格操作
doc.Tables[0].Rows[0].Cells[0].Range.Text ='123123'
worddoc.Tables[0].Rows.Add() # 增加一行

# 转换为html
wc = win32com.client.constants
w.ActiveDocument.WebOptions.RelyOnCSS = 1
w.ActiveDocument.WebOptions.OptimizeForBrowser = 1
w.ActiveDocument.WebOptions.BrowserLevel = 0 # constants.wdBrowserLevelV4
w.ActiveDocument.WebOptions.OrganizeInFolder = 0
w.ActiveDocument.WebOptions.UseLongFileNames = 1
w.ActiveDocument.WebOptions.RelyOnVML = 0
w.ActiveDocument.WebOptions.AllowPNG = 1
w.ActiveDocument.SaveAs( FileName = filenameout, FileFormat = wc.wdFormatHTML )

# 打印
doc.PrintOut()

# 关闭
# doc.Close()
w.Documents.Close(wc.wdDoNotSaveChanges)
w.Quit()

if __name__ == '__main__':
    
    
    xls = excel(r'D:/dest.xls')  
    xls.cpSheet('Sheet1')  
    xls.save()  
    xls.close() 