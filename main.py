# -*- coding: utf8 -*-

import os, sys
from win32com.client import Dispatch, constants

reload(sys)
sys.setdefaultencoding('utf8')

"""
获取所有的docx文档
"""
def fetchAllFile(path):
    files = []
    for dirpath, dirnames, filenames in os.walk(path):
        for file in filenames:
            ext = os.path.splitext(file)[1].lower()
            if ext == '.docx' or ext == '.doc':
                fullpath = os.path.join(dirpath, file)
                files.append(fullpath)
    return files

"""
将WORD文档转换为PDF文件
"""
def convertWordToPdf(docxPath, pdfPath):
    w = Dispatch("Word.Application")
    try:
        doc = w.Documents.Open(docxPath, ReadOnly=1)
        doc.ExportAsFixedFormat(pdfPath, constants.wdExportFormatPDF, Item=constants.wdExportDocumentWithMarkup, CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    except Exception, e:
        print e
    finally:
        w.Quit(constants.wdDoNotSaveChanges)
        
def main(src, dest):
    docPath = src
    if not os.path.exists(docPath):
        print "path not exists"
        return
    pdfPath = dest
    if not os.path.exists(pdfPath):
        os.makedirs(pdfPath)
    files = fetchAllFile(docPath)
    for file in files:
        identifier = file.split('\\')[-2]
        fileName = os.path.splitext(os.path.basename(file))[0] + '.pdf'
        if not os.path.exists(os.path.join(pdfPath, identifier)):
            os.mkdir(os.path.join(pdfPath, identifier))
        savePath = os.path.join(pdfPath, identifier, fileName)
        if not os.path.exists(savePath):
            print u"转换文件：", file, savePath
            convertWordToPdf(file, savePath)
        else:
            print u"文件已经存在，无需转换"
        
if __name__=='__main__':
    main('F:\\template\\', 'F:\\record\\')