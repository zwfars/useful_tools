# -*- coding: utf-8 -*-
from PyPDF2 import PdfFileWriter, PdfFileReader
import xlrd
import xlwt
import os

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )


class pdfadd:
    def __init__(self,inputname,number,offset,bookmark):
        self.name = inputname
        self.mark = bookmark
        self.num = number
        self.off = offset
    def addmark(self):
        output = PdfFileWriter() # open output
        input = PdfFileReader(open(self.name,'rb'))
        for i in range(0,self.num):
            output.addPage(input.getPage(i))
        mylen = len(self.mark)
        for i in range(0,mylen):
            if(self.mark[i][2]==1):
                tot = output.addBookmark(self.mark[i][0],int(self.mark[i][1]+self.off-1),parent=None)
            else:
                output.addBookmark(self.mark[i][0],int(self.mark[i][1]+self.off-1),parent=tot)
        outputStream = file("result.pdf","wb")
        output.write(outputStream)
        outputStream.close()


def getdata(str,rows=3):
    mylist = []
    mydata =[]
    input = xlrd.open_workbook(str)
    sheet = input.sheets()[0]
    for i in range(0,rows):
        mydata.append(sheet.col_values(i))
    mylen = len(mydata[0])
    for i in range(0,mylen):
        mylist.append([mydata[0][i],mydata[1][i],mydata[2][i]])
    return mylist


'''
excelname stores the content information.
number is the numble of pdf pages
'''

excelname = raw_input("the excel name: ")
bookmark = getdata(excelname)
number = input("the page of the pdf: ")
offset = input("the offset of the mark: ")
wf = pdfadd("test.pdf",number,offset,bookmark)
wf.addmark()
