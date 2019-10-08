# coding=utf-8
import win32com.client

import string
import os
import time

def dealPath(pathname=''):
    '''deal with windows file path'''
    if pathname:
        pathname = pathname.strip()
    if pathname:
        pathname = r'%s' % pathname
        pathname = string.replace(pathname, r'/', '\\')
        pathname = os.path.abspath(pathname)
        if pathname.find(":\\") == -1:
            pathname = os.path.join(os.getcwd(), pathname)
    return pathname


class MyTestModel(object):
    def __init__(self):
        self.m_excel = win32com.client.Dispatch('Excel.Application')

    def open(self, filename=''):
        '''open excel file'''
        if getattr(self, 'm_book', False):
            self.m_book.Close()
        self.m_filename = dealPath(filename) or ''
        self.m_exists = os.path.isfile(self.m_filename)
        if not self.m_filename or not self.m_exists:
            self.m_book = self.m_excel.Workbooks.Add()
        else:
            self.m_book = self.m_excel.Workbooks.Open(self.m_filename)

    def setCellValue(self, sheet, row, col, value):
        '''set value of one cell'''
        self.getCell(sheet, row, col).Value = value

    def getCell(self, sheet=1, row=1, col=1):
        '''get the cell object'''
        assert row > 0 and col > 0, 'the row and column index must bigger then 0'
        return self.getSheet(sheet).Cells(row, col)

    def getSheet(self, sheet=1):
        '''get the sheet object by the sheet index'''
        assert sheet > 0, 'the sheet index must bigger then 0'
        return self.m_book.Worksheets(sheet)

    def getSheetByName(self, name):
        '''get the sheet object by the sheet name'''
        for i in xrange(1, self.getSheetCount() + 1):
            sheet = self.getSheet(i)
            if name == sheet.Name:
                return sheet
        return None

    def getSheetCount(self):
        '''get the number of sheet'''
        return self.m_book.Worksheets.Count

    def save(self, newfile=''):
        '''save the excel content'''
        assert type(newfile) is str, 'filename must be type string'
        newfile = dealPath(newfile) or self.m_filename
        if not newfile or (self.m_exists and newfile == self.m_filename):
            self.m_book.Save()
            return
        pathname = os.path.dirname(newfile)
        if not os.path.isdir(pathname):
            os.makedirs(pathname)
        self.m_filename = newfile
        self.m_book.SaveAs(newfile)


    def close(self):
        '''close the application'''
        self.m_book.Close(SaveChanges=1)
        self.m_excel.Quit()
        time.sleep(2)
        self.reset()


    def reset(self):
        '''reset'''
        self.m_excel = None
        self.m_book = None
        self.m_filename = ''


if __name__ == '__main__':
    t = MyTestModel()

    t.open(filename=r"D:/111.xlsx")
    t.setCellValue("2019-07",1,1,10000)
    t.save(r"D:/222.xlsx")
    t.close()