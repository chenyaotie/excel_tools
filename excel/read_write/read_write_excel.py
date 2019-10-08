# coding=utf-8
import re

import xlrd
import xlwt
from xlutils.copy import copy

from excel.constant.constant import *


class ExcelRead:
    def __init__(self, textBrowser, path):
        self.path = path
        self.textBrowser = textBrowser
        try:
            self.excel = xlrd.open_workbook(self.path)
            self.title_index_dict = dict()  # 记录title名称对应的索引编号
        except Exception as e:
            msg = u"打开excel文件失败，文件名称：%s, 原因：%s" % (self.path, e)
            raise Exception(msg)

    def get_excel_obj(self):
        """ 获取excel表对象

        :return:
        """
        return self.excel

    def get_sheet_by_name(self, sheetname):
        """获取sheet表的对象

        :param sheetname: sheetname的正则表达式
        :return:
        """
        if not sheetname:
            msg = u"未指定sheetname"
            self.textBrowser.append(msg)
            raise Exception(msg)

        sheetnames_list = self.excel.sheet_names()

        for sheet_name in sheetnames_list:
            if re.findall(sheetname, sheet_name):
                sheet = self.excel.sheet_by_name(sheet_name)
                self._get_first_row_values(sheet)
                return sheet

    def get_sheet_by_keywords(self, keywords=list()):
        """
         根据表中sheet页第一行的关键字匹配目标sheet页
        :param keywords: 关键字列表
        :return: sheet页对象
        """
        sheets_list = self.excel.sheets()
        if sheets_list is None:
            msg = u"未找到Sheet页签"
            self.textBrowser.append(msg)
            raise Exception(msg)
        for sheet in sheets_list:
            if sheet.nrows == 0:
                continue
            cell_value = self._get_first_row_values(sheet)

            if all([v in cell_value for v in keywords]):
                return sheet
            else:
                continue
        msg = u"该Excel文件对应的页签中不包含关键字: %s，请检查表格是否正确，并调整格式" % keywords
        self.textBrowser.append(msg)
        raise Exception(msg)

    def _get_first_row_values(self, sheet):
        self.first_row_values = sheet.row_values(0)
        return self.first_row_values

    def get_col_index_of_title(self):
        """
        获取首行所有列的索引，并以dictionary（列名称，索引值）方式保存
        :return: dict
        """

        if not self.first_row_values:
            msg = u"当前Sheet页首行内容为空"
            self.textBrowser.append(msg)
            raise Exception(msg)
        for i in range(len(self.first_row_values)):
            if not self.first_row_values[i]:
                continue
            key = self.first_row_values[i]
            value = i
            self.title_index_dict[key] = value
        return self.title_index_dict

    def get_sheet_cell(self, sheetname, x, y):
        """获取对应sheet页的单元格内容

        :param sheetname:
        :param x:
        :param y:
        :return:
        """
        pass

    def set_sheet(self, sheetname):
        """ 设置sheet页名称

        :param sheetname:
        :return:
        """
        pass

    def set_sheet_cell(self, sheetname, x, y, value):
        """ 设置对应sheet页中单元格内容

        :param sheetname:
        :param x:
        :param y:
        :param value:
        :return:
        """
        pass


class ExcelWrite:
    def __init__(self, textBrowser, path):
        self.path = path
        self.textBrowser = textBrowser
        try:
            self.old_excel = xlrd.open_workbook(self.path)
            self.new_excel = copy(self.old_excel)
        except Exception as e:
            msg = u"打开excel文件失败，文件名称：%s, 原因：%s" % (self.path, e)
            raise Exception(msg)

    def get_table_index(self, sheet_name):
        index = self.old_excel.sheet_names().index(sheet_name)
        return index

    def get_xutils_table(self, index):
        return self.new_excel.get_sheet(index)

    def get_table(self, index):
        return self.old_excel.sheet_by_index(index)

    def save(self, new_file_path):
        self.new_excel.save(new_file_path)

    @staticmethod
    def get_rowsn(table):
        return table.nrows

    @staticmethod
    def get_colsn(table):
        return table.ncols


class WritingExcel(object):
    ROW = [CD, XA, HZ, SZ, BJ, NJ, SZHOU]
    COL = [REVENUE_ZH, LABOR_COST, REIMBURSEMENT_COST, WORKSTATION_COST, TAX]

    def __init__(self, textBrowser, path, sheet_name):
        self.path = path
        self.textBrowser = textBrowser
        try:
            self.workbook = xlwt.Workbook(encoding='utf-8')
            self.worksheet = self.workbook.add_sheet(
                sheet_name, cell_overwrite_ok=True)
            self.inital_excel()
        except Exception as e:
            msg = u"打开excel文件失败，文件名称：%s, 原因：%s" % (self.path, e)
            raise Exception(msg)

    def inital_excel(self):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        style.font = font

        for j in range(len(WritingExcel.COL)):
            self.worksheet.write(j + 1, 0, WritingExcel.COL[j], style)

        for i in range(len(WritingExcel.ROW)):
            self.worksheet.write(0, i + 1, WritingExcel.ROW[i], style)

    def save(self):
        self.workbook.save(self.path)

    def write(self, row_name, col_name, value):

        # 由于表格前面有空格，因此要+1
        x = WritingExcel.ROW.index(row_name) + 1
        y = WritingExcel.COL.index(col_name) + 1

        # 由于行列索引号和实际是反的，因此要倒换x,y的顺序
        self.worksheet.write(y, x, value)
