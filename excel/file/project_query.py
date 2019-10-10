# coding=utf-8
import re

from excel.constant.constant import *
from excel.read_write.read_write_excel import ExcelRead
from excel.util.log import get_logger

LOG = get_logger(__name__)

class QueryProject:
    """
    项目综合查询 表 处理类
    """

    def __init__(self,textBrowser, path):

        self.textBrowser = textBrowser
        self.excel = ExcelRead(self.textBrowser,path)
        self.sheet = self.excel.get_sheet_by_keywords([PROJECT_ID, SS_COST_CETER])
        self.index_of_title = self.excel.get_col_index_of_title()
        self.nrows = self.sheet.nrows
        self.textBrowser.append(u"初始化项目查询报表完成")

    def get_cost_ceter_name(self, project_id):
        """
        通过项目ID获取项目对应的成本中心
        :param project_id: 项目ID
        :return: 成本中心
        """
        cost_ceter_name_index = self.index_of_title.get(SS_COST_CETER)
        for i in range(self.nrows):
            rows_data = self.sheet.row_values(i)
            if project_id in rows_data:
                return self._ceter_name(project_id,rows_data[cost_ceter_name_index])

    def _ceter_name(self,project_id,name):
        """
        根据筛选出的实施成本中心名称，获得对应成本中心名称
        :param name:实施成本中心名称
        :return:成本中心名称
        """
        for ceter_name, re_pattern in HW_RE.items():
            if re.findall(re_pattern, name):
                return ceter_name
        msg =u"未找到对应的成本中心：%s，project id: %s" % (name,project_id)
        self.textBrowser.append(msg)
        return None


