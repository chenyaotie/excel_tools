# coding=utf-8
from excel.constant.constant import *
from excel.read_write.read_write_excel import ExcelRead
from excel.util.log import get_logger


class ProfitData:
    """
    利润中心表
    """

    def __init__(self, textBrowser, path):

        self.textBrowser = textBrowser
        self.LOG = get_logger(__name__)
        self.profit_excel = ExcelRead(self.textBrowser, path)
        self.sheet = self.profit_excel.get_sheet_by_name(SHEET_NAME)
        self.index_of_title = self.profit_excel.get_col_index_of_title()
        self.rowsn = self.sheet.nrows
        self.colsn = self.sheet.ncols

        self.bc_template_index = self.index_of_title.get(BC_TEMPLATE)
        self.cost_center_index = self.index_of_title.get(COST_CETER)
        self.project_id_index = self.index_of_title.get(EXTENAL_ORDER_NUM)
        self.money_index = self.index_of_title.get(MONEY)
        self.profit_dict = dict()

        self.textBrowser.append(u"初始化利润报表完成")

    def get_all_projects_data_list(self):

        projects_datalist = []
        try:
            for i in range(self.rowsn):
                row = self.sheet.row_values(i)

                projects_datalist.append(row)
        except Exception as e:
            self.LOG.info(e)
            self.textBrowser.append(e)

        return projects_datalist

    def get_project_data_class(self, project_id):
        if not project_id:
            raise Exception("proejct id 为空")
        datalist = self.get_all_projects_data_list()
        current_profit = dict()
        try:

            for row in datalist:
                if project_id != row[self.project_id_index]:
                    continue
                key = row[self.bc_template_index]
                if key in BC_CLASS:
                    value = current_profit.get(key, 0)
                    #####################未取绝对值相加########计算后，总收入需要变成正数，其他的保持不变#################
                    value += row[self.money_index]
                    current_profit[key] = value
        except Exception as e:
            self.textBrowser.append(e.message)
            self.LOG.info(e)
        return current_profit
