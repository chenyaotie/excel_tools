# coding=utf-8
from excel.read_write.read_write_excel import ExcelWrite
from excel.util.log import get_logger

LOG = get_logger(__name__)


class BCReport:
    """
    华为事业群BC报表
    """

    def __init__(self, textBrowser, path):

        self.textBrowser = textBrowser
        self.path = path
        self.excel = ExcelWrite(self.textBrowser, self.path)

        self.monthly_table = None
        self.monthly_xutils_table = None

        self.costceter_table = None

        self.tiltle_index = dict()  # 月度sheet 行首 索引
        self.monthly_col_title_index = dict()  # 月度报名 列首 索引
        self.textBrowser.append(u"初始化BC报表完成")

    def init_monthly_table(self, month):
        """
        获取月度表
        :return:
        """
        try:
            index = self.excel.get_table_index(month)
        except Exception as e:
            msg = u"获取BC报表中 %s sheet页失败：%s" % (month, e)
            LOG.error(msg)
            self.textBrowser.append(msg)
            raise Exception
        self.monthly_table = self.excel.get_table(index)
        self.monthly_xutils_table = self.excel.get_xutils_table(index)

    def init_costcenter_table(self, costceter):
        """
        获取 成本中心表
        :return:
        """
        index = self.excel.get_table_index(costceter)
        self.monthly_table = self.excel.get_table(index)

    def get_row_title_index(self):

        rowsn = ExcelWrite.get_rowsn(self.monthly_table)
        for i in range(rowsn):
            row_value = self.monthly_table.row_values(i)
            if "P&L Summary RMB" not in row_value:
                continue

            for cell_value in row_value:
                self.tiltle_index[cell_value] = row_value.index(cell_value)
            break
        return self.tiltle_index

    def get_col_title_index(self):
        colsn = ExcelWrite.get_colsn(self.monthly_table)
        for i in range(colsn):
            col_value = self.monthly_table.col_values(i)
            if "P&L Summary RMB" not in col_value:
                continue

            for cell_value in col_value:
                self.monthly_col_title_index[cell_value] = col_value.index(cell_value)
            break
        return self.monthly_col_title_index

    def get_monthly_cell_value(self, x, y):
        # table.cell(rowx,colx)
        return self.monthly_table.cell(x, y).value

    def update_value(self, x, y, v):
        self.monthly_xutils_table.write(x, y, v)
