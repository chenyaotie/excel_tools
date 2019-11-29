# coding=utf-8
from excel.read_write.read_write_excel import ExcelRead
from excel.util.log import get_logger
from excel.config.config import Config
import sys
reload(sys)
sys.setdefaultencoding('utf8')
LOG = get_logger(__name__)


class ProfitData(object):
    """
    利润中心表
    """

    def __init__(self, textBrowser, path):
        self.textBrowser = textBrowser
        msg = u"开始初始化读取利润表"
        self.textBrowser.append(msg)
        LOG.info(msg)
        self.conf = Config()
        excel = ExcelRead(self.textBrowser, path)
        sheet_name = self.__get_sheet_name()
        self.sheet = excel.get_sheet_by_name(sheet_name)
        self.index_of_title = excel.get_col_index_of_title()
        excel.clear()
        self.rowsn = self.sheet.nrows
        self.colsn = self.sheet.ncols

        self.bc_template_index = self.__get_index(u"BC模板分类")
        self.cost_center_index = self.__get_index(u"成本中心")
        self.project_id_index = self.__get_index(u"外部订单号")
        self.money_index = self.__get_index(u"公司货币金额")
        self.profit_dict = dict()

        self.all_row_datalist = self.__get_all_row_datalist()
        self.textBrowser.append(u"初始化利润报表完成")

    def __get_all_row_datalist(self):
        projects_datalist = []
        for i in range(self.rowsn):
            try:
                row = self.sheet.row_values(i)
                projects_datalist.append(row)
            except Exception as e:
                msg = "获取行内容失败行号:%d, 原因：%s" % (i, e)
                LOG.info(msg)
                self.textBrowser.append(str(msg))
                raise Exception
        return projects_datalist

    def get_project_costitem_dict(self, project_id):
        if not project_id:
            raise Exception("proejct id 为空")
        current_profit = dict()
        # 获取分类字典
        cost_class_list = self.conf.get_cost_class_list()
        for row in self.all_row_datalist:
            if project_id != row[self.project_id_index]:
                continue
            key = row[self.bc_template_index]
            if key in cost_class_list:
                try:
                    value = current_profit.get(key, 0)
                    value += row[self.money_index]
                    current_profit[key] = value
                except Exception as e:
                    msg = "获取行数据，失败原因：%s，行：%s,列标题%e" % (e, row, key)
                    LOG.error(msg)
                    raise Exception
        return current_profit

    def __get_index(self, title):
        index = self.index_of_title.get(title)
        if not index:
            msg = "利润表中不存在:%s列，请检查利润表中标题行" % title
            LOG.error(msg)
            self.textBrowser.append(msg)
            raise Exception
        msg = u"获取利润行首：%s 所在索引号：%d" % (title, index)
        LOG.info(msg)
        return index

    def __get_sheet_name(self):
        return self.conf.get_value("SHEET_NAME", "profit_sheet_name")
