# coding=utf-8
from __future__ import division
import threading
from excel.util.log import get_logger
from excel.read_write.read_write_excel import ExcelRead
from excel.config.config import Config
from excel.util import Util
import sys
reload(sys)
sys.setdefaultencoding('utf8')

LOG = get_logger(file.__name__)


class WeightData(object):
    '''
    获取权重数据表数据
    '''
    instance_flag = False
    _instance_lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        if not hasattr(WeightData, "_instance"):
            with WeightData._instance_lock:
                if not hasattr(WeightData, "_instance"):
                    WeightData._instance = object.__new__(cls)
        return WeightData._instance

    def __init__(self):
        if not WeightData.instance_flag:
            self.config = Config()
            self.textBrowser = self.config.get_textBrowser()

            self.sheet = None
            self.path = self.config.get_weight_path()
            if self.path:
                excel = ExcelRead(self.textBrowser, self.path)
                sheet_name = self.get_sheet_name()
                self.sheet = excel.get_sheet_by_name(sheet_name)
                self.title_index_dict = excel.get_col_index_of_title()

                self.cost_center_name_list = self.get_cost_center_name()
                self.weight_data_dict = self._get_weight_data_dict()
                WeightData.instance_flag = True

            else:  # 权重路径未设置的情况下，初始化为空
                self.weight_data_dict = dict()
                WeightData.instance_flag = True

    def get_sheet_name(self):
        try:
            sheet_name = self.config.get_value("Weight", "weight_sheet_name")
        except Exception as e:
            msg = u"从配置文件获取 weight_sheet_name 出错： %s" % e
            LOG.error(msg)
            self.textBrowser.append(str(msg))
            return None
        return sheet_name

    def get_cost_center_name(self):
        return self.config.get_cost_center_name_list()

    def _get_weight_data_dict(self):
        """
        获取 权重 数据对象的字典集合，key为project id
        :return: 数据字典
        """
        if not self.title_index_dict.has_key(u"项目编号"):
            msg = u"未找到项目编号列"
            self.textBrowser.append(msg)
            raise Exception(msg)

        index_of_project_col_num = self.title_index_dict.get(u"项目编号")

        zone_dict = self.config.get_zone_dict()

        # 取交集，判断权重表 与 配置文件中的所有成本中心，是否存在交集
        if not Util.has_intersection(self.title_index_dict.keys(), zone_dict.values()):
            msg = u"权重表中的成本中心名称和配置的成本中心名称无法对应。"
            LOG.error(msg)
            raise Exception(msg)

        weight_data_dict = dict()
        for row in range(1, self.sheet.nrows):
            tmp_dict = dict()
            proejct_id = self.sheet.cell_value(row, index_of_project_col_num)
            # 获取各成本中心权重
            for zone_US, zone_ZH in zone_dict.items():
                col_index = self.title_index_dict.get(zone_ZH)
                # 权重表中不存在该条成本中心，则继续循环
                if not col_index:
                    continue
                value = self.sheet.cell_value(row, col_index)
                if value:
                    tmp_dict[zone_ZH] = value
            weight_data_dict[proejct_id] = tmp_dict
        return weight_data_dict

    def get_weight_data_dict(self):
        return self.weight_data_dict