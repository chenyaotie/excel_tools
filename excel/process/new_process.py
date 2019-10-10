# coding=utf-8
from excel.constant.constant import *
from excel.file.bc_report_data import BCReport
from excel.file.data_model.cost_center_data_model import CostCenter
from excel.file.monthly_data import MonthlyDataExcel
from excel.file.data_model.new_profit_data import ProfitData
from excel.file.project_query import QueryProject
from excel.read_write.read_write_excel import WritingExcel
from excel.util import Util
from excel.util.log import get_logger


class Caculating(object):
    def __init__(
            self,
            textBrowser,
            time_sheet_report_path,
            profit_report_path,
            project_qurey_report_path,
            bc_report_path,
            bc_sheetname,
            output_path):
        self.timesheet_report_path = time_sheet_report_path
        self.profit_report_path = profit_report_path
        self.project_query_report_path = project_qurey_report_path
        self.bc_report_path = bc_report_path
        self.bc_sheet_name = bc_sheetname
        self.output_path = output_path

        LOG = get_logger(__name__)

        self.textBrowser = textBrowser

        self.monthly_obj = MonthlyDataExcel(
            textBrowser, self.timesheet_report_path)
        self.profit_obj = ProfitData(textBrowser, self.profit_report_path)
        self.project_query_obj = QueryProject(
            textBrowser, self.project_query_report_path)

        self.bc_report_obj = BCReport(textBrowser, self.bc_report_path)
        self.bc_report_obj.init_monthly_table(self.bc_sheet_name)
        self.monthly_row_title_index = self.bc_report_obj.get_monthly_row_title_index()
        self.monthly_col_title_index = self.bc_report_obj.get_monthly_col_title_index()
        self.write_obj = WritingExcel(
            textBrowser, self.output_path, self.bc_sheet_name)
        self.projects_list = self.monthly_obj.get_projects_data

        # 初始化成本中心对象，初始化所有成本中心的各项cost
        self.cost_center_obj = CostCenter()
        self.init_cost_center()

        # 通过比率计算后的更新后的成本中心数据
        self.new_cost_center_obj = CostCenter()

    def update(self):
        LOG.info(u"开始更新BC报表数据")
        for project in self.projects_list:
            # 需要拆分的部分项目
            profit_data = self.profit_obj.get_project_data(
                project.get_project_id())

            # 该项目对应的正确成本中心
            center_name = self.project_query_obj.get_cost_ceter_name(
                project.get_project_id())

            if center_name is None:
                msg = u"存在脏数据，项目id: %s 在<项目综合查询表>中不存在，请确认项目月度timesheet报表中的项目ID" \
                      u"在项目综合查询中是否存在" % project.get_project_id()
                self.textBrowser.append(msg)
                LOG.info(msg)
                continue

            master_zone = Util.get_zone_key(center_name)

            for center, ratio in project.get_ratio().items():

                # 成本中心为同一个，比值为1.0,不需要更新
                if center == center_name and ratio == 1.0:
                    continue

                if center_name == center:
                    # 主要成本中心 和 遍历出的成本中心相等时，不需要计算主要成本中心的成本变化
                    continue

                # 获取成本中心对应的英文键
                zone = Util.get_zone_key(center)

                cost = profit_data.get_revenue_cost()
                temp_cost = cost * ratio

                master_cost = getattr(
                    self.cost_center_obj,
                    'get_' + master_zone + '_' + 'revenue_cost')()

                # 次要成本中心默认值
                slave_cost = getattr(
                    self.cost_center_obj,
                    'get_' + zone + '_' + 'revenue_cost')()

                # 从主要成本中心扣除成本
                new_cost = master_cost - temp_cost

                new_cost_add = slave_cost + temp_cost
                # 从主要成本中心扣除
                getattr(
                    self.cost_center_obj,
                    'set_' +
                    master_zone +
                    '_' +
                    item_en)(new_cost)
                # 次要成本中心增加
                getattr(
                    self.cost_center_obj,
                    'set_' +
                    zone +
                    '_' +
                    item_en)(new_cost_add)

                for item, value in profit_data.items():
                    """

                    如果有成本中心（x,y），总收入分别为A,B

                    主要成本中心为x

                    占比为x:y=0.2:0.8

                    某项目的总收入为a

                    那么

                    成本中心x 需要减去 x-0.8a

                    成本中心y 总收入：B+0.8a"""
                    temp_value = value * ratio

                    if item == REVENUE:
                        temp_value = abs(temp_value)

                    item_en = Util.get_cost_item_key(item)
                    # 主要成本中心默认的总成本
                    master_value = getattr(
                        self.cost_center_obj,
                        master_zone + '_' + item_en)
                    # 次要成本中心默认值
                    slave_value = getattr(
                        self.cost_center_obj,
                        zone + '_' + item_en)
                    # 从主要成本中心扣除成本
                    new_value = master_value - temp_value
                    new_value_add = slave_value + temp_value
                    # 从主要成本中心扣除
                    setattr(
                        self.cost_center_obj,

                        master_zone +
                        '_' +
                        item_en, new_value)
                    # 次要成本中心增加
                    setattr(
                        self.cost_center_obj,
                        zone +
                        '_' +
                        item_en, new_value_add)

        self.write_to_excel()

    def write_to_excel(self):
        LOG.info(u"将更新后的BC报表数据保存")
        for zone, center_name in ZONE.items():
            for item, cost_item in COST_ITEM.items():
                value = getattr(
                    self.cost_center_obj,
                    "get_" + zone + "_" + item)()
                self.write_obj.write(center_name, cost_item, value)
        # 保存文件
        self.write_obj.save()

    def get_x_index(self, name):
        """
        获取横向坐标
        :param name:
        :return:
        """
        return self.monthly_row_title_index.get(name)

    def get_y_index(self, name):
        """
        获取纵向坐标 字典
        :param name:
        :return:
        """

        return self.monthly_col_title_index.get(name)

    def init_cost_center(self):
        """
        初始化所有的成本中心值
        :param center_obj:
        :return:
        """
        for zone_key, zone in ZONE.items():
            colx_index = self.get_x_index(zone)
            for costitem_key, item in COST_ITEM.items():
                value = 0
                rowx_index = self.get_y_index(item)
                if rowx_index:
                    # 获取bc报表单元格值
                    value = self.bc_report_obj.get_monthly_cell_value(
                        rowx_index, colx_index)
                if not value:
                    value = 0

                # 初始化对应成本中心，各项值
                getattr(
                    self.cost_center_obj,
                    "set_" +
                    zone_key +
                    "_" +
                    costitem_key)(value)

    def update_cell(self, row_name, col_name, value):
        # 获取需要更新的成本中心横向坐标
        temp_index_of_center_cost = self.get_x_index(row_name)

        index_col = self.get_y_index(col_name)

        self.bc_report_obj.update_value(
            temp_index_of_center_cost, index_col, value)

    def update_cost(self, temp_cost_dict, center, profit_data, ratio):
        try:
            for key, total_value in temp_cost_dict.items():
                cost = abs(profit_data.get(key, 0))
                cost_ratio = cost * ratio
                total_value = total_value - cost_ratio

                zone = Util.get_zone_key(center)
                cost_item = Util.get_cost_item_key(key)

                #
                getattr(
                    self.cost_center_obj,
                    "set_" +
                    zone +
                    "_" +
                    cost_item)(total_value)
        except KeyError as e:
            self.textBrowser.append(e.message)
            self.textBrowser.append(e)
