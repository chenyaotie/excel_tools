# coding=utf-8
from excel.constant.constant import *
from excel.file.bc_report_data import BCReport
from excel.file.data_model.cost_center_data_model import CostCenter
from excel.file.monthly_data import MonthlyDataExcel
from excel.file.profit_data import ProfitData
from excel.file.project_query import QueryProject
from excel.read_write.read_write_excel import WritingExcel
from excel.util import Util
from excel.util.log import get_logger
from excel.file.weight_data import WeightData
from excel.config.properties_util import Config

LOG = get_logger(__name__)


class Caculating(object):
    def __init__(self):
        self.config = Config()
        self.timesheet_report_path = self.config.get_timesheet_report_path()
        self.profit_report_path = self.config.get_profit_report_path()
        self.project_query_report_path = self.config.get_project_query_report_path()
        self.bc_report_path = self.config.get_bc_report_path()
        self.bc_sheet_name = self.config.get_bc_sheet_name()
        self.output_path = self.config.get_output_path()
        self.textBrowser = self.config.get_textBrowser()

        self.monthly_obj = MonthlyDataExcel(self.textBrowser, self.timesheet_report_path)
        self.profit_obj = ProfitData(self.textBrowser, self.profit_report_path)
        self.project_query_obj = QueryProject(self.textBrowser, self.project_query_report_path)

        self.weight_data_dict = dict()

        self.weight_data_dict = WeightData().get_weight_data_dict()

        self.bc_report_obj = BCReport(self.textBrowser, self.bc_report_path)
        self.bc_report_obj.init_monthly_table(self.bc_sheet_name)
        self.monthly_row_title_index = self.bc_report_obj.get_monthly_row_title_index()
        self.monthly_col_title_index = self.bc_report_obj.get_monthly_col_title_index()
        self.write_obj = WritingExcel(self.textBrowser, self.output_path, self.bc_sheet_name)
        self.projects_list = self.monthly_obj.get_projects_data

        # 初始化成本中心对象，初始化所有成本中心的各项cost
        self.cost_center_obj = CostCenter()
        self.init_cost_center()

        # 通过比率计算后的更新后的成本中心数据
        self.new_cost_center_obj = CostCenter()

    def update(self):
        LOG.info(u"开始更新BC报表数据")
        for project in self.projects_list:
            project_id = project.get_project_id()

            # 需要拆分的部分项目
            profit_data = self.profit_obj.get_project_data_class(project_id)

            # 该项目对应的正确成本中心
            main_center_name = self.project_query_obj.get_cost_ceter_name(project_id)

            # 获取权重字典{u'IM1904917862': {u'成都企业IT部': 1.2, u'西安企业IT部': 0.8, u'杭州企业IT部': 1}}，或为{}
            weight = self.weight_data_dict.get(project_id)

            # if weight:

            self._caculation(project, profit_data, main_center_name)

        self.write_to_excel()

    def _is_weight_error(self, weight, center_name):
        """
        权重异常判断
        :return:
        """
        if weight.has_key(center_name):
            del (weight[center_name])
        if weight.keys():
            msg = u"存在项目id: %s 权重表异常"
            self.textBrowser.append(msg)

    def write_to_excel(self):
        LOG.info(u"将更新后的BC报表数据保存")
        for zone, center_name in ZONE.items():
            for item, cost_item in COST_ITEM.items():
                value = getattr(self.cost_center_obj, zone + "_" + item)
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
            try:
                colx_index = self.get_x_index(zone)
            except Exception as e:
                msg = u"未从表格中解析到%s所在列的列号, error: %s" % (zone, str(e))
                LOG.error(msg)
                self.textBrowser.append(msg)
                raise Exception(msg)
            for costitem_key, item in COST_ITEM.items():
                try:
                    value = 0
                    rowx_index = self.get_y_index(item)
                    if rowx_index:
                        # 获取bc报表单元格值
                        value = self.bc_report_obj.get_monthly_cell_value(rowx_index, colx_index)
                    if not value:
                        value = 0
                    # 初始化对应成本中心，各项值
                    setattr(self.cost_center_obj, zone_key + "_" + costitem_key, value)
                except Exception as err:
                    msg = u"总成本中心解析出错, 总成本中心：%s, 项目：%s, error: %s" % (zone, item, str(err))
                    LOG.error(msg)
                    self.textBrowser.append(msg)
                    raise Exception(msg)

        LOG.info(u"总成本中心数据初始化完毕。")

    def update_cell(self, row_name, col_name, value):
        # 获取需要更新的成本中心横向坐标
        temp_index_of_center_cost = self.get_x_index(row_name)

        index_col = self.get_y_index(col_name)

        self.bc_report_obj.update_value(temp_index_of_center_cost, index_col, value)

    def update_cost(self, temp_cost_dict, center, profit_data, ratio):
        try:
            for key, total_value in temp_cost_dict.items():
                cost = abs(profit_data.get(key, 0))
                cost_ratio = cost * ratio
                total_value = total_value - cost_ratio

                zone = Util.get_zone_key(center)
                cost_item = Util.get_cost_item_key(key)

                #
                getattr(self.cost_center_obj, "set_" + zone + "_" + cost_item)(total_value)
        except KeyError as e:
            self.textBrowser.append(e.message)
            self.textBrowser.append(e)

    def _caculation(self, project, profit_data, main_center_name):
        master_zone = Util.get_zone_key(main_center_name)
        ratio_dict = project.get_ratio().get("ratio")
        weight_ratio_dict = project.get_ratio().get("weight_ratio")
        for slave_center, ratio in ratio_dict.items():
            if main_center_name == slave_center:
                # 遍历出的主成本中心时，不需要计算成本变化
                continue

            weight_ratio = 0
            if weight_ratio_dict:
                weight_ratio =weight_ratio_dict.get(slave_center)

            if not weight_ratio:  # 当获取不到权重占比时，就采用一般占比
                weight_ratio = ratio

            zone = Util.get_zone_key(slave_center)
            for item, value in profit_data.items():
                """
                如果有成本中心（x,y），总收入分别为A,B
                主要成本中心为x
                
                总金额占比：x:y = 0.4:0.6
                
                权重占比为x:y=0.2:0.8

                成本中心x 总收入和税金 x-0.8*a , 其他收入  x-0.6a

                成本中心y 总收入和税金：B+0.8a, 其他收入  x+0.6a
                """
                if item == REVENUE or item == TAX:
                    temp_value = value * weight_ratio
                else:
                    temp_value = value * ratio

                if item == REVENUE:
                    temp_value = abs(temp_value)

                item_en = Util.get_cost_item_key(item)
                # 主要成本中心默认的总成本
                master_value = getattr(self.cost_center_obj, master_zone + '_' + item_en)
                # 次要成本中心默认值
                slave_value = getattr(self.cost_center_obj, zone + '_' + item_en)
                # 从主要成本中心扣除成本
                new_value = master_value - temp_value
                # 次要成本中心增加
                new_value_add = slave_value + temp_value

                # 更新成本中心
                setattr(self.cost_center_obj, master_zone + '_' + item_en, new_value)

                setattr(self.cost_center_obj, zone + '_' + item_en, new_value_add)
