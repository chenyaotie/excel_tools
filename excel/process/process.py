# coding=utf-8
from excel.file.bc_report_data import BCReport
from excel.file.data_model.cost_center_data_model import CostCenter
from excel.file.monthly_data import MonthlyDataExcel
from excel.file.profit_data import ProfitData
from excel.file.project_query import QueryProject
from excel.read_write.read_write_excel import WritingExcel
from excel.util import Util
from excel.util.log import get_logger
from excel.file.weight_data import WeightData
from excel.config.config import Config

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

        self.monthly_obj = MonthlyDataExcel(
            self.textBrowser, self.timesheet_report_path)
        self.project_ratio_datalist = self.monthly_obj.get_project_ratio_datalist()

        self.profit_obj = ProfitData(self.textBrowser, self.profit_report_path)
        self.project_query_obj = QueryProject(
            self.textBrowser, self.project_query_report_path)

        self.weight_data_dict = dict()

        self.weight_data_dict = WeightData().get_weight_data_dict()

        self.bc_report_obj = BCReport(self.textBrowser, self.bc_report_path)
        self.bc_report_obj.init_monthly_table(self.bc_sheet_name)
        self.bc_row_title_index = self.bc_report_obj.get_row_title_index()
        self.bc_col_title_index = self.bc_report_obj.get_col_title_index()
        self.write_obj = WritingExcel(
            self.textBrowser,
            self.output_path,
            self.bc_sheet_name)

        # 初始化成本中心对象，初始化所有成本中心的各项cost
        self.cost_center_obj = CostCenter()
        self.__init_cost_center()

        # 通过比率计算后的待更新的成本中心数据
        self.new_cost_center_obj = CostCenter()

    def update(self):
        LOG.info(u"开始更新BC报表数据")
        for project in self.project_ratio_datalist:
            project_id = project.get_project_id()

            # 需要拆分的部分项目
            profit_data = self.profit_obj.get_project_costitem_dict(project_id)

            # 该项目对应的正确成本中心
            main_center_name = self.project_query_obj.get_cost_ceter_name(
                project_id)
            self.__caculation(project, profit_data, main_center_name)
        self.__write_to_excel()

    def _is_weight_error(self, weight, center_name):
        """
        权重异常判断
        :return:
        """
        if center_name in weight:
            del (weight[center_name])
        if weight.keys():
            msg = u"存在项目id: %s 权重表异常"
            self.textBrowser.append(msg)

    def __write_to_excel(self):
        LOG.info(u"将更新后的BC报表数据保存")
        zone_dict = self.config.get_zone_dict()
        costitem_dict = self.config.get_cost_class_dict()
        for zone, center_name in zone_dict.items():
            for item, cost_item in costitem_dict.items():
                value = getattr(self.cost_center_obj, zone + "_" + item)
                self.write_obj.write(center_name, cost_item, value)
        # 保存文件
        self.write_obj.save()

    def __get_x_index(self, name):
        """
        获取横向坐标
        :param name:
        :return:
        """
        return self.bc_row_title_index.get(name)

    def __get_y_index(self, name):
        """
        获取纵向坐标 字典
        :param name:
        :return:
        """

        return self.bc_col_title_index.get(name)

    def __init_cost_center(self):
        """
        初始化所有的成本中心值
        :return:
        """
        zone_dict = self.config.get_zone_dict()
        for zone_en, zone_zh in zone_dict.items():
            try:
                colx_index = self.__get_x_index(zone_zh)
                if not colx_index:
                    raise Exception("BC报表中不存在%s列" % zone_zh)
            except Exception as e:
                msg = u"未从表格中解析到%s所在列的列号, error: %s" % (zone_zh, str(e))
                LOG.error(msg)
                self.textBrowser.append(msg)
                raise Exception(msg)
            costitem_dict = self.config.get_cost_class_dict()
            for costitem_en, costitem_zh in costitem_dict.items():
                try:
                    value = 0
                    if costitem_zh == "Revenue":
                        costitem_zh = u"总收入"
                    rowx_index = self.__get_y_index(costitem_zh)
                    if rowx_index:
                        # 获取bc报表单元格值
                        value = self.bc_report_obj.get_monthly_cell_value(
                            rowx_index, colx_index)
                    if not value:
                        value = 0
                    # 初始化对应成本中心，各项值
                    if costitem_zh == "Revenue" or costitem_zh == u"总收入":
                        value = abs(value)
                    setattr(self.cost_center_obj, zone_en + "_" + costitem_en, value)
                except Exception as err:
                    msg = u"总成本中心解析出错, 总成本中心：%s, 项目：%s, error: %s" % (
                        zone_zh, costitem_zh, str(err))
                    LOG.error(msg)
                    self.textBrowser.append(msg)
                    raise Exception(msg)
        LOG.info(u"总成本中心数据初始化完毕。")

    def __caculation(self, project, profit_data, master_costcenter_zh):
        zone_dict = self.config.get_zone_dict()
        costitem_dict = self.config.get_cost_class_dict()
        master_zone_en = Util.get_zone_en(master_costcenter_zh, zone_dict)
        ratio_dict = project.get_ratio().get("ratio")
        weight_ratio_dict = project.get_ratio().get("weight_ratio")
        for slave_csostcenter_zh, ratio in ratio_dict.items():
            if master_costcenter_zh == slave_csostcenter_zh:
                # 遍历出的主成本中心时，不需要计算成本变化
                continue

            weight_ratio = 0
            if weight_ratio_dict:
                weight_ratio = weight_ratio_dict.get(slave_csostcenter_zh)

            if not weight_ratio:  # 当获取不到权重占比时，就采用一般占比
                weight_ratio = ratio

            slave_costcenter_en = Util.get_zone_en(slave_csostcenter_zh, zone_dict)
            for costitem, value in profit_data.items():
                """
                如果有成本中心（x,y），总收入分别为A,B
                主要成本中心为x

                总金额占比：x:y = 0.4:0.6

                权重占比为x:y=0.2:0.8

                成本中心x 总收入和税金 x-0.8*a , 其他收入  x-0.6a

                成本中心y 总收入和税金：B+0.8a, 其他收入  x+0.6a
                """
                if costitem == "Revenue" or costitem == u"税金":
                    temp_value = value * weight_ratio
                else:
                    temp_value = value * ratio

                if costitem == "Revenue":
                    temp_value = abs(temp_value)

                costitem_en = Util.get_costitem_en(costitem, costitem_dict)
                # 主要成本中心默认的总成本
                master_value = getattr(
                    self.cost_center_obj,
                    master_zone_en + '_' + costitem_en)
                # 次要成本中心默认值
                slave_value = getattr(
                    self.cost_center_obj, slave_costcenter_en + '_' + costitem_en)
                # 从主要成本中心扣除成本
                new_master_value = master_value - temp_value
                # 次要成本中心增加
                new_slave_value = slave_value + temp_value

                # 更新成本中心
                setattr(
                    self.cost_center_obj,
                    master_zone_en + '_' + costitem_en,
                    new_master_value)

                setattr(
                    self.cost_center_obj,
                    slave_costcenter_en + '_' + costitem_en,
                    new_slave_value)
