# coding=utf-8
from __future__ import division

import re

from excel.constant import constant
from excel.file.data_model.project_data_model import Project
from excel.read_write.read_write_excel import ExcelRead
from excel.util.log import get_logger
from excel.config.properties_util import Config
from excel.file.weight_data import WeightData
from excel.file.project_query import QueryProject
import operator

# 定义需要处理列的title名称，后续如表格中的tiltle名称变化，需要修改此处的
COL_TITLE = [u"成本中心", u"项目编号", u"费率", u"成本"]

LOG = get_logger(file.__name__)


class MonthlyDataExcel(object):
    """ 该Excel sheet页名称不确定，sheet 数量不确定，因此通过关键字来匹配，
    包含”费率，成本“的页签 为目标页签

    """

    def __init__(self, textBrowser, path):
        self.config = Config()
        self.project_query_report_path = self.config.get_project_query_report_path()
        self.textBrowser = textBrowser
        self.total_cost = None
        self.path = path
        self.sheet = None
        self.first_row_values = None
        self.nrows = None
        self.ncols = None
        self.textBrowser = textBrowser
        self.all_projects_list = list()
        self.projects_dict = dict()
        self.title_index_dict = dict()

        self.exclude_project_id = self._get_exclude_project_id()

        # 初始化表格函数
        self.init_sheet_obj()
        self.textBrowser.append(u"初始化TimeSheet报表完成")

        #
        self.weight = WeightData()

        self.project_query_obj = QueryProject(textBrowser, self.project_query_report_path)

        self.is_right_ratio = False

    def init_sheet_obj(self):
        """
        获取表格数据
        :return:
        """
        excel = ExcelRead(self.textBrowser, self.path)
        self.sheet = excel.get_sheet_by_keywords([u'费率', u'成本'])

        self.title_index_dict = excel.get_col_index_of_title()

        self.nrows = self.sheet.nrows
        self.ncols = self.sheet.ncols

    def _get_all_projects_list(self):
        """
        获取所有项目编号的list
        :return: 所有项目的set
        """
        if not self.title_index_dict.has_key(u"项目编号"):
            msg = u"未找到项目编号列"
            self.textBrowser.append(msg)
            raise Exception(msg)

        index_of_project_col_num = self.title_index_dict.get(u"项目编号")
        col_list = self.sheet.col_values(index_of_project_col_num)

        # 去掉第一行Title行
        col_list.pop(0)
        self.all_projects_list = set(col_list)
        return self.all_projects_list  # 返回 项目编号列数据

    def _single_project_data(self, project_id):
        """
        通过项目编号获取当前项目关联 所有成本中心 数据行列表

        :param project_id: 项目ID
        :return: 项目数据列表
        """

        if not self.all_projects_list:
            msg = u"获取所有的项目编号为空"
            self.textBrowser.append(msg)
            raise Exception(msg)

        current_project = list()
        for i in range(self.nrows):
            row = self.sheet.row_values(i)
            if project_id in row:
                current_project.append(row)
        return current_project

    @property
    def get_projects_data(self):
        """
        获取所有项目数据，项目id -key，项目关联数据行 -value
        :return: 项目数据字典
        """
        # 筛选满足条件的项目
        self._filter_project()

        return self._calculation()

    def _filter_project(self):
        if not self.all_projects_list:
            self._get_all_projects_list()

            # 成本中心的index值
        cost_ceter_index = self.title_index_dict[constant.COST_CETER]
        # 成本 的index值
        cost_index = self.title_index_dict[constant.COST]

        for project_id in self.all_projects_list:
            all_rows_datalist = self._single_project_data(project_id)

            item = dict()
            if self._is_exclude_project_id(project_id):
                LOG.info("filter project id = %s" % project_id)
                self.textBrowser.append("filter project id = %s" % project_id)
                continue

            # 更新item 字典 {项目:{成本中心1：成本,成本中心2：成本}}
            for row in all_rows_datalist:
                self._find_data_in_row(cost_index, cost_ceter_index, row, item)

            if not item:
                continue

            # 剔除只有一个成本中心，且为主成本中心的项目
            main_center_name = self.project_query_obj.get_cost_ceter_name(project_id)
            if len(item.keys()) == 1 and item.keys()[0] == main_center_name:
                continue
            if main_center_name is None:
                msg = u"存在脏数据，项目id: %s 在<项目综合查询表>中不存在，请确认项目月度timesheet报表中的项目ID" \
                      u"在项目综合查询中是否存在" % project_id
                self.textBrowser.append(msg)
                LOG.info(msg)
                continue

            self.projects_dict[project_id] = item

    def _calculation(self):
        """
        :param
        :return: [project1,project2,project3...}
        """
        project_list = list()
        for project_id, cost_dict in self.projects_dict.items():
            project = Project()
            current_project_weight = self.get_weight(project_id)

            if current_project_weight:
                weight_ratio_dict = self.get_ratio_with_weight(project_id, current_project_weight, cost_dict)
                ratio_dict = self.get_ratio_without_weight(project_id, cost_dict)

                project.set_project_id(project_id)
                project.set_ratio({"ratio": ratio_dict, "weight_ratio": weight_ratio_dict})
            else:
                ratio_dict = self.get_ratio_without_weight(project_id, cost_dict)
                project.set_project_id(project_id)
                project.set_ratio({"ratio": ratio_dict})
            project_list.append(project)

        if self.is_right_ratio:
            msg = u"表格数据异常，请整改后重试."
            LOG.error(msg)
            self.textBrowser.append(msg)
            raise Exception(msg)

        return project_list

    def _find_data_in_row(self, cost_index, cost_ceter_index, row, item_dict):
        """ 将同一个项目下的成本中心拆分{成本中心1：totalcost,成本中心2：totalcost}
        :param cost_ceter_index:
        :param row:
        :param item_dict:
        :return:
        """
        for k, r in constant.HW_RE.items():
            if re.match(r, row[cost_ceter_index]):
                value = item_dict.get(k, 0)
                value += row[cost_index]
                item_dict[k] = value
                break

    def _get_exclude_project_id(self):
        result = self.config.get_value("Default", "Exclude_Project_ID").split("|")

        return [value for value in result if value]

    def _is_exclude_project_id(self, project_id):

        if not self.exclude_project_id:
            return False

        for r in self.exclude_project_id:
            if re.findall(r, str(project_id)):
                return True
        return False

    def get_weight(self, project_id):
        current_project_weight = self.weight.get_weight_data_dict().get(project_id)
        return current_project_weight

    def get_ratio_with_weight(self, project_id, current_project_weight, cost_dict):
        """
        获取有权重的 成本中心占比 字典
        :param project_id:
        :param current_project_weight:
        :param cost_dict:
        :return:
        """

        # 判断ts中的成本中心和权重表中成本中心是否一一对应
        if not operator.eq(current_project_weight.keys(), cost_dict.keys()):
            msg = u"项目ID：%s,TS中成本中心和权重表中成本中心无法对应!!!" % project_id
            self.is_right_ratio = True
            LOG.error(msg)
            self.textBrowser.append(msg)

        total_cost = 0
        for cost_ceter, cost in cost_dict.items():
            # 获取成本中心权重
            weight_value = current_project_weight.get(cost_ceter)
            if weight_value is None:
                msg = u"项目ID：%s, 成本中心：%s 权重为空!!!" % (project_id, cost_ceter)
                LOG.warn(msg)
                self.textBrowser.append(msg)
                weight_value = 1
                self.is_right_ratio = True
            total_cost += cost * weight_value

        if total_cost == 0:
            msg = u"项目ID：%s 总成本为0，请检查TS表格" % project_id
            LOG.error(msg)
            self.textBrowser.append(msg)
            self.is_right_ratio = True
            total_cost = -1

        weigth_ratio = dict()
        for cost_ceter, cost in cost_dict.items():
            weight_value = current_project_weight.get(cost_ceter)
            if weight_value is None:
                weight_value = 1
            weigth_ratio[cost_ceter] = cost * weight_value / total_cost
        return weigth_ratio

    def get_ratio_without_weight(self, project_id, cost_dict):
        """
        获取没有权重的成本中心占比 字典
        :param project_id:
        :param cost_dict:
        :return:
        """
        # 计算总成本，用于求比值的分母
        total_cost = 0
        for cost in cost_dict.values():
            total_cost += cost

        if total_cost == 0:
            msg = u"项目ID：%s 总成本为0，请检查TS表格" % project_id
            LOG.error(msg)
            self.textBrowser.append(msg)
            self.is_right_ratio = True
            total_cost = -1

        ratio = dict()
        for cost_ceter, cost in cost_dict.items():
            ratio[cost_ceter] = cost / total_cost
        return ratio

    def get_is_right_ratio(self):
        return self.is_right_ratio
