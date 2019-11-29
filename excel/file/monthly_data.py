# coding=utf-8
from __future__ import division

import re

from excel.file.data_model.project_data_model import Project
from excel.read_write.read_write_excel import ExcelRead
from excel.util.log import get_logger
from excel.config.config import Config
from excel.file.weight_data import WeightData
from excel.file.project_query import QueryProject
from excel.util import Util
import sys

reload(sys)
sys.setdefaultencoding('utf8')

LOG = get_logger(file.__name__)


class MonthlyDataExcel(object):
    """ 该Excel sheet页名称不确定，sheet 数量不确定，因此通过关键字来匹配，
    包含”费率，成本“的页签 为目标页签

    """

    def __init__(self, textBrowser, path):
        self.textBrowser = textBrowser
        msg = u"开始初始化读取timesheet."
        LOG.info(msg)
        self.textBrowser.append(msg)
        self.config = Config()
        self.project_query_report_path = self.config.get_project_query_report_path()
        self.total_cost = None
        self.path = path
        self.sheet = None
        self.first_row_values = None
        self.nrows = None
        self.ncols = None
        self.textBrowser = textBrowser
        self.all_projectid_list = list()
        self.projects_dict = dict()
        self.title_index_dict = dict()

        self.exclude_project_id = self.__get_exclude_project_id()

        # 初始化表格函数
        self.__init_sheet_obj()
        self.textBrowser.append(u"初始化TimeSheet报表完成")

        #
        self.weight = WeightData()

        self.project_query_obj = QueryProject(
            textBrowser, self.project_query_report_path)

        self.is_right_ratio = False

        self.project_ratio_list = self.__calculation()

    def __init_sheet_obj(self):
        """
        获取表格数据
        :return:
        """
        excel = ExcelRead(self.textBrowser, self.path)
        self.sheet = excel.get_sheet_by_keywords([u'费率', u'成本'])
        excel.clear()
        # 首行所有数据
        self.title_index_dict = excel.get_col_index_of_title()

        self.nrows = self.sheet.nrows
        self.ncols = self.sheet.ncols

    def __get_all_projects_list(self):
        """
        获取所有项目编号的list
        :return: 所有项目的set
        """
        if u"项目编号" not in self.title_index_dict:
            msg = u"未找到项目编号列"
            self.textBrowser.append(msg)
            raise Exception(msg)

        index_of_projectid_col_num = self.title_index_dict.get(u"项目编号")
        # 所有项目编号数据
        self.all_projectid_list = self.__get_all_project_id_list(
            index_of_projectid_col_num)

    def __get_all_project_id_list(self, index_of_projectid_col_num):
        col_list = self.sheet.col_values(index_of_projectid_col_num)
        # 去掉第一行Title行
        col_list.pop(0)
        all_projectid_list = set(col_list)
        if not all_projectid_list:
            msg = u"获取所有的项目编号为空"
            self.textBrowser.append(msg)
            raise Exception(msg)
        return all_projectid_list

    def __current_project_data(
            self,
            project_id,
            cost_ceter_index,
            cost_index,
            item_dict):
        """
        通过项目编号获取当前项目关联 所有成本中心 数据行列表

        :param project_id: 项目ID
        :param cost_ceter_index: 项目ID
        :param cost_index: 项目ID

        :return: 项目数据列表
        """

        for i in range(self.nrows):
            row_value = self.sheet.row_values(i)
            if project_id in row_value:
                self.__find_data_in_row(
                    cost_ceter_index, cost_index, row_value, item_dict)

    def get_project_ratio_datalist(self):
        """
        获取所有项目数据，项目id -key，项目关联数据行 -value
        :return: 项目数据字典
        """
        return self.project_ratio_list

    def __filter_project(self):
        if not self.all_projectid_list:
            self.__get_all_projects_list()

        # 成本中心列 的index值
        cost_ceter_index = self.title_index_dict[u"成本中心"]
        # 成本列 的index值
        cost_index = self.title_index_dict[u"成本"]

        error_project_id = list()
        filter_project_id = list()
        for project_id in self.all_projectid_list:

            # 排除项目
            if self.__is_exclude_project_id(project_id):
                filter_project_id.append(project_id)
                continue

            # 当前项目所有成本中心对应的成本数据 累加字典，{成都企业IT部：1000000，杭州企业IT部：200000，...}
            item = dict()
            self.__current_project_data(project_id, cost_ceter_index, cost_index, item)
            if not item:
                continue

            # 剔除只有一个成本中心，且为主成本中心的项目
            # 从项目查询中去查询主要成本中心
            main_center_name = self.project_query_obj.get_cost_ceter_name(
                project_id)
            if main_center_name is None:
                error_project_id.append(project_id)
                continue

            # 比较当前项目只有一个成本中心，且成本中心和主要成本中心相等则不纳入计算
            if len(item.keys()) == 1 and item.keys()[0] == main_center_name:
                continue

            # 计算获得满足条件的项目
            self.projects_dict[project_id] = item

        if filter_project_id:
            msg = u"过滤掉的项目id：%s" % filter_project_id
            LOG.info(msg)
            self.textBrowser.append(msg)

        if error_project_id:
            msg = u"存在脏数据，项目id: %s 在<项目综合查询表>中不存在，请确认项目月度timesheet报表中的项目ID" \
                  u"在项目综合查询中是否存在" % error_project_id
            self.textBrowser.append(msg)
            LOG.info(msg)

    def __calculation(self):
        """
        :param
        :return: [project1,project2,project3...}
        """
        self.__filter_project()
        project_list = list()
        for project_id, cost_dict in self.projects_dict.items():
            project = Project()
            current_project_weight = self.__get_weight(project_id)

            if current_project_weight:
                weight_ratio_dict = self.__get_ratio_with_weight(
                    project_id, current_project_weight, cost_dict)
                ratio_dict = self.__get_ratio_without_weight(
                    project_id, cost_dict)

                project.set_project_id(project_id)
                project.set_ratio(
                    {"ratio": ratio_dict, "weight_ratio": weight_ratio_dict})
            else:
                ratio_dict = self.__get_ratio_without_weight(
                    project_id, cost_dict)
                project.set_project_id(project_id)
                project.set_ratio({"ratio": ratio_dict})
                LOG.info("project_id: %s, ratio_dict: %s" % (project_id, ratio_dict))

            project_list.append(project)
        if self.is_right_ratio:
            msg = u"表格数据异常，请整改后重试."
            LOG.error(msg)
            self.textBrowser.append(msg)
            raise Exception(msg)

        return project_list

    def __find_data_in_row(self, cost_ceter_index, cost_index, row, item_dict):
        """ 将同一个项目下的成本中心拆分{成本中心1：totalcost,成本中心2：totalcost, other:totalcost}
        :param cost_ceter_index:
        :param row:
        :param item_dict:
        :return:
        """
        if not row:
            raise Exception(u"该行数据为空")
        re_dict = self.config.get_cost_center_re_dict()

        flag = False
        for k, r in re_dict.items():
            if re.match(r, row[cost_ceter_index]):
                value = item_dict.get(k, 0)
                value += row[cost_index]
                item_dict[k] = value
                flag = True
                break

        if flag == False:
            value = item_dict.get("other", 0)
            value += row[cost_index]
            item_dict["other"] = value

    def __get_exclude_project_id(self):
        result = self.config.get_value(
            "Default", "Exclude_Project_ID").split("|")
        return [value for value in result if value]

    def __is_exclude_project_id(self, project_id):
        if not self.exclude_project_id:
            return False

        for r in self.exclude_project_id:
            if re.findall(r, str(project_id)):
                return True
        return False

    def __get_weight(self, project_id):
        """
        获取当前项目id的权重数据字典
        :param project_id:
        :return:
        """
        weight_data_dict = self.weight.get_weight_data_dict()
        current_project_weight = weight_data_dict.get(project_id)
        return current_project_weight

    def __get_ratio_with_weight(
            self,
            project_id,
            current_project_weight,
            cost_dict):
        """
        获取有权重的 成本中心占比 比率 字典
        :param project_id:当前项目ID
        :param current_project_weight: 当前项目的权重字典
        :param cost_dict: 当前项目的各成本中心总收入字典
        :return:
        """

        # 判断ts中的成本中心和权重表中成本中心是否一一对应
        list1 = current_project_weight.keys()
        list2 = cost_dict.keys()
        if not Util.is_equal(list1, list2):
            msg = u"项目ID：%s,TS中成本中心和权重表中成本中心无法对应!!!" % project_id
            self.is_right_ratio = True
            LOG.error(msg)
            self.textBrowser.append(msg)

        total_cost = 0
        for cost_ceter, cost in cost_dict.items():
            # 获取成本中心权重
            weight_value = current_project_weight.get(cost_ceter)
            if cost_ceter == "other":
                weight_value = 1
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
            weight_value = current_project_weight.get(cost_ceter, 0)
            if weight_value is None:
                weight_value = 1
            weigth_ratio[cost_ceter] = cost * weight_value / total_cost
        return weigth_ratio

    def __get_ratio_without_weight(self, project_id, cost_dict):
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
