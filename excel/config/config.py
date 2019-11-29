# coding=utf-8
import configparser
import os
import string
import threading
from excel.util.log import get_logger
import re
import random
import sys

reload(sys)
import json

sys.setdefaultencoding('utf8')

path = r"C:\Users\xjchenaf\PycharmProjects\ExcelTools\config.ini"

LOG = get_logger(file.__name__)


def dealPath(pathname=''):
    '''deal with windows file path'''
    if pathname:
        pathname = pathname.strip()
    if pathname:
        pathname = r'%s' % pathname
        pathname = string.replace(pathname, r'/', '\\')
        pathname = os.path.abspath(pathname)
        if pathname.find(":\\") == -1:
            pathname = os.path.join(os.getcwd(), pathname)
    return pathname


class Config(object):
    instance_flag = False
    _instance_lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        if not hasattr(Config, "_instance"):
            with Config._instance_lock:
                if not hasattr(Config, "_instance"):
                    Config._instance = object.__new__(cls)
        return Config._instance

    def __init__(self):
        if not Config.instance_flag:
            self.timesheet_report_path = None
            self.profit_report_path = None
            self.project_query_report_path = None
            self.bc_report_path = None
            self.bc_sheet_name = None
            self.output_path = None
            self.weight_path = None
            self.textBrowser = None
            self.loadConfig()
            self.zone = self._get_zone()
            self.cost_class_dict = self.__get_cost_class_dict()
            self.cost_center_name_list = self._get_cost_center_name_list()
            Config.instance_flag = True

    def loadConfig(self):
        '''parse config file'''
        configfile = os.path.join(os.getcwd(), "config.ini")
        filepath = dealPath(configfile)
        LOG.info("config path is: %s" % filepath)
        if not os.path.isfile(filepath):
            msg = 'Can not find the config.ini'
            LOG.info(msg)
            raise Exception(msg)

        self.parser = configparser.ConfigParser()

        content = open(filepath).read()
        # Window下用记事本打开配置文件并修改保存后，编码为UNICODE或UTF-8的文件的文件头
        # 会被相应的加上\xff\xfe（\xff\xfe）或\xef\xbb\xbf，然后再传递给ConfigParser解析的时候会出错
        # ，因此解析之前，先替换掉
        content = re.sub(r"\xfe\xff", "", content)
        content = re.sub(r"\xff\xfe", "", content)
        content = re.sub(r"\xef\xbb\xbf", "", content)
        open(filepath, 'w').write(content)

        self.parser.read(filepath, encoding="UTF-8")

    def get_value(self, section, key):

        if not self.parser:
            msg = "未初始化配置文件，请调用loadConfig函数"
            LOG(msg)
            raise Exception(msg)
        result = self.parser.get(section, key)
        if not result:
            msg = u"未在配置文件中找到SECTION:%s  KEY:%s对应的值" % (section, key)
            LOG.error(msg)
            raise Exception
        LOG.info(
            u"get config info section:%s - key:%s - value:%s" %
            (section, key, result))
        return self.parser.get(section, key)

    def set_timesheet_report_path(self, path):
        self.timesheet_report_path = path

    def get_timesheet_report_path(self):
        return self.timesheet_report_path

    def set_profit_report_path(self, path):
        self.profit_report_path = path

    def get_profit_report_path(self):
        return self.profit_report_path

    def set_project_query_report_path(self, path):
        self.project_query_report_path = path

    def get_project_query_report_path(self):
        return self.project_query_report_path

    def set_bc_report_path(self, path):
        self.bc_report_path = path

    def get_bc_report_path(self):
        return self.bc_report_path

    def set_bc_sheet_name(self, name):
        self.bc_sheet_name = name

    def get_bc_sheet_name(self):
        return self.bc_sheet_name

    def set_output_path(self, path):
        self.output_path = path

    def get_output_path(self):
        return self.output_path

    def set_weight_path(self, path):
        self.weight_path = path

    def get_weight_path(self):
        return self.weight_path

    def set_textBrowser(self, textBrowser):
        self.textBrowser = textBrowser

    def get_textBrowser(self):
        return self.textBrowser

    def _get_zone(self):
        """
        获取一个关于成本中心的字典，该字典的key为随机字符串，value为成本中心
        :return:
        """
        cost_center_list = self._get_cost_center_name_list()
        random_key = get_random_strlist(len(cost_center_list))
        zone = dict(zip(random_key, cost_center_list))
        return zone

    def get_zone_dict(self):
        return self.zone

    def set_zone_dict(self, d):
        self.zone = d

    def _get_cost_center_name_list(self):
        try:
            cost_center_name_list = self.get_value(
                "Default", "Cost_Center_Name").split("|")
        except Exception as e:
            msg = u"从配置文件读取Cost_Center_Name成本中心数据解析失败：%s" % e
            LOG.error(e)
            self.textBrowser.append(str(e))
            raise Exception(msg)
        return list(set(cost_center_name_list))

    def get_cost_center_name_list(self):
        return self.cost_center_name_list

    def get_cost_class_list(self):
        cost_class_list = self.get_value("Default", "Cost_Class").split("|")
        return list(set(cost_class_list))

    def __get_cost_class_dict(self):
        cost_class_list = self.get_cost_class_list()
        random_key = get_random_strlist(len(cost_class_list))
        cost_class_dict = dict(zip(random_key, cost_class_list))
        return cost_class_dict

    def get_cost_class_dict(self):
        return self.cost_class_dict

    def get_cost_center_re_dict(self):
        re_str = self.get_value("Default", "RE").replace("\n", "").replace("\r", "")
        try:
            re_dict = json.loads(re_str)
            return re_dict
        except Exception as e:
            msg = u"转换成本中心正则表达式为dict时失败：%s" % e
            LOG.error(msg)
            self.textBrowser.append(msg)
            raise Exception


def get_random_strlist(n):
    """
    获取一个不重复的2个字符串的列表，长度为传入len
    :param n: 预期列表长度
    :return: 长度为len的不重复字符串列表
    """
    letters = "abcdefghijklmnopqrstuvwxyz"
    strlist = list()
    while len(strlist) < n:
        s = ("").join(random.sample(letters, 2))
        if s not in strlist:
            strlist.append(s)
    return strlist


if __name__ == '__main__':
    config = Config()
    data = config.get_cost_center_re_dict()
    import json

    new_data = json.loads(data)
    print new_data
    print type(new_data)
