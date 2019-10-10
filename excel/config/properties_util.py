# coding=utf-8
import configparser
import os
import string
import threading
from excel.util.log import get_logger
import re

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
        LOG.info(u"get config info section:%s - key:%s - value:%s" % (section, key, result))
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


if __name__ == '__main__':
    config = Config()
    print config.get_value("Default", "Exclude_Project_ID")
