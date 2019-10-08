# coding=utf-8
import ConfigParser
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
    _instance_lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        if not hasattr(Config, "_instance"):
            with Config._instance_lock:
                if not hasattr(Config, "_instance"):
                    Config._instance = object.__new__(cls)
        return Config._instance

    def __init__(self):
        self.loadConfig()

    def loadConfig(self):
        '''parse config file'''
        configfile = os.path.join(os.getcwd(), "config.ini")
        filepath = dealPath(configfile)
        LOG.info("config path is: %s" % filepath)
        if not os.path.isfile(filepath):
            msg = 'Can not find the config.ini'
            LOG.info(msg)
            raise Exception(msg)

        self.parser = ConfigParser.ConfigParser()

        content = open(filepath).read()
        # Window下用记事本打开配置文件并修改保存后，编码为UNICODE或UTF-8的文件的文件头
        # 会被相应的加上\xff\xfe（\xff\xfe）或\xef\xbb\xbf，然后再传递给ConfigParser解析的时候会出错
        # ，因此解析之前，先替换掉
        content = re.sub(r"\xfe\xff", "", content)
        content = re.sub(r"\xff\xfe", "", content)
        content = re.sub(r"\xef\xbb\xbf", "", content)
        open(filepath, 'w').write(content)

        self.parser.read(filepath)

    def get_value(self, section, key):

        if not self.parser:
            msg = "未初始化配置文件，请调用loadConfig函数"
            LOG(msg)
            raise Exception(msg)
        result = self.parser.get(section, key)
        LOG.info(u"get config info section:%s - key:%s - value:%s" % (section, key, result))
        return self.parser.get(section, key)


if __name__ == '__main__':
    config = Config()
    print config.get_value("Default", "Exclude_Project_ID")
