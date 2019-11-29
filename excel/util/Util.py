# coding=utf-8
import operator
from excel.util.log import get_logger
import sys
reload(sys)
sys.setdefaultencoding('utf8')

LOG = get_logger(__name__)

def get_costitem_en(name, costitem_dict):
    costitem_en = [costitem_en for costitem_en, costitem_zh in costitem_dict.items() if name == costitem_zh]
    if not costitem_en:
        msg = u"在成本项:%s中英文字典中为找到:%s对应的英文项" % (costitem_dict, name)
        LOG.error(msg)
        raise Exception(msg)
    return costitem_en[0]


def get_zone_en(name, zone_dict):
    zone_en = [zone_en for zone_en, zone_zh in zone_dict.items() if name == zone_zh]
    if not zone_en:
        msg = u"在成本中心:%s中英文字典中为找到:%s对应的英文项" % (zone_dict, name)
        LOG.error(msg)
        raise Exception(msg)
    return zone_en[0]


def has_intersection(list1, list2):
    """
    求两个list是否存在交集
    :param list1:
    :param list2:
    :return:
    """
    if not list1 or not list2:
        return False

    if list(set(list1).intersection(set(list2))):
        return True
    return False


def is_equal(list1, list2):
    """
    判断两个list是否相等
    :param list1:
    :param list2:
    :return:
    """
    if not list1 or not list2:
        return False
    list1.sort()
    list2.sort()
    return operator.eq(list1, list2)
