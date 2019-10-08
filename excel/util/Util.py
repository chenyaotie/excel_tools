# coding=utf-8
from excel.constant.constant import *


def get_cost_item_key(name):
    if name == REVENUE or name == REVENUE_ZH:
        return 'revenue_cost'
    elif name == REIMBURSEMENT_COST:
        return "reimbursement_cost"
    elif name == LABOR_COST:
        return "labor_cost"
    elif name == WORKSTATION_COST:
        return "workstation_cost"
    elif name == TAX or name == TAX_EN:
        return "tax_rate"
    else:
        msg = "不是正确的成本分类：%s" % name
        raise Exception(msg)


def get_zone_key(name):
    if name == CD:
        return 'cd'
    elif name == BJ:
        return "bj"
    elif name == HZ:
        return "hz"
    elif name == NJ:
        return "nj"
    elif name == SZHOU:
        return "szhou"
    elif name == SZ:
        return "sz"
    elif name == XA:
        return "xa"
    else:
        raise Exception("不是正确的成本中心分类：%s" % name)
