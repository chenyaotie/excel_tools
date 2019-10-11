# coding=utf-8
from excel.config.config import Config
from excel.util.log import get_logger
import json
import random

LOG = get_logger(__name__)

CD = u"成都企业IT部"
XA = u"西安企业IT部"
HZ = u"杭州企业IT部"
SZ = u"深圳企业IT部"
BJ = u"北京企业IT部"
NJ = u"南京企业IT部"
SZHOU = u"苏州企业IT部"


# BC模板分类
REVENUE = "Revenue"
#####################################
#          月度报表                  #
####################################


CD_HW_RE = u"(?:HW成都|成都HW)企业IT(?:实施|研发)部"
XA_HW_RE = u"(?:HW西安|西安HW)企业IT(?:实施|研发)部"
HZ_HW_RE = u"(?:HW杭州|杭州HW)企业IT(?:实施|研发)部"
SZ_HW_RE = u"(?:HW深圳|深圳HW)企业IT(?:实施|研发)部"
BJ_HW_RE = u"(?:HW北京|北京HW)企业IT(?:实施|研发)部"
NJ_HW_RE = u"(?:南京HW|HW南京)企业IT(?:实施|研发)部"
SZHOU_HW_RE = u"(?:HW苏州|苏州HW)企业IT(?:实施|研发)部"

HW_RE = {CD: CD_HW_RE, XA: XA_HW_RE,
         HZ: HZ_HW_RE, SZ: SZ_HW_RE,
         NJ: NJ_HW_RE, BJ: BJ_HW_RE,
         SZHOU: SZHOU_HW_RE}
PROJECT_NAME = u"项目名称"
#####################################
#          项目综合查询               #
####################################


#####################################
#          华为事业群BC报表           #
####################################
TAX = u"税金"
TAX_EN = u"Sales Tax"
REIMBURSEMENT_COST = u"实施报销成本"
LABOR_COST = u"项目实施人工成本"
WORKSTATION_COST = u"工位成本"
REVENUE_ZH = u"总收入"

ZONE = {
    "cd": CD,
    "bj": BJ,
    "hz": HZ,
    "nj": NJ,
    "szhou": SZHOU,
    "sz": SZ,
    "xa": XA
}

COST_ITEM = {
    "revenue_cost": REVENUE_ZH,
    "tax_rate": TAX,
    "labor_cost": LABOR_COST,
    "reimbursement_cost": REIMBURSEMENT_COST,
    "workstation_cost": WORKSTATION_COST
}
