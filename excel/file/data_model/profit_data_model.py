# coding=utf-8
from excel.file.data_model.cost_center_data_model import CostCenter


class ProfitProjectData(object):
    """
     存储各个项目对应的 总收入，工位成本，报销成本，税金，人工成本
    """
    def __init__(self):
        self.project_id = None

        self.costcenter = CostCenter()
        self.revenue_cost = 0
        self.labor_cost = 0
        self.workstation_cost = 0
        self.tax_rate = 0
        self.reimbursement_cost = 0

    def get_project_id(self):
        return self.project_id

    def set_project_id(self,project_id):
        self.project_id = project_id

    def get_revenue_cost(self):
        return self.revenue_cost

    def set_revenue_cost(self, value):
        self.revenue_cost = value

    def add_revenue_cost(self, value):
        self.revenue_cost = self.get_revenue_cost() + value

    def get_labor_cost(self):
        return self.labor_cost

    def set_labor_cost(self, value):
        self.labor_cost = value

    def add_labor_cost(self, value):
        self.labor_cost = self.get_labor_cost() + value

    def get_workstation_cost(self):
        return self.workstation_cost

    def set_workstation_cost(self, value):
        self.workstation_cost = value

    def add_workstation_cost(self, value):
        self.workstation_cost = self.get_workstation_cost() + value

    def get_tax_rate(self):
        return self.tax_rate

    def set_tax_rate(self, value):
        self.tax_rate = value

    def add_tax_rate(self, value):
        self.tax_rate = self.get_tax_rate() + value

    def get_reimbursement_cost(self):
        return self.reimbursement_cost

    def set_reimbursement_cost(self, value):
        self.reimbursement_cost = value

    def add_reimbursement_cost(self, value):
        self.reimbursement_cost = self.get_reimbursement_cost() + value
