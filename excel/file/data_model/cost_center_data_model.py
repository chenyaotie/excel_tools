# coding=utf-8
class CostCenter:
    def __init__(self):
        self.revenue_cost = 0.0
        self.labor_cost = 0.0
        self.reimbursement_cost = 0.0
        self.workstation_cost = 0.0
        self.tax_rate = 0.0

        # 成都
        self.cd_revenue_cost = 0.0
        self.cd_labor_cost = 0.0
        self.cd_reimbursement_cost = 0.0
        self.cd_workstation_cost = 0.0
        self.cd_tax_rate = 0.0

        # 杭州
        self.hz_revenue_cost = 0.0
        self.hz_labor_cost = 0.0
        self.hz_reimbursement_cost = 0.0
        self.hz_workstation_cost = 0.0
        self.hz_tax_rate = 0.0

        # 深圳
        self.sz_revenue_cost = 0.0
        self.sz_labor_cost = 0.0
        self.sz_reimbursement_cost = 0.0
        self.sz_workstation_cost = 0.0
        self.sz_tax_rate = 0.0

        # 西安
        self.xa_revenue_cost = 0.0
        self.xa_labor_cost = 0.0
        self.xa_reimbursement_cost = 0.0
        self.xa_workstation_cost = 0.0
        self.xa_tax_rate = 0.0

        # 北京
        self.bj_revenue_cost = 0.0
        self.bj_labor_cost = 0.0
        self.bj_reimbursement_cost = 0.0
        self.bj_workstation_cost = 0.0
        self.bj_tax_rate = 0.0

        # 苏州
        self.szhou_revenue_cost = 0.0
        self.szhou_labor_cost = 0.0
        self.szhou_reimbursement_cost = 0.0
        self.szhou_workstation_cost = 0.0
        self.szhou_tax_rate = 0.0

        # 南京
        self.nj_revenue_cost = 0.0
        self.nj_labor_cost = 0.0
        self.nj_reimbursement_cost = 0.0
        self.nj_workstation_cost = 0.0
        self.nj_tax_rate = 0.0

    def set_revenue_cost(self, value):
        self.revenue_cost = abs(value)

    def get_revenue_cost(self):
        return self.revenue_cost

    def set_labor_cost(self, value):
        self.labor_cost = value

    def get_labor_cost(self):
        return self.labor_cost

    def set_reimbursement_cost(self, value):
        self.reimbursement_cost = value

    def get_reimbursement_cost(self):
        return self.reimbursement_cost

    def set_workstation_cost(self, value):
        self.workstation_cost = value

    def get_workstation_cost(self):
        return self.workstation_cost

    def set_tax_rate(self, value):
        self.tax_rate = value

    def get_tax_rate(self):
        return self.tax_rate

    def set_cd_revenue_cost(self, value):
        self.cd_revenue_cost = abs(value)

    def get_cd_revenue_cost(self):
        return self.cd_revenue_cost

    def set_cd_labor_cost(self, value):
        self.cd_labor_cost = value

    def get_cd_labor_cost(self):
        return self.cd_labor_cost

    def set_cd_reimbursement_cost(self, value):
        self.cd_reimbursement_cost = value

    def get_cd_reimbursement_cost(self):
        return self.cd_reimbursement_cost

    def set_cd_workstation_cost(self, value):
        self.cd_workstation_cost = value

    def get_cd_workstation_cost(self):
        return self.cd_workstation_cost

    def set_cd_tax_rate(self, value):
        self.cd_tax_rate = value

    def get_cd_tax_rate(self):
        return self.cd_tax_rate

    def set_hz_revenue_cost(self, value):
        self.hz_revenue_cost = abs(value)

    def get_hz_revenue_cost(self):
        return self.hz_revenue_cost

    def set_hz_labor_cost(self, value):
        self.hz_labor_cost = value

    def get_hz_labor_cost(self):
        return self.hz_labor_cost

    def set_hz_reimbursement_cost(self, value):
        self.hz_reimbursement_cost = value

    def get_hz_reimbursement_cost(self):
        return self.hz_reimbursement_cost

    def set_hz_workstation_cost(self, value):
        self.hz_workstation_cost = value

    def get_hz_workstation_cost(self):
        return self.hz_workstation_cost

    def set_hz_tax_rate(self, value):
        self.hz_tax_rate = value

    def get_hz_tax_rate(self):
        return self.hz_tax_rate

    def set_nj_revenue_cost(self, value):
        self.nj_revenue_cost = abs(value)

    def get_nj_revenue_cost(self):
        return self.nj_revenue_cost

    def set_nj_labor_cost(self, value):
        self.nj_labor_cost = value

    def get_nj_labor_cost(self):
        return self.nj_labor_cost

    def set_nj_reimbursement_cost(self, value):
        self.nj_reimbursement_cost = value

    def get_nj_reimbursement_cost(self):
        return self.nj_reimbursement_cost

    def set_nj_workstation_cost(self, value):
        self.nj_workstation_cost = value

    def get_nj_workstation_cost(self):
        return self.nj_workstation_cost

    def set_nj_tax_rate(self, value):
        self.nj_tax_rate = value

    def get_nj_tax_rate(self):
        return self.nj_tax_rate

    def set_bj_revenue_cost(self, value):
        self.bj_revenue_cost = abs(value)

    def get_bj_revenue_cost(self):
        return self.bj_revenue_cost

    def set_bj_labor_cost(self, value):
        self.bj_labor_cost = value

    def get_bj_labor_cost(self):
        return self.bj_labor_cost

    def set_bj_reimbursement_cost(self, value):
        self.bj_reimbursement_cost = value

    def get_bj_reimbursement_cost(self):
        return self.bj_reimbursement_cost

    def set_bj_workstation_cost(self, value):
        self.bj_workstation_cost = value

    def get_bj_workstation_cost(self):
        return self.bj_workstation_cost

    def set_bj_tax_rate(self, value):
        self.bj_tax_rate = value

    def get_bj_tax_rate(self):
        return self.bj_tax_rate

    def set_xa_revenue_cost(self, value):
        self.xa_revenue_cost = abs(value)

    def get_xa_revenue_cost(self):
        return self.xa_revenue_cost

    def set_xa_labor_cost(self, value):
        self.xa_labor_cost = value

    def get_xa_labor_cost(self):
        return self.xa_labor_cost

    def set_xa_reimbursement_cost(self, value):
        self.xa_reimbursement_cost = value

    def get_xa_reimbursement_cost(self):
        return self.xa_reimbursement_cost

    def set_xa_workstation_cost(self, value):
        self.xa_workstation_cost = value

    def get_xa_workstation_cost(self):
        return self.xa_workstation_cost

    def set_xa_tax_rate(self, value):
        self.xa_tax_rate = value

    def get_xa_tax_rate(self):
        return self.xa_tax_rate

    def set_szhou_revenue_cost(self, value):
        self.szhou_revenue_cost = abs(value)

    def get_szhou_revenue_cost(self):
        return self.szhou_revenue_cost

    def set_szhou_labor_cost(self, value):
        self.szhou_labor_cost = value

    def get_szhou_labor_cost(self):
        return self.szhou_labor_cost

    def set_szhou_reimbursement_cost(self, value):
        self.szhou_reimbursement_cost = value

    def get_szhou_reimbursement_cost(self):
        return self.szhou_reimbursement_cost

    def set_szhou_workstation_cost(self, value):
        self.szhou_workstation_cost = value

    def get_szhou_workstation_cost(self):
        return self.szhou_workstation_cost

    def set_szhou_tax_rate(self, value):
        self.szhou_tax_rate = value

    def get_szhou_tax_rate(self):
        return self.szhou_tax_rate

    def set_sz_revenue_cost(self, value):
        self.sz_revenue_cost = abs(value)

    def get_sz_revenue_cost(self):
        return self.sz_revenue_cost

    def set_sz_labor_cost(self, value):
        self.sz_labor_cost = value

    def get_sz_labor_cost(self):
        return self.sz_labor_cost

    def set_sz_reimbursement_cost(self, value):
        self.sz_reimbursement_cost = value

    def get_sz_reimbursement_cost(self):
        return self.sz_reimbursement_cost

    def set_sz_workstation_cost(self, value):
        self.sz_workstation_cost = value

    def get_sz_workstation_cost(self):
        return self.sz_workstation_cost

    def set_sz_tax_rate(self, value):
        self.sz_tax_rate = value

    def get_sz_tax_rate(self):
        return self.sz_tax_rate
