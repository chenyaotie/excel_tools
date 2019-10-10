# coding=utf-8
class WeightDataModel(object):
    """
    权重表中每行数据的 数据模型
    """

    def __init__(self):
        self.project_id = None
        self.cd_costcenter_value = 0
        self.sz_costcenter_value = 0
        self.szhou_costcenter_value = 0
        self.xa_costcenter_value = 0
        self.nj_costcenter_value = 0
        self.bj_costcenter_value = 0
        self.hz_costcenter_value = 0

    def set_project_id(self, project_id):
        self.project_id = project_id

    def get_project_id(self):
        return self.project_id

    def set_cd_costcenter_value(self, value):
        self.cd_costcenter_value = value

    def get_cd_costcenter_value(self):
        return self.cd_costcenter_value

    def set_sz_costcenter_value(self, value):
        self.sz_costcenter_value = value

    def get_sz_costcenter_value(self):
        return self.sz_costcenter_value

    def set_szhou_costcenter_value(self, value):
        self.szhou_costcenter_value = value

    def get_szhou_costcenter_value(self):
        return self.szhou_costcenter_value

    def set_xa_costcenter_value(self, value):
        self.xa_costcenter_value = value

    def get_xa_costcenter_value(self):
        return self.xa_costcenter_value

    def set_nj_costcenter_value(self, value):
        self.nj_costcenter_value = value

    def get_nj_costcenter_value(self):
        return self.nj_costcenter_value

    def set_bj_costcenter_value(self, value):
        self.bj_costcenter_value = value

    def get_bj_costcenter_value(self):
        return self.bj_costcenter_value

    def set_hz_costcenter_value(self, value):
        self.hz_costcenter_value = value

    def get_hz_costcenter_value(self):
        return self.hz_costcenter_value
