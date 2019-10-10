# coding=utf-8
class Project:
    """
    存储各个项目对应的 项目名称，总收入，各个成本中心占比
    """

    def __init__(self):
        self.project_id = None
        self.total_cost = 0.0
        self.ratio = dict()
        self.project_name = None

    def set_project_id(self, projectid):
        self.project_id = projectid

    def get_project_id(self):
        return self.project_id

    def set_total_cost(self, value):
        self.total_cost = value

    def get_total_cost(self):
        return self.total_cost

    def set_ratio(self, d):
        """
        {IM1904917862:{ratio:{cd:0.2,hz:0.8},weight_ratio:{cd:0.4,hz:0.6}}
        :param d:
        :return:
        """
        self.ratio = d

    def get_ratio(self):
        return self.ratio

    def get_project_name(self):
        return self.project_name

    def set_project_name(self, name):
        self.project_name = name
