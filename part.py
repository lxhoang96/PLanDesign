from PyQt5.QtCore import *


class Part:
    def __init__(self, name='', line='', priority=0, day=0, startDate=QDate(), endDate=QDate()):
        self.name = name
        self.line = line
        self.startdate = startDate
        self.day = day
        self.priority = priority
        self.endDate = endDate

    # def cal_endDate(self, days=0):
    #     self.endDate = self.startdate.addDays(days)
    #     return self.endDate
