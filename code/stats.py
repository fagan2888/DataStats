from my_date import MyDate


class Stats(object):
    """
    以一个机构为实例，返回计算报表相关数据
    """
    def __init__(self, name, risk, sql, level="机构"):
        self._jian_cheng = name                                 # 机构简称
        self._date = MyDate()                                     # 日期对象
        self._sql = sql                                                  # 数据库对象
        self._risk = risk                                                 # 险种
        self._level = level                                             # 机构层级（中心支公司或机构）

    @property    
    def jian_cheng(self):
        """返回机构简称"""
        return self._jian_cheng

    @property
    def sql(self):
        """返回数据库对象"""
        return self._sql

    @property
    def risk(self):
        """返回险种"""            
        return self._risk
    
    @property
    def level(self):
        """返回机构层级"""
        return self._level

    @property
    def name(self):
        """返回机构全称"""
        if self.jian_cheng == "分公司整体":
            self._name = "分公司整体"
        elif self.jian_cheng == "沾益":
            self._name = "沾益"
        elif self.jian_cheng == "祥云":
            self._name = "祥云"
        elif self.jian_cheng == "曲靖营业一部":
            self._name = "曲靖营业一部"
        elif self.jian_cheng =="曲靖中支本部":
            self._name = "曲靖中支本部"
        elif self.jian_cheng =="分公司本部":
            self._name = "分公司本部"
        elif self.jian_cheng =="大理中支本部":
            self._name = "大理中支本部"
        else:
            sql_str = ("SELECT {0} FROM [{0}] WHERE {0}简称 = '{1}'".format(self.level, self.jian_cheng))
            self._name = self.sql.exec_query(sql_str)

        return self._name

    @property
    def ji_gou_fan_wei(self):
        """
        返回机构获取数据的范围
        """
        if self.name == "分公司整体":
            value = "[中心支公司] like '%'"
        elif self.name == "沾益":
            value = "[销售团队] = '01141101010002张笑晓团队' \
                    or [销售团队] =  '01141127010001沾益支公司团队' \
                    or [销售团队] = '01141127010002沾益支公司' \
                    or [销售团队] = '01141127010003沾益支公司1'"
        elif self.name == "祥云":
            value = "[业务员] = '114193962苏宝荣' \
                    or [业务员] = '214000161李红' \
                    or [业务员] = '214000268马宏春' \
                    or [业务员] = '214000272李朝勇' \
                    or [业务员] = '214000276杨炳莲'"
        elif self.name == "曲靖营业一部":
            value = "[销售团队] = '01141101010001李雄团队' \
                    or [销售团队] =  '01141101010004曲靖中支销售一部'"
        elif self.name == "曲靖中支本部":
            value = "[机构] = '0114110101曲靖中心支公司营业一部（虚拟）' \
                    and [销售团队] <> '01141101010002张笑晓团队' \
                    and [销售团队] <>  '01141127010001沾益支公司团队' \
                    and [销售团队] <> '01141127010002沾益支公司' \
                    and [销售团队] <> '01141127010003沾益支公司1' \
                    and [销售团队] <> '01141101010001李雄团队' \
                    and [销售团队] <>  '01141101010004曲靖中支销售一部'"
        elif self.name == "大理中支本部":
            value = "[机构] = '0114170101大理中心支公司营业一部（虚拟）' \
                    and [业务员] <> '114193962苏宝荣' \
                    and [业务员] <> '214000161李红' \
                    and [业务员] <> '214000268马宏春' \
                    and [业务员] <> '214000272李朝勇' \
                    and [业务员] <> '214000276杨炳莲'"
        elif self.name == "分公司本部":
            value = "[机构] = '0114010101云南分公司营业一部（虚拟）' \
                    or [机构] =  '0114010110云南分公司航旅出行项目（虚拟）'"
        else:
            value = "[{0}] = '{1}'".format(self.level, self.name)
        return value

    @property
    def this_week(self):
        """返回今年本周保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and (([{0}年].{2}))".format(self._date.year, self._date.weeknum, self.ji_gou_fan_wei)) 
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and (([{0}年].{2})) and [车险/财产险/人身险] <> '车险'".format(self._date.year, self._date.weeknum, self.ji_gou_fan_wei))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and (([{0}年].{2})) and [车险/财产险/人身险] = '{3}'".format(self._date.year, self._date.weeknum, self.ji_gou_fan_wei, self.risk)) 

        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0
    
    @property
    def last_year_week(self):
        """返回去年本周的保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2})".format(self._date.last_year, self._date.weeknum, self.ji_gou_fan_wei))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险'".format(self._date.last_year, self._date.weeknum, self.ji_gou_fan_wei))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}'".format(self._date.last_year, self._date.weeknum, self.ji_gou_fan_wei, self.risk))
            
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_week(self):
        """返回今年上周的保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2})".format(self._date.year, self._date.last_weeknum, self.ji_gou_fan_wei))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险'".format(self._date.year, self._date.last_weeknum, self.ji_gou_fan_wei))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}'".format(self._date.year, self._date.last_weeknum, self.ji_gou_fan_wei, self.risk))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0
    
    @property
    def last_year_last_week(self):
        """返回去年上周的保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2})".format(self._date.last_year, self._date.last_weeknum, self.ji_gou_fan_wei))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险'".format(self._date.last_year, self._date.last_weeknum, self.ji_gou_fan_wei))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}'"
                    .format(self._date.last_year, self._date.last_weeknum, self.ji_gou_fan_wei, self.risk))

        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def previous_week(self):
        """返回今年前一周的保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2})".format(self._date.year, self._date.previous_week, self.ji_gou_fan_wei))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险'".format(self._date.year, self._date.previous_week, self.ji_gou_fan_wei))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[周数] = {1} \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}'"
                    .format(self._date.year, self._date.previous_week, self.ji_gou_fan_wei, self.risk))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_week_tong_bi(self):
        """返回上一周的保费周同比"""
        if self.last_year_last_week == 0:
            return 0
        else:
            return (self.last_week / self.last_year_last_week) -1

    @property
    def last_week_huan_bi(self):
        """返回上周的保费周环比"""
        if self.previous_week == 0:
            return 0
        else:
            return (self.last_week / self.previous_week) -1

    @property
    def this_month(self):
        """返回今年本月保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [日期].[日期] < '{3}'".format(self._date.year, self._date.month, self.ji_gou_fan_wei, self._date.day))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险' and [日期].[日期] < '{3}'"
                    .format(self._date.year, self._date.month, self.ji_gou_fan_wei,  self._date.day))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}' and [日期].[日期] < '{4}'"
                    .format(self._date.year, self._date.month, self.ji_gou_fan_wei,  self.risk, self._date.day))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_year_month(self):
        """返回去年本月保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [日期].[日期] < '{3}'"
                    .format(self._date.last_year, self._date.month, self.ji_gou_fan_wei, self._date.day))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险' and [日期].[日期] < '{3}'"
                    .format(self._date.last_year, self._date.month, self.ji_gou_fan_wei, self._date.day))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}' and [日期].[日期] < '{4}'"
                    .format(self._date.last_year, self._date.month, self.ji_gou_fan_wei,  self.risk, self._date.day))
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_month(self):
        """返回今年上月保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [日期].[日期] < '{3}'"
                    .format(self._date.year, self._date.last_month, self.ji_gou_fan_wei, self._date.day))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险' and [日期].[日期] < '{3}'"
                    .format(self._date.year, self._date.last_month, self.ji_gou_fan_wei, self._date.day))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}' and [日期].[日期] < '{4}'"
                    .format(self._date.year, self._date.last_month, self.ji_gou_fan_wei,  self.risk, self._date.day))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def month_tong_bi(self):
        """返回今年本月保费同比"""
        if self.last_year_month == 0:
            return 0
        else:
            return (self.this_month / self.last_year_month) -1

    @property
    def month_huan_bi(self):
        """返回今年本月保费环比"""
        if self.last_month == 0:
            return 0
        else:
            return (self.this_month / self.last_month) -1

    @property
    def this_year(self):
        """返回今年保费累计"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [投保确认日期] < '{2}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self._date.date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] <> '车险' and [投保确认日期] < '{2}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self._date.date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] = '{2}' and [投保确认日期] < '{3}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self.risk, self._date.date))
         
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_year(self):
        """返回去年同期保费累计"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [投保确认日期] < '{2}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_date)) 
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] <> '车险' and [投保确认日期] < '{2}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_date)) 
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] = '{2}' and [投保确认日期] < '{3}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self.risk, self._date.last_date)) 
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def year_tong_bi(self):
        """返回年保费同比"""
        if self.last_year == 0:
            return 0
        else:
            return (self.this_year / self.last_year) -1

    @property
    def task(self):
        """返回机构计划任务"""
        if self.risk == "整体":
            sql_str = "SELECT [整体任务] FROM [计划任务] WHERE [机构] = '{0}'".format(self.jian_cheng)
        else:
            sql_str = "SELECT [{0}任务] FROM [计划任务] WHERE [机构] = '{1}'".format(self.risk, self.jian_cheng)
            
        value = self.sql.exec_query(sql_str)
        return value

    @property
    def planned_task_completion_rate(self):
        """返回机构的计划任务达成率"""
        value = self.this_year / self.task
        return value
    
    @property
    def time_progress_rate(self):
        """返回机构的时间进度达成率"""
        if self.planned_task_completion_rate == 0:
            return 0
        else:
            value = self.planned_task_completion_rate / self._date.time_progress
            return value

    @property
    def this_year_limit(self):
        """返回同比增长率统计表中今年保费累计"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [投保确认日期] <= '{2}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self._date.end_date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] <> '车险' and [投保确认日期] <= '{2}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self._date.end_date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] = '{2}' and [投保确认日期] <= '{3}'"
                   .format(self._date.year, self.ji_gou_fan_wei, self.risk, self._date.end_date))
         
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_year_limit(self):
        """返回同比增长率统计表中去年同期保费累计"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [投保确认日期] <= '{2}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_end_date)) 
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] <> '车险' and [投保确认日期] <= '{2}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_end_date)) 
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ({1}) and [车险/财产险/人身险] = '{2}' and [投保确认日期] <= '{3}'"
                   .format(self._date.last_year, self.ji_gou_fan_wei, self.risk, self._date.last_end_date)) 
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def this_year_month_limit(self):
        """返回今年本月保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [日期].[投保确认日期] <= '{3}'".format(self._date.year, self._date.end_month, self.ji_gou_fan_wei, self._date.end_date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险' and [日期].[投保确认日期] <= '{3}'"
                    .format(self._date.year, self._date.end_month, self.ji_gou_fan_wei,  self._date.end_date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}' and [日期].[投保确认日期] <= '{4}'"
                    .format(self._date.year, self._date.end_month, self.ji_gou_fan_wei,  self.risk, self._date.end_date))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_year_month_limit(self):
        """返回去年本月保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [日期].[投保确认日期] <= '{3}'"
                    .format(self._date.last_year, self._date.end_month, self.ji_gou_fan_wei, self._date.end_date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] <> '车险' and [日期].[投保确认日期] <= '{3}'"
                    .format(self._date.last_year, self._date.end_month, self.ji_gou_fan_wei, self._date.end_date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] join [日期] on [{0}年].[投保确认日期] = [日期].[投保确认日期] where [日期].[月份] = '{1}' \
                    and ([{0}年].{2}) and [车险/财产险/人身险] = '{3}' and [日期].[投保确认日期] <= '{4}'"
                    .format(self._date.last_year, self._date.end_month, self.ji_gou_fan_wei,  self.risk, self._date.end_date))
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def this_ten_days(self):
        """返回今年旬保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [投保确认日期] >= '{2}' and [投保确认日期] <= '{3}'"
                       .format(self._date.year, self.ji_gou_fan_wei, self._date.begin_date, self._date.end_date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [车险/财产险/人身险] <> '车险' and \
                       [投保确认日期] >= '{2}' and [投保确认日期] <= '{3}'"
                       .format(self._date.year, self.ji_gou_fan_wei, self._date.begin_date, self._date.end_date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [车险/财产险/人身险] = '{2}' and \
                       [投保确认日期] >= '{3}' and [投保确认日期] <= '{4}'"
                       .format(self._date.year, self.ji_gou_fan_wei, self.risk, self._date.begin_date, self._date.end_date))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def last_ten_days(self):
        """返回今年旬保费"""
        if self.risk == "整体":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [投保确认日期] >= '{2}' and [投保确认日期] <= '{3}'"
                       .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_begin_date, self._date.last_end_date))
        elif self.risk == "非车险":
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [车险/财产险/人身险] <> '车险' and \
                       [投保确认日期] >= '{2}' and [投保确认日期] <= '{3}'"
                       .format(self._date.last_year, self.ji_gou_fan_wei, self._date.last_begin_date, self._date.last_end_date))
        else:
            sql_str = ("select sum([签单保费/批改保费]) from [{0}年] where ([{0}年].{1}) and [车险/财产险/人身险] = '{2}' and \
                       [投保确认日期] >= '{3}' and [投保确认日期] <= '{4}'"
                       .format(self._date.last_year, self.ji_gou_fan_wei, self.risk, self._date.last_begin_date, self._date.last_end_date))
        
        value = self.sql.exec_query(sql_str)
        if value is not None :
            return value/10000
        else:
            return 0

    @property
    def year_tong_bi_limit(self):
        """返回同比增长率统计表中的年度保费同比"""
        value = (self.this_year_limit / self.last_year_limit) -1
        return value

    @property
    def month_tong_bi_limit(self):
        if self.last_year_month_limit == 0:
            value = -1
        else:
            value = (self.this_year_month_limit / self.last_year_month_limit) - 1
        return value

    @property
    def ten_days_tong_bi(self):
        if self.last_ten_days == 0:
            value = -1
        else:
            value = (self.this_ten_days / self.last_ten_days) -1
        return value
