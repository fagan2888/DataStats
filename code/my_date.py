from datetime import date
from datetime import timedelta


class MyDate(object):
    """
        定义程序所用的日期变量，无参数。
        
        属性：
        date： 返回今年当天的日期，格式为：‘YYYY-MM-DD’
        year： 返回今年的年份，格式为： ‘YYYY’
        month： 返回今年当月的月份，格式为： ‘MM’
        last_month：返回今年上个月的月份，格式为：‘MM’
        day：返回当年今天的日，格式为： ‘DD’
        weeknum：返回今年本周的周数，格式为：‘WW’
        last_weeknum：返回今年上周的周数，格式为：‘WW’
        previous_week：返回今年前一周的周数，格式为：‘WW’
        last_date：返回去年同期的日期，格式为：‘YYYY-MM-DD’
        last_year：返回去年同期的年份，格式为：‘YYYY’
    """

    @property
    def date(self):
        """返回今年当天的日期，格式为：‘YYYY-MM-DD’"""
        value = date.today().strftime("%Y-%m-%d")
        return value

    @property
    def year(self):
        """返回今年的年份，格式为： ‘YYYY’"""
        value = date.today().strftime("%Y")
        return value

    @property
    def month(self):
        """返回今年当月的月份，格式为： ‘MM’"""
        value = date.today().strftime("%m")
        return value

    @property
    def last_month(self):
        """返回今年上个月的月份，格式为：‘MM’"""
        value = '{:0>2}'.format(int(self.month) -1)
        return value

    @property
    def day(self):
        """返回当年今天的日，格式为： ‘DD’"""
        value = date.today().strftime("%d")
        return value

    @property
    def weeknum(self):
        """返回今年本周的周数，格式为：‘WW’"""
        value = '{:0>2}'.format(date.today().isocalendar()[1])
        return value

    @property
    def last_weeknum(self):
        """返回今年上周的周数，格式为：‘WW’"""
        return '{:0>2}'.format(int(self.weeknum) -1)

    @property
    def previous_week(self):
        """返回今年前一周的周数，格式为：‘WW’"""
        return '{:0>2}'.format(int(self.weeknum) -2)

    @property
    def last_date(self):
        """返回去年同期的日期，格式为：‘YYYY-MM-DD’"""
        value = date(int(self.year)-1, int(self.month), int(self.day)).isoformat()
        return value

    @property
    def last_year(self):
        """返回去年同期的年份，格式为：‘YYYY’"""
        return self.last_date[:4]
    
    @property
    def yday(self):
        """返回今年已过天数，int变量"""
        value= int(date.today().strftime("%j")) - 1
        return value
    
    @property
    def time_progress(self):
        """返回今年的时间进度，float变量"""
        ydays = date(int(self.year), 12, 31).strftime("%j")
        value = self.yday / int(ydays)
        return value

    @property
    def end_date(self):
        """返回同比增长率统计表中需要的数据截至时间"""
        if int(self.day)<=10:
            value = (date(2019, int(self.month), 1) - timedelta(days = 1)).strftime("%Y-%m-%d")
        elif int(self.day)<=20:
            value = date(2019, int(self.month), 10) .strftime("%Y-%m-%d")
        else:
            value = date(2019, int(self.month), 20) .strftime("%Y-%m-%d")
        return value

    @property
    def end_month(self):
        """返回同比增长率统计表中月统计的月份"""
        value = self.end_date[5:7]
        return value

    @property
    def begin_date(self):
        """返回同比增长率统计表中旬统计需要的数据起始时间"""
        if int(self.day)<=10:
            value = date(2019, int(self.end_month), 21) .strftime("%Y-%m-%d")
        elif int(self.day)<=20:
            value = date(2019, int(self.end_month), 1) .strftime("%Y-%m-%d")
        else:
            value = date(2019, int(self.end_month), 11) .strftime("%Y-%m-%d")
        return value

    @property
    def last_end_date(self):
        """返回同比增长率统计表中同期数据需要的数据截至时间"""
        value = self.last_year + self.end_date[4:]
        return value

    @property
    def last_begin_date(self):
        """返回同比增长率统计表中同期数据需要的数据起始时间"""
        value = self.last_year + self.begin_date[4:]
        return value
