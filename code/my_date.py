from datetime import date


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
        value = date.today()
        return value

    @property
    def year(self):
        """返回今年的年份，格式为： ‘YYYY’"""
        value = '{:0>2}'.format(date.today().strftime("%Y"))
        return value

    @property
    def month(self):
        """返回今年当月的月份，格式为： ‘MM’"""
        value = '{:0>2}'.format(date.today().strftime("%m"))
        return value

    @property
    def last_month(self):
        """返回今年上个月的月份，格式为：‘MM’"""
        value = '{:0>2}'.format(int(self.month) -1)
        return value

    @property
    def day(self):
        """返回当年今天的日，格式为： ‘DD’"""
        value = '{:0>2}'.format(date.today().strftime("%d"))
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
