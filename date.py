import sqlite3


class IDate():
    '''
    使用当前数据表中最大日期初始化当前日期
    对日期进行计算
    '''
    def __init__(self, year):
        self._conn = sqlite3.connect(r"Data\data.db")
        self._cur = self._conn.cursor()

        sql_str = f"ATTACH DATABASE 'Data\\{year}年.db' AS [{year}年]"
        self._cur.execute(sql_str)

        str_sql = f"SELECT [年份], [月份], [日数], [星期], \
                    [周数], [季度], [旬]\
                    FROM [日期] \
                    WHERE [投保确认日期] = \
                    (SELECT MAX([投保确认日期]) \
                    FROM [{year}年])"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()

        self._nian = year
        self._yue = value[1]
        self._ri = value[2]
        self._xin_qi = value[3]
        self._zhou = value[4]
        self._ji_du = value[5]
        self._xun = value[6]

        self._year = year
        self._month = value[1]
        self._day = value[2]
        self._weekday = value[3]
        self._weeknum = value[4]
        self._quarter = value[5]
        self._ten_days = value[6]

    @property
    def nian(self):
        '''
        返回数据库中最大日期的年份
        '''
        return self._nian

    @property
    def yue(self):
        '''
        返回数据库中最大日期的月份
        '''
        return self._yue

    @property
    def ri(self):
        '''
        返回数据库中最大日期的日数
        '''
        return self._ri

    @property
    def zhou(self):
        '''
        返回数据库中最大日期的周数
        '''
        return self._zhou

    @property
    def xin_qi(self):
        '''
        返回数据库中最大日期的星期
        '''
        return self._xin_qi

    @property
    def ji_du(self):
        '''
        返回数据库中最大日期的季度
        '''
        return self._ji_du

    @property
    def xun(self):
        '''
        返回数据库中最大日期的旬数
        '''
        return self._xun

    def long_ri_qi(self, year=None, month=None, day=None):
        '''
        返回长日期，格式为：YYYY-MM-DD，如果无参数则返回当前数据库中最大日期

            参数:
                year：
                    int，一个4位数的年份，如果未输入则返回当前数据库中最大年份
                month：
                    int，一个月份，如果未输入则返回当前数据库中最大月份
                day：
                    int，一个日数，如果未输入则返回当前数据库中最大日数
            返回值：
                str，YYYY-MM-DD格式的日期
        '''
        if year is None:
            year = self.nian

        if month is None:
            month = self.yue

        if day is None:
            day = self.ri

        return f'{year}-{month:02d}-{day:02d}'

    def duan_ri_qi(self, month=None, day=None):
        '''
        返回短日期

            格式为：MM-DD
            如果无参数则返回当前数据库中最大日期
        '''

        if month is None:
            month = self.yue

        if day is None:
            day = self.ri

        return f'{month:02d}-{day:02d}'

    @property
    def year(self):
        '''
        返回数据库中最大日期的年份
        '''
        return self._year

    @property
    def month(self):
        '''
        返回数据库中最大日期的月份
        '''
        return self._month

    @property
    def day(self):
        '''
        返回数据库中最大日期的日数
        '''
        return self._day

    @property
    def weeknum(self):
        '''
        返回数据库中最大日期的周数
        '''
        return self._weeknum

    @property
    def weekday(self):
        '''
        返回数据库中最大日期的星期
        '''
        return self._weekday

    @property
    def quarter(self):
        '''
        返回数据库中最大日期的季度
        '''
        return self._quarter

    @property
    def ten_days(self):
        '''
        返回数据库中最大日期的旬数
        '''
        return self._ten_days

    def long_date(self, year=None, month=None, day=None):
        '''
        返回长日期，格式为：YYYY-MM-DD，如果无参数则返回当前数据库中最大日期

            参数:
                year：
                    int，一个4位数的年份，如果未输入则返回当前数据库中最大年份
                month：
                    int，一个月份，如果未输入则返回当前数据库中最大月份
                day：
                    int，一个日数，如果未输入则返回当前数据库中最大日数
            返回值：
                str，YYYY-MM-DD格式的日期
        '''
        if year is None:
            year = self.year

        if month is None:
            month = self.month

        if day is None:
            day = self.day

        return f'{year}-{month:02d}-{day:02d}'

    def short_date(self, month=None, day=None):
        '''
        返回短日期

            格式为：MM-DD
            如果无参数则返回当前数据库中最大日期
        '''

        if month is None:
            month = self.month

        if day is None:
            day = self.day

        return f'{month:02d}-{day:02d}'
