import sqlite3
from datetime import datetime


class Tong_ji(object):
    """
    以一个机构为实例，返回计算报表相关数据
    """
    def __init__(self, name, risk):
        self._jian_cheng = name  # 机构简称
        self._xian_zhong = risk  # 险种
        self._conn = sqlite3.connect(r'Data\data.db')
        self._cur = self._conn.cursor()

    @property
    def jin_tian(self):
        '''
        返回今天日期
        '''
        return datetime.today().strftime('%Y-%m-%d')

    @property
    def jin_nian(self):
        '''
        返回今天年份
        '''
        return self.jin_tian[:4]

    @property
    def jin_yue(self):
        '''
        返回今天月份
        '''
        return self.jin_tian[5:7]

    @property
    def jin_ri(self):
        '''
        返回今天日数
        '''
        return self.jin_tian[8:10]

    @property
    def jian_cheng(self):
        '''
        返回机构简称
        '''
        return self._jian_cheng

    @property
    def xian_zhong(self):
        '''
        返回险种
        '''
        return self._xian_zhong

    @property
    def lei_ji_bao_fei(self):
        '''
        返回年底累计保费
        '''
        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        WHERE [投保确认日期] < '{self.jin_tian}'"
            elif self.xian_zhong == '车险':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] < '{self.jin_tian}'"
            elif self.xian_zhong == '非车险':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] < '{self.jin_tian}'"
        else:
            if self.xian_zhong == '整体':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        JOIN [机构] \
                        ON [{self.jin_nian}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{self.jin_tian}'"
            elif self.xian_zhong == '车险':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        JOIN [机构] \
                        ON [{self.jin_nian}年].[机构] = [机构].[机构] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{self.jin_tian}'"
            elif self.xian_zhong == '非车险':
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.jin_nian}年] \
                        JOIN [机构] \
                        ON [{self.jin_nian}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{self.jin_tian}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yi_nian_lei_ji_bao_fei(self):
        '''
        返回1年同期累计保费
        '''
        year = int(self.jin_nian) - 1
        date = f'{year}-{self.jin_yue}-{self.jin_ri}'

        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE [投保确认日期] < '{date}'"
            elif self.xian_zhong == '车险':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] < '{date}'"
            elif self.xian_zhong == '非车险':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] < '{date}'"
        else:
            if self.xian_zhong == '整体':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{date}'"
            elif self.xian_zhong == '车险':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{date}'"
            elif self.xian_zhong == '非车险':
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] < '{date}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yi_nian_tong_bi(self):
        if self.yi_nian_lei_ji_bao_fei == 0:
            return 1
        else:
            return self.lei_ji_bao_fei / self.yi_nian_lei_ji_bao_fei - 1


if __name__ == '__main__':
    ji_gou = Tong_ji('罗平', '车险')
    print(ji_gou.jian_cheng,
          ji_gou.lei_ji_bao_fei,
          ji_gou.yi_nian_lei_ji_bao_fei)
