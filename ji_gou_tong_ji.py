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
    def wan_zheng_ri_qi(self):
        '''
        返回数据表中最大日期
        '''
        str_sql = "SELECT MAX([投保确认日期]) \
                   FROM   [2020年]"
        self._cur.execute(str_sql)
        value = self._cur.fetchone()
        return value[0]

    @property
    def nian(self):
        '''
        返回数据表中的年份
        '''
        return self.wan_zheng_ri_qi[:4]

    @property
    def yue(self):
        '''
        返回数据表中最大的月份
        '''
        return self.wan_zheng_ri_qi[5:7]

    @property
    def ri(self):
        '''
        返回数据表中最大日数
        '''
        return self.wan_zheng_ri_qi[8:10]

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
    def nian_bao_fei(self):
        '''
        返回年底累计保费
        '''
        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                # 查询分公司整体保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            elif self.xian_zhong == '车险':
                # 查询分公司车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询分公司非车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            else:
                # 查询分公司分险种保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [险种名称] \
                        ON [{self.nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
        else:
            if self.xian_zhong == '整体':
                # 查询机构整体保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            elif self.xian_zhong == '车险':
                # 查询机构车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询机构非车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            else:
                # 查询机构驾意险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        JOIN [险种名称] \
                        ON [{self.nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yi_nian_bao_fei(self):
        '''
        返回1年同期累计保费
        '''
        year = int(self.nian) - 1
        date = f'{year}-{self.yue}-{self.ri}'

        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                # 查询分公司整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE [投保确认日期] <= '{date}'"
            elif self.xian_zhong == '车险':
                # 查询分公司车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{date}'"
            elif self.xian_zhong == '非车险':
                # 查询分公司非车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] <= '{date}'"
            else:
                # 查询分公司驾意险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [险种名称] \
                        ON [{year}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{date}'"
        else:
            if self.xian_zhong == '整体':
                # 查询机构整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"
            elif self.xian_zhong == '车险':
                # 查询机构车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"
            elif self.xian_zhong == '非车险':
                # 查询机构非车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"
            else:
                # 查询机构驾意险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [机构] \
                        ON [{year}年].[机构] = [机构].[机构] \
                        JOIN [险种名称] \
                        ON [{year}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yi_nian_tong_bi(self):
        if self.yi_nian_bao_fei == 0:
            return '——'
        else:
            return self.nian_bao_fei / self.yi_nian_bao_fei - 1


if __name__ == '__main__':
    ji_gou = Tong_ji('罗平', '车险')
    print(ji_gou.jian_cheng,
          ji_gou.nian_bao_fei,
          ji_gou.yi_nian_lei_ji_bao_fei)
