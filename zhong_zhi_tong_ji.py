import sqlite3
from datetime import datetime
from datetime import date
import calendar


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
    def ren_wu(self):
        '''
        返回计划任务
        '''
        sql_str = f"SELECT [{self.xian_zhong}任务] \
                FROM [计划任务] \
                WHERE [机构] = '{self.jian_cheng}'"
        self._cur.execute(sql_str)

        for value in self._cur.fetchone():
            if value is None:
                return '——'
            else:
                return int(value)

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
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询分公司大险种保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
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
                # 查询中心支公司整体保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [中心支公司] \
                        ON [{self.nian}年].[中心支公司] = [中心支公司].[中心支公司] \
                        WHERE [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询中心支公司分大险种保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [中心支公司] \
                        ON [{self.nian}年].[中心支公司] = [中心支公司].[中心支公司] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{self.wan_zheng_ri_qi}'"
            else:
                # 查询中心支公司分险种保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [中心支公司] \
                        ON [{self.nian}年].[中心支公司] = [中心支公司].[中心支公司] \
                        JOIN [险种名称] \
                        ON [{self.nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
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
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询分公司大险种保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{date}'"
            else:
                # 查询分公司分险种保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [险种名称] \
                        ON [{year}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{date}'"
        else:
            if self.xian_zhong == '整体':
                # 查询中心支公司整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [中心支公司] \
                        ON [{year}年].[中心支公司] = [中心支公司].[中心支公司] \
                        WHERE [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询中心支公司整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [中心支公司] \
                        ON [{year}年].[中心支公司] = [中心支公司].[中心支公司] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{date}'"
            else:
                # 查询中心支公司整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{year}年] \
                        JOIN [中心支公司] \
                        ON [{year}年].[中心支公司] = [中心支公司].[中心支公司] \
                        JOIN [险种名称] \
                        ON [{year}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [中心支公司].[中心支公司简称] = '{self.jian_cheng}' \
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

    @property
    def ren_wu_da_cheng_lv(self):
        if self.ren_wu == '——':
            return '——'
        else:
            return self.nian_bao_fei / self.ren_wu

    @property
    def shi_jian_jin_du(self):
        if calendar.isleap(int(self.nian)):
            zong_tian_shu = 366
        else:
            zong_tian_shu = 365
        tian_shu = date(int(self.nian), int(self.yue),
                        int(self.ri)).strftime('%j')
        return int(tian_shu) / zong_tian_shu

    @property
    def shi_jian_da_cheng_lv(self):
        if self.ren_wu_da_cheng_lv == '——':
            return '——'
        else:
            return self.ren_wu_da_cheng_lv / self.shi_jian_jin_du


if __name__ == '__main__':
    zhong_zhi = Tong_ji('曲靖', '车险')
    print(zhong_zhi.jian_cheng,
          zhong_zhi.nian_bao_fei,
          zhong_zhi.ren_wu,
          zhong_zhi.ren_wu_da_cheng_lv,
          zhong_zhi.shi_jian_da_cheng_lv)
