import sqlite3
from datetime import datetime
from datetime import date
import calendar


class Tong_ji(object):
    '''
    数据统计的基类
    '''

    def __init__(self, name, risk):
        self._ji_gou = name  # 机构简称
        self._xian_zhong = risk  # 险种
        self._nian = None
        self._yue = None
        self._ri = None
        self._xin_qi = None
        self._zhou = None
        self._ji_du = None
        self._xun = None
        self._biao_ming = None
        self._conn = sqlite3.connect(r'Data\data.db')
        self._cur = self._conn.cursor()
        self.set_ri_qi()

    @property
    def biao_ming(self):
        return self._biao_ming

    @biao_ming.setter
    def biao_ming(self, value):
        self._biao_ming = 2020 - value

    def set_ri_qi(self):
        '''
        使用当前数据表中最大日期初始化当前日期
        '''
        self.biao_ming = 0

        str_sql = f"SELECT [年份], [月份], [日数], [星期], \
                    [周数], [季度], [旬] \
                    FROM [日期] \
                    WHERE [投保确认日期] = \
                    (SELECT MAX([投保确认日期]) \
                    FROM [{self.biao_ming}年])"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()
        self._nian = value[0]
        self._yue = value[1]
        self._ri = value[2]
        self._xin_qi = value[3]
        self._zhou = value[4]
        self._ji_du = value[5]
        self._xun = value[6]

    @property
    def nian(self):
        '''
        返回数据表中的年份
        '''
        return self._nian

    @property
    def yue(self):
        '''
        返回数据表中最大的月份
        '''
        return self._yue

    @property
    def ri(self):
        '''
        返回数据表中最大日数
        '''
        return self._ri

    @property
    def xin_qi(self):
        '''
        返回数据表中星期数
        '''
        return self._xin_qi

    @property
    def zhou(self):
        '''
        返回数据表中周数
        '''
        return self._zhou

    @property
    def ji_du(self):
        '''
        返回数据表中季度
        '''
        return self._ji_du

    @property
    def xun(self):
        '''
        返回数据表中旬
        '''
        return self._xun

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
            if (value is None
               or value == 0):
                return '——'
            else:
                return int(value)

    @property
    def ji_gou_join(self):
        '''
        返回SQL语句关联表中机构名称部分
        '''
        if self.ji_gou_ji_bie == '分公司':
            str_sql = ''
        elif self.ji_gou_ji_bie == '中心支公司':
            str_sql = f"JOIN [中心支公司] \
                       ON [{self.biao_ming}年].[中心支公司] = [中心支公司].[中心支公司]"
        elif self.ji_gou_ji_bie == '机构':
            str_sql = f"JOIN [机构] \
                       ON [{self.biao_ming}年].[机构] = [机构].[机构]"
        elif self.ji_gou_ji_bie == '销售团队':
            str_sql = f"JOIN [销售团队] \
                       ON [{self.biao_ming}年].[销售团队] = [销售团队].[销售团队]"
        return str_sql

    @property
    def ji_gou_where(self):
        '''
        返回SQL语句查询条件中机构名称部分
        '''
        if self.ji_gou_ji_bie == '分公司':
            str_sql = ''
        elif self.ji_gou_ji_bie == '中心支公司':
            str_sql = f"AND [中心支公司].[中心支公司简称] = {self.jian_cheng}"
        elif self.ji_gou_ji_bie == '机构':
            str_sql = f"AND [机构].[机构简称] = {self.jian_cheng}"
        elif self.ji_gou_ji_bie == '销售团队':
            str_sql = f"AND [销售团队].[销售团队简称] = {self.jian_cheng}"
        return str_sql

    @property
    def nian_bao_fei(self):
        '''
        返回本年度累计保费
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
