import sqlite3
from datetime import date
import calendar

from code.date import IDate


class Tong_Ji(object):
    '''
    数据统计的基类
    '''
    def __init__(self,
                 name='分公司整体',
                 name_leve='分公司',
                 risk='整体',
                 risk_leve='整体'):
        self._ji_gou = name  # 机构简称
        self._xian_zhong = risk  # 险种
        self._xian_zhong_ji_bie = risk_leve  # 险种类型
        self._ji_gou_ji_bie = name_leve  # 机构类型
        self._conn = sqlite3.connect(r'Data\data.db')
        self._cur = self._conn.cursor()
        self.d = IDate(2020)

    def __del__(self):
        self._cur.close()
        self._conn.close()

    @property
    def cur(self):
        return self._cur

    @property
    def ji_gou(self):
        '''
        返回机构简称
        '''
        return self._ji_gou

    @property
    def xian_zhong(self):
        '''
        返回险种
        '''
        return self._xian_zhong

    @property
    def xian_zhong_ji_bie(self):
        '''
        返回险种类型
        '''
        return self._xian_zhong_ji_bie

    @property
    def ji_gou_ji_bie(self):
        '''
        返回机构类型
        '''
        return self._ji_gou_ji_bie

    @property
    def ren_wu(self):
        '''
        返回计划任务
        '''
        sql_str = f"SELECT [{self.xian_zhong}任务] \
                FROM [计划任务] \
                WHERE [机构] = '{self.ji_gou}'"

        self.cur.execute(sql_str)

        for value in self.cur.fetchone():
            if (value is None or value == 0):
                return '——'
            else:
                return int(value)

    @property
    def shi_jian_jin_du(self):
        '''
        返回当前时间的时间进度
        '''
        if calendar.isleap(int(self.d.nian)):
            zong_tian_shu = 366
        else:
            zong_tian_shu = 365
        tian_shu = date(self.d.nian, self.d.yue,
                        self.d.ri).strftime('%j')
        return int(tian_shu) / zong_tian_shu

    @property
    def ren_wu_jin_du(self):
        '''
        返回年度保费的计划任务达成率
        '''
        if self.ren_wu is None \
           or self.ren_wu == 0 \
           or self.ren_wu == '——':

            return "——"
        else:
            return self.nian_bao_fei / self.ren_wu

    @property
    def shi_jian_da_cheng(self):
        '''
        返回当前累计保费的时间进度达成率
        '''
        if self.ren_wu_jin_du == '——':
            return '——'
        else:
            return self.ren_wu_jin_du / self.shi_jian_jin_du

    def ji_gou_join(self, n):
        '''
        返回SQL语句关联表中机构名称部分
        '''
        nian = self.d.nian - n

        if self.ji_gou_ji_bie == '分公司':
            str_sql = ''

        elif self.ji_gou_ji_bie == '中心支公司':
            str_sql = f"JOIN [中心支公司] \
                       ON [{nian}年].[中心支公司] = [中心支公司].[中心支公司]"

        elif self.ji_gou_ji_bie == '机构':
            str_sql = f"JOIN [机构] \
                       ON [{nian}年].[机构] = [机构].[机构]"

        elif self.ji_gou_ji_bie == '销售团队':
            str_sql = f"JOIN [销售团队] \
                       ON [{nian}年].[销售团队] = [销售团队].[销售团队]"

        return str_sql

    def ji_gou_where(self):
        '''
        返回SQL语句查询条件中机构名称部分
        '''
        if self.ji_gou_ji_bie == '分公司':
            str_sql = ''
        elif self.ji_gou_ji_bie == '中心支公司':
            str_sql = f"AND [中心支公司].[中心支公司简称] = '{self.ji_gou}'"
        elif self.ji_gou_ji_bie == '机构':
            str_sql = f"AND [机构].[机构简称] = '{self.ji_gou}'"
        elif self.ji_gou_ji_bie == '销售团队':
            str_sql = f"AND [销售团队].[销售团队简称] = '{self.ji_gou}'"
        return str_sql

    def xian_zhong_join(self, n):
        '''
        返回SQL语句关联表个中险种名称部分
        '''

        nian = self.d.nian - n

        if self.xian_zhong_ji_bie == '险种名称':
            str_sql = f"JOIN [险种名称] \
                       ON [{nian}年].[险种名称] = [险种名称].[险种名称]"

        else:
            str_sql = ''

        return str_sql

    def xian_zhong_where(self, n):
        '''
        返回SQL语句查询条件中险种名称部分
        '''

        nian = self.d.nian - n

        if self.xian_zhong_ji_bie == '整体':
            str_sql = ''
        elif self.xian_zhong_ji_bie == '险种':
            if self.xian_zhong == '非车险':
                str_sql = f"AND [{nian}年].[车险/财产险/人身险] != '车险'"
            else:
                str_sql = f"AND [{nian}年].[车险/财产险/人身险] = '{self.xian_zhong}'"

        elif self.xian_zhong_ji_bie == '险种大类':
            str_sql = f"AND [{nian}年].[险种大类] = '{self.xian_zhong}'"

        elif self.xian_zhong_ji_bie == '险种名称':
            str_sql = f"AND [险种名称].[险种简称] = '{self.xian_zhong}'"

        return str_sql

    @property
    def nian_bao_fei(self):
        '''
        返回本年度累计保费
        '''

        ri_qi = f"{self.d.nian}-{self.d.yue:02d}-{self.d.ri:02d}"

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{self.d.nian}年] \
            {self.ji_gou_join(0)} \
            {self.xian_zhong_join(0)} \
            WHERE [投保确认日期] <= '{ri_qi}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(0)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def tong_nian_bao_fei(self, n):
        '''
        返回同期本年度累计保费
        '''
        nian = self.d.nian - n
        ri_qi = f"{nian}-{self.d.yue:02d}-{self.d.ri:02d}"

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{nian}年] \
            {self.ji_gou_join(n)} \
            {self.xian_zhong_join(n)} \
            WHERE [投保确认日期] <= '{ri_qi}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(n)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yue_bao_fei(self):
        '''
        返回月度累计保费
        '''

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{self.d.nian}年] \
            {self.ji_gou_join(0)} \
            {self.xian_zhong_join(0)} \
            JOIN [日期] \
            ON [{self.d.nian}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[月份] = '{self.d.yue}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(0)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def tong_yue_bao_fei(self, year, month, day):
        '''
        返回同期月度累计保费
        '''
        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(0)} \
            {self.xian_zhong_join(0)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[月份] = '{month}' \
            AND [日期].[日数] <= '{day}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(0)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def nian_tong_bi(self, n):
        '''
        '''
        if self.tong_nian_bao_fei(1) == 0:
            return '——'
        value = self.nian_bao_fei / self.tong_nian_bao_fei(1) - 1
        return value



if __name__ == '__main__':
    ji_gou = Tong_Ji('曲靖', '中心支公司', '驾意险', '险种名称')
    print(ji_gou.nian_bao_fei)
    print(ji_gou.tong_nian_bao_fei(1))
    print(ji_gou.tong_nian_bao_fei(2))
    print(ji_gou.tong_nian_bao_fei(3))
    print(ji_gou.tong_nian_bao_fei(4))
    print(ji_gou.tong_nian_bao_fei(5))
