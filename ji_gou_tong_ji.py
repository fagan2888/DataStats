import sqlite3
from datetime import date
import calendar


class Tong_ji(object):
    """
    以一个机构为实例，返回计算报表相关数据
    """
    def __init__(self, name, risk):
        self._jian_cheng = name  # 机构简称
        self._xian_zhong = risk  # 险种
        self._nian = None
        self._yue = None
        self._ri = None
        self._xin_qi = None
        self._zhou = None
        self._ji_du = None
        self._xun = None
        self._conn = sqlite3.connect(r'Data\data.db')
        self._cur = self._conn.cursor()
        self.set_ri_qi()

    def set_ri_qi(self):
        str_sql = f"SELECT [年份], [月份], [日数], [星期], \
                    [周数], [季度], [旬] \
                    FROM [日期] \
                    WHERE [投保确认日期] = \
                    (SELECT MAX([投保确认日期]) \
                    FROM [2020年])"

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
        返回数据表中最大日数
        '''
        return self._xin_qi

    @property
    def zhou(self):
        '''
        返回数据表中最大日数
        '''
        return self._zhou

    @property
    def ji_du(self):
        '''
        返回数据表中最大日数
        '''
        return self._ji_du

    @property
    def xun(self):
        '''
        返回数据表中最大日数
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
    def xian_zhong_ji_bie(self):
        '''
        返回险种级别
        '''
        return self._xian_zhong_ji_bie

    @property
    def ji_gou_ji_bie(self):
        '''
        返回机构级别
        '''
        return self._ji_gou_ji_bie

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
    def nian_bao_fei(self):
        '''
        返回年底累计保费
        '''
        ri_qi = f'{self.nian}-{self.yue:02d}-{self.ri:02d}'

        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                # 查询分公司整体保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询分公司车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询分公司非车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] <= '{ri_qi}'"
            else:
                # 查询分公司分险种保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [险种名称] \
                        ON [{self.nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{ri_qi}'"
        else:
            if self.xian_zhong == '整体':
                # 查询机构整体保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询机构车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询机构非车险保费数据
                str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                        FROM [{self.nian}年] \
                        JOIN [机构] \
                        ON [{self.nian}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
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
                        AND [投保确认日期] <= '{ri_qi}'"

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
        nian = int(self.nian) - 1
        ri_qi = f'{nian}-{self.yue:02d}-{self.ri:02d}'

        if self.jian_cheng == '分公司整体':
            if self.xian_zhong == '整体':
                # 查询分公司整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        WHERE [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询分公司车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        WHERE [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询分公司非车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [投保确认日期] <= '{ri_qi}'"
            else:
                # 查询分公司驾意险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        JOIN [险种名称] \
                        ON [{nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [投保确认日期] <= '{ri_qi}'"
        else:
            if self.xian_zhong == '整体':
                # 查询机构整体保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        JOIN [机构] \
                        ON [{nian}年].[机构] = [机构].[机构] \
                        WHERE [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong in ['车险', '财产险', '人身险']:
                # 查询机构车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        JOIN [机构] \
                        ON [{nian}年].[机构] = [机构].[机构] \
                        WHERE  [车险/财产险/人身险] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            elif self.xian_zhong == '非车险':
                # 查询机构非车险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        JOIN [机构] \
                        ON [{nian}年].[机构] = [机构].[机构] \
                        WHERE ([车险/财产险/人身险] = '人身险' \
                        OR [车险/财产险/人身险] = '财产险') \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"
            else:
                # 查询机构驾意险保费数据
                str_sql = f"SELECT SUM([签单保费/批改保费]) \
                        FROM [{nian}年] \
                        JOIN [机构] \
                        ON [{nian}年].[机构] = [机构].[机构] \
                        JOIN [险种名称] \
                        ON [{nian}年].[险种名称] = [险种名称].[险种名称] \
                        WHERE  [险种名称].[险种简称] = '{self.xian_zhong}' \
                        AND [机构].[机构简称] = '{self.jian_cheng}' \
                        AND [投保确认日期] <= '{ri_qi}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def yue_bao_fei(self):
        '''
        返回月累计保费
        '''
        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{self.nian}年] \
                    JOIN [机构] \
                    ON [{self.nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{self.nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{self.nian}' \
                    AND [日期].[月份] = '{self.yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{self.nian}年] \
                    JOIN [机构] \
                    ON [{self.nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{self.nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{self.nian}' \
                    AND [日期].[月份] = '{self.yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def tong_yue_bao_fei(self):
        '''
        返回同期月累计保费
        '''
        nian = self.nian - 1

        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[月份] = '{self.yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[月份] = '{self.yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"

        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def huan_yue_bao_fei(self):
        '''
        返回同期月累计保费
        '''
        if self.yue == 1:
            nian = self.nian - 1
            yue = 12
        else:
            nian = self.nian
            yue = self.yue - 1

        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[月份] = '{yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[月份] = '{yue}' \
                    AND [日期].[日数] <= '{self.ri}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        
        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def zhou_bao_fei(self):
        '''
        返回周累计保费
        '''
        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{self.nian}年] \
                    JOIN [机构] \
                    ON [{self.nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{self.nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{self.nian}' \
                    AND [日期].[周数] <= '{self.zhou}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{self.nian}年] \
                    JOIN [机构] \
                    ON [{self.nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{self.nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{self.nian}' \
                    AND [日期].[周数] <= '{self.zhou}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        
        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def tong_zhou_bao_fei(self):
        '''
        返回同期月累计保费
        '''
        nian = self.nian - 1

        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[周数] <= '{self.zhou}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[周数] <= '{self.zhou}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        
        self._cur.execute(str_sql)
        for value in self._cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @property
    def huan_zhou_bao_fei(self):
        '''
        返回同期月累计保费
        '''
        if self.zhou == 1:
            nian = self.nian - 1
            zhou = 52
        else:
            nian = self.nian
            zhou = self.zhou - 1

        if self.xian_zhong == '整体':
            # 查询机构整体保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[周数] <= '{zhou}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"
        else:
            # 查询机构车险保费数据
            str_sql = f"SELECT SUM ([签单保费/批改保费]) \
                    FROM [{nian}年] \
                    JOIN [机构] \
                    ON [{nian}年].[机构] = [机构].[机构] \
                    JOIN [日期] \
                    ON [{nian}年].[投保确认日期] = [日期].[投保确认日期] \
                    WHERE  [日期].[年份] = '{nian}' \
                    AND [日期].[周数] <= '{zhou}' \
                    AND [车险/财产险/人身险] = '{self.xian_zhong}' \
                    AND [机构].[机构简称] = '{self.jian_cheng}'"

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
    def yue_tong_bi(self):
        if self.tong_yue_bao_fei == 0:
            return '——'
        else:
            return self.yue_bao_fei / self.tong_yue_bao_fei - 1

    @property
    def zhou_tong_bi(self):
        if self.tong_zhou_bao_fei == 0:
            return '——'
        else:
            return self.zhou_bao_fei / self.tong_zhou_bao_fei - 1

    @property
    def yue_huan_bi(self):
        if self.huan_yue_bao_fei == 0:
            return '——'
        else:
            return self.yue_bao_fei / self.huan_yue_bao_fei - 1

    @property
    def zhou_huan_bi(self):
        if self.huan_zhou_bao_fei == 0:
            return '——'
        else:
            return self.zhou_bao_fei / self.huan_zhou_bao_fei - 1

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
        tian_shu = date(self.nian, self.yue,
                        self.ri).strftime('%j')
        return int(tian_shu) / zong_tian_shu

    @property
    def shi_jian_da_cheng_lv(self):
        if self.ren_wu_da_cheng_lv == '——':
            return '——'
        else:
            return self.ren_wu_da_cheng_lv / self.shi_jian_jin_du


if __name__ == '__main__':
    ji_gou = Tong_ji('罗平', '车险')
    print(ji_gou.jian_cheng,
          ji_gou.yue_bao_fei,
          ji_gou.tong_yue_bao_fei,
          ji_gou.yue_tong_bi)
# ji_gou.huan_yue_bao_fei,
#           ji_gou.yue_huan_bi,
#           ji_gou.zhou_bao_fei,
#           ji_gou.tong_zhou_bao_fei,
#           ji_gou.huan_zhou_bao_fei,
#           ji_gou.zhou_tong_bi,
#           ji_gou.zhou_huan_bi
