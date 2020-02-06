import sqlite3
from functools import lru_cache
from datetime import date
import calendar

from code.date import IDate


class Tong_Ji(object):
    """
    数据统计的基类
    """

    def __init__(self, name=None, risk=None):
        """
        初始化统计类

        参数
            name:
                机构名称
                可以是机构名称或业务员姓名（暂不支持业务员）
            risk：
                险种名称
                无需设置险种类型，系统将根据险种名称自动判断险种类型
        """
        self._ming_cheng = name  # 机构简称
        # self._name_type = name_type  # 机构类型
        self._xian_zhong = risk  # 险种
        self._conn = sqlite3.connect(r"Data\data.db")
        self._cur = self._conn.cursor()
        self.d = IDate(2020)

    @property
    def cur(self) -> object:
        """
        返回数据库链接的游标对象

            特别说明：
                这是一个内部对象，在类的内部使用，请不要在类外部调用
            类型：
                属性
            参数：
                无
            返回：
                sqlite3.connect.cursor()
        """
        return self._cur

    @property
    def ming_cheng(self) -> str:
        """
        根据类初始化时输入的name参数，返回机构的简称

            类型：
                属性
            参数：
                无
            返回：
                str，机构简称
        """
        return self._ming_cheng

    @property
    def xian_zhong(self) -> str:
        """
        根据类初始化时输入的risk参数，返回险种名称

            类型：
                属性
            参数：
                无
            返回值：
                str，险种名称
        """
        return self._xian_zhong

    @property
    def xian_zhong_lei_xing(self) -> str:
        """
        根据险种名称判断险种类型，并返回险种类型

            类型：
                属性
            参数：
                无
            返回值：
                str，险种类型
        """
        xian_zhong = ["车险", "财产险", "人身险", "非车险"]

        xian_zhong_da_lei = [
            "保证险",
            "船舶险",
            "工程险",
            "货物运输保险",
            "家庭财产保险",
            "农业保险",
            "其他",
            "企业财产保险",
            "特殊风险保险",
            "责任保险",
            "综合保险",
            "交强险",
            "商业保险",
            "健康险",
            "意外伤害保险",
        ]

        if self.xian_zhong == "整体":
            return "整体"
        elif self.xian_zhong in xian_zhong:
            return "险种"
        elif self.xian_zhong in xian_zhong_da_lei:
            return "险种大类"
        else:
            return "险种名称"

    @property
    def ji_gou_lei_xing(self) -> str:
        """
        根据机构名称判断机构类型，并返回机构类型

            类型：
                属性
            参数：
                无
            返回值：
                str，机构类型
        """

        zhong_zhi = ["昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "昭通"]

        ji_gou = [
            "分公司本部",
            "百大国际",
            "春怡雅苑",
            "香榭丽园",
            "春之城",
            "东川",
            "宜良",
            "安宁",
            "航旅项目",
            "曲靖中支本部",
            "师宗",
            "宣威",
            "陆良",
            "沾益",
            "罗平",
            "会泽",
            "文山中支本部",
            "丘北",
            "马关",
            "广南",
            "麻栗坡",
            "富宁",
            "大理中支本部",
            "砚山",
            "祥云",
            "云龙",
            "宾川",
            "弥渡",
            "漾濞",
            "洱源",
            "版纳中支本部",
            "勐海",
            "勐腊",
            "保山中支本部",
            "施甸",
            "腾冲",
            "怒江中支本部",
            "兰坪",
        ]

        if self.ming_cheng == "分公司整体" or self.ming_cheng == "分公司":
            return "分公司"
        elif self.ming_cheng in zhong_zhi:
            return "中心支公司"
        elif self.ming_cheng in ji_gou:
            return "机构"
        else:
            return "销售团队"

    @lru_cache(maxsize=12)
    def ren_wu(self) -> int:
        """
        返回机构的计划任务，如果无任务则返回'——'

            类型:
                属性
            参数：
                无
            返回值：
                int, 计划任务
        """
        str_sql = f"SELECT [{self.xian_zhong}任务] \
            FROM [计划任务] \
            WHERE [机构] = '{self.ming_cheng}'"

        self.cur.execute(str_sql)

        for value in self.cur.fetchone():
            if value is None or value == 0:
                return "——"
            else:
                return int(value)

    @property
    def shi_jian_jin_du(self) -> float:
        """
        根据数据库中最大日期，返回当前时间的时间进度

            类型:
                属性
            参数：
                无
            返回值：
                float，时间进度
        """
        if calendar.isleap(int(self.d.nian)):
            zong_tian_shu = 366
        else:
            zong_tian_shu = 365
        tian_shu = date(self.d.nian, self.d.yue, self.d.ri).strftime("%j")
        return int(tian_shu) / zong_tian_shu

    @property
    def ren_wu_jin_du(self) -> float:
        """
        返回年度保费的计划任务达成率，如果上年度保费为零或任务为零则返回'——'

            类型:
                属性
            参数：
                无
            返回值：
                float，任务进度
        """
        if (
            self.ren_wu() is None
            or self.ren_wu() == 0
            or self.ren_wu() == "——"
        ):

            return "——"

        else:
            return self.nian_bao_fei() / self.ren_wu()

    @property
    def shi_jian_da_cheng(self) -> float:
        """
        返回当前累计保费的时间进度达成率，如果任务进度为零则返回'——'

            类型:
                属性
            参数：
                无
            返回值：
                float，时间进度达成率
        """
        if self.ren_wu_jin_du == "——":
            return "——"
        else:
            return self.ren_wu_jin_du / self.shi_jian_jin_du

    def ji_gou_join(self, year: int) -> str:
        """
        根据年份（参数year）及机构类型返回SQL语句的join语句中机构名称部分

            特别说明：
                这是一个类内部调用对象，请不要在类外部调用

            类型：
                方法
            参数：
                year：int值，一个4位数的年份
            返回值：
                str，一串SQL语言的join语句
        """

        if self.ji_gou_lei_xing == "分公司":
            str_sql = ""
        elif self.ji_gou_lei_xing == "中心支公司":
            str_sql = f"JOIN [中心支公司] \
                       ON [{year}年].[中心支公司] = [中心支公司].[中心支公司]"
        elif self.ji_gou_lei_xing == "机构":
            str_sql = f"JOIN [机构] \
                       ON [{year}年].[机构] = [机构].[机构]"
        elif self.ji_gou_lei_xing == "销售团队":
            str_sql = f"JOIN [销售团队] \
                       ON [{year}年].[销售团队] = [销售团队].[销售团队]"

        return str_sql

    @property
    def ji_gou_where(self) -> str:
        """
        根据机构类型返回SQL语句中的WHERE语句中的机构名称部分

            特别说明：
                这是一个类的内部对象，请不要在类外部调用
            类型：
                属性
            参数：
                无
            返回值：
                str，一串SQL语言的where语句
        """
        if self.ji_gou_lei_xing == "分公司":
            str_sql = ""
        elif self.ji_gou_lei_xing == "中心支公司":
            str_sql = f"AND [中心支公司].[中心支公司简称] = '{self.ming_cheng}'"
        elif self.ji_gou_lei_xing == "机构":
            str_sql = f"AND [机构].[机构简称] = '{self.ming_cheng}'"
        elif self.ji_gou_lei_xing == "销售团队":
            str_sql = f"AND [销售团队].[销售团队简称] = '{self.ming_cheng}'"

        return str_sql

    def xian_zhong_join(self, year: int) -> str:
        """'
        根据年份（参数year）及险种类型返回SQL语句的join语句中险种名称部分

            特别说明：
                这是一个类内部调用对象，请不要在类外部调用

            类型：
                方法
            参数：
                year：int值，一个4位数的年份
            返回值：
                str，一串SQL语言的join语句
        """

        if self.xian_zhong_lei_xing == "险种名称":
            str_sql = f"JOIN [险种名称] \
                       ON [{year}年].[险种名称] = [险种名称].[险种名称]"
        else:
            str_sql = ""

        return str_sql

    def xian_zhong_where(self, year: int) -> str:
        """
        根据险种类型返回SQL语句中的WHERE语句中的险种名称部分

            特别说明：
                这是一个类的内部对象，请不要在类外部调用
            类型：
                方法
            参数：
                year：int值，一个4位数的年份
            返回值：
                str，一串SQL语言的where语句
        """

        if self.xian_zhong_lei_xing == "整体":
            str_sql = ""
        elif self.xian_zhong_lei_xing == "险种":
            if self.xian_zhong == "非车险":
                str_sql = f"AND [{year}年].[车险/财产险/人身险] != '车险'"
            else:
                str_sql = f"AND [{year}年].[车险/财产险/人身险] = '{self.xian_zhong}'"
        elif self.xian_zhong_lei_xing == "险种大类":
            str_sql = f"AND [{year}年].[险种大类] = '{self.xian_zhong}'"
        elif self.xian_zhong_lei_xing == "险种名称":
            str_sql = f"AND [险种名称].[险种简称] = '{self.xian_zhong}'"

        return str_sql

    @lru_cache(maxsize=32)
    def nian_bao_fei(self) -> float:
        """
        返回机构的本年度累计保费

            类型：
                属性
            参数：
                无
            返回值：
                float, 机构的年累计保费
        """
        year = self.d.nian

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            WHERE [投保确认日期] <= '{self.d.long_ri_qi()}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @lru_cache(maxsize=32)
    def wang_nian_bao_fei(
        self, ny: int = None, year: int = None, tong: bool = True
    ) -> float:
        """
        返回往年的度累计保费

            类型：
                方法
            参数：
                ny: int值，用于计算需要返回倒数多少年的保费
                year：int值，一个4位数的年份值，用于计算需要返回倒数多少年的保费
                ny和year只需要输入一个参数，输入year参数后，ny参数将失效。
                tong：bool值，默认值为True，用于判断是否统计同期保费，输入False则统计全年累计保费
            返回值：
                float, 机构的年累计保费
        """
        if ny is not None:
            year = self.d.nian - ny

        if tong is False:
            ri_qi = self.d.long_ri_qi(year=year, month=12, day=31)
        else:
            ri_qi = self.d.long_ri_qi(year=year)

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year=year)} \
            {self.xian_zhong_join(year=year)} \
            WHERE [投保确认日期] <= '{ri_qi}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year=year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def nian_tong_bi(self, ny: int = None, year: int = None) -> float:
        """
        返回当年保费同比增长率

            类型：
                方法
            参数：
                ny: int值，用于计算需要返回倒数多少年的保费
                year：int值，一个4位数的年份值，用于计算需要返回倒数多少年的保费
                ny和year只需要输入一个参数，输入year参数后，ny参数将失效。
            返回值：
                float, 机构的当年保费同比增长率
        """
        if ny is not None:
            year = self.d.nian - ny

        if self.wang_nian_bao_fei(year=year) == 0:
            return "——"
        else:
            value = self.nian_bao_fei() / self.wang_nian_bao_fei(year=year) - 1
            return value

    def wang_nian_tong_bi(
        self, first_year: int, last_year: int, tong: bool = True
    ) -> float:
        """
        返回两年保费同比增长率

            类型：
                方法
            参数：
                first_year: int值，一个4位数的年份值，用于计算同比的第一年年保费
                last_year：int值，一个4位数的年份值，用于计算同比的第二年年保费
                tong：bool值，默认值为True，用于判断是否进行同期保费对比，输入False则进行全年累计保费对比
            返回值：
                float, 机构两年保费的同比增长率
        """

        if first_year == self.d.nian:
            tong = True

        if self.wang_nian_bao_fei(year=first_year, tong=tong) == 0:
            return "——"
        else:
            value = (
                self.wang_nian_bao_fei(year=first_year, tong=tong)
                / self.wang_nian_bao_fei(year=last_year, tong=tong)
                - 1
            )
            return value

    @lru_cache(maxsize=32)
    def ji_bao_fei(self) -> float:
        """
        返回机构的本季度累计保费

            类型：
                属性
            参数：
                无
            返回值：
                float, 机构的季度累计保费
        """
        year = self.d.nian
        quarter = self.d.ji_du

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[季度] = '{quarter}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def ji_tong_bi(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
    ):
        """
        返回季度保费同比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """

        if ny is not None:
            year = self.d.nian - ny

        if nq is not None:
            quarter = self.d.ji_du - nq

        if self.wang_ji_bao_fei(year=year, quarter=quarter) == 0:
            return "——"

        value = (
            self.ji_bao_fei()
            / self.wang_ji_bao_fei(year=year, quarter=quarter)
            - 1
        )
        return value

    @lru_cache(maxsize=32)
    def wang_ji_bao_fei(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
        tong: bool = True,
        short_date: str = None,
    ):
        """
        返回往期季度累计保费
        """
        if ny is not None:
            year = self.d.nian - ny

        if nq is not None:
            quarter = self.d.ji_du - nq

        if short_date is None:
            if tong is True:
                short_date = self.d.duan_ri_qi
            else:
                short_date = "12-31"

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[季度] = '{quarter}' \
            AND [日期].[日期] <= '{short_date}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    # def wang_ji_tong_bi(
    #     self,
    #     first_year: int = None,
    #     last_year: int = None,
    #     first_quarter: int = None,
    #     last_quarter: int = None,
    #     tong: bool = True,
    #     first_short_date: str = None,
    #     last_short_date: str = None,
    # ):
    #     """
    #     返回往年季度保费同比增长率

    #     参数：

    #         ny: 同比倒数年数
    #             与一年前数据对比则设置为1
    #             与两年前数据对比则设置为2
    #             以此类推
    #     """

    #     if (
    #         self.wang_ji_bao_fei(
    #             year=first_year,
    #             quarter=first_quarter,
    #             short_date=first_short_date,
    #         )
    #         == 0
    #         or self.wang_ji_bao_fei(
    #             year=last_year,
    #             quarter=last_quarter,
    #             short_date=last_short_date,
    #         )
    #         == 0
    #     ):

    #         return "——"

    #     value = (
    #         self.wang_ji_bao_fei(
    #             year=first_year,
    #             quarter=first_quarter,
    #             short_date=first_short_date,
    #         )
    #         / self.wang_ji_bao_fei(
    #             year=last_year,
    #             quarter=last_quarter,
    #             short_date=last_short_date,
    #         )
    #         - 1
    #     )
    #     return value

    # def wang_ji_huan_bi(
    #     self,
    #     year: int = None,
    #     quarter: int = None,
    #     tong: bool = True
    # ):
    #     """
    #     返回往年季度保费环比增长率

    #     参数：

    #         ny: 同比倒数年数
    #             与一年前数据对比则设置为1
    #             与两年前数据对比则设置为2
    #             以此类推
    #     """

    #     if tong is True:
    #         if querter == 1:
    #             month = 12 - self.d.yue - 3
    #             loat_short_date = f"{month:02d}-{self.d.ri:02d}"


    #     if (
    #         self.wang_ji_bao_fei(
    #             year=first_year,
    #             quarter=first_quarter,
    #             short_date=first_short_date,
    #         )
    #         == 0
    #         or self.wang_ji_bao_fei(
    #             year=last_year,
    #             quarter=last_quarter,
    #             short_date=last_short_date,
    #         )
    #         == 0
    #     ):

    #         return "——"

    #     value = (
    #         self.wang_ji_bao_fei(
    #             year=first_year,
    #             quarter=first_quarter,
    #             short_date=first_short_date,
    #         )
    #         / self.wang_ji_bao_fei(
    #             year=last_year,
    #             quarter=last_quarter,
    #             short_date=last_short_date,
    #         )
    #         - 1
    #     )
    #     return value

    def ji_huan_bi(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
    ):
        """
        返回季度保费环比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """

        if ny is not None:
            year = self.d.nian - ny

        if nq is not None:
            quarter = self.d.yue - nq

        if nq >= self.d.ji_du:
            year = year - 1
            quarter = 4 - (nq - self.d.ji_du)

        if self.wang_ji_bao_fei(year=year, quarter=quarter) == 0:
            return "——"
        else:
            value = (
                self.ji_bao_fei()
                / self.wang_ji_bao_fei(year=year, quarter=quarter)
                - 1
            )
            return value

    @lru_cache(maxsize=32)
    def yue_bao_fei(self):
        """
        返回月度累计保费
        """

        year = self.d.nian
        month = self.d.yue

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[月份] = '{month}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @lru_cache(maxsize=32)
    def wang_yue_bao_fei(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
        tong: bool = True,
    ):
        """
        返回往期月度累计保费
        """
        if ny is not None:
            year = self.d.nian - ny

        if nm is not None:
            month = self.d.yue - nm

        if tong is True:
            day = self.d.ri
        else:
            day = 31

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[月份] = '{month}' \
            AND [日期].[日数] <= '{day}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def yue_tong_bi(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
    ):
        """
        返回月度保费同比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """

        if ny is not None:
            year = self.d.nian - ny

        if nm is not None:
            month = self.d.yue - nm

        if self.wang_yue_bao_fei(year=year, month=month) == 0:
            return "——"

        value = (
            self.yue_bao_fei() / self.wang_yue_bao_fei(year=year, month=month)
            - 1
        )
        return value

    def yue_huan_bi(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
    ):
        """
        返回月度保费环比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """

        if ny is not None:
            year = self.d.nian - ny

        if nm is not None:
            month = self.d.yue - nm

        if nm >= self.d.yue:
            year = year - 1
            month = 12 - (nm - self.d.yue)

        if self.wang_yue_bao_fei(year=year, month=month) == 0:
            return "——"
        else:
            value = (
                self.yue_bao_fei()
                / self.wang_yue_bao_fei(year=year, month=month)
                - 1
            )
            return value

    @lru_cache(maxsize=32)
    def zhou_bao_fei(self):
        """
        返回周累计保费
        """

        year = self.d.nian
        weeknum = self.d.zhou

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[周数] = '{weeknum}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @lru_cache(maxsize=32)
    def wang_zhou_bao_fei(
        self,
        ny: int = None,
        year: int = None,
        nw: int = None,
        weeknum: int = None,
    ):
        """
        返回往期周累计保费
        """
        if ny is not None:
            year = self.d.nian - ny

        if nw is not None:
            weeknum = self.d.zhou - nw

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[周数] = '{weeknum}' \
            {self.ji_gou_where} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def zhou_tong_bi(
        self,
        ny: int = None,
        year: int = None,
        nw: int = None,
        weeknum: int = None,
    ):
        """
        返回周保费同比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """

        if ny is not None:
            year = self.d.nian - ny

        if nw is not None:
            weeknum = self.d.zhou - nw

        if self.wang_zhou_bao_fei(year=year, weeknum=weeknum) == 0:
            return "——"
        else:
            value = (
                self.zhou_bao_fei()
                / self.wang_zhou_bao_fei(year=year, weeknum=weeknum)
                - 1
            )
            return value

    def zhou_huan_bi(
        self,
        ny: int = None,
        year: int = None,
        nw: int = None,
        weeknum: int = None,
    ):
        """
        返回周保费环比增长率

        参数：

            ny: 同比倒数年数
                与一年前数据对比则设置为1
                与两年前数据对比则设置为2
                以此类推
        """
        if ny is not None:
            year = self.d.nian - ny

        if nw is not None:
            weeknum = self.d.zhou - nw

        if self.wang_zhou_bao_fei(year=year, weeknum=weeknum) == 0:
            return "——"
        else:
            value = (
                self.zhou_bao_fei()
                / self.wang_zhou_bao_fei(year=year, weeknum=weeknum)
                - 1
            )
            return value

            self.ji


if __name__ == "__main__":
    data = Tong_Ji(name="分公司", risk="整体")
    data.ji
