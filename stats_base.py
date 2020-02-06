import sqlite3
from functools import lru_cache
from datetime import date
import calendar

from date import IDate


class Stats_Base(object):
    """
    数据统计类

        数据统计中所需要的基础信息在本类中进行初始化

    属性：

        cur:
            返回一个数据库的游标对象，便于其他属性中操作数据
        name:
            返回机构简称
        risk:
            返回统计的险种
        risk_type：
            返回险种类型
        company_type：
            返回机构类型
        task：
            返回本年度任务
        time_progress：
            返回指定日期的的时间进度
        company_join：
            返回查询当前机构所需的SQL语句中的JOIN部分
        company_where：
            返回查询当前机构所需的SQL语句中的WHERE部分
        risk_join：
            返回查询当前险种所需的SQL语句中的JOIN部分
        risk_where：
            返回查询当前险种所需的SQL语句中的WHERE部分
    """

    def __init__(
        self, name: str = None, risk: str = None, conn: sqlite3.Connection = None
    ):
        """
        初始化统计类

            参数
                name:
                    str, 机构名称
                risk：
                    str, 险种名称，无需设置险种类型，系统将根据险种名称自动判断险种类型
        """
        self._name = name  # 机构简称
        self._risk = risk  # 险种

        self._conn = conn
        self._cur = self._conn.cursor()

        self.d = IDate(2020)

        # 设置该机构对应险种的计划任务
        self._set_task()

        # 初始化日期列表
        self._set_day_list()

    @property
    def name(self) -> str:
        """
        返回机构名称

            参数：
                    无
                返回值：
                    str
        """
        return self._name

    @property
    def risk(self) -> str:
        """
        返回险种名称

            参数：
                无
            返回值：
                str
        """
        return self._risk

    @property
    def cur(self) -> object:
        """
        返回数据库链接的游标对象

            特别说明：
                这是一个内部对象，在类的内部使用，请不要在类外部调用
            参数：
                无
            返回值：
                sqlite3.Cursor
        """
        return self._cur

    @property
    def risk_type(self) -> str:
        """
        根据险种名称判断险种类型，并返回险种类型

            参数：
                无
            返回值：
                str
        """
        risk_list = ["车险", "财产险", "人身险", "非车险"]

        risk_type_list = [
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

        if self.risk == "整体":
            return "整体"
        elif self.risk in risk_list:
            return "险种"
        elif self.risk in risk_type_list:
            return "险种大类"
        else:
            return "险种名称"

    @property
    def company_type(self) -> str:
        """
        根据机构名称判断机构类型，并返回机构类型

            参数：
                无
            返回值：
                str
        """

        center_branch_list = ["昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "昭通"]

        company_list = [
            "分公司本部",
            "航旅项目",
            "百大国际",
            "春怡雅苑",
            "香榭丽园",
            "春之城",
            "东川",
            "宜良",
            "安宁",
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

        if self.name == "分公司整体" or self.name == "分公司":
            return "分公司"
        elif self.name in center_branch_list:
            return "中心支公司"
        elif self.name in company_list:
            return "机构"
        else:
            return "销售团队"

    @property
    def task(self) -> int:
        """
        返回机构的计划任务，如果无任务则返回"——"

            参数：
                无
            返回值：
                int, str("——")
        """
        return self._task

    @lru_cache(maxsize=32)
    def _set_task(self) -> int:
        """
        返回机构的计划任务，如果无任务则返回"——"

            参数：
                无
            返回值：
                int, str("——")
        """
        sql_str = f"SELECT [{self.risk}任务] \
            FROM [计划任务] \
            WHERE [机构] = '{self.name}'"

        self.cur.execute(sql_str)

        for value in self.cur.fetchone():
            if value is None or value == 0:
                self._task = "——"
            else:
                self._task = int(value)

    def time_progress(
        self, year: int = None, month: int = None, day: int = None
    ) -> float:
        """
        根据指定日期，返回当前时间的时间进度

            参数：
                year:
                    int, 指定需要计算的年份，默认值为数据库中最大的年份
                month:
                    int, 指定需要计算的月份，默认值为数据库中最大的月份
                day:
                    int, 指定需要计算的日期，默认值为数据库中最大的日期
            返回值：
                float
        """

        # 判断是否是闰年
        if calendar.isleap(int(self.d.year)):
            num_day_sum = 366
        else:
            num_day_sum = 365

        if year is None:
            year = self.d.year

        if month is None:
            month = self.d.month

        if day is None:
            day = self.d.day

        # 获取已经过去的天数
        past_num_day = date(year, month, day).strftime("%j")

        return int(past_num_day) / num_day_sum

    def company_join(self, year: int) -> str:
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

        if self.company_type == "分公司":
            sql_str = ""
        elif self.company_type == "中心支公司":
            sql_str = f"JOIN [中心支公司] \
                       ON [{year}年].[中心支公司] = [中心支公司].[中心支公司]"
        elif self.company_type == "机构":
            sql_str = f"JOIN [机构] \
                       ON [{year}年].[机构] = [机构].[机构]"
        elif self.company_type == "销售团队":
            sql_str = f"JOIN [销售团队] \
                       ON [{year}年].[销售团队] = [销售团队].[销售团队]"

        return sql_str

    def company_where(self) -> str:
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
        if self.company_type == "分公司":
            sql_str = ""
        elif self.company_type == "中心支公司":
            sql_str = f"AND [中心支公司].[中心支公司简称] = '{self.name}'"
        elif self.company_type == "机构":
            sql_str = f"AND [机构].[机构简称] = '{self.name}'"
        elif self.company_type == "销售团队":
            sql_str = f"AND [销售团队].[销售团队简称] = '{self.name}'"

        return sql_str

    def risk_join(self, year: int) -> str:
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

        if self.risk_type == "险种名称":
            sql_str = f"JOIN [险种名称] \
                       ON [{year}年].[险种名称] = [险种名称].[险种名称]"
        else:
            sql_str = ""

        return sql_str

    def risk_where(self, year: int) -> str:
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

        if self.risk_type == "整体":
            sql_str = ""
        elif self.risk_type == "险种":
            if self.risk == "非车险":
                sql_str = f"AND [{year}年].[车险/财产险/人身险] != '车险'"
            else:
                sql_str = f"AND [{year}年].[车险/财产险/人身险] = '{self.risk}'"
        elif self.risk_type == "险种大类":
            sql_str = f"AND [{year}年].[险种大类] = '{self.risk}'"
        elif self.risk_type == "险种名称":
            sql_str = f"AND [险种名称].[险种简称] = '{self.risk}'"

        return sql_str

    @property
    def day_list(self):
        """
        获取一年中所有的日期

            返回值：list
        """

        return self._day_list

    def _set_day_list(self):
        """
        获取一年中所有的日期

            返回值：list
        """
        # 2020年为闰年，可获取到2月29日的日期
        sql_str = f"SELECT [日期].[日期] \
            FROM [日期] \
            WHERE [日期].[年份] = 2020 \
            ORDER BY [日期].[日期]"

        self.cur.execute(sql_str)

        data_list = []

        for keys in self.cur.fetchall():
            for key in keys:
                data_list.append(key)

        self._day_list = data_list
