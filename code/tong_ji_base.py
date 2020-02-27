import sqlite3
from io import StringIO
from functools import lru_cache
from datetime import date
import calendar

from code.date import IDate


class Tong_Ji_Base(object):
    """
    数据统计类

        数据统计中所需要的基础信息在本类中进行初始化

    属性：

        cur:
            返回一个数据库的游标对象，便于其他属性中操作数据
        ming_cheng:
            返回机构简称
        xian_zhong:
            返回统计的险种
        xian_zhong_lei_xing：
            返回险种类型
        ji_gou_lei_xing：
            返回机构类型
        ren_wu：
            返回本年度任务
        shi_jian_jin_du：
            返回当前时间的时间进度
        ji_gou_join：
            返回查询当前机构所需的SQL语句中的JOIN部分
        ji_gou_where：
            返回查询当前机构所需的SQL语句中的WHERE部分
        xian_zhong_join：
            返回查询当前险种所需的SQL语句中的JOIN部分
        xian_zhong_where：
            返回查询当前险种所需的SQL语句中的WHERE部分
    """

    def __init__(
        self,
        name: str = None,
        risk: str = None,
        conn: sqlite3.Connection = None
    ):
        """
        初始化统计类

            参数
                name:
                    str, 机构名称
                risk：
                    str, 险种名称，无需设置险种类型，系统将根据险种名称自动判断险种类型
        """
        self._ming_cheng = name  # 机构简称
        self._xian_zhong = risk  # 险种

        self._conn = conn
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

        if self.ming_cheng == "分公司整体" or self.ming_cheng == "分公司":
            return "分公司"
        elif self.ming_cheng in zhong_zhi:
            return "中心支公司"
        elif self.ming_cheng in ji_gou:
            return "机构"
        else:
            return "销售团队"

    @lru_cache(maxsize=32)
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

    def shi_jian_jin_du(self) -> float:
        """
        根据数据库中最大日期，返回当前时间的时间进度

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
