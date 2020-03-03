import sqlite3
from functools import lru_cache


class Stats_0621():
    """
    APP出单激励方案统计类

        关于APP出单激励方案的各项统计功能

        # 属性：
        #     get_day_premium:
        #         返回指定年份的全部日期的保费，返回的是从数据库中提出的一个日保费列表
        #     day_premium:
        #         返回指定年份的全部日期的保费，返回是进行二次匹配的日保费字典，处理了闰年的问题
        #     day_yoy:
        #         返回指定年份的日保费同比增长率，返回的是一个字典，包含指定年份的每一天数据
        #     day_sum:
        #         返回指定年份的日累计保费，返回的是一个字典，包含指定年份的每一天数据
        #     day_sum_yoy：
        #         返回指定年份的日累计保费的同比增长率，返回的是一个字典，包含指定年份的每一天数据
    """
    def __init__(self):
        self._set_conn()
        self._set_cur()

    def _set_conn(self):
        """
        设置数据库链接对象

            特别说明：
                这是一个内部对象，在类的内部使用，请不要在类外部调用
            参数：
                无
            返回值：
                无
        """
        self._conn = sqlite3.connect(r"Data\data.db")

    def _set_cur(self):
        """
        设置数据库链接的游标对象

            特别说明：
                这是一个内部对象，在类的内部使用，请不要在类外部调用
            参数：
                无
            返回值：
                无
        """
        self._cur = self.conn.cursor()

    @property
    def conn(self) -> sqlite3.Connection:
        """
        返回数据库链接对象

            特别说明：
                这是一个内部对象，在类的内部使用，请不要在类外部调用
            参数：
                无
            返回值：
                sqlite3.Connection
        """
        return self._conn

    @property
    def cur(self) -> sqlite3.Cursor:
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

    def attach_db(self):
        """
        附加数据库
        """

        sql_str = r"ATTACH DATABASE 'Data\2020年.db' AS [2020年]"
        self.cur.execute(sql_str)
        sql_str = r"ATTACH DATABASE 'Data\车险清单.db' AS [车险清单]"
        self.cur.execute(sql_str)

    def detach_db(self):
        """
        分离数据库
        """
        sql_str = "DETACH DATABASE [2020年]"
        self.cur.execute(sql_str)
        sql_str = "DETACH DATABASE [车险清单]"
        self.cur.execute(sql_str)

    @lru_cache(maxsize=32)
    def get_center_branch(self, month: int = None, week: int = None):
        """
        返回中心支公司的APP签单数据

            返回值：
                list
        """

        if month is None and week is None:
            sum_sql_str = ""
            risk_sql_str = ""
        elif month is not None:
            sum_sql_str = f"AND [日期].[月份] = '{month}'"
            risk_sql_str = f"AND [日期].[月份] = '{month}'"
        elif week is not None:
            sum_sql_str = f"AND [日期].[周数] = '{week}'"
            risk_sql_str = f"AND [日期].[周数] = '{week}'"

        sql_str = f"CREATE TEMP VIEW [总件数] \
            AS \
            SELECT \
                [中心支公司].[ROWID], \
                [中心支公司].[中心支公司简称] AS [中心支公司], \
                COUNT (DISTINCT [车架号]) AS [车辆数] \
            FROM   [车险清单] \
                JOIN [中心支公司] \
                ON [车险清单].[中心支公司] = [中心支公司].[中心支公司] \
                JOIN [日期] \
                ON [车险清单].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [使用性质] = '非营业' \
                AND [机动车种类] IN ('客车', '货车') \
                AND [车辆类型] IN ('二吨以下货车', '六座以下客车', '六座至十座客车') \
                AND [座位数] < 8 \
                {sum_sql_str} \
            GROUP  BY [车险清单].[中心支公司] \
            ORDER  BY [中心支公司].[ROWID]"
        self.cur.execute(sql_str)

        sql_str = f"CREATE TEMP VIEW [驾意险] \
            AS \
            SELECT \
                [中心支公司].[ROWID], \
                [中心支公司].[中心支公司简称] AS [中心支公司], \
                COUNT ([保单笔数]) AS [签单数量], \
                SUM ([签单保费/批改保费]) / 10000 AS [保费] \
            FROM   [2020年] \
                JOIN [中心支公司] \
                ON [2020年].[中心支公司] = [中心支公司].[中心支公司] \
                JOIN [日期] \
                ON [2020年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [险种名称] = '0621驾乘人员人身意外伤害保险(B款)' \
                {risk_sql_str} \
            GROUP BY [中心支公司].[中心支公司简称] \
            ORDER BY [中心支公司].[ROWID]"
        self.cur.execute(sql_str)

        sql_str = "SELECT \
                [总件数].[中心支公司], \
                [驾意险].[签单数量], \
                [驾意险].[保费], \
                [总件数].[车辆数], \
                [驾意险].[签单数量] * 1.0 / [总件数].[车辆数] AS [联动率] \
            FROM   [总件数] \
                LEFT JOIN [驾意险] \
                ON [驾意险].[中心支公司] = [总件数].[中心支公司]"
        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "DROP VIEW [驾意险]"
        self.cur.execute(sql_str)
        sql_str = "DROP VIEW [总件数]"
        self.cur.execute(sql_str)

        return value

    @lru_cache(maxsize=32)
    def get_company(self, month: int = None, week: int = None):
        """
        返回机构的APP签单数量

            参数：
                name:
                    str, 中心支公司名称

            返回值：
                list
        """

        if month is None and week is None:
            sum_sql_str = ""
            risk_sql_str = ""
        elif month is not None:
            sum_sql_str = f"AND [日期].[月份] = '{month}'"
            risk_sql_str = f"AND [日期].[月份] = '{month}'"
        elif week is not None:
            sum_sql_str = f"AND [日期].[周数] = '{week}'"
            risk_sql_str = f"AND [日期].[周数] = '{week}'"

        sql_str = f"CREATE TEMP VIEW [总件数] \
            AS \
            SELECT \
                [机构].[机构简称] AS [机构], \
                COUNT (DISTINCT [车架号]) AS [车辆数] \
            FROM   [车险清单] \
                JOIN [机构] \
                ON [车险清单].[机构] = [机构].[机构] \
                JOIN [日期] \
                ON [车险清单].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [使用性质] = '非营业' \
                AND [机动车种类] IN ('客车', '货车') \
                AND [车辆类型] IN ('二吨以下货车', '六座以下客车', '六座至十座客车') \
                AND [座位数] < 8 \
                {sum_sql_str} \
            GROUP  BY [车险清单].[机构] \
            ORDER  BY [机构].[机构]"
        self.cur.execute(sql_str)

        sql_str = f"CREATE TEMP VIEW [驾意险] \
            AS \
            SELECT \
                [机构].[机构简称] AS [机构], \
                COUNT ([保单笔数]) AS [签单数量], \
                SUM ([签单保费/批改保费]) / 10000 AS [保费] \
            FROM   [2020年] \
                JOIN [机构] \
                ON [2020年].[机构] = [机构].[机构] \
                JOIN [日期] \
                ON [2020年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE  [险种名称] = '0621驾乘人员人身意外伤害保险(B款)' \
                {risk_sql_str} \
            GROUP BY [机构].[机构简称] \
            ORDER BY [机构].[机构]"
        self.cur.execute(sql_str)

        sql_str = "SELECT \
                [总件数].[机构], \
                [驾意险].[签单数量], \
                [驾意险].[保费], \
                [总件数].[车辆数], \
                [驾意险].[签单数量] * 1.0 / [总件数].[车辆数] AS [联动率] \
            FROM   [总件数] \
                LEFT JOIN [驾意险] \
                ON [驾意险].[机构] = [总件数].[机构]"
        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "DROP VIEW [驾意险]"
        self.cur.execute(sql_str)
        sql_str = "DROP VIEW [总件数]"
        self.cur.execute(sql_str)

        return value


if __name__ == "__main__":
    app = Stats_0621()

    app.attach_db()
    value = app.get_center_branch()

    for value in value:
        print(value)

    app.detach_db()
