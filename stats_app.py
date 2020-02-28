import sqlite3
from functools import lru_cache


class Stats_App():
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

        sql_str = r"ATTACH DATABASE 'Data\掌上宝APP出单统计.db' AS [掌上宝APP出单统计]"
        self.cur.execute(sql_str)
        sql_str = r"ATTACH DATABASE 'Data\车险机动车类型.db' AS [车险机动车类型]"
        self.cur.execute(sql_str)

    def detach_db(self):
        """
        分离数据库
        """
        sql_str = "DETACH DATABASE [掌上宝APP出单统计]"
        self.cur.execute(sql_str)
        sql_str = "DETACH DATABASE [车险机动车类型]"
        self.cur.execute(sql_str)

    @lru_cache(maxsize=32)
    def get_salesman(self, week: int = None) -> list:
        """
        返回业务员APP的出单数据

            返回值：
                list
        """

        if week is None:
            sum_sql_str = ""
            app_sql_str = ""
        else:
            sum_sql_str = f"AND [日期].[周数] = '{week}'"
            app_sql_str = f"AND [日期].[周数] = '{week}'"

        sql_str = f"CREATE TEMP VIEW [总签单件数] \
            AS \
            SELECT \
                [机构].[中心支公司简称], \
                [机构].[机构简称], \
                [业务员], \
                COUNT ([掌上宝APP出单统计].[保单号]) AS [总签单件数], \
                SUM ([签单保费/批改保费]) AS[保费], \
                ROUND (ABS ([保险期限]) / 86400) AS [保期] \
            FROM 掌上宝APP出单统计.掌上宝APP出单统计 \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
                WHERE NOT ([车险/财产险/人身险] = '车险' AND 保期 < 360) \
            {sum_sql_str} \
            GROUP  BY [业务员] \
            ORDER  BY [总签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = f"CREATE TEMP VIEW [APP签单件数] \
            AS \
            SELECT \
                [机构].[中心支公司简称], \
                [机构].[机构简称], \
                [业务员], \
                COUNT ([保单号]) AS [APP签单件数], \
                SUM ([签单保费/批改保费]) AS [保费] \
            FROM   [掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE  [终端来源] = '0106移动展业(App)' \
            {app_sql_str} \
            GROUP  BY [业务员] \
            ORDER  BY [APP签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = f"SELECT \
                [总签单件数].[中心支公司简称] AS [中支], \
                [总签单件数].[机构简称] AS [机构], \
                [总签单件数].[业务员] AS [业务员], \
                [APP签单件数], \
                [APP签单件数].[保费]/10000 AS [保费], \
                [总签单件数].[总签单件数], \
                [APP签单件数] * 1.0 / [总签单件数] AS [APP出单占比] \
            FROM   [总签单件数] \
                LEFT JOIN [APP签单件数] ON [总签单件数].[业务员] = [APP签单件数].[业务员] \
            ORDER  BY [APP签单件数] DESC, [保费] DESC, [APP出单占比] DESC, [总签单件数] DESC"

        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "DROP VIEW [APP签单件数]"
        self.cur.execute(sql_str)
        sql_str = "DROP VIEW 总签单件数"
        self.cur.execute(sql_str)

        return value

    @lru_cache(maxsize=32)
    def get_terminal(self, week: int = None):
        """
        返回终端类型的签单数据

            返回值：
                list
        """

        if week is None:
            sum_sql_str = ""
        else:
            sum_sql_str = f"WHERE [日期].[周数] = '{week}'"

        sql_str = f"SELECT [终端来源], COUNT([保单号]) AS [签单数量] \
            FROM [掌上宝APP出单统计].[掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
            {sum_sql_str} \
            GROUP BY [终端来源] \
            UNION \
            SELECT '合计', COUNT([保单号]) AS [合计] \
            FROM [掌上宝APP出单统计].[掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
            {sum_sql_str} \
            ORDER BY [终端来源]"

        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "SELECT COUNT([保单号]) AS [合计] \
            FROM [掌上宝APP出单统计].[掌上宝APP出单统计]"

        self.cur.execute(sql_str)
        value_sum = self.cur.fetchone()[0]

        return value, value_sum

    @lru_cache(maxsize=32)
    def get_center_branch(self, week: int = None):
        """
        返回中心支公司的APP签单数据

            返回值：
                list
        """

        if week is None:
            sum_sql_str = ""
            app_sql_str = ""
        else:
            sum_sql_str = f"AND [日期].[周数] = '{week}'"
            app_sql_str = f"AND [日期].[周数] = '{week}'"

        sql_str = f"CREATE TEMP VIEW [APP签单件数] \
            AS \
            SELECT \
                [机构].[中心支公司简称], \
                COUNT ([保单号]) AS [APP签单件数] \
            FROM   [掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE  [终端来源] = '0106移动展业(App)' \
            {app_sql_str} \
            GROUP  BY [机构].[中心支公司简称] \
            ORDER  BY [APP签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = f"CREATE TEMP VIEW [总签单件数] \
            AS \
            SELECT \
                [机构].[中心支公司简称], \
                COUNT ([保单号]) AS [总签单件数], \
                ROUND (ABS ([保险期限]) / 86400) AS [保期] \
            FROM   掌上宝APP出单统计.掌上宝APP出单统计 \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE [机构].机构简称 != '航旅项目' \
            IS NOT ([车险/财产险/人身险] = '车险' AND 保期 < 360) \
            {sum_sql_str} \
            GROUP  BY [机构].[中心支公司简称] \
            ORDER  BY [总签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = "SELECT \
                [总签单件数].[中心支公司简称] AS [中支], \
                [APP签单件数], \
                [总签单件数].[总签单件数], \
                [APP签单件数] * 1.0 / [总签单件数] AS [APP出单占比] \
            FROM   [总签单件数] \
                JOIN [APP签单件数] ON [总签单件数].[中心支公司简称] = [APP签单件数].[中心支公司简称] \
            ORDER  BY [APP签单件数] DESC"
        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "DROP VIEW [APP签单件数]"
        self.cur.execute(sql_str)
        sql_str = "DROP VIEW 总签单件数"
        self.cur.execute(sql_str)

        return value

    @lru_cache(maxsize=32)
    def get_company(self, name, week: int = None):
        """
        返回机构的APP签单数量

            参数：
                name:
                    str, 中心支公司名称

            返回值：
                list
        """

        if week is None:
            sum_sql_str = ""
            app_sql_str = ""
        else:
            sum_sql_str = f"AND [日期].[周数] = '{week}'"
            app_sql_str = f"AND [日期].[周数] = '{week}'"

        sql_str = f"CREATE TEMP VIEW [APP签单件数] \
            AS \
            SELECT \
                [机构].[机构简称] AS [机构简称], \
                COUNT ([保单号]) AS [APP签单件数] \
            FROM   [掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE [终端来源] = '0106移动展业(App)' \
            AND [机构].[中心支公司简称] = '{name}' \
            AND [机构].[机构简称] != '航旅项目' \
            {app_sql_str} \
            GROUP  BY [机构].[机构简称] \
            ORDER  BY [APP签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = f"CREATE TEMP VIEW [总签单件数] \
            AS \
            SELECT \
                [机构].[机构简称] AS [机构简称], \
                COUNT ([保单号]) AS[总签单件数], \
                ROUND (ABS ([保险期限]) / 86400) AS [保期] \
            FROM   掌上宝APP出单统计.掌上宝APP出单统计 \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE [机构].[中心支公司简称] = '{name}' \
            AND [机构].[机构简称] != '航旅项目' \
            IS NOT ([车险/财产险/人身险] = '车险' AND 保期 < 360) \
            {sum_sql_str} \
            GROUP  BY [机构].[机构简称] \
            ORDER  BY [总签单件数] DESC"
        self.cur.execute(sql_str)

        sql_str = "SELECT \
                [总签单件数].[机构简称] AS [机构], \
                [APP签单件数], \
                [总签单件数].[总签单件数], \
                [APP签单件数] * 1.0 / [总签单件数] AS [APP出单占比] \
            FROM   [总签单件数] \
                LEFT JOIN [APP签单件数] ON [总签单件数].[机构简称] = [APP签单件数].[机构简称] \
            ORDER  BY [APP签单件数] DESC"
        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = "DROP VIEW [APP签单件数]"
        self.cur.execute(sql_str)
        sql_str = "DROP VIEW 总签单件数"
        self.cur.execute(sql_str)

        return value


if __name__ == "__main__":
    app = Stats_App()

    app.attach_db()
    value, value_sum = app.get_terminal()

    for value in value:
        print(value)

    print(value_sum)

    app.detach_db()
