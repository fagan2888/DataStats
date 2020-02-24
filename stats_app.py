from functools import lru_cache
from stats_base import Stats_Base


class Stats_App(Stats_Base):
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

    @lru_cache(maxsize=32)
    def get_Salesman_sum(self) -> list:
        """
        返回指定年份的全部日期的保费

            返回的是从数据库中提出的一个日保费列表

            参数：
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    默认值为本年年份
            返回值：
                list
        """

        sql_str = f"ATTACH DATABASE 'Data\\掌上宝APP出单统计.db' AS [掌上宝APP出单统计]"
        self.cur.execute(sql_str)

        sql_str = f"SELECT \
            [APP签单件数].[中心支公司简称], \
            [APP签单件数].[机构简称], \
            [APP签单件数].[业务员], \
            [APP签单件数], \
            [APP签单件数].[保费], \
            [总签单件数].[总签单件数], \
            [APP签单件数] * 1.0 / [总签单件数] AS [APP出单占比] \
            FROM   (SELECT \
                [机构].[中心支公司简称], \
                [机构].[机构简称], \
                [业务员], \
                COUNT ([保单号]) AS [APP签单件数], \
                SUM ([签单保费/批改保费]) AS [保费] \
            FROM   [掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            WHERE  [终端来源] = '0106移动展业(App)' \
            GROUP  BY [业务员]) AS [APP签单件数] \
            JOIN   (SELECT \
                [机构].[中心支公司简称], \
                [机构].[机构简称], \
                [业务员], \
                COUNT ([保单号]) AS [总签单件数], \
                SUM ([签单保费/批改保费]) AS [保费] \
            FROM   [掌上宝APP出单统计] \
                JOIN [日期] ON [掌上宝APP出单统计].[投保确认日期] = [日期].[投保确认日期] \
                JOIN [机构] ON [掌上宝APP出单统计].[机构] = [机构].[机构] \
            GROUP  BY [业务员]) AS [总签单件数] \
            ON [APP签单件数].[业务员] = [总签单件数].[业务员] \
            ORDER  BY [APP签单件数] DESC;"

        self.cur.execute(sql_str)
        value = self.cur.fetchall()

        sql_str = f"DETACH DATABASE [掌上宝APP出单统计]"
        self.cur.execute(sql_str)

        return value

    @lru_cache(maxsize=32)
    def day_premium(self, year: int = None) -> dict:
        """
        返回指定年份的全部日期的保费

            返回是进行二次匹配的日保费字典，处理了闰年的问题

            参数：
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    默认值为本年年份
            返回值：
                dict
        """

        # 默认值为本年年份
        if year is None:
            year = self.d.year

        # 2020年为闰年，可获取到2月29日的日期
        sql_str = f"SELECT [日期].[日期] \
            FROM [日期] \
            WHERE [日期].[年份] = 2020 \
            ORDER BY [日期].[日期]"

        self.cur.execute(sql_str)

        data_list = dict()

        for keys in self.cur.fetchall():
            for key in keys:
                data_list[key] = 0

        premium_list = self.get_day_premium(year=year)
        for key, value in premium_list:
            if value is not None:
                data_list[key] = value / 10000

        return data_list

    def day_yoy(self, year: int = None) -> dict:
        """
        返回指定年份的日保费同比增长率

            参数：
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的数据
                    默认值为本年年份
            返回值：
                dict
        """

        if year is None:
            year = self.d.year

        first_data = self.day_premium(year)
        last_data = self.day_premium(year - 1)
        data_list = dict()

        for key in list(first_data):
            if last_data[key] == 0 or first_data[key] is None:
                data_list[key] = 0
            else:
                data_list[key] = first_data[key] / last_data[key] - 1

        return data_list

    def day_sum(self, year: int = None) -> dict:
        """
        返回指定年份的日累计保费

            参数：
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的数据
                    默认值为本年年份
            返回值：
                dict
        """

        if year is None:
            year = self.d.year

        premium_data = self.day_premium(year)
        data_list = dict()
        sum_permium = 0

        for key in list(premium_data):
            if premium_data[key] is not None:
                sum_permium += premium_data[key]
                data_list[key] = sum_permium
            else:
                data_list[key] = None

        return data_list

    def day_sum_yoy(self, year: int = None) -> dict:
        """
        返回指定年份的日累计保费的同比增长率

            参数：
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的数据
                    默认值为本年年份
            返回值：
                dict
        """

        if year is None:
            year = self.d.year

        first_data = self.day_sum(year)
        last_data = self.day_sum(year - 1)
        data_list = dict()

        for key in list(first_data):
            if last_data[key] == 0 or first_data[key] is None:
                data_list[key] = 0
            else:
                data_list[key] = first_data[key] / last_data[key] - 1

        return data_list

    def day_task_progress_rate(self):
        """
        返回当前年份的日累计保费的计划任务达成率

            参数：
                无
            返回值：
                dict
        """

        day_sum = self.day_sum()
        data_list = dict()

        for key in list(day_sum):
            if day_sum[key] is not None:
                data_list[key] = day_sum[key] / self.task
            else:
                data_list[key] = 0

        return data_list

    def day_time_progress_rate(self):
        """
        返回当前年份的日累计保费的计划任务达成率

            参数：
                无
            返回值：
                dict
        """

        task_progress = self.day_task_progress_rate()
        data_list = dict()

        for key in list(task_progress):
            if task_progress[key] is not None:
                data_list[key] = task_progress[key] / self.time_progress(
                    month=key[0:2], day=key[3:]
                )
            else:
                data_list[key] = 0

        return data_list


if __name__ == "__main__":
    days = Stats_Day("分公司", "整体")

    value = days.day_premium(2019)
    print(value)
