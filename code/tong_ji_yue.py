from functools import lru_cache

from code.tong_ji_base import Tong_Ji_Base


class Tong_Ji_Yue(Tong_Ji_Base):
    """
    月度数据统计类

        关于月度数据的各项统计功能

        属性：
            yue_bao_fei:
                返回指定月份的累计保费
            yue_tong_bi:
                返回指定月份的保费同比增长率
            yue_huan_bi:
                返回指定月份的保费环比增长率
    """

    def year_month(func):
        """
        用于判断年、月的装饰器

            该装饰器通过装饰方法所带的参数判断正确的年份和月份并返回结果
        """

        def calc_year_month(
            self,
            ny: int = None,
            nm: int = None,
            year: int = None,
            month: int = None,
            tong: bool = True
        ):

            # 只有当year为空时ny参数才生效
            if ny is not None and year is None:
                year = self.d.nian - ny
            elif year is None:
                year = self.d.nian

            # 只有当month为空时nm参数才生效
            if nm is not None and month is None:
                # 当倒数月份跨年时进行特殊计算
                if year == self.d.nian and nm >= self.d.yue:
                    month = 12 - (nm - self.d.yue)
                    year -= 1
                else:
                    month = self.d.yue - nm
            elif month is None:
                month = self.d.yue

            if month == 0:
                year -= 1
                month = 12

            return func(self, year=year, month=month, tong=tong)

        return calc_year_month

    @year_month
    @lru_cache(maxsize=32)
    def yue_bao_fei(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回指定月份的累计保费

            默认返回当前月份的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nm:
                    int，设置需要返回倒数多少月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                month：
                    int，设置需要返回那一月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

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
            {self.ji_gou_where()} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @year_month
    def yue_tong_bi(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
        tong: bool = True,
    ):
        """
        返回指定月份的保费同比增长率

            默认返回当前月份的同比增长率

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nm:
                    int，设置需要返回倒数多少月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                month：
                    int，设置需要返回那一月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.yue_bao_fei(year=year, month=month, tong=tong) == 0
            or self.yue_bao_fei(year=year - 1, month=month, tong=tong) == 0
        ):
            return "——"

        value = (
            self.yue_bao_fei(year=year, month=month, tong=tong)
            / self.yue_bao_fei(year=year - 1, month=month, tong=tong)
            - 1
        )
        return value

    @year_month
    def yue_huan_bi(
        self,
        ny: int = None,
        nm: int = None,
        year: int = None,
        month: int = None,
        tong: bool = None,
    ) -> float:
        """
        返回指定月份的保费环比增长率

            默认返回当前月份的环比增长率

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nm:
                    int，设置需要返回倒数多少月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                month：
                    int，设置需要返回那一月的保费
                    nm参数与month参数不应该同时使用，当设置了month参数后nm参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.yue_bao_fei(year=year, month=month, tong=tong) == 0
            or self.yue_bao_fei(year=year, month=month - 1, tong=tong) == 0
        ):
            return "——"

        value = (
            self.yue_bao_fei(year=year, month=month, tong=tong)
            / self.yue_bao_fei(year=year, month=month - 1, tong=tong)
            - 1
        )
        return value
