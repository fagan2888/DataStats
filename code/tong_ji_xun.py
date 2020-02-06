from functools import lru_cache

from code.tong_ji_base import Tong_Ji_Base


class Tong_Ji_Xun(Tong_Ji_Base):
    """
    旬数据统计类

        关于旬数据的各项统计功能

        属性：
            xun_bao_fei:
                返回指定旬的累计保费
            xun_tong_bi:
                返回指定旬的保费同比增长率
            xun_huan_bi:
                返回指定旬的保费环比增长率
    """

    def year_month_ten_days(func):
        """
        用于判断年、月、旬的装饰器

            该装饰器通过装饰方法所带的参数判断正确的年份、月份和旬并返回结果
        """

        def calc_year_month_ten_days(
            self,
            ny: int = None,
            nm: int = None,
            nt: int = None,
            year: int = None,
            month: int = None,
            ten_days: int = None,
            tong: bool = True,
        ):

            # 只有当year为空时ny参数才生效
            if ny is not None and year is None:
                year = self.d.nian - ny
            elif year is None:
                year = self.d.nian

            # 只有当month为空时nm参数才生效
            if nm is not None and month is None:
                # 当倒数旬跨年时进行特殊计算
                if year == self.d.nian and nm >= self.d.yue:
                    month = 12 - (nm - self.d.yue)
                    year -= 1
                else:
                    month = self.d.yue - nm
            elif month is None:
                month = self.d.yue

            if nt is not None and ten_days is None:
                if month == self.d.yue and ny >= self.d.xun:
                    ten_days = 3 - (nt - self.d.xun)
                    month -= 1
                else:
                    ten_days = self.d.xun - ny
            elif ten_days is None:
                ten_days = self.d.xun

            if ten_days == 0:
                month -= 1
                ten_days = 3

            if month == 0:
                year -= 1
                month = 12

            return func(self, year=year, month=month, ten_days=ten_days, tong=tong)

        return calc_year_month_ten_days

    @year_month_ten_days
    @lru_cache(maxsize=32)
    def xun_bao_fei(
        self,
        ny: int = None,
        nm: int = None,
        nt: int = None,
        year: int = None,
        month: int = None,
        ten_days: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回指定旬的累计保费

            默认返回当前旬的保费

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
                nt:
                    int，设置需要返回倒数多少旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                ten_days:
                    int, 设置需要返回那一旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                    默认值为本年本旬
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        # 设置旬列表，上旬与中旬在末尾追加一天以匹配下旬的31日
        ten_days_list = [
            (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 10),
            (11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 20),
            (21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31),
        ]

        for days in ten_days_list:
            if self.d.ri in days:
                index = days.index(self.d.ri)

        if tong is True:
            day = ten_days_list[ten_days - 1][index]
        else:
            day = 31

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[月份] = '{month}' \
            AND [日期].[旬] = '{ten_days}' \
            AND [日期].[日数] <= '{day}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @year_month_ten_days
    def xun_tong_bi(
        self,
        ny: int = None,
        nm: int = None,
        nt: int = None,
        year: int = None,
        month: int = None,
        ten_days: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回指定旬的累计保费

            默认返回当前旬的保费

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
                nt:
                    int，设置需要返回倒数多少旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                ten_days:
                    int, 设置需要返回那一旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                    默认值为本年本旬
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.xun_bao_fei(year=year, month=month, ten_days=ten_days, tong=tong) == 0
            or self.xun_bao_fei(
                year=year - 1, month=month, ten_days=ten_days, tong=tong
            )
            == 0
        ):
            return "——"

        value = (
            self.xun_bao_fei(year=year, month=month, ten_days=ten_days, tong=tong)
            / self.xun_bao_fei(year=year - 1, month=month, ten_days=ten_days, tong=tong)
            - 1
        )
        return value

    @year_month_ten_days
    def xun_huan_bi(
        self,
        ny: int = None,
        nm: int = None,
        nt: int = None,
        year: int = None,
        month: int = None,
        ten_days: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回指定旬的累计保费

            默认返回当前旬的保费

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
                nt:
                    int，设置需要返回倒数多少旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                ten_days:
                    int, 设置需要返回那一旬的保费
                    nt参数与tan_days参数不应该同时使用，当设置了tan_days参数后nt参数将失效
                    默认值为本年本旬
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.xun_bao_fei(year=year, month=month, ten_days=ten_days, tong=tong) == 0
            or self.xun_bao_fei(
                year=year, month=month, ten_days=ten_days - 1, tong=tong
            )
            == 0
        ):
            return "——"

        value = (
            self.xun_bao_fei(year=year, month=month, ten_days=ten_days, tong=tong)
            / self.xun_bao_fei(year=year, month=month, ten_days=ten_days - 1, tong=tong)
            - 1
        )
        return value
