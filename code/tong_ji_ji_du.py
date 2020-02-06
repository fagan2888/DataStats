from functools import lru_cache

from code.tong_ji_base import Tong_Ji_Base


class Tong_Ji_Ji_Du(Tong_Ji_Base):
    """
    季度数据统计类

        关于季度数据的各项统计功能

        属性：
            ji_bao_fei:
                返回指定季度的累计保费
            ji_tong_bi:
                返回指定季度的保费同比增长率
            ji_huan_bi:
                返回指定季度的保费环比增长率
    """

    def year_quarter(func):
        """
        用于判断年、季度的装饰器

            该装饰器通过装饰方法所带的参数判断正确的年份和季度并返回结果
        """
        def calc_year_quarter(
            self,
            ny: int = None,
            nq: int = None,
            year: int = None,
            quarter: int = None,
            tong: bool = True,
        ):

            # 只有当year为空时ny参数才生效
            if ny is not None and year is None:
                year = self.d.nian - ny
            elif year is None:
                year = self.d.nian

            # 只有当quarter为空时nq参数才生效
            if nq is not None and quarter is None:
                # 当倒数季度跨年时进行特殊计算
                if year == self.d.nian and nq >= self.d.ji_du:
                    quarter = 4 - (nq - self.d.ji_du)
                    year -= 1
                else:
                    quarter = self.d.ji_du - nq
            elif quarter is None:
                quarter = self.d.ji_du

            if quarter == 0:
                year -= 1
                quarter = 4

            return func(self, year=year, quarter=quarter, tong=tong)

        return calc_year_quarter

    @year_quarter
    @lru_cache(maxsize=32)
    def ji_bao_fei(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回指定季度的累计保费

            默认返回当前季度的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nq:
                    int，设置需要返回倒数多少月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                quarter：
                    int，设置需要返回那一月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        month_data = [(1, 4, 7, 10), (2, 5, 8, 11), (3, 6, 9, 12)]
        for ji_du in month_data:
            if self.d.yue in ji_du:
                data = ji_du

        if tong is True:
            short_date = f"{data[quarter-1]:02d}-{self.d.ri:02d}"
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
            {self.ji_gou_where()} \
            {self.xian_zhong_where(year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @year_quarter
    def ji_tong_bi(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
        tong: bool = None,
    ):
        """
        返回指定季度的累计保费

            默认返回当前季度的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nq:
                    int，设置需要返回倒数多少月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                quarter：
                    int，设置需要返回那一月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.ji_bao_fei(year=year, quarter=quarter, tong=tong) == 0
            or self.ji_bao_fei(year=year - 1, quarter=quarter, tong=tong) == 0
        ):
            return "——"
        else:
            value = (
                self.ji_bao_fei(year=year, quarter=quarter, tong=tong)
                / self.ji_bao_fei(year=year - 1, quarter=quarter, tong=tong)
                - 1
            )
            return value

    @year_quarter
    def ji_huan_bi(
        self,
        ny: int = None,
        nq: int = None,
        year: int = None,
        quarter: int = None,
        tong: bool = None,
    ):
        """
        返回指定季度的累计保费

            默认返回当前季度的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nq:
                    int，设置需要返回倒数多少月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                quarter：
                    int，设置需要返回那一月的保费
                    nq参数与quarter参数不应该同时使用，当设置了quarter参数后nq参数将失效
                    默认值为本年本月
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整月累计保费
        """

        if (
            self.ji_bao_fei(year=year, quarter=quarter, tong=tong) == 0
            or self.ji_bao_fei(year=year, quarter=quarter - 1, tong=tong) == 0
        ):
            return "——"
        else:
            value = (
                self.ji_bao_fei(year=year, quarter=quarter, tong=tong)
                / self.ji_bao_fei(year=year, quarter=quarter - 1, tong=tong)
                - 1
            )
            return value
