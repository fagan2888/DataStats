import logging
from functools import lru_cache


from code.tong_ji_base import Tong_Ji_Base


logging.disable(logging.NOTSET)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Tong_Ji_Zhou(Tong_Ji_Base):
    """
    周数据统计类

        关于周数据的各项统计功能
    """

    def year_zhou(func):
        """
        """

        def calc_year_zhou(
            self,
            ny: int = None,
            nw: int = None,
            year: int = None,
            weeknum: int = None,
            tong: bool = True,
        ):

            if ny is not None and year is None:
                year = self.d.nian - ny
            elif year is None:
                year = self.d.nian

            if nw is not None and weeknum is None:
                # 当倒数周数跨年时进行特殊计算
                if year == self.d.nian and nw >= self.d.zhou:
                    weeknum = 53 - (nw - self.d.zhou)
                    year -= 1
                else:
                    weeknum = self.d.zhou - nw
            elif weeknum is None:
                weeknum = self.d.zhou

            if weeknum == 0:
                year -= 1
                weeknum = 53

            return func(self, year=year, weeknum=weeknum, tong=tong)

        return calc_year_zhou

    @year_zhou
    @lru_cache(maxsize=32)
    def zhou_bao_fei(
        self,
        ny: int = None,
        nw: int = None,
        year: int = None,
        weeknum: int = None,
        tong: bool = True,
    ):
        """
        返回指定周的累计保费

            默认返回本周的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nw:
                    int，设置需要返回倒数多少周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                weeknum：
                    int，设置需要返回那一周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                    默认值为本年本周
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整周累计保费
        """

        if tong is True:
            weekday = self.d.xin_qi
        else:
            weekday = 7

        logging.debug(f"{weekday=}")

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year)} \
            {self.xian_zhong_join(year)} \
            JOIN [日期] \
            ON [{year}年].[投保确认日期] = [日期].[投保确认日期] \
            WHERE [日期].[周数] = '{weeknum}' \
            AND [日期].[星期] <= '{weekday}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(year)}"

        logging.debug(f"{str_sql=}")

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            logging.debug(f"{value=}")
            if value is None:
                return 0
            else:
                return float(value) / 10000

    @year_zhou
    def zhou_tong_bi(
        self,
        ny: int = None,
        nw: int = None,
        year: int = None,
        weeknum: int = None,
        tong: bool = True,
    ):
        """
        返回指定周的保费同比增长率

            默认返回本周的保费同比增长率

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nw:
                    int，设置需要返回倒数多少周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                weeknum：
                    int，设置需要返回那一周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                    默认值为本年本周
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整周累计保费
        """

        if (
            self.zhou_bao_fei(year=year, weeknum=weeknum, tong=tong) == 0
            or self.zhou_bao_fei(year=year - 1, weeknum=weeknum, tong=tong)
            == 0
        ):
            return "——"
        else:
            value = (
                self.zhou_bao_fei(year=year, weeknum=weeknum, tong=tong)
                / self.zhou_bao_fei(year=year - 1, weeknum=weeknum, tong=tong)
                - 1
            )
            return value

    @year_zhou
    def zhou_huan_bi(
        self,
        ny: int = None,
        nw: int = None,
        year: int = None,
        weeknum: int = None,
        tong: bool = True,
    ):
        """
        返回指定周的保费环比增长率

            默认返回本周的保费环比增长率

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                nw:
                    int，设置需要返回倒数多少周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                weeknum：
                    int，设置需要返回那一周的保费
                    nw参数与weeknum参数不应该同时使用，当设置了weeknum参数后nw参数将失效
                    默认值为本年本周
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计整周累计保费
        """

        if (
            self.zhou_bao_fei(year=year, weeknum=weeknum, tong=tong) == 0
            or self.zhou_bao_fei(year=year, weeknum=weeknum - 1, tong=tong)
            == 0
        ):
            return "——"
        else:
            value = (
                self.zhou_bao_fei(year=year, weeknum=weeknum, tong=tong)
                / self.zhou_bao_fei(year=year, weeknum=weeknum - 1, tong=tong)
                - 1
            )
            return value
