from functools import lru_cache

from code.tong_ji_base import Tong_Ji_Base


class Tong_Ji_Nian(Tong_Ji_Base):
    """
    年度数据统计类

        关于年度数据的各项统计功能

        属性：
            ren_wu_jin_du:
                返回当前年度的年度保费的计划任务达成率
            shi_jian_da_cheng:
                返回当前年度的年度保费的时间进度达成率
            nian_bao_fei:
                返回机构的制定年度的年度累计保费
            nian_tong_bi:
                返回制定年份的保费同比增长率
    """

    def ren_wu_jin_du(self) -> float:
        """
        返回当前年度的年度保费的计划任务达成率

            如果上年度保费为零或任务为零则返回"——"

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

    def shi_jian_da_cheng(self) -> float:
        """
        返回当前年度的年度保费的时间进度达成率

            如果任务进度为零则返回 "——"

            参数：
                无
            返回值：
                float，时间进度达成率
        """

        if self.ren_wu_jin_du() == "——":
            return "——"
        else:
            return self.ren_wu_jin_du() / self.shi_jian_jin_du()

    @lru_cache(maxsize=32)
    def nian_bao_fei(
        self, ny: int = None, year: int = None, tong: bool = True
    ) -> float:
        """
        返回机构的制定年度的年度累计保费

            默认返回本年的保费

            参数：
                ny:
                    int，设置需要返回倒数多少年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                year：
                    int，一个4位数的年份值，用于计算需要返回那一年的保费
                    ny参数与year参数不应该同时使用，当设置了year参数后ny参数将失效
                    默认值为本年年份
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计全年累计保费
            返回值：
                float, 机构的年累计保费
        """

        if ny is not None and year is None:
            year = self.d.nian - ny
        elif year is None:
            year = self.d.nian

        if tong is True:
            ri_qi = self.d.long_ri_qi(year=year)
        else:
            ri_qi = self.d.long_ri_qi(year=year, month=12, day=31)

        str_sql = f"SELECT SUM([签单保费/批改保费]) \
            FROM [{year}年] \
            {self.ji_gou_join(year=year)} \
            {self.xian_zhong_join(year=year)} \
            WHERE [投保确认日期] <= '{ri_qi}' \
            {self.ji_gou_where()} \
            {self.xian_zhong_where(year=year)}"

        self.cur.execute(str_sql)
        for value in self.cur.fetchone():
            if value is None:
                return 0
            else:
                return float(value) / 10000

    def nian_tong_bi(
        self,
        year: int = None,
        first_year: int = None,
        last_year: int = None,
        tong: bool = True,
    ) -> float:
        """
        返回制定年份的保费同比增长率

            参数：
                year：
                    int，一个4位数的年份值
                    系统将计算指定年份与上一年年度的保费同比
                    year参数，first_year与last_year参数，应该只选择其中的一种组合
                    如果同时使用了year参数和first_year，last_year参数，则year参数将被忽视
                    默认值为本年年份
                first_year:
                    int, 一个4位数的年份值
                    用于计算指定年份与last_year参数指定年份的保费同比
                    first_year参数应该与last_year参数共同使用，单独使用无效
                    公式为：first_year / last_year - 1
                last_year:
                    int, 一个4位数的年份值
                    用于计算指定年份与first_year参数指定年份的保费同比
                    first_year参数应该与last_year参数共同使用，单独使用无效
                    公式为：first_year / last_year - 1
                tong：
                    bool，默认值为True，用于判断是否统计同期保费，输入False则统计全年累计保费
            返回值：
                float, 机构的当年保费同比增长率
        """

        if year is None and (first_year is None or last_year is None):
            first_year = self.d.nian
            last_year = self.d.nian - 1
        elif year is not None and (first_year is None or last_year is None):
            first_year = year
            last_year = year - 1

        if (
            self.nian_bao_fei(year=first_year, tong=tong) == 0
            or self.nian_bao_fei(year=last_year, tong=tong) == 0
        ):
            return "——"
        else:
            value = (
                self.nian_bao_fei(year=first_year, tong=tong)
                / self.nian_bao_fei(year=last_year, tong=tong)
                - 1
            )
            return value
