import logging
import xlsxwriter
import sqlite3
from datetime import datetime

# from io import StringIO

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji
from style import Style
from date import IDate
from stats_0621 import Stats_0621

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Excel_Write_0621:
    """
    Excel表格数据写入的基类

        负责初始化写入数据时所需的各项数据的公共变量
    """

    def __init__(self, wb: xlsxwriter.Workbook):
        """
        初始化基础变量

            参数：
                wb:
                    xlsxwriter.Workbook, 需要写入数据的工作簿
        """

        # 工作薄对象
        self._wb = wb

        # 工作表对象
        self._ws: xlsxwriter.Workbook.worksheet_class = None

        # 数据库连接对象
        self._set_conn()

        # 数据库操作的游标对象
        self._set_cur()

        # 数据库内数据的最大日期
        self._idate = IDate(year=2020)

        self._style = Style(wb=self.wb)

        # 操作的工作表名称
        self._table_name: str = None

    @property
    def table_name(self):
        """
        返回目前正在操作的工作表名称
        """
        return self._table_name

    def set_table_name(self, table_name: str):
        """
        设置需要操作的工作表名称

            一旦设置了工作表，实际上也就等于新建立的一个工作表，后续的操作将在新工作表中完成

            参数：
                table_name:
                    str, 工作表名称
        """

        self._table_name = table_name
        self.add_worksheet(self._table_name)

    def _set_conn(self):
        """
        设置数据库链接对象

            参数：
                无
            返回值：
                无
        """
        self._conn = sqlite3.connect(r"Data\data.db")

    @property
    def conn(self) -> sqlite3.Connection:
        """
        返回数据库连接对象

            参数：
                无
            返回值：
                sqlite3.Connection
        """
        return self._conn

    def _set_cur(self):
        """
        设置数据库的游标对象

            参数：
                无
            返回值：
                无
        """
        self._cur = self.conn.cursor()

    @property
    def cur(self) -> sqlite3.Cursor:
        """
        返回数据库操作的游标对象

            参数：
                无
            返回值：
                sqlite3.Cursor
        """
        return self._cur

    @property
    def ws(self) -> xlsxwriter.Workbook.worksheet_class:
        """
        返回需要操作的工作表对象
        """
        return self._ws

    def add_worksheet(self, table_name: str):
        """
        添加一个工作表对象

            一旦添加了一个工作表，则后线的操作将在新工作表中完成

            参数：
                table_name:
                    str, 工作表名称
        """
        self._ws = self.wb.add_worksheet(name=table_name)
        return self._ws

    @property
    def date(self) -> IDate:
        """
        返回数据库中当前的时间对象
        """
        return self._idate

    @property
    def wb(self) -> xlsxwriter.Workbook:
        """
        返回需要操作的工作簿对象
        """
        return self._wb

    @property
    def style(self) -> Style:
        """
        返回Excel表格所需的单元格样式对象
        """
        return self._style

    def write_center_branch(self, risk=None):
        """
        写入中心支公司数据
        """

        nrow = 0
        ncol = 0

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 12,
            data="驾意险业务统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 12,
            data=f"数据统计范围：2020-1-1 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        company = ["中心支公司", "昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "昭通", "合计"]
        self.ws.write(nrow, ncol, "", self.style.string)
        self.ws.write_column(nrow + 1, ncol, company, self.style.header)
        ncol += 1

        week = self.date.weeknum
        month = self.date.month
        cycle = [(week, None), (None, month), (None, None)]

        for w, m in cycle:
            values = risk.get_center_branch(week=w, month=m)
            nrow = 2

            if w is None and m is None:
                date = "年度累计"
            elif w is not None:
                date = f"第{w}周"
            elif m is not None:
                date = f"{m}月"

            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + 3,
                data=f"{date}统计数据",
                cell_format=self.style.header,
            )
            nrow += 1

            header = ["驾意险件数", "驾意险保费", "车辆数", "驾意险联动率"]
            self.ws.write_row(nrow, ncol, header, self.style.header)
            nrow += 1

            app_sum = 0
            sum_sum = 0
            car_sum = 0

            i = 1
            for value in values:
                if value[1] is None:
                    self.ws.write(nrow, ncol, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol, value[1], self.style.string)
                if value[2] is None:
                    self.ws.write(nrow, ncol + 1, 0, self.style.number)
                else:
                    self.ws.write(nrow, ncol + 1, value[2], self.style.number)
                if value[3] is None:
                    self.ws.write(nrow, ncol + 2, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol + 2, value[3], self.style.string)
                if value[4] is None:
                    self.ws.write(nrow, ncol + 3, 0, self.style.percent)
                else:
                    self.ws.write(nrow, ncol + 3, value[4], self.style.percent)

                if value[1] is not None:
                    app_sum += value[1]
                if value[2] is not None:
                    sum_sum += value[2]
                if value[3] is not None:
                    car_sum += value[3]
                nrow += 1
                i += 1

            # self.ws.write(nrow, ncol, "合计", self.style.string)
            self.ws.write(nrow, ncol, app_sum, self.style.string)
            self.ws.write(nrow, ncol + 1, sum_sum, self.style.number)
            self.ws.write(nrow, ncol + 2, car_sum, self.style.string)
            if car_sum == 0 or car_sum
            self.ws.write(nrow, ncol + 3, app_sum / car_sum, self.style.percent)
            ncol += 4
        nrow += 1

        ncol = 0
        self.ws.set_column(first_col=ncol, last_col=ncol + 16, width=14)

        logging.info(f"中心支公司数据表写入完成")

        return nrow

    def write_center_branch_cycle(self, risk=None, type=None):
        """
        写入中心支公司数据
        """

        nrow = 0
        ncol = 0

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 4,
            data="驾意险业务统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 4,
            data=f"数据统计范围：2020-1-1 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        company = ["中心支公司", "昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "昭通", "合计"]

        if type == "week":
            cycle = range(1, 54)
        elif type == "month":
            cycle = range(1, 13)

        for w in cycle:
            if type == "week":
                date = f"第{w}周"
                values = risk.get_center_branch(week=w)
            elif type == "month":
                date = f"第{w}月"
                values = risk.get_center_branch(month=w)
            ncol = 0

            self.ws.write(nrow, ncol, "", self.style.string)
            self.ws.write_column(nrow + 1, ncol, company, self.style.header)
            ncol += 1

            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + 3,
                data=f"{date}统计数据",
                cell_format=self.style.header,
            )
            nrow += 1

            header = ["驾意险件数", "驾意险保费", "车辆数", "驾意险联动率"]
            self.ws.write_row(nrow, ncol, header, self.style.header)
            nrow += 1

            app_sum = 0
            sum_sum = 0
            car_sum = 0

            i = 1
            for value in values:
                if value[1] is None:
                    self.ws.write(nrow, ncol, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol, value[1], self.style.string)
                if value[2] is None:
                    self.ws.write(nrow, ncol + 1, 0, self.style.number)
                else:
                    self.ws.write(nrow, ncol + 1, value[2], self.style.number)
                if value[3] is None:
                    self.ws.write(nrow, ncol + 2, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol + 2, value[3], self.style.string)
                if value[4] is None:
                    self.ws.write(nrow, ncol + 3, 0, self.style.percent)
                else:
                    self.ws.write(nrow, ncol + 3, value[4], self.style.percent)

                if value[1] is not None:
                    app_sum += value[1]
                if value[2] is not None:
                    sum_sum += value[2]
                if value[3] is not None:
                    car_sum += value[3]
                nrow += 1
                i += 1

            self.ws.write(nrow, ncol, app_sum, self.style.string)
            self.ws.write(nrow, ncol + 1, sum_sum, self.style.number)
            self.ws.write(nrow, ncol + 2, car_sum, self.style.string)
            if app_sum == 0:
                rate = 0
            else:
                rate = app_sum / car_sum
            self.ws.write(nrow, ncol + 3, rate, self.style.percent)
            nrow += 2

            logging.info(f"{date}数据写入完成")
        nrow += 1

        ncol = 0
        self.ws.set_column(first_col=ncol, last_col=ncol + 4, width=14)

        logging.info(f"中心支公司周数据表写入完成")

        return nrow

    def write_company(self, risk=None, nrow=None):
        """
        写入中心支公司数据
        """
        nrow += 1
        ncol = 0

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 12,
            data="驾意险业务统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 12,
            data=f"数据统计范围：2020-1-1 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        company = [
            "机构",
            "分公司本部",
            "百大国际",
            "春之城",
            "香榭丽园",
            "东川",
            "宜良",
            "安宁",
            "春怡雅苑",
            "曲靖中支本部",
            "陆良",
            "师宗",
            "宣威",
            "会泽",
            "沾益",
            "罗平",
            "保山中支本部",
            "施甸",
            "腾冲",
            "昭通",
            "文山中支一部",
            "文山中支二部",
            "砚山",
            "麻栗坡",
            "马关",
            "丘北",
            "广南",
            "富宁",
            "版纳中支本部",
            "勐海",
            "勐腊",
            "大理中支本部",
            "漾濞",
            "祥云",
            "宾川",
            "弥渡",
            "云龙",
            "洱源",
            "怒江中支本部",
            "兰坪",
            "合计",
        ]
        self.ws.write(nrow, ncol, "", self.style.string)
        self.ws.write_column(nrow + 1, ncol, company, self.style.header)
        ncol += 1

        week = self.date.weeknum
        month = self.date.month
        cycle = [(week, None), (None, month), (None, None)]

        for w, m in cycle:
            values = risk.get_company(week=w, month=m)
            tnrow = nrow

            if w is None and m is None:
                date = "年度累计"
            elif w is not None:
                date = f"第{w}周"
            elif m is not None:
                date = f"{m}月"

            self.ws.merge_range(
                first_row=tnrow,
                first_col=ncol,
                last_row=tnrow,
                last_col=ncol + 3,
                data=f"{date}统计数据",
                cell_format=self.style.header,
            )
            tnrow += 1

            header = ["驾意险件数", "驾意险保费", "车辆数", "驾意险联动率"]
            self.ws.write_row(tnrow, ncol, header, self.style.header)
            tnrow += 1

            app_sum = 0
            sum_sum = 0
            car_sum = 0

            i = 1
            for value in values:
                if value[1] is None:
                    self.ws.write(tnrow, ncol, 0, self.style.string)
                else:
                    self.ws.write(tnrow, ncol, value[1], self.style.string)
                if value[2] is None:
                    self.ws.write(tnrow, ncol + 1, 0, self.style.number)
                else:
                    self.ws.write(tnrow, ncol + 1, value[2], self.style.number)
                if value[3] is None:
                    self.ws.write(tnrow, ncol + 2, 0, self.style.string)
                else:
                    self.ws.write(tnrow, ncol + 2, value[3], self.style.string)
                if value[4] is None:
                    self.ws.write(tnrow, ncol + 3, 0, self.style.percent)
                else:
                    self.ws.write(tnrow, ncol + 3, value[4], self.style.percent)

                if value[1] is not None:
                    app_sum += value[1]
                if value[2] is not None:
                    sum_sum += value[2]
                if value[3] is not None:
                    car_sum += value[3]
                tnrow += 1
                i += 1

            # self.ws.write(nrow, ncol, "合计", self.style.string)
            self.ws.write(tnrow, ncol, app_sum, self.style.string)
            self.ws.write(tnrow, ncol + 1, sum_sum, self.style.number)
            self.ws.write(tnrow, ncol + 2, car_sum, self.style.string)
            self.ws.write(tnrow, ncol + 3, app_sum / car_sum, self.style.percent)
            ncol += 4
        nrow += tnrow + 1

        ncol = 0
        self.ws.set_column(first_col=ncol, last_col=ncol + 16, width=14)

        logging.info(f"机构数据表写入完成")

        return nrow

    def write_company_cycle(self, risk=None, type=None):
        """
        写入中心支公司数据
        """
        nrow = 0
        ncol = 0

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 4,
            data="驾意险业务统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 4,
            data=f"数据统计范围：2020-1-1 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        company = [
            "机构",
            "分公司本部",
            "百大国际",
            "春之城",
            "香榭丽园",
            "东川",
            "宜良",
            "安宁",
            "春怡雅苑",
            "曲靖中支本部",
            "陆良",
            "师宗",
            "宣威",
            "会泽",
            "沾益",
            "罗平",
            "保山中支本部",
            "施甸",
            "腾冲",
            "昭通",
            "文山中支一部",
            "文山中支二部",
            "砚山",
            "麻栗坡",
            "马关",
            "丘北",
            "广南",
            "富宁",
            "版纳中支本部",
            "勐海",
            "勐腊",
            "大理中支本部",
            "漾濞",
            "祥云",
            "宾川",
            "弥渡",
            "云龙",
            "洱源",
            "怒江中支本部",
            "兰坪",
            "合计",
        ]

        if type == "week":
            cycle = range(1, 54)
        elif type == "month":
            cycle = range(1, 13)

        for w in cycle:
            if type == "week":
                date = f"第{w}周"
                values = risk.get_company(week=w)
            elif type == "month":
                date = f"第{w}月"
                values = risk.get_company(month=w)
            ncol = 0

            self.ws.write(nrow, ncol, "", self.style.string)
            self.ws.write_column(nrow + 1, ncol, company, self.style.header)
            ncol += 1

            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + 3,
                data=f"{date}统计数据",
                cell_format=self.style.header,
            )
            nrow += 1

            header = ["驾意险件数", "驾意险保费", "车辆数", "驾意险联动率"]
            self.ws.write_row(nrow, ncol, header, self.style.header)
            nrow += 1

            app_sum = 0
            sum_sum = 0
            car_sum = 0

            i = 1
            for value in values:
                if value[1] is None:
                    self.ws.write(nrow, ncol, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol, value[1], self.style.string)
                if value[2] is None:
                    self.ws.write(nrow, ncol + 1, 0, self.style.number)
                else:
                    self.ws.write(nrow, ncol + 1, value[2], self.style.number)
                if value[3] is None:
                    self.ws.write(nrow, ncol + 2, 0, self.style.string)
                else:
                    self.ws.write(nrow, ncol + 2, value[3], self.style.string)
                if value[4] is None:
                    self.ws.write(nrow, ncol + 3, 0, self.style.percent)
                else:
                    self.ws.write(nrow, ncol + 3, value[4], self.style.percent)

                if value[1] is not None:
                    app_sum += value[1]
                if value[2] is not None:
                    sum_sum += value[2]
                if value[3] is not None:
                    car_sum += value[3]
                nrow += 1
                i += 1

            self.ws.write(nrow, ncol, app_sum, self.style.string)
            self.ws.write(nrow, ncol + 1, sum_sum, self.style.number)
            self.ws.write(nrow, ncol + 2, car_sum, self.style.string)
            if app_sum == 0:
                rate = 0
            else:
                rate = app_sum / car_sum
            self.ws.write(nrow, ncol + 3, rate, self.style.percent)
            nrow += 2
            logging.info(f"{date}数据写入完成")
        nrow += 1

        ncol = 0
        self.ws.set_column(first_col=ncol, last_col=ncol + 16, width=14)

        logging.info(f"机构数据表写入完成")

        return nrow

    def write_sum(self):
        """
        累计数据控制函数
        """
        self.set_table_name("汇总数据统计表")
        risk = Stats_0621()
        risk.attach_db()

        nrow = self.write_center_branch(risk=risk)
        nrow = self.write_company(risk=risk, nrow=nrow)

        self.set_table_name("三级机构周数据统计表")
        self.write_center_branch_cycle(risk=risk, type="week")
        self.set_table_name("三级机构月数据统计表")
        self.write_center_branch_cycle(risk=risk, type="month")
        self.set_table_name("四级机构周数据统计表")
        self.write_company_cycle(risk=risk, type="week")
        self.set_table_name("四级机构月数据统计表")
        self.write_company_cycle(risk=risk, type="month")

        risk.detach_db()


if __name__ == "__main__":
    a = datetime.now()
    wb = xlsxwriter.Workbook(r"驾意险业务统计表.xlsx")

    risk = Excel_Write_0621(wb=wb)
    risk.write_sum()

    wb.close()
    b = datetime.now()
    print(f"date {b-a=}")
