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
from stats_app import Stats_App

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Excel_Write_App:
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

    def write_salesman(self, dtype, app, week: int = None):
        """
        写入业务员统计表
        """

        nrow = 0
        ncol = 0

        if week is None:
            first_date = "2020-02-17"
            last_date = self.date.short_date()
        else:
            first_date = self.date.week_first_date(week=week)
            last_date = self.date.week_last_date(week=week)

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 7,
            data="掌上宝APP出单数量统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 7,
            data=f"数据统计范围：{first_date} 至 {last_date}",
            cell_format=self.style.explain,
        )
        nrow += 1

        header = ["序号", "中心支公司", "机构", "业务员", "保单件数", "保费", "总保单件数", "APP占比"]

        self.ws.write_row(nrow, ncol, header, self.style.header)
        nrow += 1

        if dtype == "sum":
            values = app.get_salesman()
        elif dtype == "week":
            values = app.get_salesman(week=week)

        i = 1
        for value in values:
            self.ws.write(nrow, ncol, i, self.style.string)
            self.ws.write(nrow, ncol + 1, value[0], self.style.string)
            self.ws.write(nrow, ncol + 2, value[1], self.style.string)
            self.ws.write(nrow, ncol + 3, value[2][9:], self.style.string)
            self.ws.write(nrow, ncol + 4, value[3], self.style.string)
            self.ws.write(nrow, ncol + 5, value[4], self.style.number)
            self.ws.write(nrow, ncol + 6, value[5], self.style.string)
            self.ws.write(nrow, ncol + 7, value[6], self.style.percent)

            nrow += 1
            i += 1

        self.ws.set_column(first_col=ncol, last_col=ncol, width=6)
        self.ws.set_column(first_col=ncol + 1, last_col=ncol + 1, width=12)
        self.ws.set_column(first_col=ncol + 2, last_col=ncol + 2, width=14)
        self.ws.set_column(first_col=ncol + 3, last_col=ncol + 5, width=10)
        self.ws.set_column(first_col=ncol + 6, last_col=ncol + 7, width=12)

        logging.info(f"业务员数据表写入完成")

    def write_terminal(self, dtype, app, week=None):
        """
        写入终端数据
        """

        nrow = 0
        ncol = 10

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 2,
            data="掌上宝APP出单数量统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 2,
            data=f"数据统计范围：2020-2-17 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        header = ["终端来源", "保单件数", "件数占比"]

        self.ws.write_row(nrow, ncol, header, self.style.header)
        nrow += 1

        if dtype == "sum":
            values, value_sum = app.get_terminal()
        elif dtype == "week":
            values, value_sum = app.get_terminal(week=week)

        i = 1
        for value in values:
            self.ws.write(nrow, ncol, value[0], self.style.string)
            self.ws.write(nrow, ncol + 1, value[1], self.style.string)
            self.ws.write(nrow, ncol + 2, value[1] / value_sum, self.style.percent)

            nrow += 1
            i += 1

        self.ws.set_column(first_col=ncol, last_col=ncol, width=24)
        self.ws.set_column(first_col=ncol + 1, last_col=ncol + 2, width=10)

        logging.info(f"终端类型数据表写入完成")

        return nrow

    def write_center_branch(self, dtype, app, nrow, week=None):
        """
        写入中心支公司数据
        """

        nrow = nrow
        ncol = 10

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 3,
            data="掌上宝APP出单数量统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 3,
            data=f"数据统计范围：2020-2-17 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        header = ["中心支公司", "APP件数", "总件数", "APP件数占比"]

        self.ws.write_row(nrow, ncol, header, self.style.header)
        nrow += 1

        if dtype == "sum":
            values = app.get_center_branch()
        elif dtype == "week":
            values = app.get_center_branch(week=week)

        app_num = 0
        sum_num = 0

        i = 1
        for value in values:
            self.ws.write(nrow, ncol, value[0], self.style.string)
            self.ws.write(nrow, ncol + 1, value[1], self.style.string)
            self.ws.write(nrow, ncol + 2, value[2], self.style.string)
            self.ws.write(nrow, ncol + 3, value[3], self.style.percent)
            app_num += value[1]
            sum_num += value[2]
            nrow += 1
            i += 1

        self.ws.write(nrow, ncol, "合计", self.style.string)
        self.ws.write(nrow, ncol + 1, app_num, self.style.string)
        self.ws.write(nrow, ncol + 2, sum_num, self.style.string)
        self.ws.write(nrow, ncol + 3, app_num / sum_num, self.style.percent)
        nrow += 1

        self.ws.set_column(first_col=ncol, last_col=ncol, width=24)
        self.ws.set_column(first_col=ncol + 1, last_col=ncol + 2, width=10)
        self.ws.set_column(first_col=ncol + 3, last_col=ncol + 3, width=14)

        logging.info(f"中心支公司数据表写入完成")

        return nrow

    def write_company(self, dtype, name, app, nrow, week=None):
        """
        写入中心支公司数据
        """

        nrow = nrow
        ncol = 10

        # 写入表标题
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 3,
            data="掌上宝APP出单数量统计表",
            cell_format=self.style.title,
        )
        nrow += 1

        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + 3,
            data=f"数据统计范围：2020-2-17 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        nrow += 1

        header = ["机构", "APP件数", "总件数", "APP件数占比"]

        self.ws.write_row(nrow, ncol, header, self.style.header)
        nrow += 1

        if dtype == "sum":
            values = app.get_company(name=name)
        elif dtype == "week":
            values = app.get_company(name=name, week=week)

        app_num = 0
        sum_num = 0

        i = 1
        for value in values:
            self.ws.write(nrow, ncol, value[0], self.style.string)
            self.ws.write(nrow, ncol + 1, value[1], self.style.string)
            self.ws.write(nrow, ncol + 2, value[2], self.style.string)
            self.ws.write(nrow, ncol + 3, value[3], self.style.percent)
            if value[1] is not None:
                app_num += value[1]
            if value[2] is not None:
                sum_num += value[2]
            nrow += 1
            i += 1

        self.ws.write(nrow, ncol, "合计", self.style.string)
        self.ws.write(nrow, ncol + 1, app_num, self.style.string)
        self.ws.write(nrow, ncol + 2, sum_num, self.style.string)
        self.ws.write(nrow, ncol + 3, app_num / sum_num, self.style.percent)
        nrow += 1

        self.ws.set_column(first_col=ncol, last_col=ncol, width=24)
        self.ws.set_column(first_col=ncol + 1, last_col=ncol + 2, width=10)
        self.ws.set_column(first_col=ncol + 3, last_col=ncol + 3, width=14)

        logging.info(f"{name}中心支公司数据表写入完成")

        return nrow

    def write_sum(self):
        """
        累计数据控制函数
        """
        self.set_table_name("累计数据统计表")
        app = Stats_App()
        app.attach_db()

        self.write_salesman("sum", app)
        nrow = self.write_terminal("sum", app)
        nrow = self.write_center_branch("sum", app, nrow + 1)
        nrow = self.write_company("sum", "昆明", app, nrow + 1)
        nrow = self.write_company("sum", "曲靖", app, nrow + 1)
        nrow = self.write_company("sum", "文山", app, nrow + 1)
        nrow = self.write_company("sum", "大理", app, nrow + 1)
        nrow = self.write_company("sum", "保山", app, nrow + 1)
        nrow = self.write_company("sum", "版纳", app, nrow + 1)
        nrow = self.write_company("sum", "怒江", app, nrow + 1)

        self.set_table_name("第8周数据统计表")
        self.write_salesman("week", app, week=8)
        nrow = self.write_terminal("week", app, week=8)
        nrow = self.write_center_branch("week", app, nrow + 1, week=8)
        nrow = self.write_company("week", "昆明", app, nrow + 1, week=8)
        nrow = self.write_company("week", "曲靖", app, nrow + 1, week=8)
        nrow = self.write_company("week", "文山", app, nrow + 1, week=8)
        nrow = self.write_company("week", "大理", app, nrow + 1, week=8)
        nrow = self.write_company("week", "保山", app, nrow + 1, week=8)
        nrow = self.write_company("week", "版纳", app, nrow + 1, week=8)
        nrow = self.write_company("week", "怒江", app, nrow + 1, week=8)

        self.set_table_name("第9周数据统计表")
        self.write_salesman("week", app, week=9)
        nrow = self.write_terminal("week", app, week=9)
        nrow = self.write_center_branch("week", app, nrow + 1, week=9)
        nrow = self.write_company("week", "昆明", app, nrow + 1, week=9)
        nrow = self.write_company("week", "曲靖", app, nrow + 1, week=9)
        nrow = self.write_company("week", "文山", app, nrow + 1, week=9)
        nrow = self.write_company("week", "大理", app, nrow + 1, week=9)
        nrow = self.write_company("week", "保山", app, nrow + 1, week=9)
        nrow = self.write_company("week", "版纳", app, nrow + 1, week=9)
        nrow = self.write_company("week", "怒江", app, nrow + 1, week=9)

        app.detach_db()


if __name__ == "__main__":
    a = datetime.now()
    wb = xlsxwriter.Workbook(r"APP出单业务竞赛统计表.xlsx")

    app = Excel_Write_App(wb=wb)
    app.write_sum()

    wb.close()
    b = datetime.now()
    print(f"date {b-a=}")
