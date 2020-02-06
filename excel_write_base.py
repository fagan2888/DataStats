import logging
import xlsxwriter
from datetime import datetime
import sqlite3

# from io import StringIO

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji

from style import Style
from date import IDate
from stats import Stats

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Excel_Write_Base(object):
    """
    Excel表格数据写入的基类

        负责初始化写入数据时所需的各项数据的公共变量
    """

    def __init__(self, wb: xlsxwriter.Workbook, name: str, conn: sqlite3.Connection):
        """
        初始化基础变量

            参数：
                wb:
                    xlsxwriter.Workbook, 需要写入数据的工作簿
                name:
                    str, 机构名称
        """

        # 工作薄对象
        self._wb = wb

        # 工作表对象
        self._ws: xlsxwriter.worksheet = None

        # 机构名称
        self._name = name

        # 数据库连接对象
        self._conn = conn

        # 数据库操作的游标对象
        self._cur = self._conn.cursor()

        # 数据库内数据的最大日期
        self._idate = IDate(year=2020)

        self._style = Style(wb=self.wb)

        # 行计数器
        self.nrow: int = 5

        # 列计数器
        self.ncol: int = 0

        # 操作的工作表名称
        self._table_name: str = None

        # 设置菜单的锚和文本信息
        self._menu: list = []

        # 首列的列名称
        self._first_col_name: str = None

        # 表个第一行表头的信息
        self.set_header_row_1()

        # 表个第二行表头的信息
        self._header_row_2: list = None

        # 当前正在操作的险种名称
        self._risk: str = None

        # 需要统计的险种名称列表
        self._risks: list = []

    def set_risk(self, risk: str):
        """
        当前正在操作的险种

            参数：
                risk：
                    str，险种名称
        """
        self._risk = risk

    @property
    def risk(self):
        """
        返回当前正在操作的险种
        """
        return self._risk

    def set_risks(self, risks: list):
        """
        设置统计的险种列表

            参数：
                risks：
                    list，一个包含整张表中需要统计的险种名称列表
        """
        self._risks = risks

    @property
    def risks(self):
        """
        返回统计的险种列表

            返回值：
                list
        """

        return self._risks

    def set_menu(self, nrow: int, risk: str):
        """
        设置表顶部的快捷菜单内容

            参数：
                nrow：
                    int，快捷菜单跳转的锚信息，一个行号
                risk：
                    str，快捷菜单中显示的文本信息，应该是一个险种名称
            返回值：
                无
        """
        self._menu.append((nrow, risk))

    @property
    def menu(self):
        """
        返回快捷菜单栏中的锚和文本信息

            返回一个列表，列表中的每一项为一个元组,元组中记录了快捷菜单中每一项的锚和文本信息
            列表的结构为[(nrow, risk), (nrow, risk),……]
        """
        return self._menu

    @property
    def years(self):
        """
        返回统计的年份列表

            返回值：
                list
        """

        value = [2020, 2019, 2018, 2017, 2016]

        return value

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

    @property
    def conn(self):
        """
        返回数据库连接对象
        """
        return self._conn

    @property
    def cur(self):
        """
        返回数据库操作的游标对象
        """
        return self._cur

    @property
    def ws(self):
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
    def date(self):
        """
        返回数据库中当前的时间对象
        """
        return self._idate

    @property
    def name(self):
        """
        返回机构名称
        """
        return self._name

    @property
    def wb(self):
        """
        返回需要操作的工作簿对象
        """
        return self._wb

    @property
    def style(self):
        """
        返回Excel表格所需的单元格样式对象
        """
        return self._style

    def freeze_panes(self):
        """
        冻结表个的首行，首列
        """
        nrow = 3
        ncol = 1
        # 冻结快捷工具条所在行（第一行）
        self.ws.freeze_panes(row=nrow, col=ncol, top_row=3, left_col=1)

    def set_column_width(self):
        """
        设置列宽
        """
        self.ws.set_column(first_col=self.ncol, last_col=self.ncol, width=10)
        self.ws.set_column(first_col=self.ncol + 1, last_col=self.ncol + 25, width=13)

    def set_first_col_name(self, name: str):
        """
        设置首列的列名称
        """
        self._first_col_name = name

    @property
    def first_col_name(self):
        """
        返回首列的列名称
        """
        return self._first_col_name

    def set_header_row_1(self):
        """
        设置表格表头的第一行信息

            通常第一行信息为年份信息
        """

        header = []

        for value in self.years:
            header.append(f"{value}年")

        self._header_row_1 = header

    @property
    def header_row_1(self):
        """
        表格表头的第一行信息

            返回的信息是一个包含第一行表头所有列名的一个列表
        """
        return self._header_row_1

    def set_header_row_2(self, header: list):
        """
        设置表格表头的第二行信息

            通常第二行信息为年份信息

            参数：
                header：
                    list，包含第二行表头所有列名的一个列表
        """

        self._header_row_2 = header

    @property
    def header_row_2(self):
        """
        表格表头的第二行信息

            返回的信息是一个包含第二行表头所有列名的一个列表
        """

        return self._header_row_2

    def write_menu(self) -> None:
        """
        写入表顶部的快捷菜单
        """
        # 创建表头的快捷工具条
        nrow = 0
        ncol = 0

        self.ws.write_url(
            row=nrow,
            col=ncol,
            url="internal:目录!A1",
            string="返回目录",
            cell_format=self.style.menu,
        )
        ncol += 1

        for col, data in self.menu:
            self.ws.write_url(
                row=nrow,
                col=ncol,
                url=f"internal:{self.table_name}!A{col+1}",
                string=data,
                cell_format=self.style.menu,
            )
            ncol += 1
        nrow += 1

        self.write_header(nrow=nrow)

    def write_title(self):
        """
        写入相应险种统计表的表标题

            参数：
                risk：
                    str，统计表所对应的险种名称
        """
        # 计算表头所占宽度
        title_width = len(self.header_row_1) * len(self.header_row_2)

        # 列计数器归零
        ncol = 0

        # 写入表标题
        # 表标题采用富文本格式，用浅蓝色突出险种类型
        self.ws.merge_range(
            first_row=self.nrow,
            first_col=ncol,
            last_row=self.nrow,
            last_col=ncol + title_width,
            data="",
            cell_format=self.style.title,
        )

        # 写入富文本格式数据
        self.ws.write_rich_string(
            self.nrow,
            ncol,
            self.name,
            self.style.deep_sky_blue,
            f"{self.risk}业务",
            self.style.black,
            "保费数据统计表",
            self.style.title,
        )

        # 设置表标题行高为字体的1.5倍
        self.ws.set_row(row=self.nrow, height=18)
        self.nrow += 1
        logging.info(f"{self.name}{self.risk}表标题写入完成")

        # 写入说明性文字，数据统计的时间范围
        self.ws.merge_range(
            first_row=self.nrow,
            first_col=ncol,
            last_row=self.nrow,
            last_col=ncol + title_width,
            data=f"数据统计范围：01-01 至 {self.date.short_date()}",
            cell_format=self.style.explain,
        )
        self.nrow += 1
        logging.info(f"{self.name}{self.risk}表统计范围说明性文字写入完成")

    def write_header(self, nrow: int = None):
        """
        写入统计表的表头
        """

        if nrow is None:
            nrow = self.nrow
            risk = self.risk
        else:
            risk = ""

        ncol = 0

        # 表头第一行写入时的列数偏移量
        ncol_offset = len(self.header_row_2) - 1

        # 写入表头中首列的列名，首列列名采用上下单元格合并的方式
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow + 1,
            last_col=ncol,
            data=self.first_col_name,
            cell_format=self.style.string_bold_gray,
        )
        ncol += 1

        # 开始写入表头第二列开始的其他内容
        # 通过判断表头第一行中的数据来判断不同的单元格背景色，两种颜色交替出现
        for row1_data in self.header_row_1:
            i = self.header_row_1.index(row1_data)
            if i % 2 == 0:
                style = self.style.string_bold_orange
            else:
                style = self.style.string_bold_green

            # 写入表头第一行信息
            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + ncol_offset,
                data=row1_data,
                cell_format=style,
            )

            # 写入表头第二行数据
            for row2_data in self.header_row_2:
                self.ws.write_string(
                    row=nrow + 1, col=ncol, string=row2_data, cell_format=style
                )
                ncol += 1

        # 表头占两行
        self.nrow += 2

        logging.info(f"{self.name}{risk}数据表表头写入完成")

    def write_data(self):
        """
        写入指定统计表中的数据
        """

        self.set_menu(self.nrow, f"{self.risk}")

        # 获取机构数据统计的对象
        data = Stats(name=self.name, risk=self.risk, conn=self.conn)

        # 写入首列信息
        ncol = 0
        nrow_val = self.nrow
        for value in list(data.day_list):
            if nrow_val % 2 == 1:
                string = self.style.string
            else:
                string = self.style.string_gray
            self.ws.write_string(
                row=nrow_val, col=ncol, string=value, cell_format=string
            )
            nrow_val += 1
        ncol += 1

        # 写入表数据
        for year in self.years:
            nrow_val = self.nrow
            day_premium = data.day_premium(year=year)
            day_sum = data.day_sum(year=year)
            day_sum_yoy = data.day_sum_yoy(year=year)

            for key in list(data.day_list):
                ncol_val = ncol
                if nrow_val % 2 == 1:
                    number = self.style.number
                    percent = self.style.percent
                else:
                    number = self.style.number_gray
                    percent = self.style.percent_gray

                self.ws.write(nrow_val, ncol_val, day_premium[key], number)
                ncol_val += 1
                self.ws.write(nrow_val, ncol_val, day_sum[key], number)
                ncol_val += 1
                self.ws.write(nrow_val, ncol_val, day_sum_yoy[key], percent)
                nrow_val += 1
            ncol += 3

        # 在下一个统计表之前增长一个空行
        self.nrow += 367

        logging.info(f"{self.name}{self.risk}数据表写入完成")
        logging.info("-" * 60)

    def write_day_data(self):
        """
        写入日数据的逻辑控制函数

        参数：


        返回值：
            无
        """

        logging.info("开始写入日数据统计表")

        day.set_table_name(table_name=f"{self.name}日数据统计表")
        day.set_risks(["整体", "车险", "人身险", "财产险", "非车险"])
        day.set_first_col_name("日期")
        day.set_header_row_2(["单日保费", "累计保费", "累计保费同比"])

        # 写入数据统计表并记录快捷菜单栏不同快捷菜单的锚信息和现实文本信息
        for risk in self.risks:
            self.set_risk(risk)
            # 写入表标题
            self.write_title()
            # 写入表头
            self.write_header()
            # 写入表数据
            self.write_data()

        # 写入快捷菜单栏
        self.write_menu()
        self.freeze_panes()
        self.set_column_width()


if __name__ == "__main__":
    wb = xlsxwriter.Workbook(r"2020年机构数据统计详细报告.xlsx")
    ws = wb.add_worksheet("目录")
    a = datetime.now()
    conn = sqlite3.connect(r"Data\data.db")

    # str_buffer = StringIO()

    # for line in conn.iterdump():
    #     str_buffer.write(f"{line}\n")

    # conn.close()

    # conn = sqlite3.connect(":memory:")
    # cur = conn.cursor()
    # cur.executescript(str_buffer.getvalue())
    # b = datetime.now()
    # print(f"date {b-a=}")

    a = datetime.now()
    day = Excel_Write_Base(wb=wb, name="分公司", conn=conn)
    day.write_day_data()
    b = datetime.now()
    wb.close()
    print(f"date {b-a=}")
