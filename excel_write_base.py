import logging
import xlsxwriter
import sqlite3

# from io import StringIO

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji

from style import Style
from date import IDate

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

    def __init__(self, wb: xlsxwriter.Workbook, name: str):
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
        self._ws: xlsxwriter.Workbook.worksheet_class = None

        # 机构名称
        self._name = name

        # 数据库连接对象
        self._conn = None

        self._set_conn()

        # 数据库操作的游标对象
        self._cur = None

        self._set_cur()

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

        # 首年的第一行信息
        self.set_first_year_header_1()

        # 表个第一行表头的信息
        self.set_other_year_header_1()

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

    def set_risk_list(self, risks: list):
        """
        设置统计的险种列表

            参数：
                risks：
                    list，一个包含整张表中需要统计的险种名称列表
        """
        self._risks = risks

    @property
    def risk_list(self):
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
    def menu(self) -> list:
        """
        返回快捷菜单栏中的锚和文本信息

            返回一个列表，列表中的每一项为一个元组,元组中记录了快捷菜单中每一项的锚和文本信息
            列表的结构为[(nrow, risk), (nrow, risk),……]
        """
        return self._menu

    @property
    def year_list(self) -> list:
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
    def name(self):
        """
        返回机构名称
        """
        return self._name

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

    def set_first_year_header_1(self):
        """
        设置第一年的表头第一行信息

            第一年的表头第一行信息与其他信息不同，因为第一年为当前年份，信息保护任务进度
            时间进度，时间进度达成率等信息，所以第一年的跨行更多
        """

        self._first_year_header_1 = self.year_list[0]

    @property
    def first_year_header_1(self):
        """
        表格表头的第一行信息

            返回的信息是一个包含第一行表头所有列名的一个列表
        """
        return self._first_year_header_1

    def set_other_year_header_1(self):
        """
        设置表格表头的第一行信息

            通常第一行信息为年份信息
        """

        header = []

        for value in self.year_list[1:]:
            header.append(f"{value}年")

        self._other_year_header_1 = header

    @property
    def other_year_header_1(self):
        """
        表格表头的第一行信息

            返回的信息是一个包含第一行表头所有列名的一个列表
        """
        return self._other_year_header_1

    def set_first_year_header_2(self, header: list):
        """
        设置表格表头中第一年的第二行信息

            参数：
                header：
                    list，包含第二行表头所有列名的一个列表
        """

        self._first_year_header_2 = header

    def set_other_year_header_2(self, header: list):
        """
        设置表格表头中除第一年外的第二行信息

            参数：
                header：
                    list，包含第二行表头所有列名的一个列表
        """

        self._other_year_header_2 = header

    @property
    def first_year_header_2(self):
        """
        表格表格表头中除第一年外的第二行信息

            返回的信息是一个包含第二行表头所有列名的一个列表
        """

        return self._first_year_header_2

    @property
    def other_year_header_2(self):
        """
        表格表格表头中除第一年外的第二行信息

            返回的信息是一个包含第二行表头所有列名的一个列表
        """

        return self._other_year_header_2

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
        title_width = len(self.first_year_header_2) + len(
            self.other_year_header_1
        ) * len(self.other_year_header_2)

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

    def freeze_panes(self):
        """
        冻结表个的首行，首列
        """
        nrow = 3
        ncol = 1
        # 冻结快捷工具条所在行（第一行）
        self.ws.freeze_panes(row=nrow, col=ncol, top_row=3, left_col=1)

    def write_header(self, nrow: int = None):
        """
        写入统计表的表头
        """
        raise NotImplementedError

    def write_data(self):
        """
        写入指定统计表中的数据
        """
        raise NotImplementedError

    def set_column_width(self):
        """
        设置列宽
        """
        raise NotImplementedError

    def write_data(self):  # noqa
        """
        写入日数据的逻辑控制函数
        """
        raise NotImplementedError
