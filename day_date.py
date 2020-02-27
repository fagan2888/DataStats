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

#
# 以下为设置全局变量，便于在各函数间调用
#
# 工作簿对象
wb: xlsxwriter.Workbook = None

# 工作表对象
ws: xlsxwriter.worksheet = None

# 机构名称
name: str = None

# 单元格样式对象
sy: Style = None

# 行计数器
nrow: int = 1

# 列计数器
ncol: int = 0

idate: IDate = None


def header_write(risk: str):
    """表头文字写入"""

    # 设置全局变量
    global nrow
    global ncol

    # 表头第一列名称
    first_col_name = "日期"

    # 表头第一行
    header_row_1 = ["2020年", "2019年", "2018年", "2017年", "2016年"]

    header_row_2 = ["单日保费", "累计保费", "累计保费同比"]

    # 计算表头所占宽度
    title_width = len(header_row_1) * len(header_row_2)

    # 列计数器归零
    ncol = 0

    # 写入表标题
    # 表标题采用富文本格式，用浅蓝色突出险种类型
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + title_width,
        data="",
        cell_format=sy.title,
    )

    # 写入富文本格式数据
    ws.write_rich_string(
        nrow,
        ncol,
        name,
        sy.deep_sky_blue,
        f"{risk}业务",
        sy.black,
        "保费数据统计表",
        sy.title,
    )

    # 设置表标题行高为字体的1.5倍
    ws.set_row(row=nrow, height=18)
    nrow += 1
    logging.info(f"{name}{risk}表标题写入完成")

    # 写入说明性文字，数据统计的时间范围
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + title_width,
        data=f"数据统计范围：01-01 至 {idate.duan_ri_qi()}",
        cell_format=sy.explain,
    )
    nrow += 1
    logging.info(f"{name}{risk}表统计范围说明性文字写入完成")

    # 开始写入表头

    # 表头第一行写入时的列数偏移量
    ncol_offset = len(header_row_2) - 1

    # 表头第一列采用上下单元格合并的方式呈现，与其他所有单元格均不同
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow + 1,
        last_col=ncol,
        data=first_col_name,
        cell_format=sy.string_bold_gray,
    )
    ncol += 1

    # 开始写入表头第二列开始的其他内容
    # 通过判断表头第一行中的数据来判断不同的单元格背景色，两种颜色交替出现
    for row1_data in header_row_1:
        i = header_row_1.index(row1_data)
        if i % 2 == 0:
            style = sy.string_bold_orange
        else:
            style = sy.string_bold_green

        # 写入表头第一行输入
        ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + ncol_offset,
            data=row1_data,
            cell_format=style,
        )

        # 写入表头第二行数据
        for row2_data in header_row_2:
            ws.write_string(row=nrow + 1, col=ncol, string=row2_data, cell_format=style)
            ncol += 1

    # 表头占两行
    nrow += 2

    logging.info(f"{name}{risk}数据表表头写入完成")


def data_write(risk: str, conn: sqlite3.Connection):
    """表数据写入"""

    # 设置全局变量
    global nrow
    global ncol

    # 获取机构数据统计的对象
    data = Stats(name=name, risk=risk, conn=conn)

    # 需要统计的年份
    years = [2020, 2019, 2018, 2017, 2016]

    ncol = 0
    day_premium = data.day_premium(year=2020)

    nrow_val = nrow
    for key in list(day_premium):
        if nrow_val % 2 == 1:
            string = sy.string
        else:
            string = sy.string_gray
        ws.write_string(row=nrow_val, col=ncol, string=key, cell_format=string)
        nrow_val += 1
    ncol += 1

    for year in years:
        nrow_val = nrow
        day_premium = data.day_premium(year=year)
        day_sum = data.day_sum(year=year)
        day_sum_yoy = data.day_sum_yoy(year=year)

        for key in list(day_premium):
            ncol_val = ncol
            if nrow_val % 2 == 1:
                number = sy.number
                percent = sy.percent
            else:
                number = sy.number_gray
                percent = sy.percent_gray

            ws.write(nrow_val, ncol_val, day_premium[key], number)
            ncol_val += 1
            ws.write(nrow_val, ncol_val, day_sum[key], number)
            ncol_val += 1
            ws.write(nrow_val, ncol_val, day_sum_yoy[key], percent)
            nrow_val += 1
        ncol += 3

    # 在下一个统计表之前增长一个空行
    nrow += 367

    logging.info(f"{name}整体数据表写入完成")
    logging.info("-" * 60)


def menu_write(table_name: str, menu: list):
    """快捷菜单栏的写入"""

    # 创建表头的快捷工具条
    nrow = 0
    ncol = 0

    ws.write_url(
        row=nrow, col=ncol, url="internal:目录!A1", string="返回目录", cell_format=sy.menu,
    )
    ncol += 1

    for col, data in menu:
        ws.write_url(
            row=nrow,
            col=ncol,
            url=f"internal:{table_name}!A{col+1}",
            string=data,
            cell_format=sy.menu,
        )
        ncol += 1


def day_data(
    workbook: xlsxwriter.Workbook, company: str, conn: sqlite3.Connection
) -> None:
    """
    写入日数据的主函数

    参数：
        workbook: xlsxwriter.Workbook，用于确认数据写入的工作簿
        company: str，机构名称

    返回值：
        无
    """

    logging.info("开始写入日数据统计表")

    # 设置全局变量
    global name
    global wb
    global ws
    global nrow
    global ncol
    global sy
    global idate

    # 对部分全局变量进行初始化
    wb = workbook
    name = company
    sy = Style(wb)
    idate = IDate(2020)

    table_name = f"{name}日数据统计表"
    ws = wb.add_worksheet(table_name)

    # 需要统计的险种信息
    risks = ["整体", "车险", "人身险", "财产险", "非车险"]

    # 记录快捷菜单栏的项目信息
    menu = []

    # 写入数据统计表并记录快捷菜单栏不同快捷菜单的锚信息和现实文本信息
    for risk in risks:
        menu.append((nrow, f"{risk}"))
        # 写入表头
        header_write(risk=risk)
        # 写入表数据
        data_write(risk=risk, conn=conn)

    # 写入快捷菜单栏
    menu_write(table_name=table_name, menu=menu)

    nrow = 1
    ncol = 0
    # 冻结快捷工具条所在行（第一行）
    ws.freeze_panes(row=nrow, col=ncol + 1, top_row=1, left_col=1)

    # 设置列宽
    ws.set_column(first_col=ncol, last_col=ncol, width=10)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 25, width=13)


if __name__ == "__main__":
    wb = xlsxwriter.Workbook(r"2020年机构数据统计详细报告.xlsx")
    ws = wb.add_worksheet("目录")
    a = datetime.now()
    conn = sqlite3.connect(r"Data\data.db")

    a = datetime.now()
    day_data(workbook=wb, company="分公司", conn=conn)
    b = datetime.now()
    wb.close()
    print(f"date {b-a=}")
