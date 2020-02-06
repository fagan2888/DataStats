import logging
import xlsxwriter
from datetime import datetime
import sqlite3
from io import StringIO

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji

from code.style import Style
from code.date import IDate
from code.tong_ji import Tong_Ji

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


def header_write(xian_zhong: str, tong: str, idate: IDate):
    """表头文字写入"""

    # 设置全局变量
    global name
    global wb
    global ws
    global nrow
    global ncol
    global sy

    # 表头第一列名称
    first_col_name = "周"

    # 表头第一行
    biao_tou_row1 = ["2016年", "2017年", "2018年", "2019年", "2020年"]

    # 表头第二行
    if xian_zhong == "整体":
        biao_tou_row2 = ["保费", "同比", "环比"]
    else:
        biao_tou_row2 = ["保费", "同比", "环比"]

    # 计算表头所占宽度
    biao_ti_col_width = len(biao_tou_row1) * len(biao_tou_row2)

    # 列计数器归零
    ncol = 0

    # 写入表标题
    # 表标题采用富文本格式，用浅蓝色突出险种类型
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + biao_ti_col_width,
        data="",
        cell_format=sy.title,
    )

    # 写入富文本格式数据
    ws.write_rich_string(
        nrow,
        ncol,
        name,
        sy.deep_sky_blue,
        f"{xian_zhong}业务{tong}",
        sy.black,
        "保费数据统计表",
        sy.title,
    )

    # 设置表标题行高为字体的1.5倍
    ws.set_row(row=nrow, height=18)
    nrow += 1
    logging.info(f"{name}{xian_zhong}{tong}表标题写入完成")

    # 写入说明性文字，数据统计的时间范围
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + biao_ti_col_width,
        data=f"数据统计范围：01-01 至 {idate.duan_ri_qi()}",
        cell_format=sy.explain,
    )
    nrow += 1
    logging.info(f"{name}{xian_zhong}{tong}表统计范围说明性文字写入完成")

    # 开始写入表头

    # 表头第一行写入时的列数偏移量
    ncol_offset = len(biao_tou_row2) - 1

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
    for row1_data in biao_tou_row1:
        i = biao_tou_row1.index(row1_data)
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
        for row2_data in biao_tou_row2:
            ws.write_string(row=nrow + 1, col=ncol, string=row2_data, cell_format=style)
            ncol += 1

    # 表头占两行
    nrow += 2

    logging.info(f"{name}{xian_zhong}{tong}数据表表头写入完成")


def data_write(xian_zhong: str, tong: str, idate: IDate, conn: object):
    """表数据写入"""

    # 设置全局变量
    global name
    global wb
    global ws
    global nrow
    global ncol
    global sy

    # 设置周列表
    weeknum_list = [
        1,
        2,
        3,
        4,
        5,
        6,
        7,
        8,
        9,
        10,
        11,
        12,
        13,
        14,
        15,
        16,
        17,
        18,
        19,
        20,
        21,
        22,
        23,
        24,
        25,
        26,
        27,
        28,
        29,
        30,
        31,
        32,
        33,
        34,
        35,
        36,
        37,
        38,
        39,
        40,
        41,
        42,
        43,
        44,
        45,
        46,
        47,
        48,
        49,
        50,
        51,
        52,
        53
    ]

    if tong == "同期":
        tong = True
    else:
        tong = False

    # 获取不同险种的相关数据
    fen = Tong_Ji(name=name, risk=xian_zhong, conn=conn)

    # 需要统计的年份
    year_list = [2016, 2017, 2018, 2019, 2020]

    for weeknum in weeknum_list:
        # 通过周数据的变化来判断行变换，从而实现隔行换背景色
        if weeknum % 2 == 1:
            wen_zi = sy.string
            shu_zi = sy.number
            bai_fen_bi = sy.percent
        else:
            wen_zi = sy.string_gray
            shu_zi = sy.number_gray
            bai_fen_bi = sy.percent_gray

        # 列计数器归零
        ncol = 0
        # 写入周名称
        ws.write_string(
            row=nrow,
            col=ncol,
            string=f"{weeknum}周",
            cell_format=wen_zi,
        )
        ncol += 1

        for year in year_list:
            # 写入保费
            ws.write(
                nrow,
                ncol,
                fen.zhou_bao_fei(
                    year=year, weeknum=weeknum, tong=tong
                ),
                shu_zi,
            )
            ncol += 1

            # 写入同比增长率
            # 本年的同比数据只能是同期数据
            if year == idate.nian:
                value = fen.zhou_tong_bi(
                    year=year, weeknum=weeknum, tong=True
                )
            else:
                value = fen.zhou_tong_bi(
                    year=year, weeknum=weeknum, tong=tong
                )
            ws.write(
                nrow, ncol, value, bai_fen_bi,
            )
            ncol += 1

            # 写入环比增长率
            # 本年的环比数据只能是同期数据
            if year == idate.nian:
                value = fen.zhou_huan_bi(
                    year=year, weeknum=weeknum, tong=True
                )
            else:
                value = fen.zhou_huan_bi(
                    year=year, weeknum=weeknum, tong=tong
                )
            ws.write(
                nrow, ncol, value, bai_fen_bi,
            )
            ncol += 1

        nrow += 1

    # 在下一个统计表之前增长一个空行
    nrow += 1

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


def weeknum_data(workbook: xlsxwriter.Workbook, company: str, conn: object) -> None:
    """
    写入周数据的主函数

    参数：
        workbook: xlsxwriter.Workbook，用于确认数据写入的工作簿
        company: str，机构名称

    返回值：
        无
    """

    logging.info("开始写入周数据统计表")

    # 设置全局变量
    global name
    global wb
    global ws
    global nrow
    global ncol
    global sy

    # 对部分全局变量进行初始化
    wb = workbook
    name = company
    sy = Style(wb)
    idate: IDate = IDate(2020)

    table_name = f"{name}周数据统计表"
    ws = wb.add_worksheet(table_name)

    # 需要统计的险种信息
    risks = ["整体", "车险", "非车险"]
    istong = ["同期", "整周"]

    # 记录快捷菜单栏的项目信息
    menu = []

    # 写入数据统计表并记录快捷菜单栏不同快捷菜单的锚信息和现实文本信息
    for risk in risks:
        for tong in istong:
            menu.append((nrow, f"{risk}{tong}"))
            # 写入表头
            header_write(xian_zhong=risk, tong=tong, idate=idate)
            # 写入表数据
            data_write(xian_zhong=risk, tong=tong, idate=idate, conn=conn)

    # 写入快捷菜单栏
    menu_write(table_name=table_name, menu=menu)

    nrow = 1
    ncol = 0
    # 冻结快捷工具条所在行（第一行）
    ws.freeze_panes(row=nrow, col=ncol + 1, top_row=1, left_col=1)

    # 设置列宽
    ws.set_column(first_col=ncol, last_col=ncol, width=10)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 25, width=12)


if __name__ == "__main__":
    wb = xlsxwriter.Workbook(r"2020年机构数据统计详细报告.xlsx")
    ws = wb.add_worksheet("目录")
    a = datetime.now()
    con = sqlite3.connect(r"Data\data.db")

    str_buffer = StringIO()

    for line in con.iterdump():
        str_buffer.write(f"{line}\n")

    con.close()

    conn = sqlite3.connect(":memory:")
    cur = conn.cursor()
    cur.executescript(str_buffer.getvalue())
    b = datetime.now()
    print(f"date {b-a=}")

    a = datetime.now()
    weeknum_data(workbook=wb, company="分公司", conn=conn)
    b = datetime.now()
    wb.close()
    print(f"date {b-a=}")
