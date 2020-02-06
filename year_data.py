import logging
import xlsxwriter

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


def data_write(xian_zhong: str = None):
    """
    建立保费的同期、全年统计表

    参数：
        xian_zhong: str，险种名称
    """

    # 设置全局变量
    global name
    global wb
    global ws
    global nrow
    global ncol
    global sy

    # 设置统计年份
    nian_fen: list = [2016, 2017, 2018, 2019, 2020]

    # 日期对象
    idate: IDate = IDate(2020)

    # 表头第一列名称
    first_col_name = "年份"

    # 表头第一行
    biao_tou_row1 = [f"{xian_zhong}同期", f"{xian_zhong}全年"]

    # 表头第二行
    if xian_zhong == "整体":
        biao_tou_row2 = ["保费", "同比"]
    else:
        biao_tou_row2 = ["保费", "同比", "业务占比"]

    biao_ti_col_width = len(biao_tou_row1) * len(biao_tou_row2)

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
        f"{xian_zhong}业务",
        sy.black,
        "保费数据统计表",
        sy.title,
    )

    # 设置表标题行高为字体的1.5倍
    ws.set_row(row=nrow, height=18)
    nrow += 1
    logging.info(f"{name}{xian_zhong}表标题写入完成")

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
    logging.info(f"{name}{xian_zhong}表统计范围说明性文字写入完成")

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
    for data1 in biao_tou_row1:
        i = biao_tou_row1.index(data1)
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
            data=data1,
            cell_format=style,
        )

        # 写入表头第二行数据
        for data2 in biao_tou_row2:
            ws.write_string(
                row=nrow + 1, col=ncol, string=data2, cell_format=style
            )
            ncol += 1

    # 表头行占两行
    nrow += 2

    # 列计数器归零
    ncol = 0

    logging.info(f"{name}{xian_zhong}数据表表头写入完成")

    # 开始写入统计数据
    # 获取整体业务的相关数据
    zheng = Tong_Ji(name=name, risk="整体")
    # 获取不同险种的相关数据
    fen = Tong_Ji(name=name, risk=xian_zhong)

    for year in nian_fen:
        # 行交替呈现不同背景色
        i = nian_fen.index(year)
        if i % 2 == 0:
            wen_zi = sy.string
            shu_zi = sy.number
            bai_fen_bi = sy.percent
        else:
            wen_zi = sy.string_gray
            shu_zi = sy.number_gray
            bai_fen_bi = sy.percent_gray

        # 写入年份数据
        ws.write_string(
            row=nrow, col=ncol, string=f"{year}年", cell_format=wen_zi
        )
        ncol += 1

        # 写入年度同期保费
        ws.write_number(
            row=nrow,
            col=ncol,
            number=fen.nian_bao_fei(year=year, tong=True),
            cell_format=shu_zi,
        )
        ncol += 1

        # 写入同期保费同比增长率
        ws.write_number(
            row=nrow,
            col=ncol,
            number=fen.nian_tong_bi(year=year, tong=True),
            cell_format=bai_fen_bi,
        )
        ncol += 1

        # 写入同期保费在整体业务中的占比
        if xian_zhong != "整体":
            ws.write_number(
                row=nrow,
                col=ncol,
                number=fen.nian_bao_fei(year=year, tong=True)
                / zheng.nian_bao_fei(year=year, tong=True),
                cell_format=bai_fen_bi,
            )
            ncol += 1

        # 写入年度全年保费
        ws.write_number(
            row=nrow,
            col=ncol,
            number=fen.nian_bao_fei(year=year, tong=False),
            cell_format=shu_zi,
        )
        ncol += 1

        # 写入全年同比增长率
        ws.write_number(
            row=nrow,
            col=ncol,
            number=fen.nian_tong_bi(year=year, tong=False),
            cell_format=bai_fen_bi,
        )
        ncol += 1

        # 写入全年中保费在整体业务中的占比
        if xian_zhong != "整体":
            ws.write_number(
                row=nrow,
                col=ncol,
                number=fen.nian_bao_fei(year=year, tong=False)
                / zheng.nian_bao_fei(year=year, tong=False),
                cell_format=bai_fen_bi,
            )
        nrow += 1
        ncol = 0

    # 在下一个统计表之前增长一个空行
    nrow += 1

    logging.info(f"{name}整体数据表写入完成")
    logging.info("-" * 60)


def year_data(workbook: xlsxwriter.Workbook, company: str) -> None:
    """
    写入年度数据的主函数

    参数：
        workbook: xlsxwriter.Workbook，用于确认数据写入的工作簿
        company: str，机构名称

    返回值：
        无
    """

    logging.info("开始写入年度数据统计表")

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

    table_name = f"{name}年度数据统计表"
    ws = wb.add_worksheet(table_name)

    # 需要统计的险种信息
    risks = ["整体", "车险", "财产险", "人身险", "非车险"]

    # 记录快捷菜单栏的项目信息
    menu = []

    # 写入数据统计表并记录快捷菜单栏信息
    for risk in risks:
        menu.append((nrow, risk))
        data_write(xian_zhong=risk)

    # 创建表头的快捷工具条
    nrow = 0
    ncol = 0
    ws.write_url(
        row=nrow,
        col=ncol,
        url="internal:目录!A1",
        string="返回目录",
        cell_format=sy.menu,
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

    nrow += 1
    ncol = 0
    # 冻结快捷工具条所在行（第一行）

    ws.freeze_panes(row=nrow, col=ncol, top_row=1)

    # 设置列宽
    ws.set_column(first_col=ncol, last_col=ncol, width=8)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 6, width=12)


if __name__ == "__main__":
    wb = xlsxwriter.Workbook(r"2020年机构数据统计详细报告.xlsx")
    ws = wb.add_worksheet("目录")

    year_data(workbook=wb, company="分公司")

    wb.close()
