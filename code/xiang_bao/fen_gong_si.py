import logging

import xlsxwriter

from ..style import Style
from ..date import IDate
from ..tong_ji import Tong_Ji

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')


def fen_gong_si(wb: xlsxwriter.Workbook,
                ws: xlsxwriter.worksheet) -> None:

    logging.debug('开始写入分公司历年保费同比数据统计表')

    # 获取单元格样式对象
    sy = Style(wb)

    # 获取日期对象
    idate = IDate(2020)

    # 设置行、列计数器
    nrow = 0
    ncol = 0

    # 写入表标题
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + 6,
        data='分公司历年保费同比数据统计表',
        cell_format=sy.biao_ti
    )

    # 设置表标题行高为字体的两倍
    ws.set_row(row=nrow, height=24)
    nrow += 1
    logging.debug('表标题写入完成')

    # 写入说明性文字，数据统计的时间范围
    ws.merge_range(
        first_row=nrow,
        first_col=ncol,
        last_row=nrow,
        last_col=ncol + 6,
        data=f'数据统计范围：01-01 至 {idate.duan_ri_qi()}',
        cell_format=sy.shuo_ming
    )
    nrow += 1
    logging.debug('统计范围说明性文字写入完成')

    # 写入表头
    biao_ti = ['年份', '保费', '同比\n增长率']

    for value in biao_ti:
        ws.write(nrow, ncol, value, sy.wen_zi_cu_hui)
        ncol += 1

    nrow += 1
    ncol = 0
    logging.debug('表头写入完成')

    nian_fen = [2016, 2017, 2018, 2019, 2020]

    for nian in nian_fen:
        data = Tong_Ji(name='分公司整体', risk='整体')
        ws.write(nrow, ncol, f'{nian}年', sy.wen_zi)
        ws.write(
            nrow, ncol + 1,
            data.wang_nian_bao_fei(year=nian),
            sy.shu_zi
        )
        ws.write(
            nrow, ncol + 2,
            data.wang_nian_tong_bi(first_year=nian, last_year=nian - 1),
            sy.jin_du
        )
        nrow += 1

    nrow += 1
    chart = wb.add_chart({'type': 'column'})
    chart.add_series({
        'categories': ['分公司历年保费同比', 3, ncol, nrow - 2, ncol],
        'values': ['分公司历年保费同比', 3, ncol + 1, nrow - 2, ncol + 1]
    })
    ws.insert_chart('A10', chart)
