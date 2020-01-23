import logging

import xlsxwriter

from code.xiang_bao.fen_gong_si import fen_gong_si

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

wb = xlsxwriter.Workbook(r'2020年机构数据统计详细报告.xlsx')

ws = wb.add_worksheet('分公司历年保费同比')

fen_gong_si(wb, ws)

wb.close()
