import logging

import xlsxwriter

from code.xiang_bao.fen_gong_si import fen_gong_si

logging.disable(logging.DEBUG)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')
logging.basicConfig(level=logging.INFO,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

wb = xlsxwriter.Workbook(r'2020年机构数据统计详细报告.xlsx')

fen_gong_si(wb)

wb.close()
