import logging

import xlsxwriter

from code.jian_bao.zhong_zhi import zhong_zhi
from code.jian_bao.ji_gou import ji_gou
from code.jian_bao.tuan_dui import tuan_dui

from update import update


logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

# update()
logging.debug('数据库更新完成')

wb = xlsxwriter.Workbook(r'Report\2020年机构数据统计简报.xlsx')

# ws = wb.add_worksheet('三级机构数据统计表')
# zhong_zhi(wb, ws)

# ws = wb.add_worksheet('四级机构数据统计表')
# ji_gou(wb, ws)

ws = wb.add_worksheet('内部团队数据统计表')
tuan_dui(wb, ws)

wb.close()
