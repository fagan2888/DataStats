import logging
import xlsxwriter
import sqlite3
from ji_gou_tong_ji import Tong_ji

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

wb = xlsxwriter.Workbook("昆明机构周报.xlsx")
ws = wb.add_worksheet('昆明机构周报')

style_chu = wb.add_format()
style_chu.set_font_name('微软雅黑')
style_chu.set_font_size(11)
style_chu.set_bold()
style_chu.set_valign('center')
style_chu.set_text_wrap()
style_chu.set_border(1)

style_str = wb.add_format()
style_str.set_font_name('微软雅黑')
style_str.set_font_size(11)
style_str.set_valign('center')
style_str.set_text_wrap()
style_str.set_border(1)

style_shu = wb.add_format()
style_shu.set_font_name('微软雅黑')
style_shu.set_font_size(11)
style_shu.set_valign('center')
style_shu.set_num_format('0.00')
style_shu.set_border(1)

style_bi = wb.add_format()
style_bi.set_font_name('微软雅黑')
style_bi.set_font_size(11)
style_bi.set_valign('center')
style_bi.set_num_format('0.00%')
style_bi.set_border(1)


ji_gou = ['百大国际',
          '春怡雅苑',
          '香榭丽园',
          '宜良',
          '东川',
          '安宁',
          '春之城',
          '分公司本部',
          '航旅项目']

xian_zhong = ['整体', '车险', '财产险', '人身险']

row = 0

for xz in xian_zhong:
    col = 0
    ws.merge_range(row, 0, row, 8, f'昆明地区机构{xz}保费汇总表', style_chu)
    row += 1
    ws.write(row, col, '机构', style_chu)
    ws.write(row, col+1, '周保费', style_chu)
    ws.write(row, col+2, '周环比', style_chu)
    ws.write(row, col+3, '周同比', style_chu)
    ws.write(row, col+4, '月保费', style_chu)
    ws.write(row, col+5, '月环比', style_chu)
    ws.write(row, col+6, '月同比', style_chu)
    ws.write(row, col+7, '年保费', style_chu)
    ws.write(row, col+8, '年同比', style_chu)
    row += 1
    for jg in ji_gou:
        info = Tong_ji(jg, xz)
        ws.write(row, col, info.jian_cheng, style_str)
        ws.write(row, col+1, info.zhou_bao_fei, style_shu)
        ws.write(row, col+2, info.zhou_huan_bi, style_bi)
        ws.write(row, col+3, info.zhou_tong_bi, style_bi)
        ws.write(row, col+4, info.yue_bao_fei, style_shu)
        ws.write(row, col+5, info.yue_huan_bi, style_bi)
        ws.write(row, col+6, info.yue_tong_bi, style_bi)
        ws.write(row, col+7, info.nian_bao_fei, style_shu)
        ws.write(row, col+8, info.yi_nian_tong_bi, style_bi)
        logging.debug(f'{info.jian_cheng}数据写入完成')
        row += 1
    logging.debug(f'{xz}写入完成')
    logging.debug('-' * 60)
    row += 1

ws.set_column('A:A', 12)
ws.set_column('B:B', 10)
ws.set_column('C:C', 10)
ws.set_column('D:D', 10)
ws.set_column('E:E', 10)
ws.set_column('F:F', 10)
ws.set_column('G:G', 10)
ws.set_column('H:H', 10)
ws.set_column('I:I', 10)


wb.close()