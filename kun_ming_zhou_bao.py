import logging
import xlsxwriter
from code.tong_ji import Tong_Ji
from code.style import Style

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

wb = xlsxwriter.Workbook(r"Report\昆明机构周报.xlsx")
ws = wb.add_worksheet('昆明机构周报')

sy = Style(wb=wb)

ji_gou_list = [
    '百大国际', '春怡雅苑', '香榭丽园', '宜良', '东川',
    '安宁', '春之城', '分公司本部', '航旅项目', '昆明'
]

xian_zhong_list = ['整体', '车险', '财产险', '人身险']

nrow = 0

for xian_zhong in xian_zhong_list:
    ncol = 0
    ws.merge_range(nrow, 0, nrow, 8, f'昆明地区机构{xian_zhong}保费汇总表', sy.wen_zi_cu)
    nrow += 1
    ws.write(nrow, ncol, '机构', sy.wen_zi_cu)
    ws.write(nrow, ncol+1, '周保费', sy.wen_zi_cu)
    ws.write(nrow, ncol+2, '周环比', sy.wen_zi_cu)
    ws.write(nrow, ncol+3, '周同比', sy.wen_zi_cu)
    ws.write(nrow, ncol+4, '月保费', sy.wen_zi_cu)
    ws.write(nrow, ncol+5, '月环比', sy.wen_zi_cu)
    ws.write(nrow, ncol+6, '月同比', sy.wen_zi_cu)
    ws.write(nrow, ncol+7, '年保费', sy.wen_zi_cu)
    ws.write(nrow, ncol+8, '年同比', sy.wen_zi_cu)
    nrow += 1
    for ji_gou in ji_gou_list:
        data = Tong_Ji(name=ji_gou, risk=xian_zhong)
        ws.write(nrow, ncol, data.ming_cheng, sy.wen_zi)
        ws.write(nrow, ncol+1, data.zhou_bao_fei(), sy.shu_zi)
        ws.write(nrow, ncol+2, data.zhou_huan_bi(ny=0, nw=1), sy.jin_du)
        ws.write(nrow, ncol+3, data.zhou_tong_bi(ny=1, nw=0), sy.jin_du)
        ws.write(nrow, ncol+4, data.yue_bao_fei(), sy.shu_zi)
        ws.write(nrow, ncol+5, data.yue_huan_bi(ny=0, nm=1), sy.jin_du)
        ws.write(nrow, ncol+6, data.yue_tong_bi(ny=1, nm=0), sy.jin_du)
        ws.write(nrow, ncol+7, data.nian_bao_fei(), sy.shu_zi)
        ws.write(nrow, ncol+8, data.nian_tong_bi(ny=1), sy.jin_du)
        logging.debug(f'{data.ming_cheng}数据写入完成')
        nrow += 1
    logging.debug(f'{xian_zhong}写入完成')
    logging.debug('-' * 60)
    nrow += 1

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
