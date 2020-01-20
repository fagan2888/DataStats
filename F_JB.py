import logging

import xlsxwriter

from code.style import Style
from code.date import IDate
from code.tong_ji import Tong_Ji
# from update import update


logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG,
                    format=' %(asctime)s | %(levelname)s | %(message)s')

# update()

idate = IDate(2020)

wb = xlsxwriter.Workbook('2020年机构数据统计简报.xlsx')
ws = wb.add_worksheet('三级机构数据统计表')

sy = Style(wb)

nrow = 0
ncol = 0

ws.merge_range(first_row=nrow,
               first_col=ncol,
               last_row=nrow,
               last_col=ncol + 6,
               data='2020年三级机构数据统计表',
               cell_format=sy.biao_ti)

ws.set_row(row=nrow, height=24)
nrow += 1

ws.merge_range(first_row=nrow,
               first_col=ncol,
               last_row=nrow,
               last_col=ncol + 6,
               data=f'数据统计范围：2020-01-01 至 {idate.ri_qi()}',
               cell_format=sy.shuo_ming)
nrow += 1

biao_ti = ['序号', '机构', '险种', '计划任务', '累计保费',
           '时间进度\n达成率', '同比\n增长率']

for value in biao_ti:
    ws.write(nrow, ncol, value, sy.biao_tou)
    ncol += 1

nrow += 1
ncol = 0

data = Tong_Ji(name='分公司整体',
               name_leve='分公司',
               risk='车险',
               risk_leve='险种')
name = '分公司整体'
name_leve = '分公司'
risk = ['车险', '财产险', '人身险', '驾意险', '整体']

ws.merge_range(first_row=nrow,
               first_col=ncol,
               last_row=nrow + 4,
               last_col=ncol,
               data='',
               cell_format=sy.wen_zi)

ws.merge_range(first_row=nrow,
               first_col=ncol + 1,
               last_row=nrow + 4,
               last_col=ncol + 1,
               data=name,
               cell_format=sy.wen_zi)

for r in risk:
    if r == '驾意险':
        risk_leve = '险种名称'
    elif r == '整体':
        risk_leve = '整体'
    else:
        risk_leve = '险种'
    data = Tong_Ji(name, name_leve, r, risk_leve)
    ws.write(nrow, ncol + 2, data.xian_zhong, sy.wen_zi)
    ws.write(nrow, ncol + 3, data.ren_wu, sy.wen_zi)
    ws.write(nrow, ncol + 4, data.nian_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 5, data.shi_jian_da_cheng, sy.jin_du)
    ws.write(nrow, ncol + 6, data.nian_tong_bi(1), sy.jin_du)
    nrow += 1

wb.close()
