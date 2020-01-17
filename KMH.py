import sqlite3

import xlsxwriter

from KMH_TJ import KMH_TJ


class Style():
    def __init__(self, wb):
        self._wb = wb

    @property
    def biao_ti(self):
        sy_biao_ti = self._wb.add_format()
        sy_biao_ti.set_font_name('微软雅黑')
        sy_biao_ti.set_font_size(12)
        sy_biao_ti.set_bold(True)
        sy_biao_ti.set_align('center')
        sy_biao_ti.set_align('vcenter')
        return sy_biao_ti

    @property
    def shuo_ming(self):
        sy_shuo_ming = self._wb.add_format()
        sy_shuo_ming.set_font_name('微软雅黑')
        sy_shuo_ming.set_font_size(10)
        sy_shuo_ming.set_align('center')
        sy_shuo_ming.set_align('vcenter')
        return sy_shuo_ming

    @property
    def biao_tou(self):
        sy_biao_tou = self._wb.add_format()
        sy_biao_tou.set_font_name('微软雅黑')
        sy_biao_tou.set_font_size(11)
        sy_biao_tou.set_bold(True)
        sy_biao_tou.set_align('center')
        sy_biao_tou.set_align('vcenter')
        sy_biao_tou.set_text_wrap(True)
        sy_biao_tou.set_border(style=1)
        return sy_biao_tou

    @property
    def wen_zi(self):
        sy_wen_zi = self._wb.add_format()
        sy_wen_zi.set_font_name('微软雅黑')
        sy_wen_zi.set_font_size(11)
        sy_wen_zi.set_align('center')
        sy_wen_zi.set_align('vcenter')
        sy_wen_zi.set_border(style=1)
        return sy_wen_zi

    @property
    def shu_zi(self):
        sy_shu_zi = self._wb.add_format()
        sy_shu_zi.set_font_name('微软雅黑')
        sy_shu_zi.set_font_size(11)
        sy_shu_zi.set_align('center')
        sy_shu_zi.set_align('vcenter')
        sy_shu_zi.set_num_format('0.00')
        sy_shu_zi.set_border(style=1)
        return sy_shu_zi

    @property
    def jin_du(self):
        sy_jin_du = self._wb.add_format()
        sy_jin_du.set_font_name('微软雅黑')
        sy_jin_du.set_font_size(11)
        sy_jin_du.set_align('center')
        sy_jin_du.set_align('vcenter')
        sy_jin_du.set_num_format('0.00%')
        sy_jin_du.set_border(style=1)
        return sy_jin_du


def kun_ming_che(wb, ws, ri_qi):
    nrow = 0
    ncol = 0

    sy = Style(wb)

    ws.merge_range(nrow, ncol, nrow, ncol+8,
                   data='开门红昆明机构车险统计表',
                   cell_format=sy.biao_ti)

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 8,
                   data=f'数据统计范围：2020-01-01 至 {ri_qi}',
                   cell_format=sy.shuo_ming)

    nrow += 1

    biao_tou = ['机构', '一月\n任务', '二月\n任务', '三月\n任务', '一季度\n任务',
                '当月保费', '一季度\n保费', '季度任务\n时间进度', '季度\n任务达成率']
    ws.write_row(nrow, ncol,
                 data=biao_tou,
                 cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ['百大国际', '春怡雅苑', '香榭丽园', '宜良',
                   '东川', '安宁', '春之城', '分公司本部']

    for ji_gou in ji_gou_list:
        info = KMH_TJ(ji_gou, '机构', '车险', '险种')
        ws.write(nrow, ncol, info.ji_gou, sy.wen_zi)
        ws.write(nrow, ncol + 1, info.ren_wu('一月任务'), sy.wen_zi)
        ws.write(nrow, ncol + 2, info.ren_wu('二月任务'), sy.wen_zi)
        ws.write(nrow, ncol + 3, info.ren_wu('三月任务'), sy.wen_zi)
        ws.write(nrow, ncol + 4, info.ren_wu('一季度任务'), sy.wen_zi)
        ws.write(nrow, ncol + 5, info.yue_bao_fei, sy.shu_zi)
        ws.write(nrow, ncol + 6, info.nian_bao_fei, sy.shu_zi)
        ws.write(nrow, ncol + 7, info.shi_jian_da_cheng('一季度任务'), sy.jin_du)
        ws.write(nrow, ncol + 8, info.ren_wu_jin_du('一季度任务'), sy.jin_du)

        nrow += 1

    info = KMH_TJ('昆明', '中心支公司', '车险', '险种')
    ws.write(nrow, ncol, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 1, info.ren_wu('一月任务'), sy.wen_zi)
    ws.write(nrow, ncol + 2, info.ren_wu('二月任务'), sy.wen_zi)
    ws.write(nrow, ncol + 3, info.ren_wu('三月任务'), sy.wen_zi)
    ws.write(nrow, ncol + 4, info.ren_wu('一季度任务'), sy.wen_zi)
    ws.write(nrow, ncol + 5, info.yue_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 6, info.nian_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 7, info.shi_jian_da_cheng('一季度任务'), sy.jin_du)
    ws.write(nrow, ncol + 8, info.ren_wu_jin_du('一季度任务'), sy.jin_du)

    ws.set_column(ncol, ncol, width=10)
    ws.set_column(ncol + 1, ncol + 6, width=8)
    ws.set_column(ncol + 7, ncol + 8, width=12)


def gui_mo_da_cheng(wb, ws, ri_qi):

    nrow = 0
    ncol = 0

    sy = Style(wb)

    ws.merge_range(nrow, ncol, nrow, ncol + 4,
                   data='开门红非车险规模达成奖统计表',
                   cell_format=sy.biao_ti)

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 4,
                   data=f'数据统计范围：2020-01-01 至 {ri_qi}',
                   cell_format=sy.shuo_ming)

    nrow += 1

    biao_tou = ['组别', '机构', '任务目标', '季度保费', '任务进度']

    ws.write_row(nrow, ncol,
                 data=biao_tou,
                 cell_format=sy.biao_tou)

    nrow += 1

    A_zu = ['昆明', '曲靖', '昭通', '保山']
    B_zu = ['文山', '版纳', '大理', '怒江']

    A_zu_info = []
    B_zu_info = []

    for ji_gou in A_zu:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        A_zu_info.append((info.ji_gou, info.ren_wu_jin_du('一季度任务')))

    for ji_gou in B_zu:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        B_zu_info.append((info.ji_gou, info.ren_wu_jin_du('一季度任务')))

    A_zu_sort = sorted(A_zu_info, key=lambda k: k[1], reverse=True)
    B_zu_sort = sorted(B_zu_info, key=lambda k: k[1], reverse=True)

    ji_gou_list = []
    for A in A_zu_sort:
        ji_gou_list.append(A[0])
    for B in B_zu_sort:
        ji_gou_list.append(B[0])

    for ji_gou in ji_gou_list:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        ws.write(nrow, ncol + 1, info.ji_gou, sy.wen_zi)
        ws.write(nrow, ncol + 2, info.ren_wu('一季度任务'), sy.wen_zi)
        ws.write(nrow, ncol + 3, info.nian_bao_fei, sy.shu_zi)
        ws.write(nrow, ncol + 4, info.ren_wu_jin_du('一季度任务'), sy.jin_du)

        nrow += 1

    info = KMH_TJ('分公司整体', '分公司', '非车险', '险种')
    ws.merge_range(nrow, ncol, nrow, ncol + 1, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 2, info.ren_wu('一季度任务'), sy.wen_zi)
    ws.write(nrow, ncol + 3, info.nian_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 4, info.ren_wu_jin_du('一季度任务'), sy.jin_du)

    ws.merge_range(3, 0, 6, 0, 'A组', sy.wen_zi)
    ws.merge_range(7, 0, 10, 0, 'B组', sy.wen_zi)

    ws.set_column(ncol, ncol, width=5)
    ws.set_column(ncol + 1, ncol + 5, width=12)


if __name__ == '__main__':

    conn = sqlite3.connect(r"Data\data.db")
    cur = conn.cursor()

    str_sql = f"SELECT MAX([投保确认日期]) \
                FROM [2020年]"

    cur.execute(str_sql)
    ri_qi = cur.fetchone()[0]

    wb = xlsxwriter.Workbook('2020年开门红竞赛统计表.xlsx')
    ws = wb.add_worksheet('昆明机构车险统计表')

    kun_ming_che(wb, ws, ri_qi)

    ws = wb.add_worksheet('非车险规模达成奖统计表')

    gui_mo_da_cheng(wb, ws, ri_qi)

    wb.close()
