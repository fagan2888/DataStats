import sqlite3

import xlsxwriter

from KMH_TJ import KMH_TJ


class Style():
    '''
    工作簿的单元格样式类
    '''
    def __init__(self, wb):
        self._wb = wb

    @property
    def font_size(self):
        '''
        返回字体的基准大小
        '''
        return 11

    @property
    def wb(self):
        '''
        返回工作簿对象
        '''
        return self._wb

    @property
    def biao_ti(self):
        '''
        返回表标题样式
        微软雅黑，大一号字体，加粗，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size+1)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        return value

    @property
    def shuo_ming(self):
        '''
        返回说明性文字样式
        微软雅黑，小两号字体，居中对齐
        如：数据范围
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size-2)
        value.set_align('center')
        value.set_align('vcenter')
        return value

    @property
    def biao_tou(self):
        '''
        返回表头部分文字样式
        微软雅黑，加粗，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def wen_zi(self):
        '''
        返回表格中的文字样式
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_border(style=1)
        return value

    @property
    def shu_zi(self):
        '''
        返回表格中的数字样式
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00')
        value.set_border(style=1)
        return value

    @property
    def jin_du(self):
        '''
        返回表格中的百分比样式
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00%')
        value.set_border(style=1)
        return value


def kun_ming_che(sy, ws, ri_qi):
    '''
    昆明机构车险开门红统计表
    '''
    nrow = 0
    ncol = 0

    # 写入表标题
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


def gui_mo_da_cheng(sy, ws, ri_qi):
    '''
    非车险规模达成奖开门红统计表
    '''
    nrow = 0
    ncol = 0

    ws.merge_range(nrow, ncol, nrow, ncol + 4,
                   data='规模达成奖统计表',
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
        A_zu_info.append((info.ji_gou,
                          info.ren_wu('一季度任务'),
                          info.nian_bao_fei,
                          info.ren_wu_jin_du('一季度任务')))

    for ji_gou in B_zu:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        B_zu_info.append((info.ji_gou,
                          info.ren_wu('一季度任务'),
                          info.nian_bao_fei,
                          info.ren_wu_jin_du('一季度任务')))

    A_zu_sort = sorted(A_zu_info, key=lambda k: k[3], reverse=True)
    B_zu_sort = sorted(B_zu_info, key=lambda k: k[3], reverse=True)

    for A in A_zu_sort:
        ws.write(nrow, ncol + 1, A[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, A[1], sy.wen_zi)
        ws.write(nrow, ncol + 3, A[2], sy.shu_zi)
        ws.write(nrow, ncol + 4, A[3], sy.jin_du)
        nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 4, '', sy.wen_zi)
    ws.set_row(nrow, height=8)
    nrow += 1

    for B in B_zu_sort:
        ws.write(nrow, ncol + 1, B[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, B[1], sy.wen_zi)
        ws.write(nrow, ncol + 3, B[2], sy.shu_zi)
        ws.write(nrow, ncol + 4, B[3], sy.jin_du)
        nrow += 1

    info = KMH_TJ('分公司整体', '分公司', '非车险', '险种')
    ws.merge_range(nrow, ncol, nrow, ncol + 1, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 2, info.ren_wu('一季度任务'), sy.wen_zi)
    ws.write(nrow, ncol + 3, info.nian_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 4, info.ren_wu_jin_du('一季度任务'), sy.jin_du)

    ws.merge_range(3, 0, 6, 0, 'A组', sy.wen_zi)
    ws.merge_range(8, 0, 11, 0, 'B组', sy.wen_zi)

    ws.set_column(ncol, ncol, width=5)
    ws.set_column(ncol + 1, ncol + 5, width=11)


def ze_ren_xian(sy, ws, ri_qi):
    '''
    非车险责任险发展奖开门红统计表
    '''
    nrow = 0
    ncol = 0

    ws.merge_range(nrow, ncol, nrow, ncol + 2,
                   data='责任险发展奖统计表',
                   cell_format=sy.biao_ti)

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 2,
                   data=f'数据统计范围：2020-01-01 至 {ri_qi}',
                   cell_format=sy.shuo_ming)

    nrow += 1

    biao_tou = ['序号', '机构', '季度保费']

    ws.write_row(nrow, ncol,
                 data=biao_tou,
                 cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ['昆明', '曲靖', '昭通', '保山', '文山', '版纳', '大理', '怒江']

    data = []

    for ji_gou in ji_gou_list:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        data.append((info.ji_gou,
                     info.ze_ren_xian))
    data_sort = sorted(data, key=lambda k: k[1], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow-2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        nrow += 1

    info = KMH_TJ('分公司整体', '分公司', '非车险', '险种')
    ws.write(nrow, ncol, '', sy.wen_zi)
    ws.write(nrow, ncol + 1, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 2, info.ze_ren_xian, sy.shu_zi)

    ws.set_column(ncol, ncol, width=8)
    ws.set_column(ncol + 1, ncol + 1, width=12)
    ws.set_column(ncol + 2, ncol + 2, width=16)


def su_ze_xian(sy, ws, ri_qi):
    '''
    非车险诉讼保全激励奖开门红统计表
    '''
    nrow = 0
    ncol = 0

    ws.merge_range(nrow, ncol, nrow, ncol + 2,
                   data='诉讼保全激励奖统计表',
                   cell_format=sy.biao_ti)

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 2,
                   data=f'数据统计范围：2020-01-01 至 {ri_qi}',
                   cell_format=sy.shuo_ming)

    nrow += 1

    biao_tou = ['序号', '机构', '季度保费']

    ws.write_row(nrow, ncol,
                 data=biao_tou,
                 cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ['昆明', '曲靖', '昭通', '保山', '文山', '版纳', '大理', '怒江']

    data = []

    for ji_gou in ji_gou_list:
        info = KMH_TJ(ji_gou, '中心支公司', '非车险', '险种')
        data.append((info.ji_gou,
                     info.su_ze_xian))
    data_sort = sorted(data, key=lambda k: k[1], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow-2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        nrow += 1

    info = KMH_TJ('分公司整体', '分公司', '非车险', '险种')
    ws.write(nrow, ncol, '', sy.wen_zi)
    ws.write(nrow, ncol + 1, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 2, info.su_ze_xian, sy.shu_zi)

    ws.set_column(ncol, ncol, width=8)
    ws.set_column(ncol + 1, ncol + 1, width=12)
    ws.set_column(ncol + 2, ncol + 2, width=16)


def ji_gou_jia_yi_xian(sy, ws, ri_qi):
    '''
    非车险机构驾意险保费达成奖开门红统计表
    '''
    nrow = 0
    ncol = 0

    ws.merge_range(nrow, ncol, nrow, ncol + 6,
                   data='机构驾意险保费达成奖统计表',
                   cell_format=sy.biao_ti)

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(nrow, ncol, nrow, ncol + 6,
                   data=f'数据统计范围：2020-01-01 至 {ri_qi}',
                   cell_format=sy.shuo_ming)

    nrow += 1

    biao_tou = ['序号', '机构', '季度保费',
                '一阶段\n任务', '一阶段\n时间进度',
                '二阶段\n任务', '二阶段\n时间进度']

    ws.write_row(nrow, ncol,
                 data=biao_tou,
                 cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ['昆明', '曲靖', '昭通', '保山', '文山', '版纳', '大理', '怒江']

    data = []

    for ji_gou in ji_gou_list:
        info = KMH_TJ(ji_gou, '中心支公司', '驾意险', '险种名称')
        data.append((info.ji_gou,
                     info.nian_bao_fei,
                     info.ren_wu('一阶段任务'),
                     info.shi_jian_da_cheng('一阶段任务'),
                     info.ren_wu('二阶段任务'),
                     info.shi_jian_da_cheng('二阶段任务'),))
    data_sort = sorted(data, key=lambda k: k[5], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow-2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        ws.write(nrow, ncol + 3, d[2], sy.wen_zi)
        ws.write(nrow, ncol + 4, d[3], sy.jin_du)
        ws.write(nrow, ncol + 5, d[4], sy.wen_zi)
        ws.write(nrow, ncol + 6, d[5], sy.jin_du)
        nrow += 1

    info = KMH_TJ('分公司整体', '分公司', '驾意险', '险种名称')
    ws.write(nrow, ncol, '', sy.wen_zi)
    ws.write(nrow, ncol + 1, info.ji_gou, sy.wen_zi)
    ws.write(nrow, ncol + 2, info.nian_bao_fei, sy.shu_zi)
    ws.write(nrow, ncol + 3, info.ren_wu('一阶段任务'), sy.wen_zi)
    ws.write(nrow, ncol + 4, info.shi_jian_da_cheng('一阶段任务'), sy.jin_du)
    ws.write(nrow, ncol + 5, info.ren_wu('二阶段任务'), sy.wen_zi)
    ws.write(nrow, ncol + 6, info.shi_jian_da_cheng('二阶段任务'), sy.jin_du)

    ws.set_column(ncol, ncol, width=6)
    ws.set_column(ncol + 1, ncol + 2, width=12)
    ws.set_column(ncol + 3, ncol + 3, width=10)
    ws.set_column(ncol + 4, ncol + 4, width=12)
    ws.set_column(ncol + 5, ncol + 5, width=10)
    ws.set_column(ncol + 6, ncol + 6, width=12)


if __name__ == '__main__':

    conn = sqlite3.connect(r"Data\data.db")
    cur = conn.cursor()

    str_sql = f"SELECT MAX([投保确认日期]) \
                FROM [2020年]"
    cur.execute(str_sql)
    ri_qi = cur.fetchone()[0]

    wb = xlsxwriter.Workbook('2020年开门红竞赛统计表.xlsx')
    sy = Style(wb)

    # ws = wb.add_worksheet('昆明机构车险统计表')
    # kun_ming_che(sy, ws, ri_qi)

    # ws = wb.add_worksheet('规模达成奖统计表')
    # gui_mo_da_cheng(sy, ws, ri_qi)

    # ws = wb.add_worksheet('责任险成长奖统计表')
    # ze_ren_xian(sy, ws, ri_qi)

    # ws = wb.add_worksheet('诉讼保全激励奖统计表')
    # su_ze_xian(sy, ws, ri_qi)

    ws = wb.add_worksheet('机构驾意险保费达成奖统计表')
    ji_gou_jia_yi_xian(sy, ws, ri_qi)

    wb.close()
