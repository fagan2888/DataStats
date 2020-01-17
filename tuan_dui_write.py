import logging
import sqlite3

from tuan_dui_tong_ji import Tong_ji


def tuan_dui_write(ws):

    logging.debug('开始写入内部团队数据')

    conn = sqlite3.connect(r"Data\data.db")
    cur = conn.cursor()

    str_sql = "SELECT MAX([投保确认日期]) \
               FROM   [2020年]"
    cur.execute(str_sql)
    ri_qi = cur.fetchone()[0]

    month = ri_qi[5:7]
    day = ri_qi[8:10]

    row = ['2020年内部团队数据统计表']
    ws.append(row)
    logging.debug('标题行插入完成')

    row = [f'数据统计范围：2020-01-01至2020-{month}-{day}']
    ws.append(row)
    logging.debug('说明行插入完成')

    row = ['序号', '机构', '险种', '计划任务', '累计保费',
           '时间进度\n达成率', '同比增长率']
    ws.append(row)
    logging.debug('列标题行插入完成')

    ming_cheng = ['分公司整体', '分公司本部', '分公司销售二部',
                  '曲靖中支本部', '曲靖中支销售部', '曲靖中支销售一部',
                  '曲靖中支销售二部',
                  '文山中支本部', '文山中支营业一部销售一部',
                  '文山中支营业二部销售一部',
                  '大理中支本部', '大理中支销售一部', '大理中支销售二部',
                  '版纳中支本部', '版纳中支销售一部',
                  '保山中支本部', '保山中支销售一部', '保山中支销售二部',
                  '怒江中支本部', '怒江销售一部']

    xian_zhong = ['车险', '非车险', '驾意险', '整体']
    xu_hao = ''

    for name in ming_cheng:
        for risk in xian_zhong:
            tong_ji = Tong_ji(name, risk)
            row = [xu_hao,
                   tong_ji.jian_cheng,
                   tong_ji.xian_zhong,
                   tong_ji.ren_wu,
                   tong_ji.nian_bao_fei,
                   tong_ji.shi_jian_da_cheng_lv,
                   tong_ji.yi_nian_tong_bi]

            ws.append(row)

        if xu_hao == '':
            xu_hao = 1
        else:
            xu_hao += 1

        logging.debug(f'{name}数据写入完成')

    logging.debug('全部团队数据写入完成')

    # 合并 标题行
    ws.merge_cells(start_row=1,
                   start_column=1,
                   end_row=1,
                   end_column=ws.max_column)

    # 合并 说明行
    ws.merge_cells(start_row=2,
                   start_column=1,
                   end_row=2,
                   end_column=ws.max_column)

    # 合并 序号列和机构名称列
    for i in range(4, ws.max_row, 4):
        # 合并 序号列
        ws.merge_cells(start_row=i,
                       start_column=1,
                       end_row=i+3,
                       end_column=1)
        # 合并 机构名称列
        ws.merge_cells(start_row=i,
                       start_column=2,
                       end_row=i+3,
                       end_column=2)

    logging.debug('单元格合并完成')

    nrow = 1    # 行号计数变量
    fill = False    # 判断是否有背景色

    for row in ws.rows:
        col = 1     # 每行开始重置列数为 1
        for r in row:
            if nrow == 1:
                r.style = 'biao_ti'
            elif nrow == 2:
                r.style = 'shuo_ming'
            elif nrow == 3:
                r.style = 'xiao_biao_ti'
            elif nrow % 4 == 3:     # 但行号除以 4余 3时则为整体保费行，加粗处理
                if fill is True:
                    if col <= 4:
                        r.style = 'wen_zi_cu_hui'
                    elif col <= 5:
                        r.style = 'shu_zi_cu_hui'
                    elif col <= 7:
                        r.style = 'bai_fen_bi_cu_hui'
                else:
                    if col <= 4:
                        r.style = 'wen_zi_cu'
                    elif col <= 5:
                        r.style = 'shu_zi_cu'
                    elif col <= 7:
                        r.style = 'bai_fen_bi_cu'
            else:
                if fill is True:
                    if col <= 4:
                        r.style = 'wen_zi_hui'
                    elif col <= 5:
                        r.style = 'shu_zi_hui'
                    elif col <= 7:
                        r.style = 'bai_fen_bi_hui'
                else:
                    if col <= 4:
                        r.style = 'wen_zi'
                    elif col <= 5:
                        r.style = 'shu_zi'
                    elif col <= 7:
                        r.style = 'bai_fen_bi'

            col += 1
        nrow += 1

        fill_row = [8, 9, 10, 11,
                    16, 17, 18, 19,
                    32, 33, 34, 35,
                    44, 45, 46, 47,
                    56, 57, 58, 59,
                    64, 65, 66, 67,
                    76, 77, 78, 79]

        if nrow in fill_row:
            fill = True
        else:
            fill = False

    logging.debug('单元格样式写入完成')

    # 设置列宽
    ws.row_dimensions[1].height = 30
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 11
    ws.column_dimensions['C'].width = 9
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 12

    logging.debug('列宽调整完成')
    logging.debug('内部团队数据统计表写入完成')
    print("-" * 60)