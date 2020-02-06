import logging
import xlsxwriter
from ..style import Style
from ..date import IDate
from ..tong_ji import Tong_Ji


def tuan_dui(wb: xlsxwriter.Workbook,
             ws: xlsxwriter.worksheet) -> None:
    '''
    编写分公司简报中内部团队统计表
    '''

    logging.debug('开始写入内部团队数据统计表')

    # 获取单元格样式对象
    sy = Style(wb)

    # 获取日期对象
    idate = IDate(2020)

    # 设置行、列计数器
    nrow = 0
    ncol = 0

    # 写入表标题
    ws.merge_range(first_row=nrow,
                   first_col=ncol,
                   last_row=nrow,
                   last_col=ncol + 6,
                   data='2020年内部团队数据统计表',
                   cell_format=sy.title)

    # 设置表标题行高为字体的两倍
    ws.set_row(row=nrow, height=24)
    nrow += 1
    logging.debug('表标题写入完成')

    # 写入说明性文字，数据统计的时间范围
    ws.merge_range(first_row=nrow,
                   first_col=ncol,
                   last_row=nrow,
                   last_col=ncol + 6,
                   data=f'数据统计范围：2020-01-01 至 {idate.long_ri_qi()}',
                   cell_format=sy.explain)
    nrow += 1
    logging.debug('统计范围说明性文字写入完成')

    # 写入表头
    biao_ti = ['序号', '机构', '险种', '计划任务',
               '累计保费', '时间进度\n达成率', '同比\n增长率']

    for value in biao_ti:
        ws.write(nrow, ncol, value, sy.string_bold_gray)
        ncol += 1

    nrow += 1
    ncol = 0
    logging.debug('表头写入完成')

    # 设置机构名称列表
    ji_gou_list = (
        ('曲靖中支本部', '曲靖一部', '曲靖二部', '曲靖三部'),
        ('文山中支本部', '文山一部', '文山二部'),
        ('保山中支本部', '保山一部', '保山二部'),
        ('大理中支本部',),
        ('版纳中支本部',),
        ('怒江中支本部',)
    )

    # 设置险种名称列表
    risk_list = ('车险', '非车险', '驾意险', '整体')

    hui = False  # 机构名称改变的计数器
    xu_hao = 1  # 序号计数器

    # 将数据写入表中
    for name_list in ji_gou_list:
        if hui is False:
            wen_zi = sy.string
            shu_zi = sy.number
            bai_fen_bi = sy.percent
            hui = True
        else:
            wen_zi = sy.string_gray
            shu_zi = sy.number_gray
            bai_fen_bi = sy.percent_gray
            hui = False

        n = len(name_list) * 4 - 1
        ws.merge_range(nrow, ncol, nrow + n, ncol,
                       xu_hao, wen_zi)

        for name in name_list:
            if hui is False:
                wen_zi = sy.string_gray
                shu_zi = sy.number_gray
                bai_fen_bi = sy.percent_gray
            else:
                wen_zi = sy.string
                shu_zi = sy.number
                bai_fen_bi = sy.percent

            ws.merge_range(
                nrow, ncol + 1, nrow + 3, ncol + 1,
                name, wen_zi
            )

            for risk in risk_list:
                if risk == '整体':
                    if hui is False:
                        wen_zi = sy.string_bold_gray
                        shu_zi = sy.number_bold_gray
                        bai_fen_bi = sy.percent_bold_gray
                    else:
                        wen_zi = sy.string_bold
                        shu_zi = sy.number_bold
                        bai_fen_bi = sy.percent_bold
                else:
                    if hui is False:
                        wen_zi = sy.string_gray
                        shu_zi = sy.number_gray
                        bai_fen_bi = sy.percent_gray
                    else:
                        wen_zi = sy.string
                        shu_zi = sy.number
                        bai_fen_bi = sy.percent

                d = Tong_Ji(name=name, risk=risk)
                ws.write(nrow, ncol + 2, d.xian_zhong, wen_zi)
                ws.write(nrow, ncol + 3, d.ren_wu(), wen_zi)
                ws.write(nrow, ncol + 4, d.nian_bao_fei(), shu_zi)
                ws.write(nrow, ncol + 5, d.shi_jian_da_cheng(), bai_fen_bi)
                ws.write(nrow, ncol + 6, d.nian_tong_bi(), bai_fen_bi)
                nrow += 1
            logging.debug(f'{name}机构数据写入完成')
        xu_hao += 1

    risk_list = ('车险', '非车险', '驾意险')
    ws.merge_range(nrow, ncol, nrow + 4, ncol,
                   xu_hao, sy.string)

    ws.merge_range(nrow, ncol + 1, nrow + 4, ncol + 1,
                   '分公司本部', sy.string)

    for risk in risk_list:
        d = Tong_Ji(name='分公司本部', risk=risk)
        ws.write(nrow, ncol + 2, d.xian_zhong, sy.string)
        ws.write(nrow, ncol + 3, d.ren_wu(), sy.string)
        ws.write(nrow, ncol + 4, d.nian_bao_fei(), sy.number)
        ws.write(nrow, ncol + 5, d.shi_jian_da_cheng(), sy.percent)
        ws.write(nrow, ncol + 6, d.nian_tong_bi(), sy.percent)
        nrow += 1

    d = Tong_Ji(name='航旅项目', risk='整体')
    ws.write(nrow, ncol + 2, '航旅项目', sy.string)
    ws.write(nrow, ncol + 3, d.ren_wu(), sy.string)
    ws.write(nrow, ncol + 4, d.nian_bao_fei(), sy.number)
    ws.write(nrow, ncol + 5, d.shi_jian_da_cheng(), sy.percent)
    ws.write(nrow, ncol + 6, d.nian_tong_bi(), sy.percent)
    nrow += 1

    d = Tong_Ji(name='分公司本部', risk='整体')
    h = Tong_Ji(name='航旅项目', risk='整体')

    ren_wu = d.ren_wu() + h.ren_wu()
    nian_bao_fei = d.nian_bao_fei() + h.nian_bao_fei()
    shi_jian_da_cheng = nian_bao_fei / ren_wu / d.shi_jian_jin_du()
    wang_nian_bao_fei = d.nian_bao_fei(ny=1) + h.nian_bao_fei(ny=1)
    nian_tong_bi = nian_bao_fei / wang_nian_bao_fei - 1

    ws.write(nrow, ncol + 2, '整体', sy.string_bold)
    ws.write(nrow, ncol + 3, ren_wu, sy.string_bold)
    ws.write(nrow, ncol + 4, nian_bao_fei, sy.number_bold)
    ws.write(nrow, ncol + 5, shi_jian_da_cheng, sy.percent_bold)
    ws.write(nrow, ncol + 6, nian_tong_bi, sy.percent_bold)

    logging.debug(f'分公司本部机构数据写入完成')

    # 开始设置列宽
    ncol = 0
    ws.set_column(first_col=ncol, last_col=ncol, width=4)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 1, width=15)
    ws.set_column(first_col=ncol + 2, last_col=ncol + 4, width=10)
    ws.set_column(first_col=ncol + 5, last_col=ncol + 6, width=12)

    logging.debug('列宽设置完成')
    logging.debug('内部团队数据统计表写入完成')
    logging.debug('-' * 60)
