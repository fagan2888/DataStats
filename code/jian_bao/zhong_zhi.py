import logging

from ..style import Style
from ..date import IDate
from ..tong_ji import Tong_Ji


def zhong_zhi(wb, ws):
    '''
    编写分公司简报中三级机构统计表
    '''

    logging.debug('开始写入三级机构数据统计表')

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
                   data='2020年三级机构数据统计表',
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
    name_list = ['昆明', '曲靖', '文山', '大理', '保山', '版纳', '怒江', '昭通']

    datas = []

    # 对机构的年度整体保费进行统计，并在之后根据年保费进行排序
    for name in name_list:
        d = Tong_Ji(name=name, risk='整体',)

        datas.append((d.ming_cheng, d.nian_bao_fei()))

    # 根据 年保费按倒序进行排序
    datas.sort(key=lambda k: k[1], reverse=True)

    # 使用排序后的顺序重新填充机构名称列表
    name_list = []

    for d in datas:
        name_list.append(d[0])

    # 在新的机构名称列表之前添加分公司
    name_list.insert(0, '分公司整体')

    logging.debug('机构排序完成')

    # 设置险种名称列表
    risk_list = ['车险', '财产险', '人身险', '驾意险', '整体']

    # 将数据写入表中
    for name in name_list:
        # 根据机构名称设置机构类型
        if name == '分公司整体':
            xu_hao = ''  # 分公司不参与排名
            wen_zi = sy.string
            shu_zi = sy.number
            bai_fen_bi = sy.percent
        elif xu_hao == '':
            xu_hao = 1
            wen_zi = sy.string_gray
            shu_zi = sy.number_gray
            bai_fen_bi = sy.percent_gray
        # 根据序号设置单元格是否增加底色
        elif xu_hao % 2 == 0:
            xu_hao += 1
            wen_zi = sy.string_gray
            shu_zi = sy.number_gray
            bai_fen_bi = sy.percent_gray
        else:
            xu_hao += 1
            wen_zi = sy.string
            shu_zi = sy.number
            bai_fen_bi = sy.percent

        # 写入序号列，序号占5行
        ws.merge_range(first_row=nrow,
                       first_col=ncol,
                       last_row=nrow + 4,
                       last_col=ncol,
                       data=xu_hao,
                       cell_format=wen_zi)

        # 写入 机构名称列，名称占5行
        ws.merge_range(first_row=nrow,
                       first_col=ncol + 1,
                       last_row=nrow + 4,
                       last_col=ncol + 1,
                       data=name,
                       cell_format=wen_zi)

        # 根据险种名称 设置险种类型
        for risk in risk_list:
            if risk == '整体':
                if xu_hao == '' or xu_hao % 2 == 0:
                    wen_zi = sy.string_bold
                    shu_zi = sy.number_bold
                    bai_fen_bi = sy.percent_bold
                else:
                    wen_zi = sy.string_bold_gray
                    shu_zi = sy.number_bold_gray
                    bai_fen_bi = sy.percent_bold_gray

            d = Tong_Ji(name=name,
                        risk=risk)

            ws.write(nrow, ncol + 2, d.xian_zhong, wen_zi)
            ws.write(nrow, ncol + 3, d.ren_wu(), wen_zi)
            ws.write(nrow, ncol + 4, d.nian_bao_fei(), shu_zi)
            ws.write(nrow, ncol + 5, d.shi_jian_da_cheng(), bai_fen_bi)
            ws.write(nrow, ncol + 6, d.nian_tong_bi(), bai_fen_bi)
            nrow += 1
        logging.debug(f'{name}机构数据写入完成')

    # 开始设置列宽
    ncol = 0
    ws.set_column(first_col=ncol, last_col=ncol, width=4)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 1, width=12)
    ws.set_column(first_col=ncol + 2, last_col=ncol + 4, width=10)
    ws.set_column(first_col=ncol + 5, last_col=ncol + 6, width=12)

    logging.debug('列宽设置完成')
    logging.debug('三级机构数据统计表写入完成')
    logging.debug('-' * 60)
