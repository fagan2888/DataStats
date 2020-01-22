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
                   cell_format=sy.biao_ti)

    # 设置表标题行高为字体的两倍
    ws.set_row(row=nrow, height=24)
    nrow += 1
    logging.debug('表标题写入完成')

    # 写入说明性文字，数据统计的时间范围
    ws.merge_range(first_row=nrow,
                   first_col=ncol,
                   last_row=nrow,
                   last_col=ncol + 6,
                   data=f'数据统计范围：2020-01-01 至 {idate.ri_qi()}',
                   cell_format=sy.shuo_ming)
    nrow += 1
    logging.debug('统计范围说明性文字写入完成')

    # 写入表头
    biao_ti = ['序号', '机构', '险种', '计划任务',
               '累计保费', '时间进度\n达成率', '同比\n增长率']

    for value in biao_ti:
        ws.write(nrow, ncol, value, sy.wen_zi_cu_hui)
        ncol += 1

    nrow += 1
    ncol = 0
    logging.debug('表头写入完成')

    # 设置机构名称列表
    name_list = [
        '分公司整体', '分公司本部', '分公司销售二部', '直拓业务团队', '航旅项目',
        '曲靖中支本部', '曲靖中支销售部', '曲靖中支销售一部',
        '曲靖中支销售二部',
        '文山中支本部', '文山中支营业一部销售一部',
        '文山中支营业二部销售一部',
        '大理中支本部', '大理中支销售一部', '大理中支销售二部',
        '版纳中支本部', '版纳中支销售一部',
        '保山中支本部', '保山中支销售一部', '保山中支销售二部',
        '怒江中支本部', '怒江中支销售一部'
    ]

    # 设置险种名称列表
    risk_list = ['车险', '非车险', '驾意险', '整体']

    # 将数据写入表中
    for name in name_list:

        # 根据机构名称设置机构类型
        if name == '分公司整体':
            xu_hao = ''  # 分公司不参与排名
            wen_zi_temp = sy.wen_zi
            shu_zi_temp = sy.shu_zi
            jin_du_temp = sy.jin_du
        elif xu_hao == '':
            xu_hao = 1
            wen_zi_temp = sy.wen_zi_hui
            shu_zi_temp = sy.shu_zi_hui
            jin_du_temp = sy.jin_du_hui
        # 根据序号设置单元格是否增加底色
        elif xu_hao % 2 == 0:
            xu_hao += 1
            wen_zi_temp = sy.wen_zi_hui
            shu_zi_temp = sy.shu_zi_hui
            jin_du_temp = sy.jin_du_hui
        else:
            xu_hao += 1
            wen_zi_temp = sy.wen_zi
            shu_zi_temp = sy.shu_zi
            jin_du_temp = sy.jin_du

        # 写入序号列，序号占4行
        ws.merge_range(first_row=nrow,
                       first_col=ncol,
                       last_row=nrow + 3,
                       last_col=ncol,
                       data=xu_hao,
                       cell_format=wen_zi_temp)

        # 写入 机构名称列，名称占4行
        ws.merge_range(first_row=nrow,
                       first_col=ncol + 1,
                       last_row=nrow + 3,
                       last_col=ncol + 1,
                       data=name,
                       cell_format=wen_zi_temp)

        # 根据险种名称 设置险种类型
        for risk in risk_list:
            if risk == '整体':
                if xu_hao == '' or xu_hao % 2 == 0:
                    wen_zi_temp = sy.wen_zi_cu
                    shu_zi_temp = sy.shu_zi_cu
                    jin_du_temp = sy.jin_du_cu
                else:
                    wen_zi_temp = sy.wen_zi_cu_hui
                    shu_zi_temp = sy.shu_zi_cu_hui
                    jin_du_temp = sy.jin_du_cu_hui

            d = Tong_Ji(name=name, risk=risk)

            ws.write(nrow, ncol + 2, d.xian_zhong, wen_zi_temp)
            ws.write(nrow, ncol + 3, d.ren_wu(), wen_zi_temp)
            ws.write(nrow, ncol + 4, d.nian_bao_fei(), shu_zi_temp)
            ws.write(nrow, ncol + 5, d.shi_jian_da_cheng, jin_du_temp)
            ws.write(nrow, ncol + 6, d.nian_tong_bi(ny=1), jin_du_temp)
            nrow += 1
        logging.debug(f'{name}机构数据写入完成')

    # 开始设置列宽
    ncol = 0
    ws.set_column(first_col=ncol, last_col=ncol, width=4)
    ws.set_column(first_col=ncol + 1, last_col=ncol + 1, width=16)
    ws.set_column(first_col=ncol + 2, last_col=ncol + 4, width=8)
    ws.set_column(first_col=ncol + 5, last_col=ncol + 6, width=10)

    logging.debug('列宽设置完成')
    logging.debug('内部团队数据统计表写入完成')
    logging.debug('*' * 60)
