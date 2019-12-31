import xlwt
import logging

from datetime import date
from datetime import timedelta

from cell_style import cell_style
from stats import Stats


logging.disable(logging.NOTSET)
logging.basicConfig(level = logging.DEBUG, format = ' %(asctime)s | %(levelname)s | %(message)s' )

def zhong_zhi_week(sh, sql):

    # 写入表标题
    nrow = 0                                            # 行记录器，用以记录数据写入的行
    title_style = cell_style(height = 14)
    title_str = "2019年三级机构签单保费达成情况"
    sh.write_merge(nrow, nrow, 0, 8, title_str, title_style)

    # 写入表中数据的时间统计范围
    nrow += 1
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(nrow, nrow, 0, 8, date_str, date_style)

    # 写入表中数据的时间统计范围
    nrow += 1
    explanation_style = cell_style(height = 10)
    explanation_str = "说明：表格中任务达成情况“正数”表示已超额完成全年任务，“负数”表示全年任务的缺口"
    sh.write_merge(nrow, nrow, 0, 8, explanation_str, explanation_style)

    # 开始写入表头
    nrow += 1
    header_style = cell_style(bold = True, wrap = 1, borders = True)

    # 序号列列宽及标题
    sh.col(0).width = 256 * 4
    sh.write_merge(nrow, nrow, 0, 0, "序号", header_style)

    # 机构名称列列宽及标题
    sh.col(1).width = 256 * 14
    sh.write_merge(nrow, nrow, 1, 1, "机构名称", header_style)

    # 险种列列宽及标题
    sh.col(2).width = 256 * 9
    sh.row(nrow).write(2, "险种", header_style)

    # 计划任务列列宽及标题
    sh.col(3).width = 256 * 10
    sh.row(nrow).write(3, "计划任务", header_style)

    # 年度累计保费列列宽及标题
    sh.col(4).width = 256 * 12
    sh.row(nrow).write(4, "累计保费", header_style)

    # 同比增长率列列宽及标题
    sh.col(5).width = 256 * 12
    sh.row(nrow).write(5, "同比增长", header_style)

    # 时间进度达成率列列宽及标题
    sh.col(6).width = 256 * 13
    sh.row(nrow).write(6, "时间进度", header_style)

    # 计划任务达成率列列宽及标题
    sh.col(7).width = 256 * 12
    sh.row(nrow).write(7, "任务达成率", header_style)

    # 任务达成情况列列宽及标题
    sh.col(8).width = 256 * 11
    sh.row(nrow).write(8, "任务达成情况", header_style)

    # 设置中支名称列表及单元格样式
    zhong_zhi_name = ("昆明", "曲靖", "文山", "大理", "版纳", "保山", "昭通", "怒江", "分公司整体")
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    bold_task_style = cell_style(height = 12, bold=True, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    bold_num_style = cell_style(height = 12, bold=True, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')
    percent_style_red = cell_style(height=12, borders=True, font_color = 0x0A, num_format='0.00%')
    bold_percent_style_red = cell_style(height=12, bold=True, borders=True, font_color=0x0A, num_format='0.00%')
    bold_percent_style = cell_style(height=12, bold=True, borders=True, num_format='0.00%')

    pattern_gray = xlwt.Pattern()
    pattern_gray.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern_gray.pattern_fore_colour = 0x16

    pattern_default = xlwt.Pattern()
    pattern_default.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern_default.pattern_fore_colour = 64
    
    # 获取机构数据，并写入表中
    nrow += 1
    id = 1
    for name in zhong_zhi_name:
        che = Stats(name, "车险", sql, "中心支公司")
        ren = Stats(name, "人身险", sql, "中心支公司")
        cai = Stats(name, "财产险", sql, "中心支公司")
        zheng = Stats(name, "整体", sql, "中心支公司")
        datas = ((name,), (che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance)
                          , (ren.risk, ren.task, ren.this_year, ren.year_tong_bi, ren.time_progress, ren.task_progress, ren.task_balance)
                          , (cai.risk, cai.task, cai.this_year, cai.year_tong_bi, cai.time_progress, cai.task_progress, cai.task_balance)
                          , (zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance))

        if id % 2 == 1:
            task_style.pattern = pattern_gray
            bold_task_style.pattern = pattern_gray
            num_style.pattern = pattern_gray
            bold_num_style.pattern = pattern_gray
            percent_style.pattern = pattern_gray
            percent_style_red.pattern = pattern_gray
            bold_percent_style_red.pattern = pattern_gray
            bold_percent_style.pattern = pattern_gray
        else:
            task_style.pattern = pattern_default
            bold_task_style.pattern = pattern_default
            num_style.pattern = pattern_default
            bold_num_style.pattern = pattern_default
            percent_style.pattern = pattern_default
            percent_style_red.pattern = pattern_default
            bold_percent_style_red.pattern = pattern_default
            bold_percent_style.pattern = pattern_default

        ncol = 0
        for data in datas:
            for d in data:
                if ncol == 0:
                    sh.write_merge(nrow, nrow +3, ncol, ncol, id, task_style)
                    ncol += 1
                if ncol == 1:
                    sh.write_merge(nrow, nrow + 3, ncol, ncol, d, task_style)
                    ncol += 1
                elif ncol in (2, 3):
                    if (nrow +1) % 4 == 0:
                        sh.row(nrow).write(ncol, d, bold_task_style)
                    else:
                        sh.row(nrow).write(ncol, d, task_style)
                    ncol += 1
                elif ncol == 4:
                    if (nrow +1) % 4 == 0:
                        sh.row(nrow).write(ncol, d, bold_num_style)
                    else:
                        sh.row(nrow).write(ncol, d, num_style)
                    ncol += 1
                elif ncol in (5, 6):
                    if (nrow +1) % 4 == 0:
                        sh.row(nrow).write(ncol, d, bold_percent_style)
                    else:
                        sh.row(nrow).write(ncol, d, percent_style)
                    ncol += 1
                elif ncol == 7:
                    if (nrow +1) % 4 == 0:
                        if d >= 1:
                            sh.row(nrow).write(ncol, d, bold_percent_style_red)
                        else:
                            sh.row(nrow).write(ncol, d, bold_percent_style)
                    else:
                        if d >= 1:
                            sh.row(nrow).write(ncol, d, percent_style_red)
                        else:
                            sh.row(nrow).write(ncol, d, percent_style)
                    ncol += 1
                elif ncol == 8:
                    if (nrow +1) % 4 == 0:
                        sh.row(nrow).write(ncol, d, bold_num_style)
                    else:
                        sh.row(nrow).write(ncol, d, num_style)
                    ncol += 1
                else:
                    ncol = 2
                    nrow += 1
                    if (nrow +1) % 4 == 0:
                        sh.row(nrow).write(ncol, d, bold_task_style)
                    else:
                        sh.row(nrow).write(ncol, d, task_style)
                    ncol += 1
                    
        id += 1

        logging.debug("{0}信息写入完成".format(name))
        nrow += 1
