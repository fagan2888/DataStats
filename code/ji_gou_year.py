import xlwt
import logging

from cell_style import cell_style
from datetime import date
from datetime import timedelta
from stats import Stats

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s | %(levelname)s | %(message)s')

def ji_gou_year(sh, sql):
    """写入四级机构签单保费达成情况表"""
    pattern_gray = xlwt.Pattern()
    pattern_gray.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern_gray.pattern_fore_colour = 0x16

    pattern_default = xlwt.Pattern()
    pattern_default.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern_default.pattern_fore_colour = 64
    # 写入表标题
    nrow = 0                                            # 行记录器，用以记录数据写入的行
    title_style = cell_style(height = 14)
    title_str = "2019年四级机构签单保费达成情况"
    sh.write_merge(nrow, nrow, 0, 8, title_str, title_style)

    # 写入表中数据的时间统计范围
    nrow += 1
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(nrow, nrow, 0, 8, date_str, date_style)

    # 写入表中数据的时间统计范围
    nrow += 1
    explanation_style = cell_style(height = 10)
    explanation_style.pattern = pattern_gray
    explanation_str = "说明：表格中任务达成情况“正数”表示已超额完成全年任务，“负数”表示全年任务的缺口"
    sh.write_merge(nrow, nrow, 0, 8, explanation_str, explanation_style)

    # 开始写入表头
    nrow += 1
    header_style = cell_style(bold = True, wrap = 1, borders = True)

    # 序号列列宽及标题
    sh.col(0).width = 256 * 4
    sh.write_merge(nrow, nrow, 0, 0, "序号", header_style)

    # 机构名称列列宽及标题
    sh.col(1).width = 256 * 11
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

    # 四级机构名称列表
    ji_gou_name = ("百大国际", "春怡雅苑", "香榭丽园", "宜良", "东川", "安宁", "春之城", "勐海", "勐腊", "师宗", "陆良", "宣威", "罗平", 
                          "会泽", "沾益", "丘北", "砚山", "富宁", "马关", "广南", "麻栗坡", "云龙", "宾川", "漾濞", "弥渡", "洱源", "祥云", "腾冲", "施甸", 
                          "兰坪", "巧家")
    
    # 内部团队名称列表，在数据排序时机构与内部团队分开排序
    team_name = ("保山隆阳区营业部", "版纳中支本部", "怒江中支本部", "文山营业一部", "文山营业二部", "曲靖中支本部", 
                        "曲靖营业一部", "大理中支本部", "分公司本部")

    # 记录机构数据的列表
    ji_gou_datas = []

    # 获取四级机构的各项数据
    for name in ji_gou_name:
        che = Stats(name, "车险", sql)
        fei = Stats(name, "非车险", sql)
        zheng = Stats(name, "整体", sql)
        data = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
                fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
                zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)
        ji_gou_datas.append(data)
        logging.debug("{0}信息统计完成".format(name))
    
    # 记录内部团队数据的列表
    team_datas = []

    # 获取内部团队的各项数据
    for name in team_name:
        che = Stats(name, "车险", sql)
        fei = Stats(name, "非车险", sql)
        zheng = Stats(name, "整体", sql)
        data = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
                fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
                zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)
        team_datas.append(data)
        logging.debug("{0}信息统计完成".format(name))

    # 对机构、内部团队数据按年度整体保费按降序排序，四级机构与内部团队分别排序
    ji_gou_datas_sort = sorted(ji_gou_datas, key = lambda d : d[17], reverse = True)
    team_datas_sort = sorted(team_datas, key = lambda d : d[17], reverse = True)

    # 建立三种不同数据类型的数据显示样式
    name_style = cell_style(height = 12, borders = True, wrap = 1)
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    bold_task_style = cell_style(height = 12, bold = True, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    bold_num_style = cell_style(height = 12, bold = True, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')
    bold_percent_style = cell_style(height=12, bold = True, borders=True, num_format='0.00%')
    percent_style_red = cell_style(height=12, font_color = 0x0A, borders=True, num_format='0.00%')
    bold_percent_style_red = cell_style(height=12, font_color = 0x0A, bold = True, borders=True, num_format='0.00%')

    # 从第4行开始写入数据
    nrow += 1
    id = 1

    # 将排序后的四级机构数据写入表中
    for data in ji_gou_datas_sort:
        ncol = 0

        if id % 2 == 1:
            name_style.pattern = pattern_gray
            task_style.pattern = pattern_gray
            bold_task_style.pattern = pattern_gray
            num_style.pattern = pattern_gray
            bold_num_style.pattern = pattern_gray
            percent_style.pattern = pattern_gray
            bold_percent_style.pattern = pattern_gray
            percent_style_red.pattern = pattern_gray
            bold_percent_style_red.pattern = pattern_gray
        else:
            name_style.pattern = pattern_default
            task_style.pattern = pattern_default
            bold_task_style.pattern = pattern_default
            num_style.pattern = pattern_default
            bold_num_style.pattern = pattern_default
            percent_style.pattern = pattern_default
            bold_percent_style.pattern = pattern_default
            percent_style_red.pattern = pattern_default
            bold_percent_style_red.pattern = pattern_default

        for key, value in enumerate(data):
            # 第一列写入序号
            if ncol == 0:                                                                                # 序号列
                sh.write_merge(nrow, nrow + 2, ncol, ncol, id, task_style)
                ncol += 1
            
            if key == 0:                                                                                  # 机构名称列
                sh.write_merge(nrow, nrow + 2, ncol, ncol, value, name_style)
            elif key in (1, 2):                                                                           # 车险险种列，计划任务列
                sh.row(nrow).write(ncol, value, task_style)
            elif key == 3:                                                                               # 车险累计保费列
                sh.row(nrow).write(ncol, value, num_style)
            elif key in (4, 5, 6):                                                                       # 车险同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 6 and value >=1:
                    sh.row(nrow).write(ncol, value, percent_style_red)
                else:
                    sh.row(nrow).write(ncol, value, percent_style)
            elif key == 7:                                                                              # 车险计划任务差额列
                 sh.row(nrow).write(ncol, value, num_style)
            elif key in (8, 9):                                                                          # 非车险险种列，计划任务列
                sh.row(nrow + 1).write(ncol - 7, value, task_style)
            elif key == 10:                                                                            # 非车险累计保费列
                sh.row(nrow + 1).write(ncol - 7, value, num_style)
            elif key in (11, 12, 13):                                                                # 非车险同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 13 and value >=1:
                    sh.row(nrow + 1).write(ncol - 7, value, percent_style_red)
                else:
                    sh.row(nrow + 1).write(ncol - 7, value, percent_style)
            elif key == 14:                                                                            # 非车险计划任务差额列
                 sh.row(nrow + 1).write(ncol - 7, value, num_style)
            elif key in (15, 16):                                                                      # 整体险种列，计划任务列
                sh.row(nrow + 2).write(ncol - 14, value, bold_task_style)
            elif key == 17:                                                                            # 整体累计保费列
                sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)
            elif key in (18, 19, 20):                                                                # 整体同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 20 and value >=1:
                    sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style_red)
                else:
                    sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style)
            elif key == 21:                                                                            # 整体计划任务差额列
                 sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)         

            ncol += 1
        nrow += 3
        id += 1

    # 将排序后的内部团队数据写入表中
    for data in team_datas_sort:
        ncol = 0

        if id % 2 == 1:
            name_style.pattern = pattern_gray
            task_style.pattern = pattern_gray
            bold_task_style.pattern = pattern_gray
            num_style.pattern = pattern_gray
            bold_num_style.pattern = pattern_gray
            percent_style.pattern = pattern_gray
            bold_percent_style.pattern = pattern_gray
            percent_style_red.pattern = pattern_gray
            bold_percent_style_red.pattern = pattern_gray
        else:
            name_style.pattern = pattern_default
            task_style.pattern = pattern_default
            bold_task_style.pattern = pattern_default
            num_style.pattern = pattern_default
            bold_num_style.pattern = pattern_default
            percent_style.pattern = pattern_default
            bold_percent_style.pattern = pattern_default
            percent_style_red.pattern = pattern_default
            bold_percent_style_red.pattern = pattern_default

        for key, value in enumerate(data):
            # 第一列写入序号
            if ncol == 0:                                                                                # 序号列
                sh.write_merge(nrow, nrow + 2, ncol, ncol, id, task_style)
                ncol += 1
            
            if key == 0:                                                                                  # 机构名称列
                sh.write_merge(nrow, nrow + 2, ncol, ncol, value, name_style)
            elif key in (1, 2):                                                                           # 车险险种列，计划任务列
                sh.row(nrow).write(ncol, value, task_style)
            elif key == 3:                                                                               # 车险累计保费列
                sh.row(nrow).write(ncol, value, num_style)
            elif key in (4, 5, 6):                                                                       # 车险同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 6 and value >=1:
                    sh.row(nrow).write(ncol, value, percent_style_red)
                else:
                    sh.row(nrow).write(ncol, value, percent_style)
            elif key == 7:                                                                              # 车险计划任务差额列
                 sh.row(nrow).write(ncol, value, num_style)
            elif key in (8, 9):                                                                          # 非车险险种列，计划任务列
                sh.row(nrow + 1).write(ncol - 7, value, task_style)
            elif key == 10:                                                                            # 非车险累计保费列
                sh.row(nrow + 1).write(ncol - 7, value, num_style)
            elif key in (11, 12, 13):                                                                # 非车险同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 13 and value >=1:
                    sh.row(nrow + 1).write(ncol - 7, value, percent_style_red)
                else:
                    sh.row(nrow + 1).write(ncol - 7, value, percent_style)
            elif key == 14:                                                                            # 非车险计划任务差额列
                 sh.row(nrow + 1).write(ncol - 7, value, num_style)
            elif key in (15, 16):                                                                      # 整体险种列，计划任务列
                sh.row(nrow + 2).write(ncol - 14, value, bold_task_style)
            elif key == 17:                                                                            # 整体累计保费列
                sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)
            elif key in (18, 19, 20):                                                                # 整体同比增长率列、时间进度达成率列、计划任务达成率列
                if key == 20 and value >=1:
                    sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style_red)
                else:
                    sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style)
            elif key == 21:                                                                            # 整体计划任务差额列
                 sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)         

            ncol += 1
        nrow += 3
        id += 1

    # 最后统计分公司整体保费
    name = "分公司整体"
    che = Stats(name, "车险", sql)
    fei = Stats(name, "非车险", sql)
    zheng = Stats(name, "整体", sql)
    datas = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
                fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
                zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)

     # 将分公司数据写入表中
    name_style.pattern = pattern_gray
    task_style.pattern = pattern_gray
    bold_task_style.pattern = pattern_gray
    num_style.pattern = pattern_gray
    bold_num_style.pattern = pattern_gray
    percent_style.pattern = pattern_gray
    bold_percent_style.pattern = pattern_gray
    percent_style_red.pattern = pattern_gray
    bold_percent_style_red.pattern = pattern_gray

    ncol = 0
    for key, value in enumerate(datas):
        # 第一列写入序号
        if ncol == 0:                                                                                # 序号列
            sh.write_merge(nrow, nrow + 2, ncol, ncol, id, task_style)
            ncol += 1
            
        if key == 0:                                                                                  # 机构名称列
            sh.write_merge(nrow, nrow + 2, ncol, ncol, value, name_style)
        elif key in (1, 2):                                                                           # 车险险种列，计划任务列
            sh.row(nrow).write(ncol, value, task_style)
        elif key == 3:                                                                               # 车险累计保费列
            sh.row(nrow).write(ncol, value, num_style)
        elif key in (4, 5, 6):                                                                       # 车险同比增长率列、时间进度达成率列、计划任务达成率列
            if key == 6 and value >=1:
                sh.row(nrow).write(ncol, value, percent_style_red)
            else:
                sh.row(nrow).write(ncol, value, percent_style)
        elif key == 7:                                                                              # 车险计划任务差额列
                sh.row(nrow).write(ncol, value, num_style)
        elif key in (8, 9):                                                                          # 非车险险种列，计划任务列
            sh.row(nrow + 1).write(ncol - 7, value, task_style)
        elif key == 10:                                                                            # 非车险累计保费列
            sh.row(nrow + 1).write(ncol - 7, value, num_style)
        elif key in (11, 12, 13):                                                                # 非车险同比增长率列、时间进度达成率列、计划任务达成率列
            if key == 13 and value >=1:
                sh.row(nrow + 1).write(ncol - 7, value, percent_style_red)
            else:
                sh.row(nrow + 1).write(ncol - 7, value, percent_style)
        elif key == 14:                                                                            # 非车险计划任务差额列
                sh.row(nrow + 1).write(ncol - 7, value, num_style)
        elif key in (15, 16):                                                                      # 整体险种列，计划任务列
            sh.row(nrow + 2).write(ncol - 14, value, bold_task_style)
        elif key == 17:                                                                            # 整体累计保费列
            sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)
        elif key in (18, 19, 20):                                                                # 整体同比增长率列、时间进度达成率列、计划任务达成率列
            if key == 20 and value >=1:
                sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style_red)
            else:
                sh.row(nrow + 2).write(ncol - 14, value, bold_percent_style)
        elif key == 21:                                                                            # 整体计划任务差额列
                sh.row(nrow + 2).write(ncol - 14, value, bold_num_style)         

        ncol += 1
