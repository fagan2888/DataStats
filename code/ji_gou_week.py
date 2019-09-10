import xlwt
import logging

from cell_style import cell_style
from datetime import date
from datetime import timedelta
from stats import Stats

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s | %(levelname)s | %(message)s')

def ji_gou_week(sh, sql):
    """写入四级机构签单保费达成情况表"""

    # 写入表标题
    title_style = cell_style(height = 14)
    title_str = "2019年四级机构签单保费达成情况"
    sh.write_merge(0, 0, 0, 8, title_str, title_style)

    # 写入表中数据的时间统计范围
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(1, 1, 0, 8, date_str, date_style)

    # 开始写入表头
    header_style = cell_style(bold = True, wrap = 1, borders = True)

    # 序号列列宽及标题
    sh.col(0).width = 256 * 4
    sh.write_merge(2, 2, 0, 0, "序号", header_style)

    # 机构名称列列宽及标题
    sh.col(1).width = 256 * 20
    sh.write_merge(2, 2, 1, 1, "机构名称", header_style)

    # 险种列列宽及标题
    sh.col(2).width = 256 * 6
    sh.row(2).write(2, "险种", header_style)

    # 计划任务列列宽及标题
    sh.col(3).width = 256 * 10
    sh.row(2).write(3, "计划任务", header_style)

    # 年度累计保费列列宽及标题
    sh.col(4).width = 256 * 10
    sh.row(2).write(4, "年度累计保费", header_style)

    # 同比增长率列列宽及标题
    sh.col(5).width = 256 * 10
    sh.row(2).write(5, "同比增长率", header_style)

    # 时间进度达成率列列宽及标题
    sh.col(6).width = 256 * 10
    sh.row(2).write(6, "时间进度达成率", header_style)

    # 计划任务达成率列列宽及标题
    sh.col(7).width = 256 * 10
    sh.row(2).write(7, "计划任务进度达成率", header_style)

    # 任务达成情况列列宽及标题
    sh.col(8).width = 256 * 10
    sh.row(2).write(8, "任务达成情况", header_style)

    # 四级机构名称列表
    ji_gou_name = ("百大国际", "春怡雅苑", "香榭丽园", "宜良", "东川", "安宁", "春之城", "勐海", "勐腊", "师宗", "陆良", "宣威", "罗平", 
                          "会泽", "沾益", "丘北", "砚山", "富宁", "马关", "广南", "麻栗坡", "云龙", "宾川", "漾濞", "弥渡", "洱源", "祥云", "腾冲", "施甸", 
                          "兰坪", "巧家")
    
    # 内部团队名称列表，在数据排序时机构与内部团队分开排序
    team_name = ("保山隆阳区营业部", "版纳中支本部", "怒江中支本部", "文山营业一部", "文山营业二部", "曲靖中支本部", 
                        "曲靖营业一部", "大理中支本部", "分公司本部")

    # 记录机构数据的列表
    ji_gou_data = []

    # 获取四级机构的各项数据
    for name in ji_gou_name:
        che = Stats(name, "车险", sql)
        fei = Stats(name, "非车险", sql)
        zheng = Stats(name, "整体", sql)
        data = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
                fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
                zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)
        ji_gou_data.append(data)
        logging.debug("{0}信息统计完成".format(name))
    
    # 记录内部团队数据的列表
    team_data = []

    # 获取内部团队的各项数据
    for name in team_name:
        che = Stats(name, "车险", sql)
        fei = Stats(name, "非车险", sql)
        zheng = Stats(name, "整体", sql)
        data = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
                fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
                zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)
        team_data.append(data)
        logging.debug("{0}信息统计完成".format(name))

    # 对机构、内部团队数据按年度整体保费按降序排序，四级机构与内部团队分别排序
    ji_gou_data_sort = sorted(ji_gou_data, key = lambda d : d[17], reverse = True)
    team_data_sort = sorted(team_data, key = lambda d : d[17], reverse = True)
    
    # 建立三种不同数据类型的数据显示样式
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')
    percent_style_red = cell_style(height=12, font_color = 0x0A, borders=True, num_format='0.00%')

    # 从第4行开始写入数据
    nrow = 3
    id = 1

    # 将排序后的四级机构数据写入表中
    for datas in ji_gou_data_sort:
        ncol = 0
        for k, v in enumerate(datas):
            if ncol < 4:
                data_style = task_style
            elif ncol < 5:
                data_style = num_style
            elif ncol < 8:
                data_style = percent_style
            else:
                data_style = num_style
            
            # 第一列写入序号
            if ncol == 0:
                sh.write_merge(nrow, nrow + 2, ncol, ncol, id, num_style)
                ncol += 1
            
            if k == 0:
                sh.write_merge(nrow, nrow + 2, ncol, ncol, k, num_style)
            elif k < 8:
                sh.row(nrow).write(ncol, k, data_style)
            elif k < 15:
                sh.row(nrow + 1).write(ncol - 7, k, data_style)
            else:
                sh.row(nrow + 2).write(ncol - 14, k, data_style)

            ncol += 1
        nrow += 3
        id += 1

    ## 将排序后的内部团队数据写入表中，行号延续四级机构写入完的行号
    #for datas in team_data_sort:
    #    ncol = 0
    #    for value in datas:
    #        if ncol < 5:
    #            data_style = task_style
    #        elif ncol < 8:
    #            data_style = num_style
    #        elif ncol < 11:
    #            data_style = percent_style
    #        elif ncol < 14:
    #            if value >= 1:
    #                data_style = percent_style_red
    #            else:
    #                data_style = percent_style
    #        else:
    #            data_style = percent_style
            
    #        if ncol == 0:
    #            sh.row(nrow).write(ncol, nrow - 3, data_style)
    #            ncol += 1

    #        sh.row(nrow).write(ncol, value, data_style)
    #        ncol += 1
    #    nrow += 1
    
    ## 最后统计分公司整体保费
    #name = "分公司整体"
    #che = Stats(name, "车险", sql)
    #fei = Stats(name, "非车险", sql)
    #zheng = Stats(name, "整体", sql)
    #data = (name, che.risk, che.task, che.this_year, che.year_tong_bi, che.time_progress, che.task_progress, che.task_balance,
    #            fei.risk, fei.task, fei.this_year, fei.year_tong_bi, fei.time_progress, fei.task_progress, fei.task_balance,
    #            zheng.risk, zheng.task, zheng.this_year, zheng.year_tong_bi, zheng.time_progress, zheng.task_progress, zheng.task_balance)

    #ncol = 0
    #for value in datas:
    #    if ncol < 5:
    #        data_style = task_style
    #    elif ncol < 8:
    #        data_style = num_style
    #    else:
    #        data_style = percent_style
            
    #    if ncol == 0:
    #        sh.row(nrow).write(ncol, nrow - 3, data_style)
    #        ncol += 1

    #    sh.row(nrow).write(ncol, value, data_style)
    #    ncol += 1