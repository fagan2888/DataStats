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
    title_style = cell_style(height = 14, bold = True)
    title_str = "2019年三级机构签单保费达成情况"
    sh.write_merge(0, 0, 0, 20, title_str, title_style)
    logging.debug("表头写入完成")

    # 写入数据统计的时间范围
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(1, 1, 0, 20, date_str, date_style)
    logging.debug("时间范围写入完成")

    # 写入表头
    header_style = cell_style(height=12, bold=True, borders=True)

    # 设置表头行高
    row_style = xlwt.easyxf("font:height {0}".format(20 * 16))
    sh.row(3).set_style(row_style)

    # 设置机构名称列表头和列宽
    sh.col(0).width = 256 * 13
    sh.write_merge(2, 3, 0, 0, "机构名称", header_style)

    # 设置计划任务列表头和列宽
    sh.write_merge(2, 2, 1, 4, "计划任务", header_style)
    for i in range(1, 5):
        sh.col(i).width = 256 * 8

    # 设置年度累计保费列表头和列宽
    sh.write_merge(2, 2, 5, 8, "年度累计保费", header_style)
    for i in range(5, 9):
        sh.col(i).width = 256 * 11
    
    # 设置时间进度达成率列表头和列宽
    sh.write_merge(2, 2, 9, 12, "时间进度达成率", header_style)
    for i in range(9, 13):
        sh.col(i).width = 256 * 11

    # 设置同比增长率列表头和列宽
    sh.write_merge(2, 2, 13, 16, "同比增长率", header_style)
    for i in range(13, 17):
        sh.col(i).width = 256 * 11

    # 设置计划任务达成率表头和列宽
    sh.write_merge(2, 2, 17, 20, "计划任务达成率", header_style)
    for i in range(17, 21):
        sh.col(i).width = 265 * 11

    # 写入第二行表头（险种信息）
    risk_list = ("车险", "财产险", "人身险", "整体")
    sh.row(4).set_style(row_style)
    i = 0
    while i < 5:
        ncol = 0
        while ncol < 4:
            c = ncol + 1 + i * 4
            sh.row(3).write(c, risk_list[ncol], header_style)
            ncol += 1
        i += 1

    # 设置中支名称列表及单元格样式
    zhong_zhi_name = ("昆明", "曲靖", "文山", "大理", "版纳", "保山", "巧家", "怒江", "分公司整体")
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')
    
    # 获取机构数据，并写入表中
    nrow = 4
    for name in zhong_zhi_name:
        che = Stats(name, "车险", sql, "中心支公司")
        cai_chan = Stats(name, "财产险", sql, "中心支公司")
        ren_shen = Stats(name, "人身险", sql, "中心支公司")
        zheng_ti = Stats(name, "整体", sql, "中心支公司")
        zhong_zhi_data = (name, che.task, cai_chan.task, ren_shen.task, zheng_ti.task, che.this_year, cai_chan.this_year, ren_shen.this_year, 
                          zheng_ti.this_year, che.time_progress_rate, cai_chan.time_progress_rate, ren_shen.time_progress_rate, 
                          zheng_ti.time_progress_rate, che.year_tong_bi, cai_chan.year_tong_bi, ren_shen.year_tong_bi, zheng_ti.year_tong_bi, 
                          che.planned_task_completion_rate, cai_chan.planned_task_completion_rate, ren_shen.planned_task_completion_rate, 
                          zheng_ti.planned_task_completion_rate)
        ncol = 0
        for value in zhong_zhi_data:
            if ncol < 5:
                data_style = task_style
            elif ncol < 9:
                data_style = num_style
            else:
                data_style = percent_style
            
            sh.row(nrow).write(ncol, value, data_style)
            sh.row(nrow).set_style(row_style) 
            ncol += 1

        logging.debug("{0}信息写入完成".format(name))
        nrow += 1
