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
    sh.write_merge(0, 0, 0, 16, title_str, title_style)
    logging.debug("表头写入完成")

    # 写入数据统计的时间范围
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(1, 1, 0, 16, date_str, date_style)
    logging.debug("时间范围写入完成")


    header_style = cell_style(height=12, bold=True, borders=True)
    row_style = xlwt.easyxf("font:height {0}".format(20 * 16))

    sh.col(0).width = 256 * 14
    sh.row(3).set_style(row_style)
    sh.write_merge(2, 3, 0, 0, "机构名称", header_style)
    sh.write_merge(2, 2, 1, 4, "计划任务", header_style)
    sh.write_merge(2, 2, 5, 8, "年度累计保费", header_style)
    sh.write_merge(2, 2, 9, 12, "时间进度达成率", header_style)
    sh.write_merge(2, 2, 13, 16, "同比增长率", header_style)

    risk_list = ("车险", "财产险", "人身险", "整体")
    sh.row(4).set_style(row_style)
    i = 0
    while i < 4:
        ncol = 0
        while ncol < 4:
            c = ncol + 1 + i * 4
            sh.row(3).write(c, risk_list[ncol], header_style)
            ncol += 1
        i += 1

    zhong_zhi_name = ("昆明", "曲靖", "文山", "大理", "版纳", "保山", "巧家", "怒江", "分公司整体")
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')
    
    nrow = 4
    for name in zhong_zhi_name:
        che = Stats(name, "车险", sql, "中心支公司")
        cai_chan = Stats(name, "财产险", sql, "中心支公司")
        ren_shen = Stats(name, "人身险", sql, "中心支公司")
        zheng_ti = Stats(name, "整体", sql, "中心支公司")
        zhong_zhi_data = (name, che.task, cai_chan.task, ren_shen.task, zheng_ti.task, che.this_year,
            cai_chan.this_year, ren_shen.this_year, zheng_ti.this_year, che.time_progress_rate,
            cai_chan.time_progress_rate, ren_shen.time_progress_rate, zheng_ti.time_progress_rate, che.year_tong_bi, cai_chan.year_tong_bi, ren_shen.year_tong_bi, zheng_ti.year_tong_bi)
        ncol = 0
        for value in zhong_zhi_data:
            if ncol < 5:
                data_style = task_style
                if ncol > 0:
                    sh.col(ncol).width = 256 * 12
            elif ncol < 9:
                data_style = num_style
                sh.col(ncol).width = 256 * 12
            else:
                data_style = percent_style
                sh.col(ncol).width = 256 * 12
            
            sh.row(nrow).write(ncol, value, data_style)
            sh.row(nrow).set_style(row_style) 
            ncol += 1

        logging.debug("{0}信息写入完成".format(name))
        nrow += 1
