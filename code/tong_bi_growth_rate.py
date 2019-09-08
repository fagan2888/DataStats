import xlwt
import logging

from cell_style import cell_style
from my_date import MyDate
from stats import Stats


def tong_bi_growth_rate(sh, sql):
    """设置同比增长率统计表"""

    # 设置年度统计表标题
    nrow = 0
    title_str = "2019年 年度 同比增长率统计表"
    title_style = cell_style(height=14, bold=True)
    sh.write_merge(nrow, nrow, 0, 12, title_str, title_style)
    logging.debug("表标题写入完成")

    # 设置年度统计表统计时间范围
    nrow += 1
    idate = MyDate()
    date_str = "统计时间范围：2019-01-01 至 " + idate.end_date
    date_style = cell_style(height=10)
    sh.write_merge(nrow, nrow, 0, 12, date_str, date_style)
    logging.debug("统计时间范围写入完成")

    # 设置年度统计表第1行表头
    nrow += 1
    header_style = cell_style(bold=True, borders=True)
    
    sh.write_merge(nrow, nrow + 1, 0, 0, "机构名称", header_style)
    sh.col(0).width = 256 * 16

    sh.write_merge(nrow, nrow, 1, 3, "整体保费", header_style)
    for i in range(1, 4):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 4, 6, "车险保费", header_style)
    for i in range(4, 7):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 7, 9, "人身险保费", header_style)
    for i in range(7, 10):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 10, 12, "财产险保费", header_style)
    for i in range(10, 13):
        sh.col(i).width = 256 * 12

    logging.debug("第1行表头写入完成")

    # 设置年度统计表第2行表头
    nrow += 1
    header_str = ("2018年", "2019年", "增长率")
    header_style2 = cell_style(bold = True, borders = True)

    ncol = 1
    for i in range(4):
        for value in header_str:
            sh.row(nrow).write(ncol, value, header_style2)
            ncol += 1
    logging.debug("第2行表头写入完成")

    # 写入年度统计表统计数据
    nrow += 1
    names = ("昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "巧家", "分公司整体")
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')

    for name in names:
        zheng_ti = Stats(name, "整体", sql, "中心支公司")
        che = Stats(name, "车险", sql, "中心支公司")
        ren_shen = Stats(name, "人身险", sql, "中心支公司")
        cai_chan = Stats(name, "财产险", sql, "中心支公司")
        datas = (name, zheng_ti.last_year_limit, zheng_ti.this_year_limit, zheng_ti.year_tong_bi_limit, 
                          che.last_year_limit, che.this_year_limit, che.year_tong_bi_limit, 
                          ren_shen.last_year_limit, ren_shen.this_year_limit, ren_shen.year_tong_bi_limit, 
                          cai_chan.last_year_limit, cai_chan.this_year_limit, cai_chan.year_tong_bi_limit)
        ncol = 0
        for data in datas:
            if ncol % 3 == 0:
                data_style = percent_style
            else:
                data_style = num_style

            sh.row(nrow).write(ncol, data, data_style)
            ncol += 1

        logging.debug("{0}数据写入完成".format(name))
        nrow += 1
    logging.debug("年度同比增长率统计表写入完成\n")

    #********************开始写入月度统计表****************************
    # 设置月度统计表标题
    nrow += 1
    title_str = "2019年 月度 同比增长率统计表"
    title_style = cell_style(height=14, bold=True)
    sh.write_merge(nrow, nrow, 0, 12, title_str, title_style)
    logging.debug("表标题写入完成")

    # 设置月度统计表统计时间范围
    nrow += 1
    idate = MyDate()
    date_str = "统计时间范围：2019-{0}-01 至 {1}".format(idate.end_month, idate.end_date)
    date_style = cell_style(height=10)
    sh.write_merge(nrow, nrow, 0, 12, date_str, date_style)
    logging.debug("统计时间范围写入完成")

    # 设置月度统计表第1行表头
    nrow += 1
    header_style = cell_style(bold=True, borders=True)
    
    sh.write_merge(nrow, nrow + 1, 0, 0, "机构名称", header_style)
    sh.col(0).width = 256 * 16

    sh.write_merge(nrow, nrow, 1, 3, "整体保费", header_style)
    for i in range(1, 4):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 4, 6, "车险保费", header_style)
    for i in range(4, 7):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 7, 9, "人身险保费", header_style)
    for i in range(7, 10):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 10, 12, "财产险保费", header_style)
    for i in range(10, 13):
        sh.col(i).width = 256 * 12

    logging.debug("第1行表头写入完成")

    # 设置月度统计表第2行表头
    nrow += 1
    header_str = ("2018年", "2019年", "增长率")
    header_style2 = cell_style(bold = True, borders = True)

    ncol = 1
    for i in range(4):
        for value in header_str:
            sh.row(nrow).write(ncol, value, header_style2)
            ncol += 1
    logging.debug("第2行表头写入完成")

    # 写入月度统计表统计数据
    nrow += 1
    names = ("昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "巧家", "分公司整体")
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')

    for name in names:
        zheng_ti = Stats(name, "整体", sql, "中心支公司")
        che = Stats(name, "车险", sql, "中心支公司")
        ren_shen = Stats(name, "人身险", sql, "中心支公司")
        cai_chan = Stats(name, "财产险", sql, "中心支公司")
        datas = (name, zheng_ti.last_year_month_limit, zheng_ti.this_year_month_limit, zheng_ti.month_tong_bi_limit, 
                          che.last_year_month_limit, che.this_year_month_limit, che.month_tong_bi_limit, 
                          ren_shen.last_year_month_limit, ren_shen.this_year_month_limit, ren_shen.month_tong_bi_limit, 
                          cai_chan.last_year_month_limit, cai_chan.this_year_month_limit, cai_chan.month_tong_bi_limit)
        ncol = 0
        for data in datas:
            if ncol % 3 == 0:
                data_style = percent_style
            else:
                data_style = num_style

            sh.row(nrow).write(ncol, data, data_style)
            ncol += 1

        logging.debug("{0}数据写入完成".format(name))
        nrow += 1
    logging.debug("月度同比增长率统计表写入完成\n")

    #********************开始写入旬度统计表****************************
    # 设置旬度统计表标题
    nrow += 1
    title_str = "2019年 旬度 同比增长率统计表"
    title_style = cell_style(height=14, bold=True)
    sh.write_merge(nrow, nrow, 0, 12, title_str, title_style)
    logging.debug("表标题写入完成")

    # 设置旬度统计表统计时间范围
    nrow += 1
    idate = MyDate()
    date_str = "统计时间范围：{0} 至 {1}".format(idate.begin_date, idate.end_date)
    date_style = cell_style(height=10)
    sh.write_merge(nrow, nrow, 0, 12, date_str, date_style)
    logging.debug("统计时间范围写入完成")

    # 设置旬度统计表第1行表头
    nrow += 1
    header_style = cell_style(bold=True, borders=True)
    
    sh.write_merge(nrow, nrow + 1, 0, 0, "机构名称", header_style)
    sh.col(0).width = 256 * 16

    sh.write_merge(nrow, nrow, 1, 3, "整体保费", header_style)
    for i in range(1, 4):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 4, 6, "车险保费", header_style)
    for i in range(4, 7):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 7, 9, "人身险保费", header_style)
    for i in range(7, 10):
        sh.col(i).width = 256 * 12

    sh.write_merge(nrow, nrow, 10, 12, "财产险保费", header_style)
    for i in range(10, 13):
        sh.col(i).width = 256 * 12

    logging.debug("第1行表头写入完成")

    # 设置旬度统计表第2行表头
    nrow += 1
    header_str = ("2018年", "2019年", "增长率")
    header_style2 = cell_style(bold = True, borders = True)

    ncol = 1
    for i in range(4):
        for value in header_str:
            sh.row(nrow).write(ncol, value, header_style2)
            ncol += 1
    logging.debug("第2行表头写入完成")

    # 写入旬度统计表统计数据
    nrow += 1
    names = ("昆明", "曲靖", "文山", "大理", "保山", "版纳", "怒江", "巧家", "分公司整体")
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')

    for name in names:
        zheng_ti = Stats(name, "整体", sql, "中心支公司")
        che = Stats(name, "车险", sql, "中心支公司")
        ren_shen = Stats(name, "人身险", sql, "中心支公司")
        cai_chan = Stats(name, "财产险", sql, "中心支公司")
        datas = (name, zheng_ti.last_ten_days, zheng_ti.this_ten_days, zheng_ti.ten_days_tong_bi, 
                          che.last_ten_days, che.this_ten_days, che.ten_days_tong_bi, 
                          ren_shen.last_ten_days, ren_shen.this_ten_days, ren_shen.ten_days_tong_bi, 
                          cai_chan.last_ten_days, cai_chan.this_ten_days, cai_chan.ten_days_tong_bi)
        ncol = 0
        for data in datas:
            if ncol % 3 == 0:
                data_style = percent_style
            else:
                data_style = num_style

            sh.row(nrow).write(ncol, data, data_style)
            ncol += 1

        logging.debug("{0}数据写入完成".format(name))
        nrow += 1
    logging.debug("旬度同比增长率统计表写入完成\n")