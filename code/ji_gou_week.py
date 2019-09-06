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
    sh.write_merge(0, 0, 0, 11, title_str, title_style)

    # 写入表中数据的时间统计范围
    date_style = cell_style(height = 10)
    date_str = "数据统计范围：{0}至{1}".format("2019-01-01", date.today()-timedelta(days = 1))
    sh.write_merge(1, 1, 0, 11, date_str, date_style)

    # 写入表标题部分
    # 四级机构签单保费达成情况表的表标题有两行，第一、第二列分别占两行，其余列存在父标题与子标题的关系
    # 以下部分写入表标题的第一行
    # 标题中“序号”需换行处理，wrap参数控制单元格样式为自动换行
    header_style = cell_style(bold = True, wrap = 1, borders = True)

    # 序号列列宽及标题
    sh.col(0).width = 256 * 4
    sh.write_merge(2, 3, 0, 0, "序号", header_style)

    # 机构名称列列宽及标题
    sh.col(1).width = 256 * 20
    sh.write_merge(2, 3, 1, 1, "机构名称", header_style)

    # 计划任务列列宽及第一行标题
    for i in range(2, 5):
        sh.col(i).width = 256 * 10
    sh.write_merge(2, 2, 2, 4, "计划任务", header_style)

    # 年度累计保费列宽及第一行标题
    for i in range(5, 8):
        sh.col(i).width = 256 * 12
    sh.write_merge(2, 2, 5, 7, "年度累计保费", header_style)

    # 时间进度达成率列宽及第一行标题
    for i in range(8, 11):
        sh.col(i).width = 256 * 12
    sh.write_merge(2, 2, 8, 10, "时间进度达成率", header_style)
    
    # 同比增长率列宽及第一行标题
    sh.col(11).width = 256 * 14
    sh.write_merge(2, 2, 11, 11, "同比增长率", header_style)

    # 以下部分写入表标题的第二行
    header_str = ("车险", "非车险", "整体")
    ncol = 2
    # 表标题第二行为险种名称循环，变量i控制循环次数
    i = 0

    while i < 3:
        for value in header_str:
            sh.row(3).write(ncol, value, header_style)
            ncol += 1
        i += 1
    
    # 表标题第二行的最后一列名为机构整体业务的同比增长率
    sh.row(3).write(ncol, "整体", header_style)

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
        fei_che = Stats(name, "非车险", sql)
        zheng_ti = Stats(name, "整体", sql)
        data = (name, che.task, fei_che.task, zheng_ti.task, che.this_year, fei_che.this_year, zheng_ti.this_year, che.time_progress_rate, fei_che.time_progress_rate, zheng_ti.time_progress_rate,zheng_ti.year_tong_bi)
        ji_gou_data.append(data)
        logging.debug("{0}信息统计完成".format(name))
    
    # 记录内部团队数据的列表
    team_data = []

    # 获取内部团队的各项数据
    for name in team_name:
        che = Stats(name, "车险", sql)
        fei_che = Stats(name, "非车险", sql)
        zheng_ti = Stats(name, "整体", sql)
        data = (name, che.task, fei_che.task, zheng_ti.task, che.this_year, fei_che.this_year, zheng_ti.this_year, che.time_progress_rate, fei_che.time_progress_rate, zheng_ti.time_progress_rate,zheng_ti.year_tong_bi)
        team_data.append(data)
        logging.debug("{0}信息统计完成".format(name))

    # 对机构、内部团队数据按年度整体保费按降序排序，四级机构与内部团队分别排序
    ji_gou_data_sort = sorted(ji_gou_data, key = lambda d : d[6], reverse = True)
    team_data_sort = sorted(team_data, key = lambda d : d[6], reverse = True)
    
    # 建立三种不同数据类型的数据显示样式
    task_style = cell_style(height = 12, borders = True, num_format = '0')
    num_style = cell_style(height = 12, borders = True, num_format = '0.00')
    percent_style = cell_style(height=12, borders=True, num_format='0.00%')

    nrow = 4

    for datas in ji_gou_data_sort:
        ncol = 0
        for value in datas:
            if ncol < 5:
                data_style = task_style
            elif ncol < 8:
                data_style = num_style
            else:
                data_style = percent_style
            
            if ncol == 0:
                sh.row(nrow).write(ncol, nrow - 3, data_style)
                ncol += 1

            sh.row(nrow).write(ncol, value, data_style)

            ncol += 1
        nrow += 1

    for datas in team_data_sort:
        ncol = 0
        for value in datas:
            if ncol < 5:
                data_style = task_style
            elif ncol < 8:
                data_style = num_style
            else:
                data_style = percent_style
            
            if ncol == 0:
                sh.row(nrow).write(ncol, nrow - 3, data_style)
                ncol += 1

            sh.row(nrow).write(ncol, value, data_style)
            ncol += 1
        nrow += 1

    name = "分公司整体"
    che = Stats(name, "车险", sql)
    fei_che = Stats(name, "非车险", sql)
    zheng_ti = Stats(name, "整体", sql)
    datas = (name, che.task, fei_che.task, zheng_ti.task, che.this_year, fei_che.this_year, zheng_ti.this_year, che.time_progress_rate, fei_che.time_progress_rate, zheng_ti.time_progress_rate,zheng_ti.year_tong_bi)

    ncol = 0
    for value in datas:
        if ncol < 5:
            data_style = task_style
        elif ncol < 8:
            data_style = num_style
        else:
            data_style = percent_style
            
        if ncol == 0:
            sh.row(nrow).write(ncol, nrow - 3, data_style)
            ncol += 1

        sh.row(nrow).write(ncol, value, data_style)
        ncol += 1