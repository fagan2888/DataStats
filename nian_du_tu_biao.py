import logging

import xlsxwriter

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji

from code.style import Style
from code.date import IDate
from code.tong_ji import Tong_Ji

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG,
    format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO,
    format=" %(asctime)s | %(levelname)s | %(message)s"
)

    # # 设置图表的高度和宽度
    # chart_height = 460
    # chart_width = 1000

    # 设置图表所占行数
    # chart_nrow = 22

    # 开始写入图表
    nrow += 1
    zheng_tong_column = wb.add_chart({"type": "column"})
    zheng_tong_column.add_series({
        "name": [t_name, zheng_nrow - 1, ncol + 1],
        "categories": [t_name, zheng_nrow, ncol, zheng_nrow + 4, ncol],
        "values": [t_name, zheng_nrow, ncol + 1, zheng_nrow + 4, ncol + 1],
    })

    zheng_tong_line = wb.add_chart({"type": "line"})
    zheng_tong_line.add_series({
        "name": [t_name, zheng_nrow - 1, ncol + 2],
        "categories": [t_name, zheng_nrow, ncol, zheng_nrow + 4, ncol],
        "values": [t_name, zheng_nrow, ncol + 2, zheng_nrow + 4, ncol + 2],
        "marker": {
            "type": "circle"
        },
        "y2_axis":
        True,
    })
    zheng_tong_column.combine(zheng_tong_line)

    zheng_tong_column.set_y_axis({
        "name": "同期保费",
        "name_font": {
            "name": "微软雅黑",
            "rotation": 270
        }
    })
    zheng_tong_line.set_y2_axis({
        "name": "同比增长率",
        "name_font": {
            "name": "微软雅黑",
            "rotation": 270
        }
    })
    zheng_tong_column.set_title({
        "name": "分公司历年保费同期数据对比图",
        "name_font": {
            "name": "微软雅黑"
        }
    })
    zheng_tong_column.set_table({"show_keys": True})
    zheng_tong_column.set_size({"width": chart_width, "height": chart_height})
    zheng_tong_column.set_legend({"none": True})

    ws.insert_chart(1, 0, zheng_tong_column)

    # 开始写入整体分公司历年保费全年对比图
    nrow += 24
    zheng_quan_column = wb.add_chart({"type": "column"})
    zheng_quan_column.add_series({
        "name": [t_name, zheng_nrow - 1, ncol + 3],
        "categories": [t_name, zheng_nrow, ncol, zheng_nrow + 4, ncol],
        "values": [t_name, zheng_nrow, ncol + 3, zheng_nrow + 4, ncol + 3],
    })

    zheng_quan_line = wb.add_chart({"type": "line"})
    zheng_quan_line.add_series({
        "name": [t_name, zheng_nrow - 1, ncol + 4],
        "categories": [t_name, zheng_nrow, ncol, zheng_nrow + 4, ncol],
        "values": [t_name, zheng_nrow, ncol + 4, zheng_nrow + 4, ncol + 4],
        "marker": {"type": "circle"},
        "y2_axis": True,
    })
    zheng_quan_column.combine(zheng_quan_line)

    zheng_quan_column.set_y_axis({
        "name": "全年保费",
        "name_font": {
            "name": "微软雅黑",
            "rotation": 270
        }
    })
    zheng_quan_line.set_y2_axis({
        "name": "同比增长率",
        "name_font": {
            "name": "微软雅黑",
            "rotation": 270
        }
    })
    zheng_quan_column.set_title({
        "name": "分公司整体业务历年保费全年数据对比图",
        "name_font": {
            "name": "微软雅黑"
        }
    })
    zheng_quan_column.set_table({"show_keys": True})
    zheng_quan_column.set_size({"width": chart_width, "height": chart_height})
    zheng_quan_column.set_legend({"none": True})

    ws.insert_chart(24, 0, zheng_quan_column)