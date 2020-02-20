import logging
import xlsxwriter
from datetime import datetime

# from io import StringIO

# from ..style import Style
# from ..date import IDate
# from ..tong_ji import Tong_Ji

from stats import Stats
from excel_write_base import Excel_Write_Base

logging.disable(logging.DEBUG)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)
logging.basicConfig(
    level=logging.INFO, format=" %(asctime)s | %(levelname)s | %(message)s"
)


class Excel_Write_Day(Excel_Write_Base):
    def write_header(self, nrow: int = None):
        """
        写入统计表的表头
        """

        if nrow is None:
            nrow = self.nrow
            risk = self.risk
        else:
            risk = ""

        ncol = 0

        # 写入表头中首列的列名，首列列名采用上下单元格合并的方式
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow + 1,
            last_col=ncol,
            data=self.first_col_name,
            cell_format=self.style.string_bold_gray,
        )
        ncol += 1

        # 开始写入表头当前年的第一行信息
        self.ws.merge_range(
            first_row=nrow,
            first_col=ncol,
            last_row=nrow,
            last_col=ncol + len(self.first_year_header_2) - 1,
            data=self.first_year_header_1,
            cell_format=self.style.string_bold_orange,
        )

        # 写入表头当前年的第二行信息
        for row2_data in self.first_year_header_2:
            self.ws.write_string(
                row=nrow + 1,
                col=ncol,
                string=row2_data,
                cell_format=self.style.string_bold_orange,
            )
            ncol += 1

        # 表头第一行其他年份信息写入时的列数偏移量
        ncol_offset = len(self.other_year_header_2) - 1

        # 通过判断表头第一行中的数据来判断不同的单元格背景色，两种颜色交替出现
        for row1_data in self.other_year_header_1:
            i = self.other_year_header_1.index(row1_data)
            if i % 2 == 0:
                style = self.style.string_bold_green
            else:
                style = self.style.string_bold_orange

            # 写入表头其他年份的第一行信息
            self.ws.merge_range(
                first_row=nrow,
                first_col=ncol,
                last_row=nrow,
                last_col=ncol + ncol_offset,
                data=row1_data,
                cell_format=style,
            )

            # 写入表头其他年份的第二行数据
            for row2_data in self.other_year_header_2:
                self.ws.write_string(
                    row=nrow + 1, col=ncol, string=row2_data, cell_format=style
                )
                ncol += 1

        # 表头占两行
        self.nrow += 2

        logging.info(f"{self.name}{risk}数据表表头写入完成")

    def write_data(self):
        """
        写入指定统计表中的数据
        """

        self.set_menu(self.nrow, f"{self.risk}")

        # 获取机构数据统计的对象
        data = Stats(name=self.name, risk=self.risk)

        # 写入首列信息
        ncol = 0
        nrow_val = self.nrow
        for value in list(data.day_list):
            if nrow_val % 2 == 1:
                string = self.style.string
            else:
                string = self.style.string_gray
            self.ws.write_string(
                row=nrow_val, col=ncol, string=value, cell_format=string
            )
            nrow_val += 1
        ncol += 1

        # 写入表数据
        for year in self.year_list:
            nrow_val = self.nrow
            day_premium = data.day_premium(year=year)
            day_sum = data.day_sum(year=year)
            day_sum_yoy = data.day_sum_yoy(year=year)
            day_task_progress_rate = data.day_task_progress_rate()
            day_time_progress_rate = data.day_time_progress_rate()

            for key in list(data.day_list):
                ncol_val = ncol
                if nrow_val % 2 == 1:
                    number = self.style.number
                    percent = self.style.percent
                else:
                    number = self.style.number_gray
                    percent = self.style.percent_gray

                self.ws.write(nrow_val, ncol_val, day_premium[key], number)
                ncol_val += 1
                self.ws.write(nrow_val, ncol_val, day_sum[key], number)
                ncol_val += 1
                self.ws.write(nrow_val, ncol_val, day_sum_yoy[key], percent)

                # 如果年份是当前年份则需要再写入两列数据
                if year == self.first_year_header_1:
                    ncol_val += 1
                    self.ws.write(
                        nrow_val, ncol_val, day_task_progress_rate[key], percent
                    )
                    ncol_val += 1
                    self.ws.write(
                        nrow_val, ncol_val, day_time_progress_rate[key], percent
                    )
                    nrow_val += 1
                else:
                    nrow_val += 1

            if year == self.first_year_header_1:
                ncol += 5
            else:
                ncol += 3

        # 在下一个统计表之前增长一个空行
        self.nrow += 367

        logging.info(f"{self.name}{self.risk}数据表写入完成")
        logging.info("-" * 60)

    def set_column_width(self):
        """
        设置列宽
        """
        self.ws.set_column(first_col=self.ncol, last_col=self.ncol, width=10)
        self.ws.set_column(first_col=self.ncol + 1, last_col=self.ncol + 25, width=13)

    def set_conditional_format(self, col):
        """
        设置条件格式
        """

        self.ws.conditional_format(
            col,
            options={
                "type": "cell",
                "criteria": "<",
                "value": 1,
                "format": self.style.red,
            },
        )

    def make_form(self):
        """
        写入日数据的逻辑控制函数
        """

        logging.info("开始写入日数据统计表")

        day.set_table_name(table_name=f"{self.name}日数据统计表")
        day.set_risk_list(["整体", "车险", "人身险", "财产险", "非车险"])
        day.set_first_col_name("日期")
        day.set_first_year_header_2(["单日保费", "累计保费", "累计保费同比", "累计任务进度", "累计时间进度"])
        day.set_other_year_header_2(["单日保费", "累计保费", "累计保费同比"])

        # 写入数据统计表并记录快捷菜单栏不同快捷菜单的锚信息和现实文本信息
        for risk in self.risk_list:
            self.set_risk(risk)
            # 写入表标题
            self.write_title()
            # 写入表头
            self.write_header()
            # 写入表数据
            self.write_data()

        # 写入快捷菜单栏
        self.write_menu()
        self.freeze_panes()
        self.set_column_width()
        self.set_conditional_format("F10:F2000")

    def draw_chart(self):
        """
        绘制日数据的图标
        """

        self.set_table_name(table_name=f"{self.name}日数据统计图")

        column_chart = self.wb.add_chart({"type": "column"})
        column_chart.add_series(
            {
                "categories": [f"{self.name}日数据统计表", 9, 0, 374, 0],
                "values": [f"{self.name}日数据统计表", 9, 1, 374, 1],
                "name": "2020年",
                "overlap": -15,
            }
        )
        column_chart.add_series(
            {"values": [f"{self.name}日数据统计表", 9, 6, 374, 6], "name": "2019年"}
        )
        column_chart.add_series(
            {"values": [f"{self.name}日数据统计表", 9, 9, 374, 9], "name": "2018年"}
        )
        column_chart.add_series(
            {"values": [f"{self.name}日数据统计表", 9, 12, 374, 12], "name": "2017年"}
        )
        column_chart.add_series(
            {"values": [f"{self.name}日数据统计表", 9, 15, 374, 15], "name": "2016年"}
        )
        column_chart.set_title(
            {
                "name": f"{self.name}日保费数据统计图",
                "name_font": {"name": "微软雅黑", "size": 12},
                "layout": {"x": 0.025, "y": 0.01},
            }
        )
        column_chart.set_size({"width": 25000, "height": 550})
        column_chart.set_y_axis(
            {
                "name": "日保费",
                "min": 0,
                "max": 250,
                "name_font": {"name": "微软雅黑", "size": 10, "rotation": 270},
            }
        )
        column_chart.set_table(
            {"font": {"name": "Arial", "size": 9}, "show_keys": True}
        )
        column_chart.set_legend(
            {"font": {"name": "微软雅黑", "size": 10}, "position": "left"}
        )

        line_chart = self.wb.add_chart({"type": "line"})
        line_chart.add_series(
            {
                "categories": [f"{self.name}日数据统计表", 9, 0, 374, 0],
                "values": [f"{self.name}日数据统计表", 9, 3, 374, 3],
                "name": "同比增长率",
                'marker': {'type': 'circle'},
                "y2_axis": True,
            }
        )
        line_chart.add_series(
            {
                "values": [f"{self.name}日数据统计表", 9, 5, 374, 5],
                "name": "时间进度达成率",
                'marker': {'type': 'circle'},
                "y2_axis": True,
            }
        )
        column_chart.combine(line_chart)
        line_chart.set_y2_axis(
            {
                "name": "增长（达成）率",
                "min": -1,
                "max": 2,
                "name_font": {"name": "微软雅黑", "size": 10, "rotation": 270},
            }
        )

        self.ws.insert_chart("A7", column_chart)


if __name__ == "__main__":
    a = datetime.now()
    wb = xlsxwriter.Workbook(r"2020年机构数据统计详细报告.xlsx")
    ws = wb.add_worksheet("目录")

    day = Excel_Write_Day(wb=wb, name="分公司")
    day.make_form()
    day.draw_chart()

    wb.close()
    b = datetime.now()
    print(f"date {b-a=}")
