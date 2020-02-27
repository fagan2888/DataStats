import xlsxwriter


class Style:
    """
    工作簿的单元格样式类
    """

    def __init__(self, wb: xlsxwriter.Workbook):
        """
        初始化工作簿的单元格样式类

            参数：
                wb：
                    xlsxwriter.Workbook，对象的工作薄对象
        """
        self._wb = wb

    @property
    def number_format(self):
        """数字单元格格式"""
        return "?,??0.00;[红色]-?,??0.00;\"-\";"

    @property
    def percent_format(self):
        """"""
        return "0.00%;[红色]-0.00%;\"-\";"

    @property
    def black(self):
        """
        返回一个字体颜色的对象，黑色
        """
        value = self.wb.add_format()
        value.set_color(self.black_code)
        value.set_bold(True)
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size + 1)

        return value

    @property
    def black_code(self):
        """
        返回黑色的#RRGGBB编码
        """
        return "#000000"

    @property
    def deep_sky_blue(self):
        """
        返回一个字体颜色的对象，深天蓝
        """
        value = self.wb.add_format()
        value.set_color(self.deep_sky_blue_code)
        value.set_bold(True)
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size + 1)

        return value

    @property
    def deep_sky_blue_code(self):
        """
        返回深天蓝的#RRGGBB编码
        """
        return "#00BFFF"

    @property
    def red(self):
        """
        返回一个字体颜色的对象，红色
        """
        value = self.wb.add_format()
        value.set_color(self.red_code)

        return value

    @property
    def red_code(self):
        """
        返回红色的#RRGGBB编码
        """
        return "#FF0000"

    @property
    def green(self):
        """
        返回一个字体颜色的对象，绿色
        """
        value = self.wb.add_format()
        value.set_color(self.green_code)

        return value

    @property
    def green_code(self):
        """
        返回绿色的#RRGGBB编码
        """
        return "#92D050"

    @property
    def orange(self):
        """
        返回一个字体颜色的对象，橙色
        """
        value = self.wb.add_format()
        value.set_color(self.orange_code)

        return value

    @property
    def orange_code(self):
        """
        返回橙色的#RRGGBB编码
        """
        return "#FFC000"

    @property
    def gray(self):
        """
        返回一个字体颜色的对象，灰色
        """
        value = self.wb.add_format()
        value.set_color(self.gray_code)

        return value

    @property
    def gray_code(self):
        """
        返回灰色的#RRGGBB编码
        """
        return "#D9D9D9"

    @property
    def blue(self):
        """
        返回一个字体颜色的对象，蓝色
        """
        value = self.wb.add_format()
        value.set_color(self.blue_code)

        return value

    @property
    def blue_code(self):
        """
        返回蓝色的#RRGGBB编码
        """
        return "#0070C0"

    @property
    def menu(self):
        """
        返回快捷菜单里文字样式

        微软雅黑，加粗，下划线，灰色背景，蓝色字体，居中对齐
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_underline(True)
        value.set_bg_color(self.gray_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_border(style=1)
        value.set_color(self.blue_code)

        return value

    @property
    def font_size(self):
        """
        返回字体的基准大小
        """
        return 11

    @property
    def wb(self):
        """
        返回工作簿对象
        """
        return self._wb

    @property
    def title(self):
        """
        返回表标题样式

        微软雅黑，字体加大一号，加粗，居中对齐
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size + 1)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        return value

    @property
    def explain(self):
        """
        返回说明性文字样式

        微软雅黑，字体减少两号，居中对齐
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size - 2)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        return value

    @property
    def header(self):
        """
        返回表头部分文字样式

        微软雅黑，加粗，居中对齐，自动换行，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def string(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def number(self):
        """
        返回表格中的数字样式

        微软雅黑，居中对齐，采用0.00格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.number_format)
        value.set_border(style=1)
        return value

    @property
    def percent(self):
        """
        返回表格中的百分比样式

        微软雅黑，居中对齐，采用0.00%格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.percent_format)
        value.set_border(style=1)
        return value

    @property
    def string_bold(self):
        """
        返回表格中的文字样式

        微软雅黑，加粗，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def number_bold(self):
        """
        返回表格中的数字样式

        微软雅黑，加粗，居中对齐，采用0.00格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.number_format)
        value.set_border(style=1)
        return value

    @property
    def percent_bold(self):
        """
        返回表格中的百分比样式

        微软雅黑，加粗，居中对齐，采用0.00%格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.percent_format)
        value.set_border(style=1)
        return value

    @property
    def string_gray(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def string_orange(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.orange_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def string_bold_orange(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold()
        value.set_bg_color(self.orange_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def string_bold_green(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bold()
        value.set_bg_color(self.green_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def number_gray(self):
        """
        返回表格中的数字样式

        微软雅黑，居中对齐，采用0.00格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.number_format)
        value.set_border(style=1)
        return value

    @property
    def percent_gray(self):
        """
        返回表格中的百分比样式

        微软雅黑，居中对齐，采用0.00%格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.percent_format)
        value.set_border(style=1)
        return value

    @property
    def string_bold_gray(self):
        """
        返回表格中的文字样式

        微软雅黑，居中对齐，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def number_bold_gray(self):
        """
        返回表格中的数字样式

        微软雅黑，居中对齐，采用0.00格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.number_format)
        value.set_border(style=1)
        return value

    @property
    def percent_bold_gray(self):
        """
        返回表格中的百分比样式

        微软雅黑，居中对齐，采用0.00%格式，边框画线
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_bg_color(self.gray_code)
        value.set_bold(True)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_num_format(self.percent_format)
        value.set_border(style=1)
        return value

    @property
    def footnote(self):
        """
        返回脚注文字样式

        微软雅黑，字体减少两号，居中对齐
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        return value

    @property
    def remarks(self):
        """
        返回备注文字样式

        微软雅黑，字体减少两号，居中对齐
        """
        value = self.wb.add_format()
        value.set_font_name("微软雅黑")
        value.set_font_size(self.font_size)
        value.set_align("center")
        value.set_align("vcenter")
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value


if __name__ == "__main__":
    wb = xlsxwriter.Workbook("样式测试.xlsx")
    sy = Style(wb)
