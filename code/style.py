import xlsxwriter


class Style():
    '''
    工作簿的单元格样式类
    '''
    def __init__(self, wb):
        self._wb = wb

    @property
    def font_size(self):
        '''
        返回字体的基准大小
        '''
        return 11

    @property
    def wb(self):
        '''
        返回工作簿对象
        '''
        return self._wb

    @property
    def biao_ti(self):
        '''
        返回表标题样式
        微软雅黑，字体加大一号，加粗，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size+1)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        return value

    @property
    def shuo_ming(self):
        '''
        返回说明性文字样式
        微软雅黑，字体减少两号，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size-2)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        return value

    @property
    def biao_tou(self):
        '''
        返回表头部分文字样式
        微软雅黑，加粗，居中对齐，自动换行，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def wen_zi(self):
        '''
        返回表格中的文字样式
        微软雅黑，居中对齐，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def shu_zi(self):
        '''
        返回表格中的数字样式
        微软雅黑，居中对齐，采用0.00格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00')
        value.set_border(style=1)
        return value

    @property
    def jin_du(self):
        '''
        返回表格中的百分比样式
        微软雅黑，居中对齐，采用0.00%格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00%')
        value.set_border(style=1)
        return value

    @property
    def wen_zi_cu(self):
        '''
        返回表格中的文字样式
        微软雅黑，加粗，居中对齐，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def shu_zi_cu(self):
        '''
        返回表格中的数字样式
        微软雅黑，加粗，居中对齐，采用0.00格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00')
        value.set_border(style=1)
        return value

    @property
    def jin_du_cu(self):
        '''
        返回表格中的百分比样式
        微软雅黑，加粗，居中对齐，采用0.00%格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00%')
        value.set_border(style=1)
        return value

    @property
    def wen_zi_hui(self):
        '''
        返回表格中的文字样式
        微软雅黑，居中对齐，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def shu_zi_hui(self):
        '''
        返回表格中的数字样式
        微软雅黑，居中对齐，采用0.00格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00')
        value.set_border(style=1)
        return value

    @property
    def jin_du_hui(self):
        '''
        返回表格中的百分比样式
        微软雅黑，居中对齐，采用0.00%格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00%')
        value.set_border(style=1)
        return value

    @property
    def wen_zi_cu_hui(self):
        '''
        返回表格中的文字样式
        微软雅黑，居中对齐，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value

    @property
    def shu_zi_cu_hui(self):
        '''
        返回表格中的数字样式
        微软雅黑，居中对齐，采用0.00格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00')
        value.set_border(style=1)
        return value

    @property
    def jin_du_cu_hui(self):
        '''
        返回表格中的百分比样式
        微软雅黑，居中对齐，采用0.00%格式，边框画线
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_bg_color('#cccccc')
        value.set_bold(True)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_num_format('0.00%')
        value.set_border(style=1)
        return value

    @property
    def jiao_zhu(self):
        '''
        返回说明性文字样式
        微软雅黑，字体减少两号，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        return value

    @property
    def bei_zhu(self):
        '''
        返回说明性文字样式
        微软雅黑，字体减少两号，居中对齐
        '''
        value = self.wb.add_format()
        value.set_font_name('微软雅黑')
        value.set_font_size(self.font_size)
        value.set_align('center')
        value.set_align('vcenter')
        value.set_text_wrap(True)
        value.set_border(style=1)
        return value


if __name__ == '__main__':
    wb = xlsxwriter.Workbook('样式测试.xlsx')
    sy = Style(wb)
