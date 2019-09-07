import xlwt


def cell_style(height=12, bold=False, font_color = 0x7FF, borders=False, num_format='general', wrap=0):
    """
        设置单元格样式，包括字体、字号、加粗、对齐、边框、数字格式，可对字号、加粗、边框、数字格式进行自定义设置，默认为：
        字体：微软雅黑
        字号：12
        加粗：否
        水平对齐：居中对齐
        垂直对齐：居中对齐
        边框：四周均为实线边框
        数字格式：常规
    """
    style = xlwt.XFStyle()
    
    # 建立字体样式
    font = xlwt.Font()
    font.name = "微软雅黑"
    # 字体高度为字号的20倍整数表达，故此处字号值应为 20 × 字号
    font.height = 20 * height
    font.bold = bold
    font.colour_index = font_color
    
    # 设置数字格式
    if num_format !=  'general':
        style.num_format_str = num_format

    # 建立单元格的对齐方式
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    alignment.wrap = wrap
    
    #建立单元格的边框样式，THIN为实线
    _borders = xlwt.Borders()
    _borders.bottom = xlwt.Borders.THIN
    _borders.top = xlwt.Borders.THIN
    _borders.left = xlwt.Borders.THIN
    _borders.right = xlwt.Borders.THIN
    
    #将字体、对齐方式、边框加入到样式中
    style.font = font
    style.alignment = alignment
    if borders == True:
        style.borders = _borders

    return style

