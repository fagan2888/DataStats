from openpyxl import Workbook
from openpyxl.styles import (Alignment, Border, Color, Font, NamedStyle,
                             PatternFill, Side)


def style(wb):
    fill_color = 'E5E5E5'

    font_12_bold = Font(name='微软雅黑',
                        size=12,
                        bold=True)

    font_11 = Font(name='微软雅黑',
                        size=11)

    font_11_bold = Font(name='微软雅黑',
                        size=11,
                        bold=True)

    font_10 = Font(name='微软雅黑',
                        size=10)

    center = Alignment(horizontal='center',
                       vertical='center',
                       wrapText=True)

    border = Border(left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

    hui = PatternFill('solid', fgColor=fill_color)

    # 设置主标题样式
    biao_ti = NamedStyle(name='biao_ti')
    biao_ti.font = font_12_bold
    biao_ti.alignment = center

    # 设置小标题样式（列标题）
    xiao_biao_ti = NamedStyle(name='xiao_biao_ti')
    xiao_biao_ti.font = font_11_bold
    xiao_biao_ti.alignment = center
    xiao_biao_ti.border = border

    # 设置说明文字样式
    shuo_ming = NamedStyle(name='shuo_ming')
    shuo_ming.font = font_10
    shuo_ming.alignment = center

    # 设置数字格式样式 无背景色
    shu_zi = NamedStyle(name='shu_zi')
    shu_zi.font = font_11
    shu_zi.alignment = center
    shu_zi.border = border
    shu_zi.number_format = '0.00'

    # 设置数字格式样式 加粗 无背景色
    shu_zi_cu = NamedStyle(name='shu_zi_cu')
    shu_zi_cu.font = font_11_bold
    shu_zi_cu.alignment = center
    shu_zi_cu.border = border
    shu_zi_cu.number_format = '0.00'

    # 设置数字格式样式 背景色填充
    shu_zi_hui = NamedStyle(name='shu_zi_hui')
    shu_zi_hui.font = font_11
    shu_zi_hui.alignment = center
    shu_zi_hui.border = border
    shu_zi_hui.fill = hui
    shu_zi_hui.number_format = '0.00'

    # 设置数字格式样式 加粗 背景色填充
    shu_zi_cu_hui = NamedStyle(name='shu_zi_cu_hui')
    shu_zi_cu_hui.font = font_11_bold
    shu_zi_cu_hui.alignment = center
    shu_zi_cu_hui.border = border
    shu_zi_cu_hui.fill = hui
    shu_zi_cu_hui.number_format = '0.00'

    # 设置百分比格式样式
    bai_fen_bi = NamedStyle(name='bai_fen_bi')
    bai_fen_bi.font = font_11
    bai_fen_bi.alignment = center
    bai_fen_bi.border = border
    bai_fen_bi.number_format = '0.00%'

    # 设置百分比格式样式 加粗
    bai_fen_bi_cu = NamedStyle(name='bai_fen_bi_cu')
    bai_fen_bi_cu.font = font_11_bold
    bai_fen_bi_cu.alignment = center
    bai_fen_bi_cu.border = border
    bai_fen_bi_cu.number_format = '0.00%'

    # 设置百分比格式样式 背景色
    bai_fen_bi_hui = NamedStyle(name='bai_fen_bi_hui')
    bai_fen_bi_hui.font = font_11
    bai_fen_bi_hui.alignment = center
    bai_fen_bi_hui.border = border
    bai_fen_bi_hui.fill = hui
    bai_fen_bi_hui.number_format = '0.00%'

    # 设置百分比格式样式 加粗 背景色
    bai_fen_bi_cu_hui = NamedStyle(name='bai_fen_bi_cu_hui')
    bai_fen_bi_cu_hui.font = font_11_bold
    bai_fen_bi_cu_hui.alignment = center
    bai_fen_bi_cu_hui.border = border
    bai_fen_bi_cu_hui.fill = hui
    bai_fen_bi_cu_hui.number_format = '0.00%'

    # 设置文字格式样式
    wen_zi = NamedStyle(name='wen_zi')
    wen_zi.font = font_11
    wen_zi.alignment = center
    wen_zi.border = border

    # 设置文字格式样式 加粗
    wen_zi_cu = NamedStyle(name='wen_zi_cu')
    wen_zi_cu.font = font_11_bold
    wen_zi_cu.alignment = center
    wen_zi_cu.border = border

    # 设置文字格式样式 背景色
    wen_zi_hui = NamedStyle(name='wen_zi_hui')
    wen_zi_hui.font = font_11
    wen_zi_hui.alignment = center
    wen_zi_hui.border = border
    wen_zi_hui.fill = hui

    # 设置文字格式样式 背景色 加粗
    wen_zi_cu_hui = NamedStyle(name='wen_zi_cu_hui')
    wen_zi_cu_hui.font = font_11_bold
    wen_zi_cu_hui.alignment = center
    wen_zi_cu_hui.border = border
    wen_zi_cu_hui.fill = hui

    wb.add_named_style(biao_ti)
    wb.add_named_style(xiao_biao_ti)
    wb.add_named_style(shuo_ming)
    wb.add_named_style(shu_zi)
    wb.add_named_style(shu_zi_hui)
    wb.add_named_style(shu_zi_cu)
    wb.add_named_style(shu_zi_cu_hui)
    wb.add_named_style(bai_fen_bi)
    wb.add_named_style(bai_fen_bi_hui)
    wb.add_named_style(bai_fen_bi_cu)
    wb.add_named_style(bai_fen_bi_cu_hui)
    wb.add_named_style(wen_zi)
    wb.add_named_style(wen_zi_hui)
    wb.add_named_style(wen_zi_cu)
    wb.add_named_style(wen_zi_cu_hui)
