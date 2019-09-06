import xlrd
import os
import logging

logging.disable(logging.NOTSET)
logging.basicConfig(level = logging.DEBUG, format = ' %(asctime)s | %(levelname)s | %(message)s' )

def db_update(table_name, sql):
    """
    读取从MIS系统导出的Excel表中的内容
    清空数据库中今年的数据
    将今年的新数据写入数据库中
    """
       
    # 载入excel中的业务清单表,并对后期 sql语句数据插入做准备
    book = xlrd.open_workbook(table_name)
    sheet = book.sheet_by_name("page")
    # 获取投保确认时间
    tou_bao_date = sheet.col_values(0)
    # 获取险种
    risk = sheet.col_values(1)
    # 获取中心支公司
    zhong_zhi = sheet.col_values(2)
    # 获取机构
    ji_gou = sheet.col_values(3)
    # 获取团队
    tuan_dui = sheet.col_values(4)
    # 获取代理
    dai_li = sheet.col_values(5)
    # 获取客户源
    ke_hu_yuan = sheet.col_values(6)
    # 获取业务员
    ye_wu_yuan = sheet.col_values(7)
    # 获取保费
    bao_fei = sheet.col_values(8)
    logging.debug("Excel 文件读入成功")

    # 清空原数据库数据
    str_sql = "DELETE FROM [2019年]"
    sql.exec_non_query(str_sql)
    logging.debug("数据库数据清空完毕")

    # 将数据插入到sqlite数据库中
    i = 1

    while i < sheet.nrows:
        str_sql = ("INSERT INTO [2019年] VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8})".format(tou_bao_date[i], risk[i], zhong_zhi[i], ji_gou[i], 
            tuan_dui[i], dai_li[i], ke_hu_yuan[i], ye_wu_yuan[i], bao_fei[i]))
        #logging.debug("数据库数据插入语句：{0}".format(str_sql))
        sql.exec_non_query(str_sql)
        i += 1

    sql.commit()
    logging.debug("数据库数据插入完成")

