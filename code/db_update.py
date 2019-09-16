import logging

from openpyxl import load_workbook

#logging.disable(logging.NOTSET)
#logging.basicConfig(level = logging.DEBUG, format = ' %(asctime)s | %(levelname)s | %(message)s' )

def db_update(table_name, sql):
    """
    读取从MIS系统导出的Excel表中的内容
    清空数据库中今年的数据
    将今年的新数据写入数据库中
    """

   # 载入excel中的业务清单表,并对后期 sql语句数据插入做准备
    book = load_workbook(table_name)
    sheet = book.get_sheet_by_name("page")
    logging.debug("Excel 文件载入完成")

    # 清空原数据库数据
    str_sql = "DELETE FROM [2019年]"
    sql.exec_non_query(str_sql)
    logging.debug("数据库数据清空完毕")

    # 将数据插入到sqlite数据库中
    for row in sheet.iter_rows(min_row=2, values_only=True):
        str_sql = f"INSERT INTO [2019年] VALUES ('{row[0]}', '{row[1]}', '{row[2]}', '{row[3]}', '{row[4]}', '{row[5]}', '{row[6]}', '{row[7]}', {row[8]})"
        sql.exec_non_query(str_sql)

    sql.commit()
    logging.debug("数据库数据插入完成")
