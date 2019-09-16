# _*_ coding: utf-8 _*_

import xlwt
import logging
import os
import sys
import shutil
from openpyxl import Workbook

from datetime import datetime
from io import StringIO

dirname = os.getcwd()
code_path = dirname + r'\code'
sys.path.append(code_path)

from my_sql import MySql
from db_update import db_update
from kun_ming_week import kun_ming_week
from zhong_zhi_year import zhong_zhi_year
from ji_gou_year import ji_gou_year
from tong_bi_ten_days import tong_bi_ten_days
from seven_days import seven_days

logging.disable(logging.NOTSET)
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s | %(levelname)s | %(message)s')


def update_data(db_file):
    """
    更新数据库，数据库更新完毕后将原始数据文件进行备份
    同时避免多次无意义的更新数据库
    """
    table_name = "昆明机构周报.xlsx"
    if os.path.isfile(table_name):
        db_update(table_name, db_file)
        backup_path = dirname + r"\Backup"
        backup_name = backup_path + "\\" + table_name[:-5] + datetime.today().strftime("%Y%m%d") + ".xlsx"
        shutil.copyfile(table_name, backup_name)
        os.remove(table_name)

    memory_str_sql = db_file.iterdump()
    db_memory = MySql(":memory:")
    db_memory.executescript(memory_str_sql)
    db_file.close()

    logging.debug("数据库更新完成\n")
    return db_memory


def set_kun_ming_week(book, sql):
    """
     设置昆明机构周报工作表
    """
    # 添加昆明机构周报工作表
    sheet = book.create_sheet("昆明机构周报")

    logging.debug("开始写入昆明机构周报表")
    # 将数据写入昆明机构周报工作表
    kun_ming_week(sheet, sql, "车险", 1)
    kun_ming_week(sheet, sql, "人身险", 2)
    kun_ming_week(sheet, sql, "财产险", 3)
    logging.debug("昆明机构周报表写入完成\n")


def set_zhong_zhi_year(book, sql):
    """
    设置三级机构周报工作表
    """

    # 添加三级机构周报工作表
    sheet = book.add_sheet("三级机构周报")

    logging.debug("开始写入三级机构周报表")
    # 将数据写入三级机构周报表
    zhong_zhi_year(sheet, sql)
    logging.debug("三级机构周报表写入完成\n")


def set_ji_gou_year(book, sql):
    """
    设置四级机构周报工作表
    """

    # 添加四级机构周报工作表
    sheet = book.add_sheet("四级机构周报")

    logging.debug("开始写入四级机构周报表")
    # 将数据写入四级机构周报表
    ji_gou_year(sheet, sql)
    logging.debug("四级机构周报表写入完成\n")


def set_tong_bi(book, sql):
    """
    设置同比增长率统计表
    """

    # 添加同比增长率统计表
    sheet = book.add_sheet("同比增长率统计表")

    logging.debug("开始写入同比增长率统计表")
    tong_bi_ten_days(sheet, sql)
    logging.debug("同比增长率统计表写入完成\n")


def set_seven_days(book, sql):
    """
    设置7天连续保费
    """
    sheet = book.add_sheet("7天连续保费")

    logging.debug("开始写入7天连续保费统计表")
    seven_days(sheet, sql)
    logging.debug("7天连续保费统计表写入完成\n")


def main():
    logging.debug("……程序开始运行……")
    # 建立数据库连接
    db_file = MySql(r"Data\data.db")
    logging.debug("数据库连接成功")

    # 更新数据库
    sql = update_data(db_file)

    # 建立Excel工作簿
    book = Workbook()

    # 设置昆明机构周报工作表
    set_kun_ming_week(book, sql)

    ## 设置三级机构周报工作表
    # set_zhong_zhi_year(book, sql)

    ## 设置四级机构周报工作表
    # set_ji_gou_year(book, sql)

    ## 设置同比增长率统计表
    # set_tong_bi(book, sql)

    # 设置滚动7天数据统计表
    # set_seven_days(book, sql)

    # 保存数据至Excel工作表中
    book.save("数据统计表.xlsx")

    # 关闭数据库
    sql.close()
    logging.debug("数据库关闭")
    logging.debug("……程序运行结束……")


# ******启动方式判断******
if __name__ == '__main__':
    main()
