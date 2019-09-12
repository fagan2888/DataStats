import sqlite3
import logging
from io import StringIO

logging.disable(logging.DEBUG)
logging.basicConfig(level = logging.DEBUG, format = ' %(asctime)s | %(levelname)s | %(message)s' )

class MySql():
    """
    建立一个数据库对象，并执行SQL语句，该对象为不同模块间传递数据库引用提供方便。

    参数：
        path：必须，str变量，表示数据库所在路径，需包含数据库文件名及文件后缀
    
    方法：
        exec_query()：执行SQL查询语句；
        exec_non_query()：执行SQL非查询语句；
        commit()：向数据库提交事务；
        closs()：关闭数据库。
    """
    def __init__(self, path):
        # 数据库连接
        self._conn = sqlite3.connect(path)
        logging.debug("数据库链接路径：{0}".format(path))
        logging.debug("数据库连接成功")
        # 数据库游标
        self._cur = self._conn.cursor()
        logging.debug("数据库游标创建成功")

    def exec_query(self, sql):
        """
        执行SQL查询语句，并返回单个查询结果

        参数：
            sql：必须，一个字符串，表示一个需要执行的SQL查询语句
                由于本函数仅返回查询结果的第一个数据，在设计SQL查询语句是应避免查询一个多数据的结果集

        返回值：
            res：返回执行SQL查询语句所得结果的第一个数据
        """
        self._cur.execute(sql)
        res = self._cur.fetchone()[0]
        return res

    def exec_non_query(self, sql):
        """
        执行一条非查询的SQL语句，执行后不提交事务，提交事务需执行commit()函数，无返回值

        参数：
            sql：必须，一个字符串，表示一个需要执行的非查询SQL语句 
        """
        self._cur.execute(sql)
    
    def commit(self):
        """
        向数据库提交事务，无参数，无返回值
        """
        self._conn.commit()
    
    def iterdump(self):
        """
        将数据库内容以文本的形式进行保存，方便与将数据在不同数据库见转移

        返回值：
            返回一个StringIO对象
        """
        memory_sql_str = StringIO()

        for line in self._conn.iterdump():
            memory_sql_str.write(f"{line}\n")
        return memory_sql_str

    def executescript(self, memory_sql_str):
        """
        批量执行SQL语句脚本
        """
        self._cur.executescript(memory_sql_str.getvalue())

    def close(self):
        """
        关闭数据库，无参数，无返回值
        """
        self._cur.close()
        logging.debug("数据库游标关闭")
        self._conn.close()
        logging.debug("数据库链接关闭")
