import sqlite3
from datetime import date
from datetime import timedelta

conn = sqlite3.connect(r'Data\data.db')
cur = conn.cursor()

start = '2019-01-01'

sql_ji_gou = "SELECT 机构简称 FROM 任务达成情况 WHERE 险种 = '车险'"
cur.execute(sql_ji_gou)
ji_gou = []
for value in cur.fetchall():
    ji_gou.append(value[0])
print(ji_gou)
