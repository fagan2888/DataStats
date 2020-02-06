import sqlite3
import logging

import xlsxwriter

from code.kai_meng_hong.KMH_TJ import KMH_TJ
from code.style import Style
from temp_update import update


logging.disable(logging.NOTSET)
logging.basicConfig(
    level=logging.DEBUG, format=" %(asctime)s | %(levelname)s | %(message)s"
)


def kun_ming_che(sy, ws, ri_qi):
    """
    昆明机构车险开门红统计表
    """
    nrow = 0  # 行号计数器
    ncol = 0  # 列号计数器

    # 写入表标题
    ws.merge_range(
        nrow, ncol, nrow, ncol + 8, data="开门红昆明机构车险统计表", cell_format=sy.biao_ti
    )

    # 设置标题行的行高
    # 标题行行高设置为字号的两倍
    ws.set_row(nrow, height=24)

    # 每次完成一行的输入写入后，行号+1
    # 为写入下一行数据做准备
    nrow += 1

    # 写入说明性文字
    # 各统计表的说明性文字均使用数据统计的时间范围
    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 8,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    # 写入表头
    biao_tou = [
        "机构",
        "一月\n任务",
        "二月\n任务",
        "三月\n任务",
        "一季度\n任务",
        "当月保费",
        "一季度\n保费",
        "季度任务\n时间进度",
        "季度\n任务达成率",
    ]
    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    # 本表统计的机构列表
    ji_gou_list = ["百大国际", "春怡雅苑", "香榭丽园", "宜良", "东川", "安宁", "春之城", "分公司本部"]

    # 使用开门红统计类获取各机构对应的数据，并写入Excel表中
    for ji_gou in ji_gou_list:
        data = KMH_TJ(ji_gou, "车险")
        ws.write(nrow, ncol, data.ming_cheng, sy.wen_zi)
        ws.write(nrow, ncol + 1, data.ren_wu("一月任务"), sy.wen_zi)
        ws.write(nrow, ncol + 2, data.ren_wu("二月任务"), sy.wen_zi)
        ws.write(nrow, ncol + 3, data.ren_wu("三月任务"), sy.wen_zi)
        ws.write(nrow, ncol + 4, data.ren_wu("一季度任务"), sy.wen_zi)
        ws.write(nrow, ncol + 5, data.yue_bao_fei(), sy.shu_zi)
        ws.write(nrow, ncol + 6, data.nian_bao_fei(), sy.shu_zi)
        ws.write(
            nrow, ncol + 7, data.shi_jian_da_cheng("一季度任务"), sy.bai_fen_bi
        )
        ws.write(
            nrow, ncol + 8, data.ren_wu_jin_du("一季度任务"), sy.bai_fen_bi
        )

        nrow += 1

    # 写入昆明整体数据
    # 本表为统计四级机构数据，最后的合计行采用三级机构昆明的数据
    data = KMH_TJ("昆明", "车险")
    ws.write(nrow, ncol, data.ming_cheng, sy.wen_zi)
    ws.write(nrow, ncol + 1, data.ren_wu("一月任务"), sy.wen_zi)
    ws.write(nrow, ncol + 2, data.ren_wu("二月任务"), sy.wen_zi)
    ws.write(nrow, ncol + 3, data.ren_wu("三月任务"), sy.wen_zi)
    ws.write(nrow, ncol + 4, data.ren_wu("一季度任务"), sy.wen_zi)
    ws.write(nrow, ncol + 5, data.yue_bao_fei(), sy.shu_zi)
    ws.write(nrow, ncol + 6, data.nian_bao_fei(), sy.shu_zi)
    ws.write(nrow, ncol + 7, data.shi_jian_da_cheng("一季度任务"), sy.bai_fen_bi)
    ws.write(nrow, ncol + 8, data.ren_wu_jin_du("一季度任务"), sy.bai_fen_bi)

    # 设置列宽
    ws.set_column(ncol, ncol, width=10)
    ws.set_column(ncol + 1, ncol + 6, width=12)
    ws.set_column(ncol + 7, ncol + 8, width=12)


def gui_mo_da_cheng(sy, ws, ri_qi):
    """
    开门红非车险规模达成奖统计表
    """
    nrow = 0  # 行计数器
    ncol = 0  # 列计数器

    # 写入表标题
    ws.merge_range(
        nrow, ncol, nrow, ncol + 4, data="非车险规模达成奖统计表", cell_format=sy.biao_ti
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 4,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = ["组别", "机构", "任务目标", "季度保费", "任务进度"]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    # 统计航旅项目的数据
    # 昆明地区数据需要剔除航旅项目的数据，在此先进行获取
    han_lv = KMH_TJ("航旅项目", "非车险")

    # A组和B组的机构名称列表
    # 奖项激励分为A组和B组，分别进行排名，分别进行奖励
    A_zu = ["昆明", "曲靖", "昭通", "保山"]
    B_zu = ["文山", "版纳", "大理", "怒江"]

    # 分别用于记录两个组的数据，以便后面排序使用
    A_zu_data = []
    B_zu_data = []

    for ji_gou in A_zu:
        data = KMH_TJ(ji_gou, "非车险")
        if data.ming_cheng == "昆明":
            # 如果机构为昆明，需要剔除航旅项目的数据
            A_zu_data.append(
                (
                    data.ming_cheng,
                    data.ren_wu("一季度任务"),
                    data.nian_bao_fei() - han_lv.nian_bao_fei(),
                    data.ren_wu_jin_du("一季度任务"),
                )
            )
        else:
            A_zu_data.append(
                (
                    data.ming_cheng,
                    data.ren_wu("一季度任务"),
                    data.nian_bao_fei(),
                    data.ren_wu_jin_du("一季度任务"),
                )
            )

    for ji_gou in B_zu:
        data = KMH_TJ(ji_gou, "非车险")
        B_zu_data.append(
            (
                data.ming_cheng,
                data.ren_wu("一季度任务"),
                data.nian_bao_fei(),
                data.ren_wu_jin_du("一季度任务"),
            )
        )

    # 以列表中机构的时间进度达成率（第4个值）为依据对进行降序排列，
    A_zu_data.sort(key=lambda k: k[3], reverse=True)
    B_zu_data.sort(key=lambda k: k[3], reverse=True)

    for A in A_zu_data:
        ws.write(nrow, ncol + 1, A[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, A[1], sy.wen_zi)
        ws.write(nrow, ncol + 3, A[2], sy.shu_zi)
        ws.write(nrow, ncol + 4, A[3], sy.bai_fen_bi)
        nrow += 1

    # 在两组的信息之间键入一个空行，并缩小行高
    # 在视觉上更容易区分两个组的数据
    ws.merge_range(nrow, ncol, nrow, ncol + 4, "", sy.wen_zi)
    ws.set_row(nrow, height=8)
    nrow += 1

    for B in B_zu_data:
        ws.write(nrow, ncol + 1, B[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, B[1], sy.wen_zi)
        ws.write(nrow, ncol + 3, B[2], sy.shu_zi)
        ws.write(nrow, ncol + 4, B[3], sy.bai_fen_bi)
        nrow += 1

    data = KMH_TJ("分公司整体", "非车险")
    ws.merge_range(nrow, ncol, nrow, ncol + 1, data.ming_cheng, sy.wen_zi)
    ws.write(nrow, ncol + 2, data.ren_wu("一季度任务"), sy.wen_zi)
    ws.write(nrow, ncol + 3, data.nian_bao_fei(), sy.shu_zi)
    ws.write(nrow, ncol + 4, data.ren_wu_jin_du("一季度任务"), sy.bai_fen_bi)

    # 合并组别数据列，以便于在视觉上更容易区分两个组别的数据
    ws.merge_range(3, 0, 6, 0, "A组", sy.wen_zi)
    ws.merge_range(8, 0, 11, 0, "B组", sy.wen_zi)

    nrow += 2
    ncol = 0
    ws.merge_range(nrow, ncol, nrow, ncol + 4, "注：A组、B组各取前3名获奖", sy.jiao_zhu)

    ws.set_column(ncol, ncol, width=5)
    ws.set_column(ncol + 1, ncol + 5, width=11)


def ze_ren_xian(sy, ws, ri_qi):
    """
    开门红非车险责任险发展奖统计表
    """
    nrow = 0
    ncol = 0

    ws.merge_range(
        nrow, ncol, nrow, ncol + 3, data="非车险责任险发展奖统计表", cell_format=sy.biao_ti
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 3,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = ["序号", "机构", "保费", "备注"]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ["昆明", "曲靖", "昭通", "保山", "文山", "版纳", "大理", "怒江"]

    data = []

    for ji_gou in ji_gou_list:
        temp = KMH_TJ(ji_gou, "非车险")
        data.append((temp.ming_cheng, temp.ze_ren_xian))

    data_sort = sorted(data, key=lambda k: k[1], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow - 2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        nrow += 1

    temp = KMH_TJ("分公司整体", "非车险")
    ws.merge_range(nrow, ncol, nrow, ncol + 1, temp.ming_cheng, sy.wen_zi)
    ws.write(nrow, ncol + 2, temp.ze_ren_xian, sy.shu_zi)

    ws.merge_range(3, 3, 11, 3, "10万元入围\n季度前三名获奖", sy.bei_zhu)

    ws.set_column(ncol, ncol, width=8)
    ws.set_column(ncol + 1, ncol + 1, width=12)
    ws.set_column(ncol + 2, ncol + 2, width=16)
    ws.set_column(ncol + 3, ncol + 3, width=16)


def su_ze_xian(sy, ws, ri_qi):
    """
    开门红非车险诉讼保全激励奖统计表
    """
    nrow = 0
    ncol = 0

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 3,
        data="非车险诉讼保全激励奖统计表",
        cell_format=sy.biao_ti,
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 3,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = ["序号", "机构", "保费", "备注"]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ["昆明", "曲靖", "昭通", "保山", "文山", "版纳", "大理", "怒江"]

    data = []

    for ji_gou in ji_gou_list:
        temp = KMH_TJ(ji_gou, "非车险")
        data.append((temp.ming_cheng, temp.su_ze_xian))
    data_sort = sorted(data, key=lambda k: k[1], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow - 2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        nrow += 1

    temp = KMH_TJ("分公司整体", "非车险")
    ws.merge_range(nrow, ncol, nrow, ncol + 1, temp.ming_cheng, sy.wen_zi)
    ws.write(nrow, ncol + 2, temp.su_ze_xian, sy.shu_zi)
    ws.merge_range(3, 3, 11, 3, "5万元入围\n季度前三名获奖", sy.bei_zhu)

    ws.set_column(ncol, ncol, width=8)
    ws.set_column(ncol + 1, ncol + 1, width=12)
    ws.set_column(ncol + 2, ncol + 2, width=16)
    ws.set_column(ncol + 3, ncol + 3, width=16)


def ji_gou_jia_yi_xian(sy, ws, ri_qi):
    """
    开门红非车险驾意险机构保费达成奖统计表
    """
    nrow = 0
    ncol = 0

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 6,
        data="非车险机构驾意险保费达成奖统计表",
        cell_format=sy.biao_ti,
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 6,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = [
        "序号",
        "机构",
        "保费",
        "一阶段\n任务",
        "一阶段\n时间进度",
        "二阶段\n任务",
        "二阶段\n时间进度",
    ]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    ji_gou_list = ["昆明", "曲靖", "昭通", "保山", "文山", "版纳", "大理", "怒江"]

    data = []

    for ji_gou in ji_gou_list:
        temp = KMH_TJ(ji_gou, "驾意险")
        data.append(
            (
                temp.ming_cheng,
                temp.nian_bao_fei(),
                temp.ren_wu("一阶段任务"),
                temp.shi_jian_da_cheng("一阶段任务"),
                temp.ren_wu("二阶段任务"),
                temp.shi_jian_da_cheng("二阶段任务"),
            )
        )
    data_sort = sorted(data, key=lambda k: k[5], reverse=True)

    for d in data_sort:
        ws.write(nrow, ncol, nrow - 2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.shu_zi)
        ws.write(nrow, ncol + 3, d[2], sy.wen_zi)
        ws.write(nrow, ncol + 4, d[3], sy.bai_fen_bi)
        ws.write(nrow, ncol + 5, d[4], sy.wen_zi)
        ws.write(nrow, ncol + 6, d[5], sy.bai_fen_bi)
        nrow += 1

    temp = KMH_TJ("分公司整体", "驾意险")
    ws.merge_range(nrow, ncol, nrow, ncol + 1, temp.ming_cheng, sy.wen_zi)
    ws.write(nrow, ncol + 2, temp.nian_bao_fei(), sy.shu_zi)
    ws.write(nrow, ncol + 3, temp.ren_wu("一阶段任务"), sy.wen_zi)
    ws.write(nrow, ncol + 4, temp.shi_jian_da_cheng("一阶段任务"), sy.bai_fen_bi)
    ws.write(nrow, ncol + 5, temp.ren_wu("二阶段任务"), sy.wen_zi)
    ws.write(nrow, ncol + 6, temp.shi_jian_da_cheng("二阶段任务"), sy.bai_fen_bi)

    ws.set_column(ncol, ncol, width=6)
    ws.set_column(ncol + 1, ncol + 2, width=12)
    ws.set_column(ncol + 3, ncol + 3, width=10)
    ws.set_column(ncol + 4, ncol + 4, width=12)
    ws.set_column(ncol + 5, ncol + 5, width=10)
    ws.set_column(ncol + 6, ncol + 6, width=12)


def ge_ren_jia_yi_xian(sy, ws, ri_qi, cur):
    """
    开门红非车险驾意险月度个人销售奖统计表
    """
    nrow = 0
    ncol = 0

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 4,
        data="非车险驾意险月度个人销售奖统计表",
        cell_format=sy.biao_ti,
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 4,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = ["排名", "中支", "姓名", "月度保费", "签单笔数"]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    str_sql = f"SELECT [中心支公司简称], \
                [业务员], \
                SUM([签单保费/批改保费]) AS 保费, \
                SUM([保单笔数]) \
                FROM [2020年] \
                JOIN [中心支公司] \
                ON [2020年].[中心支公司] = [中心支公司].[中心支公司] \
                WHERE [险种名称] = '0621驾乘人员人身意外伤害保险(B款)' \
                GROUP BY [业务员], [中心支公司简称] \
                ORDER BY [保费] DESC \
                LIMIT 20"

    cur.execute(str_sql)

    for d in cur.fetchall():
        ws.write(nrow, ncol, nrow - 2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1][9:], sy.wen_zi)
        ws.write(nrow, ncol + 3, d[2], sy.shu_zi)
        ws.write(nrow, ncol + 4, d[3], sy.wen_zi)
        nrow += 1

    nrow += 1
    ws.merge_range(nrow, ncol, nrow, ncol + 4, "注：月度前10名获奖", sy.jiao_zhu)

    ws.set_column(ncol, ncol, width=6)
    ws.set_column(ncol + 1, ncol + 2, width=12)
    ws.set_column(ncol + 3, ncol + 3, width=12)
    ws.set_column(ncol + 4, ncol + 4, width=12)


def jia_yi_xian_lian_dong(sy, ws, ri_qi, cur):
    """
    开门红非车险驾意险月度个人联动率奖统计表
    """
    update(tb_name="车险清单", back=False)

    nrow = 0
    ncol = 0

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 5,
        data="非车险驾意险月度个人联动率奖统计表",
        cell_format=sy.biao_ti,
    )

    ws.set_row(nrow, height=24)

    nrow += 1

    ws.merge_range(
        nrow,
        ncol,
        nrow,
        ncol + 5,
        data=f"数据统计范围：2020-01-01 至 {ri_qi}",
        cell_format=sy.shuo_ming,
    )

    nrow += 1

    biao_tou = ["排名", "中支", "姓名", "驾意险\n签单笔数", "车险签单\n车辆数", "驾意险\n联动率"]

    ws.write_row(nrow, ncol, data=biao_tou, cell_format=sy.biao_tou)

    nrow += 1

    str_sql = "SELECT [车险清单].[业务员], \
               COUNT(distinct [车架号]) AS 车辆数 \
               FROM [车险清单] \
               WHERE [使用性质] = '非营业' \
               AND [机动车种类] IN ('客车', '货车') \
               AND [车辆类型] IN ('二吨以下货车', '六座以下客车', '六座至十座客车') \
               AND[座位数] < 8 \
               GROUP BY [业务员] \
               ORDER BY 车辆数 DESC"

    che = {}
    cur.execute(str_sql)
    for d in cur.fetchall():
        che[d[0]] = d[1]

    str_sql = f"SELECT [中心支公司简称], \
        [业务员], \
        [保单数] \
        FROM \
        (SELECT \
        [中心支公司简称], \
        [业务员], \
        SUM ([保单笔数]) AS [保单数] \
        FROM [2020年] \
        JOIN [中心支公司] \
        ON [2020年].[中心支公司] = [中心支公司].[中心支公司] \
        WHERE [险种名称] = '0621驾乘人员人身意外伤害保险(B款)' \
        GROUP  BY [业务员], [中心支公司简称] \
        ORDER  BY [保单数] DESC)AS T \
        WHERE T.保单数 >= 20 \
        LIMIT 20"

    cur.execute(str_sql)
    row_data = []

    for d in cur.fetchall():
        if d[1] in che:
            row_data.append(
                (d[0], d[1][9:], d[2], che[d[1]], d[2] / che[d[1]])
            )
        else:
            row_data.append((d[0], d[1][9:], d[2], 0, 0))

    row_data_sort = sorted(row_data, key=lambda k: k[4], reverse=True)

    for d in row_data_sort:
        ws.write(nrow, ncol, nrow - 2, sy.wen_zi)
        ws.write(nrow, ncol + 1, d[0], sy.wen_zi)
        ws.write(nrow, ncol + 2, d[1], sy.wen_zi)
        ws.write(nrow, ncol + 3, d[2], sy.wen_zi)
        ws.write(nrow, ncol + 4, d[3], sy.wen_zi)
        ws.write(nrow, ncol + 5, d[4], sy.bai_fen_bi)
        nrow += 1

    nrow += 1
    ws.merge_range(nrow, ncol, nrow, ncol + 4, "注：月度前10名获奖", sy.jiao_zhu)

    ws.set_column(ncol, ncol, width=6)
    ws.set_column(ncol + 1, ncol + 2, width=12)
    ws.set_column(ncol + 3, ncol + 3, width=12)


if __name__ == "__main__":

    conn = sqlite3.connect(r"Data\data.db")
    cur = conn.cursor()

    str_sql = f"SELECT MAX([投保确认日期]) \
                FROM [2020年]"
    cur.execute(str_sql)
    ri_qi = cur.fetchone()[0]

    wb = xlsxwriter.Workbook(r"Report\2020年开门红竞赛统计表.xlsx")
    sy = Style(wb)

    ws = wb.add_worksheet("昆明机构车险统计表")
    kun_ming_che(sy, ws, ri_qi)

    ws = wb.add_worksheet("规模达成奖统计表")
    gui_mo_da_cheng(sy, ws, ri_qi)

    ws = wb.add_worksheet("责任险成长奖统计表")
    ze_ren_xian(sy, ws, ri_qi)

    ws = wb.add_worksheet("诉讼保全激励奖统计表")
    su_ze_xian(sy, ws, ri_qi)

    ws = wb.add_worksheet("机构驾意险保费达成奖统计表")
    ji_gou_jia_yi_xian(sy, ws, ri_qi)

    ws = wb.add_worksheet("驾意险月度个人销售奖统计表")
    ge_ren_jia_yi_xian(sy, ws, ri_qi, cur)

    ws = wb.add_worksheet("驾意险月度个人联动率奖统计表")
    jia_yi_xian_lian_dong(sy, ws, ri_qi, cur)

    cur.close()
    conn.close()
    wb.close()
