from tong_ji import Tong_Ji
from datetime import date


class KMH_TJ(Tong_Ji):

    def ren_wu(self, jie_duan):
        str_sql = f"SELECT [任务] \
                   FROM   [2020年开门红任务] \
                   WHERE  [险种] = '{self.xian_zhong}' \
                   AND [机构] = '{self.ji_gou}' \
                   AND [阶段] = '{jie_duan}'"

        self._cur.execute(str_sql)
        value = self._cur.fetchone()
        if value is None:
            return "——"
        else:
            return float(value[0])

    def ren_wu_jin_du(self, jie_duan):
        if self.ren_wu(jie_duan) is None \
          or self.ren_wu(jie_duan) == "——":
            return '——'
        else:
            return self.nian_bao_fei / self.ren_wu(jie_duan)

    @property
    def shi_jian_jin_du(self):
        '''
        返回当前时间的时间进度
        '''
        tian_shu = date(self.nian, self.yue,
                        self.ri).strftime('%j')
        return int(tian_shu) / 91

    def shi_jian_da_cheng(self, jie_duan):
        if self.ren_wu_jin_du(jie_duan) == '——':
            return '——'
        else:
            return self.ren_wu_jin_du(jie_duan) / self.shi_jian_jin_du


if __name__ == '__main__':
    k = KMH_TJ('曲靖', '中心支公司', '非车险', '险种', '一季度任务')
    print(k.nian_bao_fei,
          k.ren_wu,
          k.ren_wu_da_cheng,
          k.shi_jian_da_cheng)
