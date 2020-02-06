from code.tong_ji_nian import Tong_Ji_Nian
from code.tong_ji_ji_du import Tong_Ji_Ji_Du
from code.tong_ji_yue import Tong_Ji_Yue
from code.tong_ji_xun import Tong_Ji_Xun
from code.tong_ji_zhou import Tong_Ji_Zhou
from code.tong_ji_ri import Tong_Ji_Ri


class Tong_Ji(
    Tong_Ji_Nian, Tong_Ji_Ji_Du, Tong_Ji_Yue, Tong_Ji_Xun, Tong_Ji_Zhou, Tong_Ji_Ri
):
    """
    统计类的统一调用接口

        本类通过继承其他类来实现一些列类的统一调用，
        本类中并没有包含实际代码，不同功能在不同的类中实现

        class Tong_Ji_Base:
            统计类的统一基类，各项基础信息的实现代码

        class Tong_Ji_Nian:
            年度数据统计类，关于年度数据的各项统计功能的实现代码

        class Tong_Ji_Yue:
            月度数据统计类，关于月度数据的各项统计功能的实现代码

        class Tong_Ji_Xun:
            旬数据统计类，关于旬数据的各项统计功能的实现代码

        class Tong_Ji_Zhou：
            周数据统计类，关于周数据的各项统计功能的实现代码

        class Tong_Ji_Ri：
            日数据统计类，关于日数据的各项统计功能的实现代码

    """
