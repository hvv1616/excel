#本测试程序用于测算字符串匹配率

import difflib
from init_table import *



def diff_str(str, list_table):
    rate = 0
    str_std = ''
    for cmp_str in list_table:
        r = difflib.SequenceMatcher(None, str, cmp_str).ratio()
        if rate < r:
            rate = r
            str_std = cmp_str
    return rate, str_std


if __name__ == '__main__':
    str1 = ' 2.5 WLAN'
    list_table1 = type_table_increase_sub
    print(diff_str(str1, list_table1))
