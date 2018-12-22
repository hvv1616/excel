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

def check_diff_of_str(cmp_str, str_list, base_rate=0.9):
    # str   是用于比较的字符串
    # list_table 是用于比较的字符串列表
    # base_rate 是判断的最低相似度，大于此值返回True，否则返回False
    rate = 0
    for str in str_list:
        r = difflib.SequenceMatcher(None, cmp_str, str).ratio()
        if rate < r:
            rate = r
    if rate >= base_rate:
        return True
    else:
        return False

if __name__ == '__main__':
    
    str_in = '7、机动资金'
    # 用于比较的字符串
    str_list = finality_classes1_list
    # 用于比较的列表
    print(diff_str(str_in, str_list))
