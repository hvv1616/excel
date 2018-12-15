from init_table import *
import sys
import difflib
from gettext import find

from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
import pandas as pd
from xlrd import open_workbook

from ui_excel整理计划表 import Ui_MainWindow


def diff_str_check(str, list_table):
    rate = 0
    for cmp_str in list_table:
        r = difflib.SequenceMatcher(None, str, cmp_str).ratio()
        if rate < r:
            rate = r
    if rate >= 0.9:
        return True
    else:
        return False


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.statusBar().showMessage('Ready')

        # 建立DataFrame数组
        self.target_df = pd.DataFrame(columns=index_table)
        # 初始化表格行数
        self.table_begin_line_num = 0
        # 初始化源文件名
        self.source_filenames = []

    def initUI(self):

        self.setWindowTitle('计划表处理')
        self.setupUi(self)

        # 显示版权信息
        self.textBrowse_info.append('胡询 2018年11月28日')

        # 按钮的处理
        self.pushButton_get_input_file.clicked.connect(
            lambda: self.read_branch_files(1))
        self.pushButton_get_output_file.clicked.connect(self.save_file)
        self.pushButton_get_input_file_all_in_1.clicked.connect(
            self.read_branch_files_all_in_1)
        self.pushButton_get_output_file_all_in_1.clicked.connect(
            self.save_file)
        self.pushButton_get_input_file_finality.clicked.connect(
            self.read_branch_files_finality)
        self.pushButton_get_output_file_finality.clicked.connect(
            self.save_file)

    def read_branch_files_finality(self):
        # 获取多个分行上报表的文件名
        source_filenames = self.get_input_file_names()

        # 对每个文件进行读取
        for filename in source_filenames:
            # 将各分行的批复文件（只有唯一sheet）读入数组
            sheet_name_str = 0
            source_df, table_line_num = self.read_sheet_all_in_1(
                filename, sheet_name_str)
            # 转换格式存放到输出数组
            self.change_format_finality(
                source_df, self.table_begin_line_num, table_line_num, filename)
            self.table_begin_line_num = self.table_begin_line_num + table_line_num

    # 处理总行下发一级分行表格
    def change_format_finality(self, df_read, append_begin_line_num, total_line_num, filename):
        type_1 = ''  # 大类
        type_2 = ''  # 小类
        type_usage = ''  # 用途
        # 取文件名作为分行名称
        branch_name = filename.split('/')[-1].split('.')[0]

        for i in range(0, total_line_num):
            self.target_df.at[i + append_begin_line_num, '分行名称'] = branch_name
            self.target_df.at[i + append_begin_line_num, '项目序号'] = i + 3
            self.target_df.at[i + append_begin_line_num, '分项目名称'] = '000无分项目名称'
            self.target_df.at[i + append_begin_line_num,
                              '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '数量'] = 0

            if str(df_read.iloc[i, 0]).strip() in finality_renewal_list:  # 判断是否存量
                type_usage = '存量'
            elif '(增量）' in str(df_read.iloc[i, 0]) or '新增' in str(df_read.iloc[i, 0]) or '机动资金' in str(
                    df_read.iloc[i, 0]):  # 判断是否增量
                type_usage = '增量'
            elif str(df_read.iloc[i, 0]).strip() in finality_classes1_list:  # 判断该行是否是大类说明
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
            elif str(df_read.iloc[i, 0]).strip() in finality_classes2_list:  # 判断该行是否是小类说明
                type_2 = df_read.iloc[i, 0].strip()
            elif str(df_read.iloc[i, 0]).strip() in index_table:  # 判断该行是否是标题
                pass  # 跳过标题行
            elif str(df_read.iloc[i, 1]).strip() == '合计' or str(df_read.iloc[i, 2]).strip() == '合计':
                pass  # 跳过合计行
            elif str(df_read.iloc[i, 1]).strip() == '总计' or str(df_read.iloc[i, 2]).strip() == '总计':
                pass  # 跳过总计行
            elif '负责人' in str(df_read.iloc[i, 7]):
                pass
            elif '联系人' in str(df_read.iloc[i, 7]):
                pass
            elif '联系电话' in str(df_read.iloc[i, 7]):
                pass
            else:
                self.target_df.at[i + append_begin_line_num, '大类'] = type_1
                self.target_df.at[i + append_begin_line_num, '小类'] = type_2
                self.target_df.at[i + append_begin_line_num, '用途'] = type_usage
                self.target_df.at[i + append_begin_line_num,
                                  '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[i + append_begin_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[i + append_begin_line_num, '分项目名称'] = self.target_df.at[
                        i + append_begin_line_num - 1, '分项目名称']
                self.target_df.at[i + append_begin_line_num,
                                  '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[i + append_begin_line_num,
                                  '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[i + append_begin_line_num,
                                  '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[i + append_begin_line_num,
                                  '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[i + append_begin_line_num,
                                  '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[i + append_begin_line_num,
                                  '数量'] = df_read.iloc[i, 7]
                self.target_df.at[i + append_begin_line_num,
                                  '备注'] = df_read.iloc[i, 8]

            # 自动核对每一行金额计算是否正确
            print(branch_name, type_usage, i, ':')
            print('自动金额', '=', self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'], '*',
                  self.target_df.at[i + append_begin_line_num, '数量'])
            self.target_df.at[i + append_begin_line_num, '自动金额'] = \
                self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] * \
                self.target_df.at[i + append_begin_line_num, '数量']
            print(self.target_df.at[i + append_begin_line_num, '自动金额'], '-',
                  self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'])
            self.target_df.at[i + append_begin_line_num, '金额核对'] = \
                self.target_df.at[i + append_begin_line_num, '自动金额'] - \
                self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）']
            print(self.target_df.at[i + append_begin_line_num, '自动金额'], '-',
                  self.target_df.at[i + append_begin_line_num,
                                    '分项计划资金（RMB万元）'], '=',
                  self.target_df.at[i + append_begin_line_num, '金额核对'])

    def read_branch_files_all_in_1(self):
        # 获取多个分行上报表的文件名
        source_filenames = self.get_input_file_names()

        # 对每个文件进行读取
        for filename in source_filenames:
            # 将各分行的sheet（存量和增量）读入数组
            for sheet_name_str in sheet_name_list:
                # 将sheet表读入
                source_df, table_line_num = self.read_sheet_all_in_1(
                    filename, sheet_name_str)
                # 转换格式存放到输出数组
                self.change_format_all_in_1(
                    source_df, self.table_begin_line_num, table_line_num, sheet_name_str)
                self.table_begin_line_num = self.table_begin_line_num + table_line_num
                
        print('读入文件完成。')

    def read_sheet_all_in_1(self, filename, sheet_name_str):
        # 读取文件的sheet到pandas的数组
        try:
            df_readin = pd.read_excel(filename, sheet_name=sheet_name_str)
        except:
            # 读取sheet表不存在，则清空读入数组和计数器
            df_readin = pd.DataFrame(columns=index_table)
            sheet_total_line_num = 0
        else:
            # 读取sheet表存在，获取数组的总行数
            sheet_total_line_num = df_readin.shape[0]
            self.textBrowse_info.append('文件%s的%s表读取共%i条记录。' % (
                filename, sheet_name_str, sheet_total_line_num))

        return df_readin, sheet_total_line_num

    # 处理审核组所有分行明细汇总表格
    def change_format_all_in_1(self, df_read, append_begin_line_num, total_line_num, sheet_name_str):

        type_1 = ''  # 大类
        type_2 = ''  # 小类
        # 取sheet名称的前部分作为分行名称
        branch_name = sheet_name_str[:-3]
        # 取sheet名称的最后一个字符判断是存量表还是增量表
        sheet_type = sheet_name_str[-2] + '量'
        # 按存量和增量表分别取大类和小类清单
        if sheet_type == '存量':
            type_table = type_table_renewal
            type_table_sub = type_table_renewal_sub
        elif sheet_type == '增量':
            type_table = type_table_increase
            type_table_sub = type_table_increase_sub

        for i in range(0, total_line_num):
            target_df_append_line_num = i + append_begin_line_num
            self.target_df.at[target_df_append_line_num, '用途'] = sheet_type
            self.target_df.at[target_df_append_line_num, '分行名称'] = branch_name
            self.target_df.at[target_df_append_line_num, '项目序号'] = i + 3
            self.target_df.at[target_df_append_line_num, '分项目名称'] = '000无分项目名称'
            self.target_df.at[target_df_append_line_num,
                              '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '数量'] = 0

            # 判断该行是否是填表说明和备注
            if diff_str_check(str(df_read.iloc[i, 0]).strip(), other_list):
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000说明'
            # 判断该行是否是大类说明
            elif diff_str_check(str(df_read.iloc[i, 0]).strip(), type_table):
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000大类'
            # 判断该行是否是小类说明
            elif diff_str_check(str(df_read.iloc[i, 0]).strip(), type_table_sub):
                type_2 = df_read.iloc[i, 0].strip()
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000小类'
            elif str(df_read.iloc[i, 0]).strip() in index_table:  # 判断该行是否是标题
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000表头'
            elif str(df_read.iloc[i, 1]).strip() == '合计' or str(df_read.iloc[i, 2]).strip() == '合计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000合计'
            elif str(df_read.iloc[i, 1]).strip() == '总计' or str(df_read.iloc[i, 2]).strip() == '总计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000总计'
            elif str(df_read.iloc[i, 1]).strip() == '小计' or str(df_read.iloc[i, 2]).strip() == '小计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000小计'
            else:
                self.target_df.at[target_df_append_line_num, '大类'] = type_1
                self.target_df.at[target_df_append_line_num, '小类'] = type_2
                self.target_df.at[target_df_append_line_num,
                                  '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[target_df_append_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[target_df_append_line_num, '分项目名称'] = self.target_df.at[
                        target_df_append_line_num - 1, '分项目名称']
                self.target_df.at[target_df_append_line_num,
                                  '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[target_df_append_line_num,
                                  '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[target_df_append_line_num,
                                  '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[target_df_append_line_num,
                                  '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[target_df_append_line_num,
                                  '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[target_df_append_line_num,
                                  '数量'] = df_read.iloc[i, 7]
                self.target_df.at[target_df_append_line_num,
                                  '现有设备情况说明'] = df_read.iloc[i, 8]
                self.target_df.at[target_df_append_line_num,
                                  '备注'] = df_read.iloc[i, 9]
                self.target_df.at[target_df_append_line_num,
                                  '审核意见'] = df_read.iloc[i, 10]
                if type_1 != '2、网络设备':  # 常规类和开放类按标准格式
                    self.target_df.at[target_df_append_line_num, '审核单价'] = df_read.iloc[i, 11]
                    self.target_df.at[target_df_append_line_num, '审核数量'] = df_read.iloc[i, 12]
                    self.target_df.at[target_df_append_line_num, '审核金额'] = df_read.iloc[i, 13]
                else:  # 网络类按特殊格式
                    self.target_df.at[target_df_append_line_num, '审核单价'] = df_read.iloc[i, 6]
                    self.target_df.at[target_df_append_line_num, '审核数量'] = df_read.iloc[i, 7]
                    self.target_df.at[target_df_append_line_num, '审核金额'] = df_read.iloc[i, 3]
                print(branch_name, sheet_type, i, '行', type_1, type_2, df_read.iloc[i, 1], '\n',
                        self.target_df.at[target_df_append_line_num, '审核单价'], '*',
                        self.target_df.at[target_df_append_line_num, '审核数量'], '-',
                        self.target_df.at[target_df_append_line_num, '审核金额'])

                self.target_df.at[target_df_append_line_num, '审核金额核对'] = self.target_df.at[target_df_append_line_num, '审核单价'] * self.target_df.at[target_df_append_line_num, '审核数量'] - self.target_df.at[target_df_append_line_num, '审核金额']

                # 自动核对每一行金额计算是否正确
                if (str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == 'nannan' or (
                        str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == '00':
                    # 判断该行单价和金额是否同时为空或0
                    # 标注无金额和单价的行
                    self.target_df.at[target_df_append_line_num,
                                      '分项目名称'] = '000单价金额均为空'
                else:
                    # print(branch_name, sheet_type, i, ':')
                    '''print('自动金额', '=', self.target_df.at[target_df_append_line_num, '单价（RMB万元）'], '*',
                        self.target_df.at[target_df_append_line_num, '数量'])'''
                    self.target_df.at[target_df_append_line_num, '自动金额'] = \
                        self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] * \
                        self.target_df.at[target_df_append_line_num, '数量']
                    '''print(self.target_df.at[target_df_append_line_num, '自动金额'], '-',
                        self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）'])'''
                    self.target_df.at[target_df_append_line_num, '金额核对'] = \
                        self.target_df.at[target_df_append_line_num, '自动金额'] - \
                        self.target_df.at[target_df_append_line_num,
                                          '分项计划资金（RMB万元）']
                    '''print(self.target_df.at[target_df_append_line_num, '自动金额'], '-',
                        self.target_df.at[target_df_append_line_num,
                                            '分项计划资金（RMB万元）'], '=',
                        self.target_df.at[target_df_append_line_num, '金额核对'])'''

    def get_input_file_names(self):
        # 打开文件对话框，读入文件
        try:
            got_filenames, filename_filter = QFileDialog.getOpenFileNames(self, "选择分行计划文件",
                                                                          '/user/Hvv/PycharmProjects/excel')
        except FileNotFoundError:
            self.textBrowse_info.append('无法找到文件。')
        else:
            return got_filenames

    def read_sheet(self, filename, sheet_order_num):
        # 读取文件的sheet到pandas的数组
        df_readin = pd.read_excel(filename, sheet_name=sheet_order_num)
        # 获取数组的总行数
        sheet_total_line_num = df_readin.shape[0]

        self.textBrowse_info.append('文件 %s %i号sheet有%i条记录。' % (filename, sheet_order_num,
                                                               sheet_total_line_num))
        return df_readin, sheet_total_line_num

    def read_branch_files(self, sheet_begin_num):
        # 获取多个分行上报表的文件名
        source_filenames = self.get_input_file_names()

        # 对每个文件进行读取
        for filename in source_filenames:
            # 分行上报表sheet0 是汇总表，sheet1 是存量明细表，sheet2 是增量明细表

            # 将存量表读入，转换格式存放到输出数组
            source_df, table_line_num = self.read_sheet(
                filename, sheet_begin_num)
            self.change_format(source_df, self.table_begin_line_num, table_line_num, '存量', type_table_renewal,
                               type_table_renewal_sub, filename)
            self.table_begin_line_num = self.table_begin_line_num + table_line_num

            # 将增量表读入，转换格式存放到输出数组
            source_df, table_line_num = self.read_sheet(
                filename, sheet_begin_num + 1)
            self.change_format(source_df, self.table_begin_line_num + 1, table_line_num, '增量', type_table_increase,
                               type_table_increase_sub, filename)
            self.table_begin_line_num = self.table_begin_line_num + table_line_num

    # 处理各一级分行上报表格
    def change_format(self, df_read, append_begin_line_num, total_line_num, sheet_type, type_table, type_table_sub,
                      filename):

        type_1 = ''  # 大类
        type_2 = ''  # 小类

        for i in range(0, total_line_num):
            target_df_append_line_num = i + append_begin_line_num
            self.target_df.at[target_df_append_line_num, '用途'] = sheet_type
            self.target_df.at[target_df_append_line_num,
                              '分行名称'] = filename.split('/')[-1].split('.')[0]
            self.target_df.at[target_df_append_line_num, '项目序号'] = i + 3
            self.target_df.at[target_df_append_line_num, '分项目名称'] = '000无分项目名称'
            self.target_df.at[target_df_append_line_num,
                              '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '数量'] = 0

            # 判断该行是否是大类说明
            if diff_str_check(str(df_read.iloc[i, 0]).strip(), type_table):
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000大类'
            # 判断该行是否是小类说明
            elif diff_str_check(str(df_read.iloc[i, 0]).strip(), type_table_sub):
                type_2 = df_read.iloc[i, 0].strip()
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000小类'
            elif str(df_read.iloc[i, 0]).strip() in index_table:  # 判断该行是否是标题
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000表头'
            elif str(df_read.iloc[i, 1]).strip() == '合计' or str(df_read.iloc[i, 2]).strip() == '合计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000合计'
            elif str(df_read.iloc[i, 1]).strip() == '总计' or str(df_read.iloc[i, 2]).strip() == '总计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000总计'
            elif str(df_read.iloc[i, 1]).strip() == '小计' or str(df_read.iloc[i, 2]).strip() == '小计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000小计'
            else:
                self.target_df.at[target_df_append_line_num, '大类'] = type_1
                self.target_df.at[target_df_append_line_num, '小类'] = type_2
                self.target_df.at[target_df_append_line_num,
                                  '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[target_df_append_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[target_df_append_line_num, '分项目名称'] = self.target_df.at[
                        target_df_append_line_num - 1, '分项目名称']
                self.target_df.at[target_df_append_line_num,
                                  '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[target_df_append_line_num,
                                  '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[target_df_append_line_num,
                                  '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[target_df_append_line_num,
                                  '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[target_df_append_line_num,
                                  '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[target_df_append_line_num,
                                  '数量'] = df_read.iloc[i, 7]
                self.target_df.at[target_df_append_line_num,
                                  '现有设备情况说明'] = df_read.iloc[i, 8]
                self.target_df.at[target_df_append_line_num,
                                  '备注'] = df_read.iloc[i, 9]

            # 自动核对每一行金额计算是否正确
            if (str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == 'nannan' or (
                    str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == '00':
                # 判断该行单价和金额是否同时为空或0
                # 标注无金额和单价的行
                self.target_df.at[target_df_append_line_num,
                                  '分项目名称'] = '000单价金额均为空'
            else:
                print(filename, sheet_type, i, ':')
                print('自动金额', '=', self.target_df.at[target_df_append_line_num, '单价（RMB万元）'], '*',
                      self.target_df.at[target_df_append_line_num, '数量'])
                self.target_df.at[target_df_append_line_num, '自动金额'] = \
                    self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] * \
                    self.target_df.at[target_df_append_line_num, '数量']
                print(self.target_df.at[target_df_append_line_num, '自动金额'], '-',
                      self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）'])
                self.target_df.at[target_df_append_line_num, '金额核对'] = \
                    self.target_df.at[target_df_append_line_num, '自动金额'] - \
                    self.target_df.at[target_df_append_line_num,
                                      '分项计划资金（RMB万元）']
                print(self.target_df.at[target_df_append_line_num, '自动金额'], '-',
                      self.target_df.at[target_df_append_line_num,
                                        '分项计划资金（RMB万元）'], '=',
                      self.target_df.at[target_df_append_line_num, '金额核对'])

    def get_output_file_name(self):
        output_file_name = ''
        try:
            output_file_name, myfilter = QFileDialog.getSaveFileName(self, '保存文件', 'Users/Hvv/PycharmProjects/excel',
                                                                     'Excel Files(*.xlsx)', None)
        except:
            self.textBrowse_info.append('选择保存文件出错！！！')
            output_file_name = ''
        else:
            return output_file_name

    def save_file(self):
        save_filename = self.get_output_file_name().split('.')[
            0] + '_Output.xlsx'

        # 写入Excel表
        if save_filename != '_Output.xlsx':
            try:
                self.target_df.to_excel(save_filename, sheet_name=save_filename.split('/')[-1].split('_')[0],
                                        index=False, engine='xlsxwriter')
            except FileExistsError:
                self.textBrowse_info.append('%s 文件未成功保存。' % save_filename)
            else:
                self.textBrowse_info.append('%s 文件已经成功保存。共 %i 条记录。' % (
                    save_filename, self.table_begin_line_num))

                # 建立DataFrame数组
                self.target_df = pd.DataFrame(columns=index_table)
                # 初始化表格行数
                self.table_begin_line_num = 0
                # 初始化源文件名
                self.source_filenames = []
        else:
            self.textBrowse_info.append('请重新选择保存文件。')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    my_mainwindow = MyMainWindow()
    my_mainwindow.show()
    sys.exit(app.exec_())
