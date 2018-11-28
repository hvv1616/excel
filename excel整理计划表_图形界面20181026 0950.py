import sys
from gettext import find

from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
import pandas as pd

from ui_excel整理计划表_图形界面20181026 import Ui_MainWindow


class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.statusBar().showMessage('Ready')
        # 公共字典表
        self.type_table_renewal = ['1、机房环境类', '2、网络设备', '3、小型机、服务器及其配套设备', '4、 分行特色应用',
                                   '5、网点建设', '6、 办公自动化。']
        self.type_table_renewal_sub = ['1.1 省、地市分行主机房设备更新', '1.2 机房安全管理',
                                       '2.1 网络（1#）安全生产', '2.2 网络（2#）安全生产',
                                       '3.1 小型机设备', '3.2 服务器设备', '3.3小型机、服务器配套设备', '3.4小型机、服务器及存储监控管理软件',
                                       '4.1 2018年递延项目',
                                       '5.1、各类机构安全生产设备更新']
        self.type_table_increase = ['1、机房环境类', '2、网络设备', '3、小型机、服务器及其配套设备', '4、 分行特色应用',
                                    '5、网点建设', '6、 办公自动化。']
        self.type_table_increase_sub = ['1.1 省、地市分行主机房设备新增', '1.2 机房安全管理', '1.3 机房搬迁/改造',
                                        '2.1 网络（1#）安全生产', '2.2 网络（2#）安全生产', '2.3 3G-4G无线网络', '2.4 WLAN网络（内外网）',
                                        '2.5 视频会议系统',
                                        '3.1 小型机设备', '3.2 服务器设备', '3.3小型机、服务器配套设备', '3.4小型机、服务器及存储监控管理软件',
                                        '4.2 2019年应用项目(2016-2018年，我行年均特色项目执行金额为160万元)',
                                        '5.2、新增柜员、机构设备新增']
        self.index_table = ('项目序号', '分项目名称', '设备名称', '分项计划资金（RMB万元）', '参考品牌', '参考型号',
                            '单价（RMB万元）', '数量', '现有设备情况说明', '备注', '大类', '小类', '需求方细分',
                            '用途', '分行名称', '自动金额', '金额核对', '审核意见', '审核单价', '审核数量',
                            '审核金额', '审核金额核对')
        self.sheet_name_list = ('安徽（存）', '北京（存）', '大连（存）', '福建（存）', '甘肃（存）', '广东（存）',
                                '广西（存）', '贵州（存）', '海南（存）', '河北（存）', '河南（存）', '黑龙江（存）',
                                '湖北（存）', '湖南（存）', '吉林（存）', '江苏（存）', '江西（存）', '辽宁（存）',
                                '内蒙（存）', '宁波（存）', '宁夏（存）', '青岛（存）', '青海（存）', '山东（存）',
                                '山西（存）', '陕西（存）', '上海（存）', '深圳（存）', '四川（存）', '苏州（存）',
                                '天津（存）', '西藏（存）', '新疆（存）', '云南（存）', '浙江（存）', '重庆（存）',
                                '安徽（增）', '北京（增）', '大连（增）', '福建（增）', '甘肃（增）', '广东（增）',
                                '广西（增）', '贵州（增）', '海南（增）', '河北（增）', '河南（增）', '黑龙江（增）',
                                '湖北（增）', '湖南（增）', '吉林（增）', '江苏（增）', '江西（增）', '辽宁（增）',
                                '内蒙（增）', '宁波（增）', '宁夏（增）', '青岛（增）', '青海（增）', '山东（增）',
                                '山西（增）', '陕西（增）', '上海（增）', '深圳（增）', '四川（增）', '苏州（增）',
                                '天津（增）', '西藏（增）', '新疆（增）', '云南（增）', '浙江（增）', '重庆（增）')
        self.finality_classes1_list = ['1、机房环境类', '2、网络设备', '3、小型机、服务器及其配套设备', '4、分行特色应用',
                                       '5、网点建设', '6、办公自动化', '7、机动资金']
        self.finality_classes2_list = ['1.1 省、地市分行主机房设备更新', '1.2 机房安全管理', '1.3 机房搬迁/改造',
                                       '2.1 网络（1#）安全生产', '2.2 网络（2#）安全生产', '2.3 3G无线网络', '2.4 视频会议',
                                       '2.5 WLAN', '2.6 总行统一推广',
                                       '3.1 小型机设备', '3.2 服务器设备', '3.3小型机、服务器配套设备',
                                       '3.4小型机、服务器及存储监控管理软件',
                                       '4.1 上年递延项目',
                                       '5.1 各类机构安全生产设备更新', '5.2 新增柜员、机构设备新增',
                                       '6.1 现有人员设备更新', '6.2 新增人员设备', '6.3 稽核人员设备']
        self.finality_renewal_list = ['1.1.1省、地市分行主机房设备更新(存量）',
                                      '1.2.1 机房安全管理(存量）',
                                      '1.3 机房搬迁/改造',
                                      '2.1.1 网络（1#）安全生产(存量）',
                                      '2.2.1 网络（2#）安全生产(存量）',
                                      '2.3 3G无线网络',
                                      '2.4 视频会议',
                                      '2.5 WLAN',
                                      '3.1.1 小型机设备(存量）',
                                      '3.2.1 服务器设备（存量）',
                                      '3.3.1小型机、服务器配套设备（存量）',
                                      '3.4.1小型机、服务器及存储监控管理软件（存量）',
                                      '4.1 2016年递延项目',
                                      '5.1 各类机构安全生产设备更新',
                                      '6.1 现有人员设备更新',
                                      '6.3 稽核人员设备']
        '''self.finality_increase_list = ['1.1.2 省、地市分行主机房设备新增(增量）',
                                       '1.2.2 机房安全管理(增量）',
                                       '2.1.2 网络（1#）安全生产(增量）',
                                       '2.2.2 网络（2#）安全生产(增量）',
                                       '3.1.2 小型机设备（增量）',
                                       '3.2.2 服务器设备（增量）',
                                       '3.3.2小型机、服务器配套设备（增量）',
                                       '3.4.2小型机、服务器及存储监控管理软件（增量）',
                                       '5.2 新增柜员、机构设备新增',
                                       '6.2 新增人员设备',
                                       '7、机动资金']'''

        # 建立DataFrame数组
        self.target_df = pd.DataFrame(columns=self.index_table)
        # 初始化表格行数
        self.table_begin_line_num = 0
        # 初始化源文件名
        self.source_filenames = []

    def initUI(self):

        self.setWindowTitle('计划表处理')
        self.setupUi(self)

        # 显示版权信息
        self.textBrowse_info.append('胡询 2018年2月7日制作完成。 ')

        # 按钮的处理
        self.pushButton_get_input_file.clicked.connect(lambda: self.read_branch_files(1))
        self.pushButton_get_output_file.clicked.connect(self.save_file)
        self.pushButton_get_input_file_all_in_1.clicked.connect(self.read_branch_files_all_in_1)
        self.pushButton_get_output_file_all_in_1.clicked.connect(self.save_file)
        self.pushButton_get_input_file_finality.clicked.connect(self.read_branch_files_finality)
        self.pushButton_get_output_file_finality.clicked.connect(self.save_file)

    def read_branch_files_finality(self):
        # 获取多个分行上报表的文件名
        source_filenames = self.get_input_file_names()

        # 对每个文件进行读取
        for filename in source_filenames:
            # 将各分行的批复文件（只有唯一sheet）读入数组
            sheet_name_str = 0
            source_df, table_line_num = self.read_sheet_all_in_1(filename, sheet_name_str)
            # 转换格式存放到输出数组
            self.change_format_finality(source_df, self.table_begin_line_num, table_line_num, filename)
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
            self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '数量'] = 0

            if str(df_read.iloc[i, 0]).strip() in self.finality_renewal_list:  # 判断是否存量
                type_usage = '存量'
            elif '(增量）' in str(df_read.iloc[i, 0]) or '新增' in str(df_read.iloc[i, 0]) or '机动资金' in str(
                    df_read.iloc[i, 0]):  # 判断是否增量
                type_usage = '增量'
            elif str(df_read.iloc[i, 0]).strip() in self.finality_classes1_list:  # 判断该行是否是大类说明
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
            elif str(df_read.iloc[i, 0]).strip() in self.finality_classes2_list:  # 判断该行是否是小类说明
                type_2 = df_read.iloc[i, 0].strip()
            elif str(df_read.iloc[i, 0]).strip() in self.index_table:  # 判断该行是否是标题
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
                self.target_df.at[i + append_begin_line_num, '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[i + append_begin_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[i + append_begin_line_num, '分项目名称'] = self.target_df.at[
                        i + append_begin_line_num - 1, '分项目名称']
                self.target_df.at[i + append_begin_line_num, '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[i + append_begin_line_num, '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[i + append_begin_line_num, '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[i + append_begin_line_num, '数量'] = df_read.iloc[i, 7]
                self.target_df.at[i + append_begin_line_num, '备注'] = df_read.iloc[i, 8]

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
                  self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'], '=',
                  self.target_df.at[i + append_begin_line_num, '金额核对'])

    def read_branch_files_all_in_1(self):
        # 获取多个分行上报表的文件名
        source_filenames = self.get_input_file_names()

        # 对每个文件进行读取
        for filename in source_filenames:
            # 将各分行的sheet（存量和增量）读入数组
            for sheet_name_str in self.sheet_name_list:
                # 将sheet表读入
                source_df, table_line_num = self.read_sheet_all_in_1(filename, sheet_name_str)
                # 转换格式存放到输出数组
                self.change_format_all_in_1(source_df, self.table_begin_line_num, table_line_num, sheet_name_str)
                self.table_begin_line_num = self.table_begin_line_num + table_line_num

    def read_sheet_all_in_1(self, filename, sheet_name_str):
        # 读取文件的sheet到pandas的数组
        try:
            df_readin = pd.read_excel(filename, sheet_name=sheet_name_str)
        except:
            # 读取sheet表不存在，则清空读入数组和计数器
            df_readin = pd.DataFrame(columns=self.index_table)
            sheet_total_line_num = 0
        else:
            # 读取sheet表存在，获取数组的总行数
            sheet_total_line_num = df_readin.shape[0]
            self.textBrowse_info.append('文件%s的%s表读取共%i条记录。' % (filename, sheet_name_str, sheet_total_line_num))

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
            type_table = self.type_table_renewal
            type_table_sub = self.type_table_renewal_sub
        elif sheet_type == '增量':
            type_table = self.type_table_increase
            type_table_sub = self.type_table_increase_sub

        for i in range(0, total_line_num):
            self.target_df.at[i + append_begin_line_num, '用途'] = sheet_type
            self.target_df.at[i + append_begin_line_num, '分行名称'] = branch_name
            self.target_df.at[i + append_begin_line_num, '项目序号'] = i + 3
            self.target_df.at[i + append_begin_line_num, '分项目名称'] = '000无分项目名称'
            self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[i + append_begin_line_num, '数量'] = 0

            if str(df_read.iloc[i, 0]).strip() in type_table:  # 判断该行是否是大类说明
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
            elif str(df_read.iloc[i, 0]).strip() in type_table_sub:  # 判断该行是否是小类说明
                type_2 = df_read.iloc[i, 0].strip()
            elif str(df_read.iloc[i, 0]).strip() in self.index_table:  # 判断该行是否是标题
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
            # elif (str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == 'nannan':  # 判断该行单价和金额是否同时为空
            #    pass  # 跳过无金额和单价的行
            else:
                self.target_df.at[i + append_begin_line_num, '大类'] = type_1
                self.target_df.at[i + append_begin_line_num, '小类'] = type_2
                self.target_df.at[i + append_begin_line_num, '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[i + append_begin_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[i + append_begin_line_num, '分项目名称'] = self.target_df.at[
                        i + append_begin_line_num - 1, '分项目名称']
                self.target_df.at[i + append_begin_line_num, '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[i + append_begin_line_num, '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[i + append_begin_line_num, '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[i + append_begin_line_num, '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[i + append_begin_line_num, '数量'] = df_read.iloc[i, 7]
                self.target_df.at[i + append_begin_line_num, '现有设备情况说明'] = df_read.iloc[i, 8]
                self.target_df.at[i + append_begin_line_num, '备注'] = df_read.iloc[i, 9]
                self.target_df.at[i + append_begin_line_num, '审核意见'] = df_read.iloc[i, 10]
                self.target_df.at[i + append_begin_line_num, '审核单价'] = df_read.iloc[i, 11]
                self.target_df.at[i + append_begin_line_num, '审核数量'] = df_read.iloc[i, 12]
                self.target_df.at[i + append_begin_line_num, '审核金额'] = df_read.iloc[i, 13]
                self.target_df.at[i + append_begin_line_num, '审核金额核对'] = df_read.iloc[i, 11] * df_read.iloc[i, 12] \
                                                                         - df_read.iloc[i, 13]

            # 自动核对每一行金额计算是否正确
            print(branch_name, sheet_type, i, ':')
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
                  self.target_df.at[i + append_begin_line_num, '分项计划资金（RMB万元）'], '=',
                  self.target_df.at[i + append_begin_line_num, '金额核对'])

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
            source_df, table_line_num = self.read_sheet(filename, sheet_begin_num)
            self.change_format(source_df, self.table_begin_line_num, table_line_num, '存量', self.type_table_renewal,
                               self.type_table_renewal_sub, filename)
            self.table_begin_line_num = self.table_begin_line_num + table_line_num

            # 将增量表读入，转换格式存放到输出数组
            source_df, table_line_num = self.read_sheet(filename, sheet_begin_num + 1)
            self.change_format(source_df, self.table_begin_line_num + 1, table_line_num, '增量', self.type_table_increase,
                               self.type_table_increase_sub, filename)
            self.table_begin_line_num = self.table_begin_line_num + table_line_num

    # 处理各一级分行上报表格
    def change_format(self, df_read, append_begin_line_num, total_line_num, sheet_type, type_table, type_table_sub,
                      filename):

        type_1 = ''  # 大类
        type_2 = ''  # 小类

        for i in range(0, total_line_num):
            target_df_append_line_num = i + append_begin_line_num
            self.target_df.at[target_df_append_line_num, '用途'] = sheet_type
            self.target_df.at[target_df_append_line_num, '分行名称'] = filename.split('/')[-1].split('.')[0]
            self.target_df.at[target_df_append_line_num, '项目序号'] = i + 3
            self.target_df.at[target_df_append_line_num, '分项目名称'] = '000无分项目名称'
            self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] = 0.0000
            self.target_df.at[target_df_append_line_num, '数量'] = 0

            if str(df_read.iloc[i, 0]).strip() in type_table:  # 判断该行是否是大类说明
                type_1 = df_read.iloc[i, 0].strip()
                type_2 = ''
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000大类'
            elif str(df_read.iloc[i, 0]).strip() in type_table_sub:  # 判断该行是否是小类说明
                type_2 = df_read.iloc[i, 0].strip()
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000小类'
            elif str(df_read.iloc[i, 0]).strip() in self.index_table:  # 判断该行是否是标题
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000表头'
            elif str(df_read.iloc[i, 1]).strip() == '合计' or str(df_read.iloc[i, 2]).strip() == '合计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000合计'
            elif str(df_read.iloc[i, 1]).strip() == '总计' or str(df_read.iloc[i, 2]).strip() == '总计':
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000总计'
            else:
                self.target_df.at[target_df_append_line_num, '大类'] = type_1
                self.target_df.at[target_df_append_line_num, '小类'] = type_2
                self.target_df.at[target_df_append_line_num, '分项目名称'] = df_read.iloc[i, 1]
                # 需要判断'分项目名称'处的合并单元格无法读取值的问题，取上一行数据即可
                if str(self.target_df.at[target_df_append_line_num, '分项目名称']) == 'nan' and i > 5:
                    self.target_df.at[target_df_append_line_num, '分项目名称'] = self.target_df.at[
                        target_df_append_line_num - 1, '分项目名称']
                self.target_df.at[target_df_append_line_num, '设备名称'] = df_read.iloc[i, 2]
                self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）'] = df_read.iloc[i, 3]
                self.target_df.at[target_df_append_line_num, '参考品牌'] = df_read.iloc[i, 4]
                self.target_df.at[target_df_append_line_num, '参考型号'] = df_read.iloc[i, 5]
                self.target_df.at[target_df_append_line_num, '单价（RMB万元）'] = df_read.iloc[i, 6]
                self.target_df.at[target_df_append_line_num, '数量'] = df_read.iloc[i, 7]
                self.target_df.at[target_df_append_line_num, '现有设备情况说明'] = df_read.iloc[i, 8]
                self.target_df.at[target_df_append_line_num, '备注'] = df_read.iloc[i, 9]

            # 自动核对每一行金额计算是否正确
            if (str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == 'nannan' or (
                    str(df_read.iloc[i, 3]) + str(df_read.iloc[i, 6])) == '00':
                # 判断该行单价和金额是否同时为空或0
                # 标注无金额和单价的行
                self.target_df.at[target_df_append_line_num, '分项目名称'] = '000单价金额均为空'
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
                    self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）']
                print(self.target_df.at[target_df_append_line_num, '自动金额'], '-',
                      self.target_df.at[target_df_append_line_num, '分项计划资金（RMB万元）'], '=',
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
        save_filename = self.get_output_file_name().split('.')[0] + '_Output.xlsx'

        # 写入Excel表
        if save_filename != '_Output.xlsx':
            try:
                self.target_df.to_excel(save_filename, sheet_name=save_filename.split('/')[-1].split('_')[0],
                                        index=False)
            except FileExistsError:
                self.textBrowse_info.append('%s 文件未成功保存。' % save_filename)
            else:
                self.textBrowse_info.append('%s 文件已经成功保存。共 %i 条记录。' % (save_filename, self.table_begin_line_num))

                # 建立DataFrame数组
                self.target_df = pd.DataFrame(columns=self.index_table)
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
