import pandas as pd
import openpyxl as op
import numpy as np
from io import StringIO
from openpyxl import Workbook


def cut_file(content):  # 读数据
    end_flag = ['|', '|'] 	 # 结束符号,包括中英文的竖杠

    content_len = len(content)

    df_data = pd.DataFrame(
        columns=['time', 'type', 'state', 'num', 'name'])  # 创建新的数据列表，并指定列索引
    words = []
    tmp_char = ''
    flag = 0

    for idx, char in enumerate(content):

        if (idx + 1) == content_len:  # 判断是否已经到了文末
            words.append(tmp_char)
            df_tmp = pd.DataFrame([[words[0], words[1], words[2], words[3], words[4]]], columns=[
                                  'time', 'type', 'state', 'num', 'name'])
            df_data = df_data._append(df_tmp)
            break

        if char in end_flag:  # 判断当前字符是否为结束符
            next_idx = idx + 1
            if not tmp_char == '':  # 防止在开头出现结束符导致最后有空的单元格出现
                if not content[next_idx] in end_flag:  # 再判断下一个字符是否为结束符，如果不是则切分文本
                    words.append(tmp_char)
                    tmp_char = ''
                    flag += 1
        else:
            tmp_char += char  # 拼接字符

        if flag > 4:  # 当遍历一行后，将该行数据加入df_data列表中
            df_tmp = pd.DataFrame([[words[0], words[1], words[2], words[3], words[4]]], columns=[
                                  'time', 'type', 'state', 'num', 'name'])
            df_data = df_data._append(df_tmp)
            words = []  # 清空暂存
            flag = 0  # 重置标识
    df_data.reset_index(drop=True, inplace=True)  # 重新生成行索引
    return df_data  # 返回包含切割完成后的数据的df结构


content = '||2021\.7.23|发明|授权|202110834536.3|越障底盘结构及移动机器人|||2020\.10.16|发明|授权|202011108207.2|通过视觉测距补偿体温检测精度的方法及装置|||2020\.11.18|发明|授权|202011295100.3|自动作业方法及系统|||2020\.7.27|实用新型|授权|202021524737.0|行走机构及机器人|||2020\.7.27|实用新型|授权|202021538043.2|喷雾消杀机器人|||2020\.8.14|实用新型|授权|202021733966.3|智能充电交互系统|||2020\.8.24|实用新型|授权|202021785152.4|消毒检疫控制系统|||2020\.10.10|实用新型|授权|202022248921.3|消毒检疫机器人|||2020\.11.10|实用新型|授权|202022592429.8|自动化散件拣选仓储系统|||2020\.11.10|实用新型|授权|202022592427.9|自动化散件拣选仓储系统|||2020\.11.10|实用新型|授权|202022592275.2|智能化散件拣选装置|||2020\.11.10|实用新型|授权|202022588423.3|智能化散件拣选装置及货物拣选仓储系统|||2020\.11.10|实用新型|授权|202022592384.4|货物拣选箱及货物拣选仓储系统|||2020\.11.10|实用新型|授权|202022592428.3|货物拣选箱及货物拣选仓储系统|||2020\.11.10|实用新型|授权|202022592430.0|货物拣选箱及货物拣选仓储系统|||2020\.12.03|实用新型|授权|202022868570.6|柔性喷雾装置|||2020\.12.24|实用新型|授权|202023166237.7|换装机器人及机器人工作站|||2021\.01.29|实用新型|授权|202120253797.1|电池快换系统|||2021\.02.08|实用新型|授权|202120355700.8|无人机收发系统|||2021\.02.08|实用新型|授权|202120355806.8|智能搬运装置|||2021\.02.08|实用新型|授权|202120366205.7|包装拆卸装置|||2021\.04.16|实用新型|授权|202120793832.9|包装拆卸装置|||2021\.02.08|实用新型|授权|202120366839.2|自动作业装置|||2021\.04.01|实用新型|授权|202120673435.8|货物拣选装置|||2021\.04.15|实用新型|授权|202120775141.6|货物搬运装置|||2021\.04.29|实用新型|授权|202120926445.8|龙门式分拣装置及仓储系统|||2021\.05.24|实用新型|授权|202121121493.6|物体物理特性采集装置及智能拆零系统|||2021\.07.06|实用新型|授权|202121544510.7|叉车搬运装置|||2021\.02.08|实用新型|授权|202120354458.2|自主清洁消毒装置|||2021\.02.25|实用新型|授权|202120421856.1|机器人集成电控柜及机器人电气控制系统|||2021\.09.09|实用新型|授权|202122182318.4|激光切割装置|||2021\.09.09|实用新型|授权|202122180887.5|激光切割装置|'
df_data = cut_file(content)


data_row = df_data.shape[0]  # 读数据的行数
data_column = df_data.shape[1]  # 读数据的列数
print('数据共有'+str(df_data.shape[0])+'行'+str(df_data.shape[1])+'列')  # 显示原始数据行列


data = df_data  # 读数据
wb = op.load_workbook(
    r'C:\Users\admin\Desktop\test\test2\清单.xlsx')  # 选中目标excel
ws = wb.worksheets[0]  # 在第一个sheet中写入

# Q1：专利‘name’是什么时候申请的？


def create_Q1():

    ws = wb.active
    for row1 in range(0, data_row):
        question_list = ['请问专利"'+df_data['name'].iloc[row1]+'"的申请日期是什么时候？',
                         '能告诉我专利"'+df_data['name'].iloc[row1]+'"是在什么时候申请的吗？',
                         '我想知道专利"'+df_data['name'].iloc[row1]+'"的申请日期。',
                         '专利"'+df_data['name'].iloc[row1]+'"是在哪一天申请的？',
                         '请提供专利"'+df_data['name'].iloc[row1]+'"的申请日期。',
                         '你知道专利"' +
                         df_data['name'].iloc[row1]+'"是什么时候申请的吗？',
                         '我可以查询到专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期吗？',
                         '有关于专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期信息吗？',
                         '我想查询关于专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',
                         '请告诉我有关于专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',

                         '请问盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期是什么时候？',
                         '能告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"是在什么时候申请的吗？',
                         '我想知道盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',
                         '盈合公司专利"'+df_data['name'].iloc[row1]+'"是在哪一天申请的？',
                         '请提供盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',
                         '你知道盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"是什么时候申请的吗？',
                         '我可以查询到盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期吗？',
                         '有关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期信息吗？',
                         '我想查询关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',
                         '请告诉我有关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的申请日期。',
                         ]
        for i in range(len(question_list)):
            ws.append([question_list[i], df_data['time'].iloc[row1]])
    print('Q1 successfully created')

# Q2：专利‘name’是什么类型的？


def create_Q2():

    ws = wb.active
    for row1 in range(0, data_row):
        question_list = ['专利"'+df_data['name'].iloc[row1]+'"的类型是什么？',
                         '请问专利"'+df_data['name'].iloc[row1]+'"的类型是什么？',
                         '请告诉我专利"'+df_data['name'].iloc[row1]+'"的类型。',
                         '专利"'+df_data['name'].iloc[row1]+'"的类型是什么呢？',
                         '能告诉我专利"'+df_data['name'].iloc[row1]+'"的类型吗？',
                         '我想知道专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '能查询一下专利"' +
                         df_data['name'].iloc[row1]+'"的类型吗？',
                         '请帮我查一下专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '我想查询关于专利"'+df_data['name'].iloc[row1]+'"的类型。',
                         '请提供专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',

                         '盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型是什么？',
                         '请问盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型是什么？',
                         '请告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '盈合公司专利"'+df_data['name'].iloc[row1]+'"的类型是什么呢？',
                         '能告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型吗？',
                         '我想知道盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '能查询一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型吗？',
                         '请帮我查一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '我想查询关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型。',
                         '请提供盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的类型。'
                         ]
        for i in range(len(question_list)):
            ws.append([question_list[i], df_data['type'].iloc[row1]])
    print('Q2 successfully created')

# Q3：专利‘name’的专利号是多少？


def create_Q3():

    ws = wb.active
    for row1 in range(0, data_row):
        question_list = ['专利"'+df_data['name'].iloc[row1]+'"的专利号是什么？',
                         '请问专利"'+df_data['name'].iloc[row1]+'"的专利号是多少？',
                         '请告诉我专利"'+df_data['name'].iloc[row1]+'"的专利号。',
                         '专利"'+df_data['name'].iloc[row1]+'"的专利号是多少呢？',
                         '能告诉我专利"'+df_data['name'].iloc[row1]+'"的专利号吗？',
                         '我想知道专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '能查询一下专利"' +
                         df_data['name'].iloc[row1]+'"的专利号吗？',
                         '请帮我查一下专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '我想查询关于专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '请提供专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',

                         '盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号是什么？',
                         '请问盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号是多少？',
                         '请告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '盈合公司专利"'+df_data['name'].iloc[row1]+'"的专利号是多少呢？',
                         '能告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号吗？',
                         '我想知道盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '能查询一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号吗？',
                         '请帮我查一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '我想查询关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。',
                         '请提供盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的专利号。'
                         ]
        for i in range(len(question_list)):
            ws.append([question_list[i], df_data['num'].iloc[row1]])
    print('Q3 successfully created')

# Q4：专利‘name’的状态是什么？


def create_Q4():

    ws = wb.active
    for row1 in range(0, data_row):
        question_list = ['专利"'+df_data['name'].iloc[row1]+'"的状态是什么？',
                         '请问专利"'+df_data['name'].iloc[row1]+'"的状态是怎样的？',
                         '请告诉我专利"'+df_data['name'].iloc[row1]+'"的状态。',
                         '专利"'+df_data['name'].iloc[row1]+'"的状态是什么呢？',
                         '能告诉我专利"'+df_data['name'].iloc[row1]+'"的状态吗？',
                         '我想知道专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '能查询一下专利"' +
                         df_data['name'].iloc[row1]+'"的状态吗？',
                         '请帮我查一下专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '我想查询关于专利"'+df_data['name'].iloc[row1]+'"的状态。',
                         '请告诉我专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',

                         '盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态是什么？',
                         '请问盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态是怎样的？',
                         '请告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '盈合公司专利"'+df_data['name'].iloc[row1]+'"的状态是什么呢？',
                         '能告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态吗？',
                         '我想知道盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '能查询一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态吗？',
                         '请帮我查一下盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '我想查询关于盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态。',
                         '请告诉我盈合公司专利"' +
                         df_data['name'].iloc[row1]+'"的状态。'
                         ]
        for i in range(len(question_list)):
            ws.append([question_list[i], df_data['state'].iloc[row1]])
    print('Q4 successfully created')

# Q5：2020年申请的专利有哪些？
# Q6：2021年申请的专利有哪些？


def create_Q5Q6():
    df_data_time_2020 = pd.DataFrame()
    df_data_time_2021 = pd.DataFrame()
    year = ''

    for index in range(0, data_row):        # 遍历原始数据的每一行
        # 逐个检测time列中的每一个字
        for idx, char in enumerate(df_data['time'].iloc[index]):
            year += char
            if (idx > 2):  # 读取前4个数字（年份）

                if year == '2020':

                    df_tmp = pd.DataFrame(
                        [[df_data['name'].iloc[index]]], columns=['2020'])
                    df_data_time_2020 = df_data_time_2020._append(df_tmp)
                    year = ''

                if year == '2021':

                    df_tmp = pd.DataFrame(
                        [[df_data['name'].iloc[index]]], columns=['2021'])
                    df_data_time_2021 = df_data_time_2021._append(df_tmp)
                    year = ''
                break

    # 重新生成行索引
    df_data_time_2020.reset_index(drop=True, inplace=True)
    df_data_time_2021.reset_index(drop=True, inplace=True)

    # 生成问题
    ws = wb.active

    question_list = ['2020年申请的专利有哪些？',
                     '哪些专利是在2020年申请的？',
                     '2020年申请了哪些专利？',
                     '请问2020年有哪些专利申请？',
                     '2020年申请的专利有哪些呢？',
                     '2020年的专利申请有哪些？',
                     '在2020年，哪些专利被申请了？',
                     '2020年，有哪些专利申请？',
                     '哪些专利在2020年被申请了？',
                     '请列出2020年的专利申请。',

                     '盈合公司2020年申请的专利有哪些？',
                     '盈合公司哪些专利是在2020年申请的？',
                     '2020年盈合公司申请了哪些专利？',
                     '请问2020年盈合公司有哪些专利申请？',
                     '盈合公司2020年申请的专利有哪些呢？',
                     '盈合公司2020年的专利申请有哪些？',
                     '在2020年，哪些专利被盈合公司申请了？',
                     '盈合公司2020年，有哪些专利申请？',
                     '哪些专利在2020年被盈合公司申请了？',
                     '请列出盈合公司2020年的专利申请。']

    data_array = np.array(df_data_time_2020)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])

    question_list = ['2021年申请的专利有哪些？',
                     '哪些专利是在2021年申请的？',
                     '2021年申请了哪些专利？',
                     '请问2021年有哪些专利申请？',
                     '2021年申请的专利有哪些呢？',
                     '2021年的专利申请有哪些？',
                     '在2021年，哪些专利被申请了？',
                     '2021年，有哪些专利申请？',
                     '哪些专利在2021年被申请了？',
                     '请列出2021年的专利申请。',

                     '盈合公司2021年申请的专利有哪些？',
                     '盈合公司哪些专利是在2021年申请的？',
                     '2021年盈合公司申请了哪些专利？',
                     '请问2021年盈合公司有哪些专利申请？',
                     '盈合公司2021年申请的专利有哪些呢？',
                     '盈合公司2021年的专利申请有哪些？',
                     '在2021年，哪些专利被盈合公司申请了？',
                     '盈合公司2021年，有哪些专利申请？',
                     '哪些专利在2021年被盈合公司申请了？',
                     '请列出盈合公司2021年的专利申请。']

    data_array = np.array(df_data_time_2021)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])

    # print(df_data_time_2020)
    # print(df_data_time_2021)
    print('Q5 and Q6 successfully created')

# Q7：性质为“发明”的专利有哪些？
# Q8：性质为“实用新型”的专利有哪些？
# Q9：性质为“外观设计”的专利有哪些？

def create_Q7Q8Q9():

    df_data_type_faming = pd.DataFrame()   #发明
    df_data_type_shiyong = pd.DataFrame()  #实用新型
    df_data_type_waiguan = pd.DataFrame()  #外观设计
    type1 = ''

    for index in range(0, data_row):        # 遍历原始数据的每一行
        # 逐个检测type列中的每一个字
        for idx, char in enumerate(df_data['type'].iloc[index]):
            type1 += char

            if type1 == '发明':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['发明'])
                df_data_type_faming = df_data_type_faming._append(df_tmp)
                type1 = ''
                break
            if type1 == '实用新型':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['实用新型'])
                df_data_type_shiyong = df_data_type_shiyong._append(df_tmp)
                type1 = ''
                break
            if type1 == '外观设计':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['外观设计'])
                df_data_type_waiguan = df_data_type_waiguan._append(df_tmp)
                type1 = ''
                break

    # 重新生成行索引
    df_data_type_faming.reset_index(drop=True, inplace=True)
    df_data_type_shiyong.reset_index(drop=True, inplace=True)
    df_data_type_waiguan.reset_index(drop=True, inplace=True)
    # 生成问题
    ws = wb.active
    question_list = ['性质为“发明”的专利有哪些?',
                     '有哪些性质为“发明”的专利？',
                     '哪些专利的性质是“发明”？',
                     '请问有哪些性质为“发明”的专利？',
                     '性质为“发明”的专利有哪些呢？',
                     '性质为“发明”的专利都有哪些？',
                     '哪些专利属于“发明”性质？',
                     '请列出性质为“发明”的专利。',
                     '性质为“发明”的专利包括哪些？',
                     '“发明”性质的专利有哪些？',

                     '盈合公司性质为“发明”的专利有哪些?',
                     '盈合公司有哪些性质为“发明”的专利？',
                     '盈合公司的哪些专利的性质是“发明”？',
                     '请问盈合公司有哪些性质为“发明”的专利？',
                     '盈合公司性质为“发明”的专利有哪些呢？',
                     '盈合公司性质为“发明”的专利都有哪些？',
                     '盈合公司的哪些专利属于“发明”性质？',
                     '请列出盈合公司的性质为“发明”的专利。',
                     '盈合公司性质为“发明”的专利包括哪些？',
                     '盈合公司“发明”性质的专利有哪些？']
    data_array = np.array(df_data_type_faming)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])

    question_list = ['性质为“实用新型”的专利有哪些?',
                     '有哪些性质为“实用新型”的专利？',
                     '哪些专利的性质是“实用新型”？',
                     '请问有哪些性质为“实用新型”的专利？',
                     '性质为“实用新型”的专利有哪些呢？',
                     '性质为“实用新型”的专利都有哪些？',
                     '哪些专利属于“实用新型”性质？',
                     '请列出性质为“实用新型”的专利。',
                     '性质为“实用新型”的专利包括哪些？',
                     '“实用新型”性质的专利有哪些？',

                     '盈合公司性质为“实用新型”的专利有哪些?',
                     '盈合公司有哪些性质为“实用新型”的专利？',
                     '盈合公司的哪些专利的性质是“实用新型”？',
                     '请问盈合公司有哪些性质为“实用新型”的专利？',
                     '盈合公司性质为“实用新型”的专利有哪些呢？',
                     '盈合公司性质为“实用新型”的专利都有哪些？',
                     '盈合公司的哪些专利属于“实用新型”性质？',
                     '请列出盈合公司的性质为“实用新型”的专利。',
                     '盈合公司性质为“实用新型”的专利包括哪些？',
                     '盈合公司“实用新型”性质的专利有哪些？']
    data_array = np.array(df_data_type_shiyong)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])
        
    question_list = ['性质为“外观设计”的专利有哪些?',
                     '有哪些性质为“外观设计”的专利？',
                     '哪些专利的性质是“外观设计”？',
                     '请问有哪些性质为“外观设计”的专利？',
                     '性质为“外观设计”的专利有哪些呢？',
                     '性质为“外观设计”的专利都有哪些？',
                     '哪些专利属于“外观设计”性质？',
                     '请列出性质为“外观设计”的专利。',
                     '性质为“外观设计”的专利包括哪些？',
                     '“外观设计”性质的专利有哪些？',

                     '盈合公司性质为“外观设计”的专利有哪些?',
                     '盈合公司有哪些性质为“外观设计”的专利？',
                     '盈合公司的哪些专利的性质是“外观设计”？',
                     '请问盈合公司有哪些性质为“外观设计”的专利？',
                     '盈合公司性质为“外观设计”的专利有哪些呢？',
                     '盈合公司性质为“外观设计”的专利都有哪些？',
                     '盈合公司的哪些专利属于“外观设计”性质？',
                     '请列出盈合公司的性质为“外观设计”的专利。',
                     '盈合公司性质为“外观设计”的专利包括哪些？',
                     '盈合公司“外观设计”性质的专利有哪些？']
    data_array = np.array(df_data_type_waiguan)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])
    print('Q7, Q8 and Q9 successfully created')

# Q10：状态为“授权”的专利有哪些？
# Q11：状态为“申请中”的专利有哪些？

def create_Q10Q11():
    df_data_state_sccess = pd.DataFrame()  #授权
    df_data_state_ongoing = pd.DataFrame()  #申请中
    df_data_state_fail = pd.DataFrame()  #驳回
    state1 = ''

    for index in range(0, data_row):        # 遍历原始数据的每一行
        # 逐个检测state列中的每一个字
        for idx, char in enumerate(df_data['state'].iloc[index]):

            state1 += char
            if state1 == '授权':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['授权'])
                df_data_state_sccess = df_data_state_sccess._append(df_tmp)
                state1 = ''
                break
            if state1 == '申请中':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['申请中'])
                df_data_state_ongoing = df_data_state_ongoing._append(df_tmp)
                state1 = ''
                break
            if state1 == '驳回':
                df_tmp = pd.DataFrame(
                    [[df_data['name'].iloc[index]]], columns=['驳回'])
                df_data_state_fail = df_data_state_fail._append(df_tmp)
                state1 = ''
                break

    # 重新生成行索引
    df_data_state_sccess.reset_index(drop=True, inplace=True)

    # 生成问题
    ws = wb.active
    question_list = ['有哪些状态为“授权”的专利？',
                     '哪些专利的状态是“授权”？',
                     '状态为“授权”的专利有哪些？',
                     '请问有哪些状态为“授权”的专利？',
                     '状态为“授权”的专利有哪些呢？',
                     '状态为“授权”的专利都有哪些？',
                     '哪些专利属于“授权”状态？',
                     '请列出状态为“授权”的专利。',
                     '状态为“授权”的专利包括哪些？',
                     '“授权”状态的专利有哪些？'

                     '盈合公司有哪些状态为“授权”的专利？',
                     '盈合公司哪些专利的状态是“授权”？',
                     '盈合公司状态为“授权”的专利有哪些？',
                     '请问盈合公司有哪些状态为“授权”的专利？',
                     '盈合公司状态为“授权”的专利有哪些呢？',
                     '盈合公司状态为“授权”的专利都有哪些？',
                     '盈合公司的哪些专利属于“授权”状态？',
                     '请列出盈合公司状态为“授权”的专利。',
                     '盈合公司状态为“授权”的专利包括哪些？',
                     '盈合公司“授权”状态的专利有哪些？']
    data_array = np.array(df_data_state_sccess)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])
        
    question_list = ['有哪些状态为“申请中”的专利？',
                     '哪些专利的状态是“申请中”？',
                     '状态为“申请中”的专利有哪些？',
                     '请问有哪些状态为“申请中”的专利？',
                     '状态为“申请中”的专利有哪些呢？',
                     '状态为“申请中”的专利都有哪些？',
                     '哪些专利属于“申请中”状态？',
                     '请列出状态为“申请中”的专利。',
                     '状态为“申请中”的专利包括哪些？',
                     '“申请中”状态的专利有哪些？',

                     '盈合公司有哪些状态为“申请中”的专利？',
                     '盈合公司哪些专利的状态是“申请中”？',
                     '盈合公司状态为“申请中”的专利有哪些？',
                     '请问盈合公司有哪些状态为“申请中”的专利？',
                     '盈合公司状态为“申请中”的专利有哪些呢？',
                     '盈合公司状态为“申请中”的专利都有哪些？',
                     '盈合公司的哪些专利属于“申请中”状态？',
                     '请列出盈合公司状态为“申请中”的专利。',
                     '盈合公司状态为“申请中”的专利包括哪些？',
                     '盈合公司“申请中”状态的专利有哪些？']
    data_array = np.array(df_data_state_ongoing)
    data_list = data_array.tolist()
    for i in range(len(question_list)):
        ws.append([question_list[i], str([data_list])])
    print('Q10 and Q11 successfully created')


create_Q1()
create_Q2()
create_Q3()
create_Q4()
create_Q5Q6()
create_Q7Q8Q9()
create_Q10Q11()

wb.save(r'C:\Users\admin\Desktop\test\test2\清单.xlsx')  # 保存

