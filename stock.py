# -*- coding: utf-8 -*-
import baostock as bs
import xlsxwriter
import numpy as np
from datetime import datetime


# 登陆系统
lg = bs.login()
# 显示登陆返回信息
print('login respond error_code:' + lg.error_code)
print('login respond error_msg:' + lg.error_msg)

code_dic = [
            ('其它梯队', '东方财富', '300059', 'sz'),
            ('其它梯队', '中信证券', '600030', 'sh'),
            ('其它梯队', '中信建投', '601066', 'sh'),
            ('其它梯队', '广发证券', '000776', 'sz'),

            ('第一梯队', '贵州茅台', '600519', 'sh'),
            ('第二梯队', '王粮液', '000858', 'sz'),
            ('其它梯队', '泸州老窖', '000568', 'sz'),
            ('其它梯队', 'ST舍得', '600702', 'sh'),

            ('第二梯队', '伊利股份', '600887', 'sh'),
            ('第二梯队', '海天味业', '603288', 'sh'),
            ('其它梯队', '千禾味业', '603027', 'sh'),
            ('其它梯队', '涪陵榨菜', '002507', 'sz'),
            ('其它梯队', '双汇发展', '000895', 'sz'),
            ('其它梯队', '中炬高新', '600872', 'sh'),
            ('其它梯队', '汤臣倍健', '300146', 'sz'),
            ('其它梯队', '安琪酵母', '600298', 'sh'),

            ('第一梯队', '恒瑞医药', '600276', 'sh'),
            ('第二梯队', '长春高新', '000661', 'sz'),
            ('第一梯队', '片仔癀', '600436', 'sh'),
            ('其它梯队', '云南白药', '000538', 'sz'),
            ('其它梯队', '同仁堂', '600085', 'sh'),
            ('其它梯队', '马应龙', '600993', 'sh'),

            ('第二梯队', '通策医疗', '600763', 'sh'),
            ('第二梯队', '爱尔眼科', '300015', 'sz'),
            ('其它梯队', '迈瑞医疗', '300760', 'sz'),
            ('其它梯队', '欧普康视', '300595', 'sz'),
            ('其它梯队', '药明康德', '603259', 'sh'),
            ('其它梯队', '泰格医药', '300347', 'sz'),
            ('其它梯队', '乐普医疗', '300003', 'sz'),
            ('其它梯队', '国际医学', '000516', 'sz'),
            ('其它梯队', '珀莱雅', '603605', 'sh'),

            ('其它梯队', '华东医药', '000963', 'sz'),
            ('其它梯队', '恩华药业', '002262', 'sz'),
            ('其它梯队', '甘李药业', '603087', 'sh'),
            ('其它梯队', '复星医药', '600196', 'sh'),
            ('其它梯队', '我武生物', '300357', 'sz'),

            ('第二梯队', '格力电器', '000651', 'sz'),
            ('其它梯队', '美的集团', '000333', 'sz'),
            ('其它梯队', '海尔智家', '600690', 'sh'),

            ('其它梯队', '招商银行', '600036', 'sh'),
            ('其它梯队', '宁波银行', '002142', 'sz'),
            ('其它梯队', '杭州银行', '600926', 'sh'),

            ('第二梯队', '中国平安', '601318', 'sh'),
            ('其它梯队', '万科A', '000002', 'sz'),
            ('其它梯队', '保利地产', '600048', 'sh'),

            ('其它梯队', '海螺水泥', '600585', 'sh'),
            ('其它梯队', '福耀玻璃', '600660', 'sh'),
            ('其它梯队', '上海机场', '600009', 'sh'),
            ('其它梯队', '中国中免', '601888', 'sh'),
            ('其它梯队', '海康威视', '002415', 'sz'),
            ('其它梯队', '恒生电子', '600570', 'sh'),
            ('其它梯队', '南极电商', '002127', 'sz'),
            ('其它梯队', '晨光文具', '603899', 'sh'),
            ('其它梯队', '深南电路', '002916', 'sz'),
            ('其它梯队', '中兴通讯', '000063', 'sz'),

            ('其它梯队', '中国铁建', '601186', 'sh'),
            ('其它梯队', '中国建筑', '601668', 'sh'),
            ]

month_data_headings = ['日期', '开盘价', '最高价', '最低价', '收盘价', '12个月均价', '24个月均价']
day_data_headings1 = ['日期', '开盘价', '最高价', '最低价', '收盘价',
                      'K12值',
                      'K24值',
                      ]
day_data_headings2 = ['日期', '开盘价', '最高价', '最低价', '收盘价',
                      '当日', '平均', '最低', '最高', '方差', '买入', '卖出', '卖出2',
                      '当日', '平均', '最低', '最高', '方差', '买入', '卖出', '卖出2',
                      ]
today_data_headings1 = ['日期', '梯队', '证券名称', '证券代码', 'K12值', 'K24值', '价格', '市值',
                        'K12值=1.1', 'K12值=1.0', 'K12值=0.9', 'K12值=1.5', 'K12值=1.6',
                        'K12值',
                        'K24值',
                        ]
today_data_headings2 = ['日期', '梯队', '证券名称', '证券代码', 'K12值', 'K24值', '价格', '市值',
                        'K12值=1.1', 'K12值=1.0', 'K12值=0.9', 'K12值=1.5', 'K12值=1.6',
                        '12个月均价', '当日', '平均', '最低', '最高', '方差', '买入', '卖出', '卖出2',
                        '24个月均价', '当日', '平均', '最低', '最高', '方差', '买入', '卖出', '卖出2',
                        ]

today_data_list = []

for index, (level, name, code, area) in enumerate(code_dic):
    print('get ' + name + code + ' data...')
    month_data_list = []
    day_data_list = []
    # 获取沪深A股历史K线数据
    # 详细指标参数，参见“历史行情指标参数”章节；“分钟线”参数与“日线”参数不同。“分钟线”不包含指数。
    # 分钟线指标：date,time,code,open,high,low,close,volume,amount,adjustflag
    # 周月线指标：date,code,open,high,low,close,volume,amount,adjustflag,turn,pctChg
    rs = bs.query_history_k_data_plus(str(area) + '.' + str(code),
                                      "date,open,high,low,close",
                                      start_date='1990-12-19', end_date='2030-12-31',
                                      frequency="m", adjustflag="2")
    print('query_history_month_k_data respond error_code:' + rs.error_code)
    print('query_history_month_k_data respond error_msg:' + rs.error_msg)
    while (rs.error_code == '0') & rs.next():
        month_data_list.append(rs.get_row_data())

    rs = bs.query_history_k_data_plus(str(area) + '.' + str(code),
                                      "date,open,high,low,close",
                                      start_date='1990-12-19', end_date='2030-12-31',
                                      frequency="d", adjustflag="2")
    print('query_history_day_k_data respond error_code:' + rs.error_code)
    print('query_history_day_k_data respond error_msg:' + rs.error_msg)
    while (rs.error_code == '0') & rs.next():
        day_data_list.append(rs.get_row_data())

    workbook = xlsxwriter.Workbook(name + str(code) + '.xlsx')
    workformat = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
    })
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

    average_month_12_dict = {}
    average_month_24_dict = {}

    homesheet_k12 = workbook.add_worksheet(name + str(code) + 'K12值择时图')
    homesheet_k24 = workbook.add_worksheet(name + str(code) + 'K24值择时图')

    worksheet = workbook.add_worksheet('month_data')
    worksheet.write_row('A1', month_data_headings, workformat)
    worksheet.set_column('A:H', 10)

    if day_data_list[-1][0] != month_data_list[-1][0]:
        month_data_list.append(day_data_list[-1])
    month_data_list = np.array(month_data_list)
    day_data_list = np.array(day_data_list)

    start_row = 1
    current_row = start_row
    for row_no, data in enumerate(month_data_list):
        if row_no >= 24:
            average_month_12 = np.average(np.array(month_data_list[row_no - 11:row_no + 1, 4], np.float))
            average_month_12_dict[data[0][0:7]] = average_month_12

            average_month_24 = np.average(np.array(month_data_list[row_no - 23:row_no + 1, 4], np.float))
            average_month_24_dict[data[0][0:7]] = average_month_24

            date = datetime.strptime(data[0], "%Y-%m-%d")
            worksheet.write_datetime(current_row, 0, date, date_format)
            worksheet.write_number(current_row, 1, np.round(float(data[1]), 2))
            worksheet.write_number(current_row, 2, np.round(float(data[2]), 2))
            worksheet.write_number(current_row, 3, np.round(float(data[3]), 2))
            worksheet.write_number(current_row, 4, np.round(float(data[4]), 2))
            worksheet.write_number(current_row, 5, np.round(float(average_month_12), 2))
            worksheet.write_number(current_row, 6, np.round(float(average_month_24), 2))
            current_row = current_row + 1

    worksheet = workbook.add_worksheet('day_data')
    worksheet.write_row('A2', day_data_headings2, workformat)
    worksheet.merge_range('A1:A2', day_data_headings1[0], workformat)
    worksheet.merge_range('B1:B2', day_data_headings1[1], workformat)
    worksheet.merge_range('C1:C2', day_data_headings1[2], workformat)
    worksheet.merge_range('D1:D2', day_data_headings1[3], workformat)
    worksheet.merge_range('E1:E2', day_data_headings1[4], workformat)
    worksheet.merge_range('F1:M1', day_data_headings1[5], workformat)
    worksheet.merge_range('N1:U1', day_data_headings1[6], workformat)
    worksheet.set_column('A:AZ', 10)

    start_row = 2
    current_row = start_row
    k_value_list = []
    last_k_value_list = []
    first_year = None
    last_year = int(day_data_list[-1][0][0:4])
    for row_no, data in enumerate(day_data_list):
        if int(data[0][0:4]) < last_year - 10:
            continue
        if first_year is None:
            first_year = int(data[0][0:4])

        average_month_12 = average_month_12_dict.get(data[0][0:7])
        average_month_24 = average_month_24_dict.get(data[0][0:7])
        if average_month_12 is not None and average_month_12 != 0 and \
                average_month_24 is not None and average_month_24 != 0:
            k12_value = float(data[4]) / average_month_12
            k24_value = float(data[4]) / average_month_24
            if row_no == len(day_data_list) - 1:
                today_data_list.append([data[0], level, name, code, data[4]])
                last_k_value_list = [average_month_12, k12_value, average_month_24, k24_value]

            date = datetime.strptime(data[0], "%Y-%m-%d")
            worksheet.write_datetime(current_row, 0, date, date_format)
            worksheet.write_number(current_row, 1, np.round(float(data[1]), 2))
            worksheet.write_number(current_row, 2, np.round(float(data[2]), 2))
            worksheet.write_number(current_row, 3, np.round(float(data[3]), 2))
            worksheet.write_number(current_row, 4, np.round(float(data[4]), 2))
            worksheet.write_number(current_row, 5, np.round(k12_value, 2))
            worksheet.write_number(current_row, 13, np.round(k24_value, 2))
            k_value_list.append([k12_value, k24_value])
            current_row = current_row + 1

    k_value_list = np.array(k_value_list)
    if len(k_value_list) == 0:
        continue

    for k in range(len(k_value_list[0])):
        average_month_12 = last_k_value_list[2 * k]
        last_k12_value = last_k_value_list[2 * k + 1]
        k12_min = np.min(k_value_list[:, k])
        k12_max = np.max(k_value_list[:, k])
        k12_avg = np.average(k_value_list[:, k])
        k12_std = np.std(k_value_list[:, k])
        k12_buy = k12_avg - k12_std
        k12_sell = k12_avg + k12_std
        k12_sell2 = k12_avg + 2 * k12_std
        today_data_list[-1].extend([average_month_12, last_k12_value, k12_avg, k12_min, k12_max,
                                    k12_std, k12_buy, k12_sell, k12_sell2])

        for i in range(start_row, current_row):
            worksheet.write_number(i, 6 + 8 * k, np.round(k12_avg, 2))

        worksheet.write_number(start_row, 7 + 8 * k, np.round(k12_min, 2))
        worksheet.write_number(start_row, 8 + 8 * k, np.round(k12_max, 2))
        worksheet.write_number(start_row, 9 + 8 * k, np.round(k12_std, 2))

        for i in range(start_row, current_row):
            worksheet.write_number(i, 10 + 8 * k, np.round(k12_buy, 2))
            worksheet.write_number(i, 11 + 8 * k, np.round(k12_sell, 2))
            worksheet.write_number(i, 12 + 8 * k, np.round(k12_sell2, 2))

    k12_chart = workbook.add_chart({'type': 'line'})

    k12_chart.add_series(
        {
            'name': '买入线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$K$' + str(start_row + 1) + ':$K$' + str(current_row),
            'line': {'color': 'green'},
        }
    )
    k12_chart.add_series(
        {
            'name': '平均线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$G$' + str(start_row + 1) + ':$G$' + str(current_row),
            'line': {'color': 'blue'},
        }
    )
    k12_chart.add_series(
        {
            'name': '卖出线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$L$' + str(start_row + 1) + ':$L$' + str(current_row),
            'line': {'color': 'purple'},
        }
    )
    k12_chart.add_series(
        {
            'name': '卖出线2',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$M$' + str(start_row + 1) + ':$M$' + str(current_row),
            'line': {'color': 'red'},
        }
    )
    k12_chart.add_series(
        {
            'name': 'K12值',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$F$' + str(start_row + 1) + ':$F$' + str(current_row),
            'line': {'color': 'black'},
        }
    )

    k12_chart.set_title({'name': name + str(code) + ' 近' + str(int(last_year) - int(first_year)) + '年K12值择时图'})
    k12_chart.set_x_axis({'visible': True,
                          'date_axis': True,
                          'major_unit_type': 'months', 'minor_unit_type': 'months',
                          'major_unit': 3, 'minor_unit': 3})
    k12_chart.set_y_axis({'visible': True,
                          'major_unit': 0.1, 'minor_unit': 0.1,
                          'min': 0.5})
    k12_chart.set_style(1)
    k12_chart.set_legend({
        'position': 'top',
        # 'none': True
    })
    k12_chart.set_size({'width': 2500, 'height': 600})

    homesheet_k12.insert_chart('A1', k12_chart, {'x_offset': 25, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

    k24_chart = workbook.add_chart({'type': 'line'})

    k24_chart.add_series(
        {
            'name': '买入线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$S$' + str(start_row + 1) + ':$S$' + str(current_row),
            'line': {'color': 'green'},
        }
    )
    k24_chart.add_series(
        {
            'name': '平均线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$O$' + str(start_row + 1) + ':$O$' + str(current_row),
            'line': {'color': 'blue'},
        }
    )
    k24_chart.add_series(
        {
            'name': '卖出线',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$T$' + str(start_row + 1) + ':$T$' + str(current_row),
            'line': {'color': 'purple'},
        }
    )
    k24_chart.add_series(
        {
            'name': '卖出线2',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$U$' + str(start_row + 1) + ':$U$' + str(current_row),
            'line': {'color': 'red'},
        }
    )
    k24_chart.add_series(
        {
            'name': 'K24值',
            'categories': '=day_data!$A$' + str(start_row + 1) + ':$A$' + str(current_row),
            'values': '=day_data!$N$' + str(start_row + 1) + ':$N$' + str(current_row),
            'line': {'color': 'black'},
        }
    )

    k24_chart.set_title({'name': name + str(code) + ' 近' + str(int(last_year) - int(first_year)) + '年K24值择时图'})
    k24_chart.set_x_axis({'visible': True,
                          'date_axis': True,
                          'major_unit_type': 'months', 'minor_unit_type': 'months',
                          'major_unit': 3, 'minor_unit': 3})
    k24_chart.set_y_axis({'visible': True,
                          'major_unit': 0.1, 'minor_unit': 0.1,
                          'min': 0.5})
    k24_chart.set_style(1)
    k24_chart.set_legend({
        'position': 'top',
        # 'none': True
    })
    k24_chart.set_size({'width': 2500, 'height': 600})

    homesheet_k24.insert_chart('A1', k24_chart, {'x_offset': 25, 'y_offset': 10, 'x_scale': 1, 'y_scale': 1})

    workbook.close()

workbook = xlsxwriter.Workbook('最新K值.xlsx')
workformat = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
cell_format_red = workbook.add_format({'bold': True, 'font_color': 'red'})
cell_format_purple = workbook.add_format({'bold': True, 'font_color': 'purple'})
cell_format_green = workbook.add_format({'bold': True, 'font_color': 'green'})

worksheet = workbook.add_worksheet('最新K值')
worksheet.write_row('A1', today_data_headings1, workformat)
worksheet.write_row('A2', today_data_headings2, workformat)
worksheet.merge_range('A1:A2', today_data_headings1[0], workformat)
worksheet.merge_range('B1:B2', today_data_headings1[1], workformat)
worksheet.merge_range('C1:C2', today_data_headings1[2], workformat)
worksheet.merge_range('D1:D2', today_data_headings1[3], workformat)
worksheet.merge_range('E1:E2', today_data_headings1[4], workformat)

worksheet.merge_range('F1:F2', today_data_headings1[5], workformat)
worksheet.merge_range('G1:G2', today_data_headings1[6], workformat)
worksheet.merge_range('H1:H2', today_data_headings1[7], workformat)
worksheet.merge_range('I1:I2', today_data_headings1[8], workformat)
worksheet.merge_range('J1:J2', today_data_headings1[9], workformat)
worksheet.merge_range('K1:K2', today_data_headings1[10], workformat)
worksheet.merge_range('L1:L2', today_data_headings1[11], workformat)
worksheet.merge_range('M1:M2', today_data_headings1[12], workformat)

worksheet.merge_range('N1:V1', today_data_headings1[13], workformat)
worksheet.merge_range('W1:AE1', today_data_headings1[14], workformat)
worksheet.set_column('A:AZ', 10)


today_data_list = np.array(today_data_list)
k12_value_sorted = np.argsort(today_data_list[:, 6])

start_row = 2
current_row = start_row
for row_no, data in enumerate(today_data_list[k12_value_sorted]):
    price = float(data[4])
    average_month_12 = float(data[5])
    average_month_24 = float(data[14])
    k12_value = price / average_month_12
    k24_value = price / average_month_24

    date = datetime.strptime(data[0], "%Y-%m-%d")
    worksheet.write_datetime(current_row, 0, date, date_format)
    worksheet.write_string(current_row, 1, data[1])
    worksheet.write_string(current_row, 2, data[2])
    worksheet.write_string(current_row, 3, data[3])

    if k12_value >= float(data[13]):
        worksheet.write_number(current_row, 4, np.round(k12_value, 2), cell_format_red)
    elif k12_value >= float(data[12]):
        worksheet.write_number(current_row, 4, np.round(k12_value, 2), cell_format_purple)
    elif k12_value <= float(data[11]) or (k12_value <= 1.0 and (data[1] == '第一梯队' or data[1] == '第二梯队')):
        worksheet.write_number(current_row, 4, np.round(k12_value, 2), cell_format_green)
    else:
        worksheet.write_number(current_row, 4, np.round(k12_value, 2))

    if k24_value >= float(data[22]):
        worksheet.write_number(current_row, 5, np.round(k24_value, 2), cell_format_red)
    if k24_value >= float(data[21]):
        worksheet.write_number(current_row, 5, np.round(k24_value, 2), cell_format_purple)
    elif k24_value <= float(data[20]):
        worksheet.write_number(current_row, 5, np.round(k24_value, 2), cell_format_green)
    else:
        worksheet.write_number(current_row, 5, np.round(k24_value, 2))

    worksheet.write_number(current_row, 6, np.round(price, 2))

    worksheet.write_number(current_row, 8, np.round(1.1 * average_month_12, 2))
    worksheet.write_number(current_row, 9, np.round(1.0 * average_month_12, 2))
    worksheet.write_number(current_row, 10, np.round(0.9 * average_month_12, 2))
    worksheet.write_number(current_row, 11, np.round(1.5 * average_month_12, 2))
    worksheet.write_number(current_row, 12, np.round(1.6 * average_month_12, 2))

    for i in range(len(data[5:])):
        worksheet.write_number(current_row, 13 + i, np.round(float(data[5 + i]), 2))
    current_row = current_row + 1

workbook.close()

# 登出系统
bs.logout()
