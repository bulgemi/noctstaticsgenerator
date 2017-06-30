# _*_ coding: utf-8 _*_

import xlsxwriter


header = list()
data = list()


def create_excel():
    workbook = xlsxwriter.Workbook('noctstatics4w.xlsx')
    worksheet_sum = workbook.add_worksheet(u'통계 요약')
    worksheet = workbook.add_worksheet(u'상세 일자별 통계')
    chart = workbook.add_chart({'type': 'column'})
    chart.set_title({'name': u'일별 처리 건수'})
    chart.set_legend({'none': True})
    chart.set_size({'width': 720, 'height': 576})
    chart_categories = [u'상세 일자별 통계', ]
    chart_values = [u'상세 일자별 통계', ]

    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
    header_format.set_pattern(1)
    header_format.set_bg_color('gray')

    number_format = workbook.add_format({'border': 1})
    number_format.set_num_format('#,##0')

    max_number_format = workbook.add_format({'bold': True, 'border': 1})
    max_number_format.set_num_format('#,##0')

    sum_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
    sum_format.set_bg_color('gray')

    # 통계 요약
    worksheet_sum.write_string(0, 0, u'주차', header_format)
    worksheet_sum.write_string(0, 1, u'1주 [' + header[2] + '~' + header[8] + ']', header_format)
    worksheet_sum.write_string(0, 2, u'2주 [' + header[9] + '~' + header[15] + ']', header_format)
    worksheet_sum.write_string(0, 3, u'3주 [' + header[16] + '~' + header[22] + ']', header_format)
    worksheet_sum.write_string(0, 4, u'4주 [' + header[23] + '~' + header[29] + ']', header_format)
    worksheet_sum.write_string(0, 5, u'총 처리건수', header_format)
    worksheet_sum.write_string(0, 6, u'평균 처리건수', header_format)
    worksheet_sum.write_string(1, 0, u'처리건수', header_format)
    worksheet_sum.set_column(0, 0, 40)
    worksheet_sum.set_column(0, 1, 40)
    worksheet_sum.set_column(0, 2, 40)
    worksheet_sum.set_column(0, 3, 40)
    worksheet_sum.set_column(0, 4, 40)
    worksheet_sum.set_column(0, 5, 40)
    worksheet_sum.set_column(0, 6, 40)

    # print header
    idx = 0
    for v in range(0, len(header)):
        worksheet.write(0, idx, header[v], header_format)
        idx += 1

    chart_categories.append(0)
    chart_categories.append(2)
    chart_categories.append(0)
    chart_categories.append(idx-1)

#    print(chart_categories)

    worksheet.write_string(0, idx, u'차수별 총합', header_format)
    idx += 1
    worksheet.write_string(0, idx, u'차수별 평균', header_format)
    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, idx, 20)

#    print(len(data), data)

    idx = 1
    for v in range(0, len(data)):
        col = 0
        data_tmp = data[v]
        number_data_tmp = [int(i) for i in data_tmp[2:]]
        max_value_index = number_data_tmp.index(max(number_data_tmp)) + 2
        cycle_sum = 0
        cycle_avg = 0

        for e in range(0, len(data[v])):
            if e < 2:
                worksheet.write(idx, col, data_tmp[e])
            else:
                if e == max_value_index:
                    worksheet.write(idx, col, int(data_tmp[e]), max_number_format)
                else:
                    worksheet.write(idx, col, int(data_tmp[e]), number_format)
                cycle_sum += int(data_tmp[e])
                cycle_avg += 1
            col += 1

        worksheet.write(idx, col, cycle_sum, number_format)
        col += 1
        worksheet.write(idx, col, cycle_sum/(cycle_avg+1), number_format)
        idx += 1

    chart_values.append(idx)
    chart_values.append(2)
    chart_values.append(idx)
    chart_values.append(col-2)

#    print(chart_values)

    worksheet.conditional_format(1, col-1, idx-1, col-1, {'type': 'data_bar'})
    worksheet.conditional_format(1, col, idx-1, col, {'type': 'data_bar'})
    worksheet.conditional_format(idx, 2, idx, col-2, {'type': 'data_bar'})
    worksheet.merge_range(idx, 0, idx, 1, u'일자별 합계', sum_format)

    day_tmp = data[0]
    day_sum = day_tmp[2:]
    total_sum = 0
    total_idx = 0

    for v in range(1, len(data)):
        data_tmp = data[v]

        for e in range(2, len(data[v])):
            sum_tmp = int(day_sum[e-2]) + int(data_tmp[e])
            day_sum[e-2] = sum_tmp

    for e in range(0, len(day_sum)):
        worksheet.write(idx, e+2, day_sum[e], number_format)
        total_sum += int(day_sum[e])
        total_idx = e + 2

    week4 = [0, 0, 0, 0]
    week4[0] = sum(int(w1) for w1 in day_sum[0:7])
    week4[1] = sum(int(w1) for w1 in day_sum[7:14])
    week4[2] = sum(int(w1) for w1 in day_sum[14:21])
    week4[3] = sum(int(w1) for w1 in day_sum[21:28])

    chart.add_series({'categories': chart_categories,
                      'values': chart_values,
                      })

    worksheet_sum.insert_chart('A7', chart)

    worksheet_sum.write(1, 1, week4[0], number_format)
    worksheet_sum.write(1, 2, week4[1], number_format)
    worksheet_sum.write(1, 3, week4[2], number_format)
    worksheet_sum.write(1, 4, week4[3], number_format)
    worksheet_sum.write(1, 5, sum(week4), max_number_format)
    worksheet_sum.write(1, 6, sum(week4)/4, max_number_format)
    worksheet_sum.conditional_format(1, 1, 1, 4, {'type': 'data_bar'})

    worksheet.merge_range(idx, total_idx+1, idx, total_idx+2, total_sum, max_number_format)
    workbook.close()


def set_header(value):
    global header
    header_length = len(value)

    if len(header) == 0:
        for v in range(1, header_length-1):
            header.append(value[v])
    else:
        for v in range(3, header_length-1):
            header.append(value[v])


def set_data(value):
    global data

    data_tmp = list()
    data_length = len(value)

    if len(data) < 48:
        for v in range(1, data_length-1):
            data_tmp.append(value[v])
        data.append(data_tmp)
    else:
        for v in range(3, data_length-1):
            data[int(value[1])-1].append(str(value[v]))

    #print(data)


if __name__ == "__main__":
    statdata = open('C:/Users/Administrator/Desktop/noct4week.data', 'r')

    while True:
        line = statdata.readline()
        val = line.split('|')

        if val[0] == 'H':
            set_header(val)
        elif val[0] == 'D':
            set_data(val)
        else:
            pass

        if not line:
            break

    statdata.close()
    create_excel()
