# _*_ coding: utf-8 _*_

import xlsxwriter


header = list()
data = list()


def create_excel():
    workbook = xlsxwriter.Workbook('noctstatics1w.xlsx')
    worksheet = workbook.add_worksheet(u'일자별 통계')

    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
    header_format.set_pattern(1)
    header_format.set_bg_color('gray')

    number_format = workbook.add_format({'border': 1})
    number_format.set_num_format('#,##0')

    max_number_format = workbook.add_format({'bold': True, 'border': 1})
    max_number_format.set_num_format('#,##0')

    sum_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
    sum_format.set_bg_color('gray')

    # print header
    idx = 0
    for v in range(0, len(header)):
        worksheet.write(0, idx, header[v], header_format)
        idx += 1

    worksheet.write_string(0, idx, u'차수별 총합', header_format)
    idx += 1
    worksheet.write_string(0, idx, u'차수별 평균', header_format)
    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, idx, 20)

    idx = 1
    for v in range(0, len(data)):
        col = 0
        data_tmp = data[v]
        number_data_tmp = [int(i) for i in data_tmp[2:]]
        max_value_index = number_data_tmp.index(max(number_data_tmp)) + 2
        cycle_sum = 0
        cycle_avg = 0

#        print(number_data_tmp)
#        print(max(number_data_tmp), number_data_tmp.index(max(number_data_tmp)))

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
    statdata = open('C:/Users/Administrator/Desktop/noct1week.data', 'r')

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
