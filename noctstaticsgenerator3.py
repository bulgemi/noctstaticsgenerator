# _*_ coding: utf-8 _*_

import xlsxwriter
import re

data = list()


def create_excel():
    total_sum = 0
    total_avg = 0
    column_index = 0

    workbook = xlsxwriter.Workbook('noctstatics.xlsx')
    worksheet_sum = workbook.add_worksheet(u'통계 요약')
    worksheet = workbook.add_worksheet(u'상세 일자별 통계')
    chart = workbook.add_chart({'type': 'column'})
    chart.set_title({'name': u'일별 처리 건수'})
    chart.set_legend({'none': True})
    chart.set_size({'width': 720, 'height': 576})
    chart_categories = [u'통계 요약', ]
    chart_values = [u'통계 요약', ]

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
    worksheet_sum.write_string(0, column_index, u'일자', header_format)
    worksheet_sum.set_column(0, column_index, 0)

    for v in data:
        if len(v) < 50:
            break
        column_index += 1

        worksheet_sum.write_string(0, column_index, v[0], header_format)
        worksheet_sum.set_column(0, column_index, 20)
#        print(len(v), v)

    column_index += 1
    worksheet_sum.write_string(0, column_index, u'총 처리건수', header_format)
    worksheet_sum.set_column(0, column_index, 20)
    column_index += 1
    worksheet_sum.write_string(0, column_index, u'평균 처리건수', header_format)
    worksheet_sum.set_column(0, column_index, 20)

    column_index = 0
    worksheet_sum.write_string(1, column_index, u'처리건수', header_format)

    for v in data:
        daysum=0

        if len(v) < 50:
            break
        column_index += 1

        for c in v:
            if c == '':
                break

            cycledata_yn = re.findall(',', c)

            if len(cycledata_yn) == 0:
                continue

            cycledata = re.split(',', c)
            daysum += int(cycledata[2])
            print(c)

        worksheet_sum.write(1, column_index, daysum, number_format)
        total_sum += daysum

    total_avg = total_sum/(len(data)-1)

    #차트 카테고리 생생
    chart_categories.append(0)
    chart_categories.append(1)
    chart_categories.append(0)
    chart_categories.append(column_index)
    chart_values.append(1)
    chart_values.append(1)
    chart_values.append(1)
    chart_values.append(column_index)

    #차트 생성
    chart.add_series({'categories': chart_categories,
                      'values': chart_values,
                      })

    worksheet_sum.insert_chart('A7', chart)

    column_index += 1
    worksheet_sum.write(1, column_index, total_sum, number_format)
    column_index += 1
    worksheet_sum.write(1, column_index, total_avg, number_format)

    workbook.close()


if __name__ == "__main__":
    filename = input(u'메타 파일명: ')
    fullpath = 'C:/Users/Administrator/Desktop/' + filename

    print(fullpath)
    statfile = open(fullpath, 'r')

    while True:
        line = statfile.readline()
        val = line.split('|')
        data.append(val)

        if not line:
            break

#    print(data)
    print(len(data))

    statfile.close()
    create_excel()
