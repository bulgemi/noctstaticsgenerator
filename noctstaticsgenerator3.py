# _*_ coding: utf-8 _*_
"""
    특정 기간동안 처리한 건수를 이용하여 일별 차수별 통계 데이터를 엑셀 포맷으로 생성하는 도구.
"""
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
    chart.set_size({'width': 1200, 'height': 600})
    chart_categories = [u'통계 요약', ]
    chart_values = [u'통계 요약', ]

    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
    header_format.set_pattern(1)
    header_format.set_bg_color('gray')

    date_format = workbook.add_format({'bold': False, 'align': 'left', 'border': 1})

    number_format = workbook.add_format({'border': 1})
    number_format.set_num_format('#,##0')

    number_format2 = workbook.add_format({'bold': True, 'border': 1})
    number_format2.set_num_format('#,##0')

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
        daysum = 0

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
#            print(c)

        worksheet_sum.write(1, column_index, daysum, number_format)
        total_sum += daysum

    total_avg = total_sum / (len(data) - 1)

    # 차트 카테고리 생생
    chart_categories.append(0)
    chart_categories.append(1)
    chart_categories.append(0)
    chart_categories.append(column_index)
    chart_values.append(1)
    chart_values.append(1)
    chart_values.append(1)
    chart_values.append(column_index)

    # 차트 생성
    chart.add_series({'categories': chart_categories,
                      'values': chart_values,
                      })

    worksheet_sum.insert_chart('A7', chart)

    column_index += 1
    worksheet_sum.write(1, column_index, total_sum, number_format2)
    column_index += 1
    worksheet_sum.write(1, column_index, total_avg, number_format2)

    # 항목별 목차 생성.
    # 통계 상세 타이틀
    column_index = 0
    worksheet.write_string(0, column_index, u'일자', header_format)
    worksheet.set_column(0, column_index, 20)
    cycleinfolist = data[0]

    # 차수 컬럼 생성.
    for cycleinfo in cycleinfolist:
        cycleinfo_yn = re.findall(',', cycleinfo)

        if len(cycleinfo_yn) == 0:
            continue

        column_index += 1

        cycle_time_list = re.split(',', cycleinfo)
        cycle_time = cycle_time_list[0] + u" 차수(" + cycle_time_list[1] + ")"
        worksheet.write_string(0, column_index, cycle_time, header_format)
        worksheet.set_column(0, column_index, 20)

    # 일별 총합 컬럼 생성.
    column_index += 1
    worksheet.write_string(0, column_index, u"일별 합계", header_format)
    worksheet.set_column(0, column_index, 20)

    # 차수 평균 컬럼 생성.
    column_index += 1
    worksheet.write_string(0, column_index, u"차수 평균", header_format)
    worksheet.set_column(0, column_index, 20)

    # 일자별 처리 데이터 생성.
    row_index = 1 # row 인덱스 초기화.

    # 차수별 처리건수 합계 리스트 생성.
    cyclesum = [0 for _ in range(48)]
#    print(cyclesum)

    for v in data:
        if len(v) < 50:
            break

        daysum = 0
        column_index = 0 # 컬럼 인덱스 초기화.

        for c in v:
            if column_index == 49:
                worksheet.write(row_index, column_index, daysum, number_format2)
                worksheet.write(row_index, column_index+1, daysum/48, number_format2)

            if len(c) <= 1:
                break

            cycledata_yn = re.findall(',', c)

            if len(cycledata_yn) == 0:
                # 일자 컬럼 생성.
                worksheet.write_string(row_index, column_index, c, date_format)
            else:
                # 차수별 처리 건수 생성.
                cycledata = re.split(',', c)
                daysum += int(cycledata[2])
                cyclesum[column_index-1] += int(cycledata[2])
                worksheet.write(row_index, column_index, int(cycledata[2]), number_format)

            column_index += 1 # 컬럼 인덱스 증가.


        row_index += 1 # row 인덱스 증가.

    # 최대 처리 건수 차수 추출.

    # 차수 평균 컬럼 생성.
    column_index = 0
    worksheet.write_string(row_index, column_index, u"차수 합계", header_format)

    total_sum = 0

    for c in cyclesum:
        column_index += 1
        worksheet.write(row_index, column_index, c, number_format2)
        total_sum += c

    column_index += 1
    worksheet.write_string(row_index, column_index, u"총합", header_format)
    column_index += 1
    worksheet.write(row_index, column_index, total_sum, number_format2)

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
#    print(len(data))

    statfile.close()
    create_excel()
