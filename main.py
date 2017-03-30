import os
import datetime

import xlrd
import xlsxwriter

final = []
workbook = xlrd.open_workbook('orbitz.xlsx')
workbook = xlrd.open_workbook('orbitz.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []  # The row where we stock the name of the column
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
orbitz = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    orbitz.append(elm)
for f in orbitz:
    print(f)

workbook = xlrd.open_workbook('downloaded.xlsx')
workbook = xlrd.open_workbook('downloaded.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []  # The row where we stock the name of the column
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
downloaded = []
for row in range(1, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
        downloaded.append(elm)
for f in downloaded:
    print(f)

for line in orbitz:
    found = False
    for row in downloaded:
        if not found:
            if line['sail_date'] == row['SailDate'] and line['return_date'] == row['ReturnDate']:
                # compare prices
                pass
            else:
                pass
    if not found:
        final.append([['Celebrity', line['ship'], str(line['sail_date']).split('.')[0], str(line['return_date']).split('.')[0], str(line['Interior']).split('.')[0], str(line['Oceanview']).split('.')[0],
                       str(line['Balcony']).split('.')[0], str(line['Suite']).split('.')[0]], ['cyan', 'cyan', 'cyan', 'cyan']])


def write_file_to_excell(data_array):

    final_workbook = xlsxwriter.Workbook('final.xlsx')

    final_worksheet = final_workbook.add_worksheet()
    bold = final_workbook.add_format({'bold': True})
    final_worksheet.set_column("A:A", 15)
    final_worksheet.set_column("B:B", 25)
    final_worksheet.set_column("C:C", 10)
    final_worksheet.set_column("D:D", 25)
    final_worksheet.set_column("E:E", 20)
    final_worksheet.set_column("F:F", 30)
    final_worksheet.set_column("G:G", 20)
    final_worksheet.set_column("H:H", 50)
    final_worksheet.set_column("I:I", 20)
    final_worksheet.set_column("J:J", 20)
    final_worksheet.set_column("K:K", 20)
    final_worksheet.set_column("L:L", 20)
    final_worksheet.set_column("M:M", 25)
    final_worksheet.set_column("N:N", 20)
    final_worksheet.set_column("O:O", 20)
    final_worksheet.write('A1', 'Company', bold)
    final_worksheet.write('B1', 'Ship', bold)
    final_worksheet.write('C1', 'SailDate', bold)
    final_worksheet.write('D1', 'Return_date', bold)
    final_worksheet.write('L1', 'InteriorBucketPrice', bold)
    final_worksheet.write('M1', 'OceanViewBucketPrice', bold)
    final_worksheet.write('N1', 'BalconyBucketPrice', bold)
    final_worksheet.write('O1', 'SuiteBucketPrice', bold)
    date_format = final_workbook.add_format({'bold': True})
    date_format.set_align("center")
    row_count = 1
    for ship_entry in data_array:
        entry = ship_entry[0]
        formats = ship_entry[1]
        column_count = 0
        format_count = 0
        for i in range(0, len(entry)):
            if column_count == 0:
                final_worksheet.write_string(row_count, column_count, entry[i], bold)
            elif column_count == 1:
                final_worksheet.write_string(row_count, column_count, entry[i], bold)
            elif column_count == 2:
                date_time = datetime.datetime.strptime(str(entry[i]), "%m/%d/%Y")
                worksheet.write_datetime(row_count, column_count, date_time, date_format)
            elif column_count == 3:
                date_time = datetime.datetime.strptime(str(entry[i]), "%m/%d/%Y")
                worksheet.write_datetime(row_count, column_count, date_time, date_format)
            elif column_count == 4:
                color = formats[format_count]
                date_format.set_bg_color(color)
                final_worksheet.write_number(row_count, column_count, entry[i], date_format)
                date_format.set_bg_color('white')
                format_count += 1
            elif column_count == 5:
                color = formats[format_count]
                date_format.set_bg_color(color)
                final_worksheet.write_number(row_count, column_count, entry[i], date_format)
                date_format.set_bg_color('white')
                format_count += 1
            elif column_count == 6:
                color = formats[format_count]
                date_format.set_bg_color(color)
                final_worksheet.write_number(row_count, column_count, entry[i], date_format)
                date_format.set_bg_color('white')
                format_count += 1
            elif column_count == 7:
                color = formats[format_count]
                date_format.set_bg_color(color)
                final_worksheet.write_number(row_count, column_count, entry[i], date_format)
                date_format.set_bg_color('white')
                format_count += 1
            column_count += 1
        row_count += 1
    final_workbook.close()
    pass
write_file_to_excell(final)