#!/usr/bin/env python
# -*- coding: utf-8 -*-

import calendar
import time
from Tkinter import Tk
from tkFileDialog import askopenfilename

from lxml import etree
import re
import xlsxwriter
import os


# params
font = 'Arial narrow'
size = 9
margin_width = 4.5
page_left_margin = 0.2
page_right_margin = 0.2
page_top_margin = 0.2
page_bottom_margin = 0.2
a_width = 11.25
b_width = 14.45
c_width = 14.45
row_height = 11.9
symbols = ['DW', 'FS', 'FZ', 'KP', 'KW', 'FSE', 'Rach', 'Rach.', 'FZP', 'K', 'N', 'Z', 'RK', 'R', 'na']
rows_per_page = 68


def main():
    Tk().withdraw()
    source = askopenfilename()
    fix = 'fix.xml'
    fix_file(source, fix)

    doc = etree.parse(fix)
    root = doc.getroot()
    document = root.find('document')
    date = get_date(document)
    month = date[0]
    year = date[1]
    decrees = []
    for table in document:
        if table.tag == 'table':
            decree = parse_table(table, month, year)
            if decree is not None:
                decrees.append(decree)
    os.remove(fix)

    workbook = xlsxwriter.Workbook('dekret_' + str(calendar.timegm(time.gmtime())) + '.xlsx')
    workbook.formats[0].set_font_size(size)
    workbook.formats[0].set_font(font)

    money_formatting = get_money_formatting(workbook, font, size)
    formatting = get_formatting(workbook, font, size)

    columns = [['A', 'B', 'C'], ['E', 'F', 'G'], ['I', 'J', 'K']]
    start = 1
    worksheet = workbook.add_worksheet('Arkusz 1')
    worksheet.set_margins(page_left_margin, page_right_margin, page_top_margin, page_bottom_margin)
    worksheet.set_default_row(row_height)
    worksheet.set_column('A:A', a_width)
    worksheet.set_column('B:B', b_width)
    worksheet.set_column('C:C', c_width)
    worksheet.set_column('D:D', margin_width)
    worksheet.set_column('E:E', a_width)
    worksheet.set_column('F:F', b_width)
    worksheet.set_column('G:G', c_width)
    worksheet.set_column('H:H', margin_width)
    worksheet.set_column('I:I', a_width)
    worksheet.set_column('J:J', b_width)
    worksheet.set_column('K:K', c_width)
    max_rows_from_three_operations = 0
    page_breaks = [0]
    for idx, decree in enumerate(decrees):
        if idx != 0 and idx % 3 == 0:
            start = start + rows_per_page / 4 * ((max_rows_from_three_operations + 3) / (rows_per_page / 4 - 1) + 1)
            max_rows_from_three_operations = 0
        max_rows_from_three_operations = max(len(decree['rows']), max_rows_from_three_operations)
        if (start - page_breaks[-1]) / rows_per_page != (
                start - page_breaks[-1] + len(decree['rows']) + 1) / rows_per_page:
            page_breaks.append(start - 1)
        if start - page_breaks[-1] > rows_per_page:  # adjust start to page start
            start = page_breaks[-1] + rows_per_page + 1
            page_breaks.append(start - 1)
        write_decree(worksheet, decree, start, columns[idx % 3], formatting, money_formatting)
    worksheet.set_h_pagebreaks(page_breaks)
    workbook.close()


def write_decree(worksheet, decree, start, columns, formatting, money_formatting):
    worksheet.write(columns[0] + str(start), 'Data:', formatting)
    worksheet.merge_range(columns[1] + str(start) + ':' + columns[2] + str(start), decree['date'], formatting)
    worksheet.write(columns[0] + str(start + 1), 'Nr dowodu', formatting)
    worksheet.merge_range(columns[1] + str(start + 1) + ':' + columns[2] + str(start + 1), decree['number'], formatting)
    worksheet.write(columns[0] + str(start + 2), 'Konto', formatting)
    worksheet.write(columns[1] + str(start + 2), 'Wn', formatting)
    worksheet.write(columns[2] + str(start + 2), 'Ma', formatting)
    for idx, row in enumerate(decree['rows']):
        worksheet.write(columns[0] + str(start + 3 + idx), row['account'], formatting)
        worksheet.write(columns[1] + str(start + 3 + idx), get_currency_value(row['wn']), money_formatting)
        worksheet.write(columns[2] + str(start + 3 + idx), get_currency_value(row['ma']), money_formatting)


def get_formatting(workbook, font, size):
    formatting = workbook.add_format()
    formatting.set_border()
    formatting.set_font(font)
    formatting.set_align('center')
    formatting.set_font_size(size)
    return formatting


def get_money_formatting(workbook, font, size):
    formatting = workbook.add_format(
        {'num_format': '_-* #,##0.00 "zł"_-;-* #,##0.00 "zł"_-;_-* "-"?? "zł"_-;_-@_-'.decode('utf-8')})
    formatting.set_border()
    formatting.set_font(font)
    formatting.set_align('center')
    formatting.set_font_size(size)
    return formatting


def get_currency_value(raw_number):
    if raw_number is None:
        return raw_number
    number = float(raw_number.replace(',', '.'))
    return number


def parse_table(table, month, year):
    decree = {'date': None, 'number': None, 'rows': []}

    for idx, row in enumerate(table):
        if row.tag == 'row':
            if idx == 0 and len(row) != 10:
                return None  # incorrect table
            if row[0][0].text == 'Lp.' or idx >= len(table) - 1:  # skip header and last row
                continue
            if decree['date'] is None:
                decree['date'] = year + '-' + month + '-' + row[1][0].text.split('.')[0]
            if decree['number'] is None:
                decree['number'] = row[2][0].text
            decree_row = {'account': row[7][0].text, 'wn': row[8][0].text, 'ma': row[9][0].text}
            decree['rows'].append(decree_row)
            if row[5][0].text not in symbols:
                return None
    return decree


def fix_file(file1, file2):
    f1 = open(file1, 'r')
    f2 = open(file2, 'w')
    for line in f1:
        f2.write(line.replace('&d', '').replace('&t', '').replace('&p', '').replace('&P', ''))
    f1.close()
    f2.close()


def get_date(document):
    for child in document:
        if child.tag == 'utext' and child.text is not None:
            m = re.match('[^\d]*(\d+)\.(\d+)', child.text)
            if m is not None and len(m.groups()) == 2:
                return m.group(1), m.group(2)


if __name__ == '__main__':
    main()
