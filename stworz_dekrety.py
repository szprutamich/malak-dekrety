#!/usr/bin/env python
# -*- coding: utf-8 -*-

import calendar
import time
from Tkinter import *
import tkFileDialog

from datetime import datetime
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
page_bottom_margin = 0.1
a_width = 11.25
b_width = 14.45
c_width = 14.45
row_height = 13.9
excluded_symbols = ['PK', 'WB']
rows_per_page = 68

# vars
file_name = None
main_window, top, center, bottom = None, None, None, None
decrees = []
symbols = []
min_date = None
max_date = None
input_min = None
input_max = None


def main():
    global main_window, top, center, bottom
    main_window = Tk()
    main_window.winfo_toplevel().title('Dekrety')

    top = Frame(main_window)
    top.pack(side=TOP, fill=X)
    center = Frame(main_window)
    center.pack(side=TOP, fill=X)
    bottom = Frame(main_window)
    bottom.pack(side=BOTTOM, fill=X)

    Label(top, width=2).pack(side=LEFT)

    choose_button = Button(top, height=1, text='Wybierz plik: ')
    choose_button.configure(command=lambda: process_file(file_label, choose_button))
    choose_button.pack(side=LEFT)
    Label(top, width=2).pack(side=LEFT)
    file_label = Label(top, height=2)
    file_label.pack(side=LEFT)
    Label(top, width=2).pack(side=LEFT)

    main_window.mainloop()


def process_file(file_label, choose_button):
    choose_button.configure(state=DISABLED)
    global file_name, decrees
    decrees = []
    file_name = tkFileDialog.askopenfilename(filetypes=[('Plik xml', '*.xml')], title='Wybierz plik')
    print(file_name)
    file_label.config(text=file_name)
    fix = 'fix.xml'
    fix_file(file_name, fix)

    doc = etree.parse(fix)
    root = doc.getroot()
    document = root.find('document')
    date = get_date(document)
    month = date[0]
    year = date[1]
    for table in document:
        if table.tag == 'table':
            decree = parse_table(table, month, year)
            if decree is not None:
                decrees.append(decree)
    os.remove(fix)
    show_rest_of_the_controls()


def show_rest_of_the_controls():
    global input_min, input_max
    convert_button = Button(top, height=1, text='Konwertuj plik ', command=lambda: convert_file())
    convert_button.pack(side=LEFT)
    Label(top, width=2).pack(side=LEFT)
    Label(center, width=2).pack(side=LEFT)
    symbols_label = Label(center, text='Symbole:')
    symbols_label.pack(side=LEFT)
    for symbol in symbols:
        check = Checkbutton(center, text=symbol[0], variable=symbol[1], height=2, command=lambda: check)
        check.pack(side=LEFT)
        if symbol[0] not in excluded_symbols:
            check.select()
    Label(center, width=2).pack(side=LEFT)

    Label(bottom, width=2).pack(side=LEFT)
    Label(bottom, text='Od:', height=2).pack(side=LEFT)
    Label(bottom, width=2).pack(side=LEFT)
    input_min = Entry(bottom)
    input_min.insert(0, min_date.strftime("%d.%m.%Y"))
    input_min.pack(side=LEFT)
    Label(bottom, width=2).pack(side=LEFT)
    Label(bottom, text='Do:', height=2).pack(side=LEFT)
    Label(bottom, width=2).pack(side=LEFT)
    input_max = Entry(bottom)
    input_max.insert(0, max_date.strftime("%d.%m.%Y"))
    input_max.pack(side=LEFT)
    Label(bottom, width=2).pack(side=LEFT)


def convert_file():
    workbook = xlsxwriter.Workbook('dekret_' + str(calendar.timegm(time.gmtime())) + '.xlsx')
    workbook.formats[0].set_font_size(size)
    workbook.formats[0].set_font(font)

    money_formatting = get_money_formatting(workbook)
    formatting = get_formatting(workbook)

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
    for idx, decree in enumerate(filter_decrees()):
        if idx != 0 and idx % 3 == 0:
            start = start + rows_per_page / 4 * ((max_rows_from_three_operations + 3) / (rows_per_page / 4 - 1) + 1)
            if (start - page_breaks[-1]) > rows_per_page / 2:
                start = start + 1
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
    main_window.destroy()


def filter_decrees():
    filtered_decrees = []
    final_min_date = parse_date(input_min.get())
    final_max_date = parse_date(input_max.get())
    dict_symbols = dict(symbols)
    for decree in decrees:
        input_date = parse_date(decree['input_date'])
        if dict_symbols.get(decree['symbol'].upper()).get() and final_min_date <= input_date <= final_max_date:
            filtered_decrees.append(decree)
    return filtered_decrees


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


def get_formatting(workbook):
    formatting = workbook.add_format()
    formatting.set_border()
    formatting.set_font(font)
    formatting.set_align('center')
    formatting.set_font_size(size)
    return formatting


def get_money_formatting(workbook):
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
    decree = {'symbol': None, 'date': None, 'input_date': None, 'number': None, 'rows': []}
    if int(table.attrib['cols']) == 5:
        parse_summary_table(table)
    if int(table.attrib['cols']) != 10:
        return None
    for idx, row in enumerate(table):
        if row.tag == 'row':
            if row[0][0].text == 'Lp.' or idx >= len(table) - 1:  # skip header and last row
                continue
            if decree['date'] is None:
                decree['date'] = year + '-' + month + '-' + row[1][0].text.split('.')[0]
            if decree['input_date'] is None or decree['input_date'] == '':
                decree['input_date'] = row[3][0].text.strip()
                update_min_max_date(decree['input_date'])
            if decree['number'] is None:
                decree['number'] = row[2][0].text
            decree_row = {'account': row[7][0].text, 'wn': row[8][0].text, 'ma': row[9][0].text}
            decree['rows'].append(decree_row)
            if decree['symbol'] is None:
                if row[5][0].text is not None:
                    decree['symbol'] = row[5][0].text.upper()
                else:
                    decree['symbol'] = ""
    return decree


def parse_summary_table(table):
    for idx, row in enumerate(table):
        if row.tag == 'row':
            if idx < 2 or idx >= len(table) - 1:  # skip irrelevant rows
                continue
            if row[1][0].text is not None:
                symbols.append((row[1][0].text.upper(), BooleanVar()))
            else:
                symbols.append(("", BooleanVar()))


def update_min_max_date(date):
    global min_date, max_date
    dt = parse_date(date)
    if dt is None:
        return
    if min_date is None or min_date > dt:
        min_date = dt
    if max_date is None or max_date < dt:
        max_date = dt


def parse_date(date):
    try:
        return datetime.strptime(date, '%d.%m.%Y')
    except ValueError:
        return None


def fix_file(file1, file2):
    f1 = open(file1, 'r')
    f2 = open(file2, 'w')
    for line in f1:
        f2.write(re.sub(r"&#?[a-zA-Z\d]*;?", "", line))
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
