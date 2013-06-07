# -*- coding: utf-8 -*-

import xlrd
import json


rb = xlrd.open_workbook('Tabl-35-12.xls', formatting_info=True)
font = rb.font_list
sheet = rb.sheet_by_index(1)

rows_number = sheet.nrows

peoples_dict = {}  # main dict
region = ''
raion = ''
for rownum in range(7, rows_number):

    cell = sheet.cell(rownum, 0)
    value = cell.value.strip()  # delete spaces at start and end
    peoples_count = sheet.cell(rownum, 1).value
    if peoples_count == 0 or peoples_count == '':  # empty row - continue
        continue
    peoples_count = int(sheet.cell(rownum, 1).value)  # from 12313.0 to integer

    cell_format = rb.xf_list[cell.xf_index]
    bold = bool(font[cell_format.font_index].bold)
    italic = bool(font[cell_format.font_index].italic)
    indent = cell_format.alignment.indent_level

    is_region = bold and not italic
    is_raion = bold and italic
    is_municipal = (indent == 2)

    # use this - replace('\n', ' ') - only because in one string exist '\n', that not needed
    if is_region:
        region = value.replace('\n', ' ')
        peoples_dict[region] = {'count': peoples_count}
    elif is_raion:
        raion = value.replace('\n', ' ')
        peoples_dict[region][raion] = {'count': peoples_count}
    elif is_municipal:
        municipal = value.replace('\n', ' ')
        peoples_dict[region][raion][municipal] = {'count': peoples_count}

# print peoples_dict['Московская область']['Истринский муниципальный район']['Городское поселение Истра']['count']

with open('peoples.json', 'w') as outfile:
    json.dump(peoples_dict, outfile)

