import pandas as pd
import json
import xlrd
import datetime
from ast import literal_eval

EXCEL_FILE = 'example_data.xlsm'
JSON_FILE = 'json_data.json'


def save_to_json(input_file, out_file):

    sheets = pd.ExcelFile(input_file).sheet_names
    json_data = {}
    for sheet in sheets:
        df = pd.read_excel(input_file, index_col=None, header=0,
                           sheet_name=sheet, na_values='null')

        if sheet == 'loco_26':
            # Определяем datemode для листа loco_26
            book = xlrd.open_workbook(EXCEL_FILE)
            xl_sheet = book.sheet_by_name(sheet)
            datemode = xl_sheet.book.datemode
            # Преобразуем поле с датой в datetime
            df['REPAIR_DATE'] = \
                pd.to_datetime(df['REPAIR_DATE']
                               .map(lambda x: datetime.datetime(*xlrd.xldate_as_tuple(x, datemode))))

        if sheet == 'acts_31L':
            # Преобразуем json-строки в словари.
            df['IT_SECTIONS'] = df['IT_SECTIONS'].apply(lambda x: literal_eval(str(x)))
            df['IT_INVENT'] = df['IT_INVENT'].apply(lambda x: literal_eval(str(x)))

        json_data[sheet] = json.loads(df.to_json(orient='records', force_ascii=False, date_format='iso'))
    with open(out_file, 'w',  encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=2, ensure_ascii=False)


save_to_json(EXCEL_FILE, JSON_FILE)
