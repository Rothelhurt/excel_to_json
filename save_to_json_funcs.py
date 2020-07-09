import pandas as pd
import json
import xlrd
import datetime
import os
import platform
from ast import literal_eval


def define_slash():
    """
    Устанавливает правильный слеш в
    зависимости от операционной системы.
    """
    if platform.platform().startswith('Windows'):
        return '\\'
    else:
        return '/'


def is_path(file_path):
    """
    Проверяет, является-ли
    введенная строка директорией.
    """
    if not file_path.endswith(r'.*') and os.path.isdir(file_path):
        return True
    else:
        return False


def file_is_valid(file):
    """Проверка на возможность создания указанного файла."""
    try:
        check_file = open(file, 'w', encoding='utf-8')
        check_file.close()
        os.remove(file)
    except FileNotFoundError:
        return False
    if not file.endswith('.json'):
        return False
    return True


def define_outfile(out_file, in_file, in_f_path, slash):
    """
    Задает имя выходного файла, в
    зависимости от введенных данных.
    """
    if slash not in out_file:
        out_f_path = in_f_path
        out_file = out_f_path + out_file
    elif is_path(out_file):
        out_f_path = out_file
        out_file = out_f_path + (in_file.rsplit('.', 1)[0] +
                                 '.json').rsplit(slash, 1)[1]
    return out_file



def file_exists(file):
    """
    Проверяет, существует ли файл.
    Принимает строку с путем и именем файла.
    """
    return os.path.exists(file) and os.path.isfile(file)


def save_to_json(input_file, out_file):
    """
    Считывает данные из excel-файла <input_file>.
    Сохраняет данные в json-файл <out_file>.
    """
    sheets = pd.ExcelFile(input_file).sheet_names
    json_data = {}
    for sheet in sheets:
        df = pd.read_excel(input_file, index_col=None, header=0,
                           sheet_name=sheet, na_values='null')

        if sheet == 'loco_26':
            # Определяем datemode для листа loco_26
            book = xlrd.open_workbook(input_file)
            xl_sheet = book.sheet_by_name(sheet)
            datemode = xl_sheet.book.datemode
            # Преобразуем поле с датой в datetime
            df['REPAIR_DATE'] = \
                pd.to_datetime(df['REPAIR_DATE']
                               .map(lambda x: datetime.datetime(*xlrd.
                                                                xldate_as_tuple(x, datemode))))

        if sheet == 'acts_31L':
            # Преобразуем json-строки в словари.
            df['IT_SECTIONS'] = df['IT_SECTIONS'].apply(lambda x: literal_eval(str(x)))
            df['IT_INVENT'] = df['IT_INVENT'].apply(lambda x: literal_eval(str(x)))

        json_data[sheet] = json.loads(df.to_json(orient='records', force_ascii=False,
                                                 date_format='iso'))
    with open(out_file, 'w',  encoding='utf-8') as json_file:
        json.dump(json_data, json_file, indent=2, ensure_ascii=False)
