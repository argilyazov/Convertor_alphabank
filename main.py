import pandas as pd
import numpy as np
import openpyxl


def get_func(command):
    if command == "SPLIT":
        return split_column
    elif command == "ZIP":
        return zip_columns
    else:  # command == "RENAME":
        return empty_method


def split_column(df):
    splitter = ' '
    values = df[0]
    length = len(corr)
    result = []
    for i in range(length):
        result.append([])
    for value in values:
        # проверка на длину
        splitted_row = value.split(splitter)
        for i in range(len(splitted_row)):
            result[i].append(splitted_row[i])
    return result


def zip_columns(df):
    pass


def empty_method():
    pass


# входные данные
commands_data = [('SPLIT', ['ФИО'], ['Фамилия', 'Имя', 'Отчество']),
                 ('RENAME', ['Диагноз (расшифровка)'], ['Диагноз']),
                 ('RENAME', ['Диагноз (код)'], ['Код диагноза']),
                 ('RENAME', ['Тип исследования'], ['Категория исследования']),
                 ('RENAME', ['Адрес прописки пациента'], ['Адрес проживания'])]
filename = 'C:\\Users\\Арсель\\Downloads\\файл для проекта по конвертации.xlsx'
df_input = pd.read_excel(filename, 'исходный формат')
corr_fields = pd.read_excel(filename, 'нужный формат').columns
result = pd.DataFrame()
skip = []
rename_data = []
# выполнение команд
for command_data in commands_data:
    command = command_data[0]
    input = command_data[1]
    corr = command_data[2]
    if command == 'RENAME':
        rename_data.append(command_data)
        continue
    for item in input:
        skip.append(item)
    func = get_func(command)
    df = [df_input[x] for x in input]
    corr_columns = func(df)
    for i in range(len(corr)):
        result[corr[i]] = corr_columns[i]

# rename
dict_rename = {x[1][0]: x[2][0] for x in rename_data}
df_input.rename(columns=dict_rename, inplace=True)
# заполнение дата фрейма
for field in corr_fields:
    if not result.__contains__(field):
        if df_input.__contains__(field):
            result[field] = df_input[field]
        else:
            result[field] = [' ' for i in range(len(result))]
result.to_excel('C:\\Users\\Арсель\\PycharmProjects\\Convertor\\res.xlsx', 'нужный формат')
