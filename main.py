import math
from datetime import datetime
from datetime import timedelta
import pandas as pd
import numpy as np
import openpyxl
import timestamps as timestamps
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


class Convertor():
    def __init__(self, path):
        self.original = pd.read_excel(path, 'исходный формат')
        self.result = pd.read_excel(path, 'нужный формат')
        self.corr_fields =self.result.columns
        self.between = pd.DataFrame()

    def execute(self, command_data):
        command = command_data[0]
        input = command_data[1]
        corr = command_data[2]
        self.corr = corr
        if command == 'RENAME':
            self.rename(command_data)
            return
        changer = self.get_func(command)
        columns_for_change = [self.original[x] for x in input]
        corr_columns = changer(columns_for_change)
        #заполняет только первый столб, исправить!!!!
        for i in range(len(corr_columns[0]), self.original.shape[0]):
            for column in corr_columns:
                column.append(' ')
        for i in range(len(corr)):
            self.between[corr[i]] = corr_columns[i]
        self.fill_result()
        self.fix_date()
    def rename(self, command_data):
        dict_rename = {command_data[1][0]: command_data[2][0]}
        self.original.rename(columns=dict_rename, inplace=True)
    def fill_result(self):
        result = pd.DataFrame()
        for field in self.corr_fields:
            if self.between.__contains__(field):
                result[field] = self.between[field]
            elif self.original.__contains__(field):
                result[field] = self.original[field]
            else:
                result[field] = [' ' for i in range(len(result))]
        self.result=result
    def fix_date(self):
        for key in self.result.columns:
            if type(self.result[key][0]) is pd.Timestamp:
                list = []
                for item in self.result[key]:
                    a = str(item).replace('00:00:00', '')
                    list.append(a if a != 'NaT' else ' ')
                self.result[key] = list
    def get_func(self, command):
        if command == "SPLIT":
            return self.split_column
        elif command == "ZIP":
            return self.zip_columns
        else:  # command == "RENAME":
            return self.empty_method

    # [["фамилия"],["имя"],["отчество"]] -> [["ФИО"]] #corr_columns[i]
    def split_column(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.split_date(columns_for_change)
        splitter = ' '
        values = columns_for_change[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            # проверка на длину
            splitted_row = value.split(splitter)
            for i in range(len(splitted_row)):
                result[i].append(splitted_row[i])
        return result

    def split_date(self, columns):
        values = columns[0]
        length = len(self.corr)
        result = []
        for i in range(length):
            result.append([])
        for value in values:
            # проверка на длину
            date = [datetime.date(value), datetime.time(value)]
            # (value.year,value.month,value.day)
            for i in range(len(date)):
                result[i].append(date[i])
        return result

    def zip_date(self, columns_for_change):
        result = []
        for i in range(len(columns_for_change[0])):
            date = columns_for_change[0][i] if type(columns_for_change[0][i]) is pd.Timestamp else \
                columns_for_change[1][i]
            time = columns_for_change[0][i] if type(columns_for_change[0][i]) is not pd.Timestamp else \
                columns_for_change[1][i]
            if pd.isna(date) or pd.isna(time):
                result.append(' ')
                continue
            delta = timedelta(hours=time.hour, minutes=time.minute, seconds=time.second)
            result.append(date + delta)
        return [result]

    def zip_columns(self, columns_for_change):
        if type(columns_for_change[0][0]) in [pd.Timestamp, datetime]:
            return self.zip_date(columns_for_change)
        result = []
        values = []
        cars = pd.concat(columns_for_change).dropna().sort_index().astype('str').to_list()
        for car in cars:
            values.append(car)
            if (car.replace('.', '').isdigit()):
                model = ' '.join(values)
                values.clear()
                result.append(model)
        result = [result]  # -> список [["",""]]
        # result = [[result[x]] for x in range(len(result))] #-> список списков [[""],[""]]
        return result

    def empty_method(self):
        pass

    def as_text(self, val):
        if val is None:
            return ""
        return str(val)

    def to_exel(self):
        writer = pd.ExcelWriter('result1.xlsx',
                                engine='openpyxl')
        self.result.to_excel(writer, 'нужный формат', index=False)

        wb = Workbook()
        ws = wb.active
        # добавляем строчки дфа в опенпайексель
        for r in dataframe_to_rows(self.result, index=False, header=True):
            ws.append(r)
        # серега что-то тут нашаманил
        for column in ws.columns:
            length = max(len(self.as_text(cell.value)) for cell in column)
            ws.column_dimensions[column[0].column_letter].width = length + 2
        wb.save("result1.xlsx")

if __name__ == '__main__':
    commands_data = [('SPLIT', ['ФИО'], ['Фамилия', 'Имя', 'Отчество']),
                                       ('SPLIT', ['split_date'], ['date', 'time']),
                                       ('RENAME', ['Диагноз (расшифровка)'], ['Диагноз']),
                                       ('RENAME', ['Диагноз (код)'], ['Код диагноза']),
                                       ('RENAME', ['Тип исследования'], ['Категория исследования']),
                                       ('RENAME', ['Возвраст пациента'], ['Возраст']),
                                       ('RENAME', ['Адрес прописки пациента'], ['Адрес проживания']),
                                       ('ZIP', ["марка", "модель", "год"], ["машины"]),
                                       ('ZIP', ["Дата взятия анализа", "Время взятия анализа"], ["Дата и время взятия анализа"]),
                                       ('ZIP', ["Дата выполнения", "Время выполнения анализа"], ["Дата и время выполнения"])]
    convertor= Convertor('start.xlsx')
    for command in commands_data:
        convertor.execute(command)
    convertor.to_exel()
# # входные данные
# commands_data = [('SPLIT', ['ФИО'], ['Фамилия', 'Имя', 'Отчество']),
#                  ('SPLIT', ['split_date'], ['date', 'time']),
#                  ('RENAME', ['Диагноз (расшифровка)'], ['Диагноз']),
#                  ('RENAME', ['Диагноз (код)'], ['Код диагноза']),
#                  ('RENAME', ['Тип исследования'], ['Категория исследования']),
#                  ('RENAME', ['Возвраст пациента'], ['Возраст']),
#                  ('RENAME', ['Адрес прописки пациента'], ['Адрес проживания']),
#                  ('ZIP', ["марка", "модель", "год"], ["машины"]),
#                  ('ZIP', ["Дата взятия анализа", "Время взятия анализа"], ["Дата и время взятия анализа"]),
#                  ('ZIP', ["Дата выполнения", "Время выполнения анализа"], ["Дата и время выполнения"])]
# filename = 'start.xlsx'
#
# df_input = pd.read_excel(filename, 'исходный формат')
# corr_fields = pd.read_excel(filename, 'нужный формат').columns
# between = pd.DataFrame()
# skip = []
# rename_data = []
# # выполнение команд
# for command_data in commands_data:
#     command = command_data[0]
#     input = command_data[1]
#     corr = command_data[2]
#     if command == 'RENAME':
#         rename_data.append(command_data)
#         continue
#     for item in input:
#         skip.append(item)
#     func = get_func(command)
#     columns_for_change = [df_input[x] for x in input]
#     corr_columns = func(columns_for_change)
#     for i in range(len(corr_columns[0]), df_input.shape[0]):
#         for column in corr_columns:
#             column.append(' ')
#     for i in range(len(corr)):
#         between[corr[i]] = corr_columns[i]
#
# # rename
# dict_rename = {x[1][0]: x[2][0] for x in rename_data}
# df_input.rename(columns=dict_rename, inplace=True)
# # заполнение результирующего дата фрейма
# result = pd.DataFrame()
# for field in corr_fields:
#     if between.__contains__(field):
#         result[field] = between[field]
#     elif df_input.__contains__(field):
#         result[field] = df_input[field]
#     else:
#         result[field] = [' ' for i in range(len(result))]
# # приводим дату к человеческому виду
# for key in corr_fields:
#     if type(result[key][0]) is pd.Timestamp:
#         list = []
#         for item in result[key]:
#             a = str(item).replace('00:00:00', '')
#             list.append(a if a != 'NaT' else ' ')
#         result[key] = list
#
# writer = pd.ExcelWriter('result.xlsx',
#                         engine='openpyxl')
# result.to_excel(writer, 'нужный формат', index=False)
#
# wb = Workbook()
# ws = wb.active
# # добавляем строчки дфа в опенпайексель
# for r in dataframe_to_rows(result, index=False, header=True):
#     ws.append(r)
# # серега что-то тут нашаманил
# for column in ws.columns:
#     length = max(len(as_text(cell.value)) for cell in column)
#     ws.column_dimensions[column[0].column_letter].width = length + 2
# wb.save("result.xlsx")

# workbook = writer.book
# worksheet = writer.sheets['нужный формат']
# (max_row, max_col) = result.shape
# percent = workbook.add_format({'num_format': '0%'})
# worksheet.set_column(0, max_col, 20, None)
# worksheet.set_column(18, 18, 20, percent)
# phone = workbook.add_format({'num_format': '[<=9999999]###-####;(###) ###-####'})
# worksheet.set_column(6, 6, 10, date)
# writer.save()
