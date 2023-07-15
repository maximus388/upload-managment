LINK_MAIN = 'https://docs.google.com/spreadsheets/d/1x78-8URsts_gZDAJYVM2d_x_whbDB-AUKbZOoN9KM1A/edit#gid=0'


import pip
import sys
import time
from datetime import datetime, timedelta
from googleapiclient.errors import HttpError
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

import google.auth
from google.colab import auth
auth.authenticate_user()

# Установка библиотеки для работы с Google-таблицами
try:                                        
  import pygsheets
except:
  !pip install pygsheets
  import pygsheets

import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

offset = 4
credentials, _ = google.auth.default()
gc = pygsheets.client.Client(credentials)
sh = gc.open_by_url(LINK_MAIN)
WKS_LISTS = sh.worksheet_by_title('ЛИСТЫ')
DF_LISTS = WKS_LISTS.get_as_df(start = 'A1', end = 'B')
LISTS_TO_UPDATE = list(DF_LISTS[DF_LISTS['Нужно обновлять?'] == 'Да']['Листы с таблицами'])

def update():
  
    # Получение инструкций для копирования данных и проверка введенных данных на корректность 
  
    for LIST in LISTS_TO_UPDATE:
        wks = sh.worksheet_by_title(LIST)
        df = wks.get_as_df(start = 'A1', end = 'Q')
        df = df[2:].reset_index(drop = True)
        for index, row in df.iterrows():
            if (row['update'] == 'Да') & (('ОШИБКА' in row['status']) | ('OK' not in row['status'])):
                try:
                    try:
                        SH_FROM = gc.open_by_url(row['link_from'])
                        SH_TO = gc.open_by_url(row['link_to'])
                    except pygsheets.exceptions.NoValidUrlKeyFound:
                        wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                        now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                        wks.update_value(f'Q{index + offset}', now, True)
                        print(f'Лист - "{LIST}", cтрока {index + offset}: Нет ссылки на гугл-документ!')
                    except HttpError:
                        if str(sys.exc_info()[1]).split(' ')[1] == '403':
                            wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f'Лист - "{LIST}", cтрока {index + offset}: Ошибка доступа к гугл-документам!')
                            print()
                            print('Необходимо к каждому гугл-документу получить доступ')
                        elif str(sys.exc_info()[1]).split(' ')[1] == '404':
                            wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f'Лист - "{LIST}", cтрока {index + offset}: Ошибка доступа к гугл-документам!')
                            print()
                            print('Необходимо проверить ссылки на гугл-документы на корректность')
                    try:
                        LIST_FROM = SH_FROM.worksheet_by_title(row['list_from'])
                        LIST_TO = SH_TO.worksheet_by_title(row['list_to'])
                    except pygsheets.exceptions.WorksheetNotFound:
                        wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                        now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                        wks.update_value(f'Q{index + offset}', now, True)
                        print(f'Лист - "{LIST}", cтрока {index + offset}: Необходимо проверить названия листов!')
                    ARRAY_FROM = row['array_from']

                    # Этап копирования данных

                    if ARRAY_FROM == '':
                        df_array = LIST_FROM.get_as_df()
                    else:
                        array_list = ARRAY_FROM.split(':')
                        try:
                            df_array = LIST_FROM.get_as_df(start = array_list[0], end = array_list[1])
                        except IndexError:
                            wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f'Лист - "{LIST}", cтрока {index + offset}: Ошибка в написании адреса массива копирования!')
                    if row['filter'] != '':
                        try:
                            df_array = df_array[eval(row['filter'])]
                        except:
                            wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f'Лист - "{LIST}", cтрока {index + offset}: Ошибка в тексте фильтра!')
                    if row['array_columns'] != '':
                        try:
                            cols_array_from = [int(i.replace(' ','')) - 1 for i in str(row['array_columns']).split(',')]
                            df_array = df_array.iloc[:, cols_array_from]
                        except IndexError:
                            wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f'Лист - "{LIST}", cтрока {index + offset}: Номер колонки за пределами выбранного диапазона!')

                    # Этап вставки данных

                    if row['marker'] != '':
                        df_array.insert(0, row['marker'], row['marker'])
                    if row['paste_type'] == 'Обычная':
                        if row['clear'] == 'Да':
                            if row['array_to_clear'] != '':
                                array_to_clear_list = row['array_to_clear'].split(':')
                                try:
                                    LIST_TO.clear(start = array_to_clear_list[0], end = array_to_clear_list[1])
                                except IndexError:
                                    wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                                    now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                                    wks.update_value(f'Q{index + offset}', now, True)
                                    print(f'Лист - "{LIST}", cтрока {index + offset}: Ошибка в написании адреса массива очистки!')
                            else:
                                LIST_TO.clear()
                            copy_head = True if row['include_headers'] == 'Да' else False
                            len_df = len(df_array) + 1 if copy_head == True else len(df_array)
                            try:
                              LIST_TO.set_dataframe(df_array, row['cell_to'], copy_head = copy_head, extend = True, nan = '')
                            except:
                              print(df_array)
                            wks.update_value(f'P{index + offset}', f'OK ({len_df})', True)
                            now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                            wks.update_value(f'Q{index + offset}', now, True)
                            print(f"""Скопировано: Лист - '{LIST}', Название - '{row['name']}'""")
                    elif row['paste_type'] == 'Дополнение':
                        last_row_num = len(LIST_TO.get_as_df())
                        copy_head = True if row['include_headers'] == 'Да' else False
                        copy_head_offset = 1 if last_row_num == 0 else 2
                        len_df = len(df_array) + 1 if copy_head == True else len(df_array)
                        LIST_TO.set_dataframe(df_array, f'A{last_row_num + copy_head_offset}', copy_head = copy_head, extend = True, nan = '')
                        wks.update_value(f'P{index + offset}', f'OK ({len_df})', True)
                        now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                        wks.update_value(f'Q{index + offset}', now, True)
                        print(f"""Скопировано: Лист - '{LIST}', Название - '{row['name']}'""")
                    else:
                        wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                        now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                        wks.update_value(f'Q{index + offset}', now, True)
                        print(f'Лист - "{LIST}", cтрока {index + offset}: Некорректный тип вставки!')
                
                except UnboundLocalError:
                    wks.update_value(f'P{index + offset}', f'ОШИБКА!', True)
                    now = (datetime.now() + timedelta(hours = 5)).strftime('%Y-%m-%d %H:%M:%S')
                    wks.update_value(f'Q{index + offset}', now, True)
                    pass
                    
    

def main():
    update()
    print()
    print('Скрипт выполнен!')

if __name__ == '__main__':
    main()
