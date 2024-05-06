import openpyxl as xl
from openpyxl import Workbook
from datetime import datetime
import os
import sys
import pandas as pd

# Отчет по терсхеме


def combine_files(files_dir):
    print('start')
    files = os.listdir(files_dir)

    if not files:
        sys.exit(1)
    else:
        xlsx_files = [e for e in files if e.endswith('.xlsx')]
        if not xlsx_files:
            sys.exit(1)

    df1 = pd.read_excel(f'{files_dir}{xlsx_files[0]}', skiprows=3)
    print(1)
    df2 = pd.read_excel(f'{files_dir}{xlsx_files[1]}', skiprows=3)
    print(2)

    combined_df = pd.concat([df1, df2], ignore_index=True)
    print(3)

    schedules_for_delete = ['субботник', 'под загрузку', 'под погрузку', 'установка']

    cleaned_df = combined_df[~combined_df['График вывоза'].astype(str).apply(
        lambda x: any(schedule.lower() in x.lower() for schedule in schedules_for_delete))]
    cleaned_df = cleaned_df.loc[~cleaned_df['Код КП'].astype(str).str.lower().str.startswith('789')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Адрес'].astype(str).str.lower().str.contains('сигнальный')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Адрес'].astype(str).str.lower().str.contains('вкп')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Примечание'].astype(str).str.lower().str.contains('контейнер выходного дня')]
    cleaned_df['Район'] = cleaned_df['Район'].replace('Невский (Левый)', 'Невский')
    cleaned_df['Район'] = cleaned_df['Район'].replace('Невский (Правый)', 'Невский')

    print(4)

    pivot_table = cleaned_df.pivot_table(index='Район', values='Код КП', aggfunc='count').reset_index()

    with pd.ExcelWriter(f'{files_dir}Итог/Реестр КП {datetime.now().strftime("%d.%m.%Y %H_%M")}.xlsx') as writer:
        cleaned_df.to_excel(writer, sheet_name='Реестр КП', index=False)
        pivot_table.to_excel(writer, sheet_name='Сводная', index=False)

    # cleaned_df.to_excel(f'{files_dir}Итог/Реестр КП {datetime.now().strftime("%d.%m.%Y %H_%M")}.xlsx', index=False)


def delete_rows():
    print('start')
    files_dir = 'files/disp/'
    files = os.listdir(files_dir)

    if not files:
        sys.exit(1)
    else:
        xlsx_files = [e for e in files if e.endswith('.xlsx')]
        if not xlsx_files:
            sys.exit(1)

    df = pd.read_excel(f'{files_dir}{xlsx_files[0]}', skiprows=3)
    print(1)

    schedules_for_delete = ['субботник', 'под загрузку', 'под погрузку', 'установка']

    # cleaned_df = df.loc[~df['График вывоза'].astype(str).str.lower().str.contains('субботник')]
    cleaned_df = df[~df['График вывоза'].astype(str).apply(
        lambda x: any(schedule.lower() in x.lower() for schedule in schedules_for_delete))]
    cleaned_df = cleaned_df.loc[~cleaned_df['Код КП'].astype(str).str.lower().str.startswith('789')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Адрес'].astype(str).str.lower().str.contains('сигнальный')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Адрес'].astype(str).str.lower().str.contains('вкп')]
    cleaned_df = cleaned_df.loc[~cleaned_df['Примечание'].astype(str).str.lower().str.contains('контейнер выходного дня')]

    print(2)
    cleaned_df.to_excel(f'{files_dir}Итог/Реестр КП test {datetime.now().strftime("%d.%m.%Y %H_%M")}.xlsx', index=False)
    print(3)


combine_files('files/disp/')
# delete_rows()
