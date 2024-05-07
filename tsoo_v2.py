# import openpyxl as xl
# from openpyxl import Workbook
from datetime import datetime
import os
import sys
import pandas as pd
import sqlite3

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

    schedules_for_delete = ['Субботник', 'Под загрузку', 'Под погрузку', 'Установка']
    org_types_for_delete = ['Юридическое лицо', 'Бюджетное учреждение', 'СНТ', 'РЖД', '']

    cleaned_df_1 = combined_df[~combined_df['График вывоза'].astype(str).apply(
        lambda x: any(schedule.lower() in x.lower() for schedule in schedules_for_delete))]
    cleaned_df_1 = cleaned_df_1.loc[~cleaned_df_1['Код КП'].astype(str).str.lower().str.startswith('789')]
    cleaned_df_1 = cleaned_df_1.loc[~cleaned_df_1['Адрес'].astype(str).str.contains('сигнальный', case=False)]
    cleaned_df_1 = cleaned_df_1.loc[~cleaned_df_1['Адрес'].astype(str).str.contains('вкп', case=False)]
    cleaned_df_1 = cleaned_df_1.loc[~cleaned_df_1['Примечание'].astype(str).str.contains('контейнер выходного дня', case=False)]
    cleaned_df_1['Район'] = cleaned_df_1['Район'].replace('Невский (Левый)', 'Невский')
    cleaned_df_1['Район'] = cleaned_df_1['Район'].replace('Невский (Правый)', 'Невский')

    cleaned_df_2 = cleaned_df_1.loc[cleaned_df_1['Тариф'].astype(str).str.lower() != 'факт']
    cleaned_df_2 = cleaned_df_2.loc[~cleaned_df_2['Категория'].astype(str).apply(
        lambda x: any(org_type.lower() == x.lower() for org_type in org_types_for_delete))]

    # sql start
    # conn = sqlite3.connect('files/test_db.db')
    # cleaned_df_2.to_sql('df_22594', conn, if_exists='replace', index=False)
    # conn.close()
    # sql end

    cleaned_df_2 = cleaned_df_2.loc[cleaned_df_2['Объём'].astype(float) >= 14]
    cleaned_df_2 = cleaned_df_2.loc[~cleaned_df_2['Адрес'].astype(str).str.contains('смет', case=False)]
    cleaned_df_2 = cleaned_df_2.drop_duplicates(subset=['Код КП'])
    print(4)

    pivot_table_1 = cleaned_df_1.pivot_table(index='Район', values='Код КП', aggfunc='count').reset_index()
    pivot_table_2 = cleaned_df_2.pivot_table(index='Район', values='Код КП', aggfunc='count').reset_index()

    can_sum_all = cleaned_df_1.groupby('Район')['Количество'].sum().reset_index()
    can_sum_all.rename(columns={'Количество': 'Количество всего'}, inplace=True)
    pivot_table_1 = pivot_table_1.merge(can_sum_all, on='Район', how='left')

    can_volumes = [0.36, 0.66, 0.75, 0.77, 1.1, 6, 8, 14, 27]
    can_sum = []
    counter = 0
    for e in can_volumes:
        can_sum.append(cleaned_df_1[cleaned_df_1['Объём'] == e].groupby('Район')['Количество'].sum().reset_index())
        can_sum[counter].rename(columns={'Количество': f'Количество {str(e).replace(".", ",")}'}, inplace=True)
        pivot_table_1 = pivot_table_1.merge(can_sum[counter], on='Район', how='left')
        counter += 1

    cleaned_df_1_filtered = cleaned_df_1[~cleaned_df_1['Объём'].isin(can_volumes)]
    can_sum_others = cleaned_df_1_filtered.groupby('Район')['Количество'].sum().reset_index()
    can_sum_others.rename(columns={'Количество': 'Количество другие'}, inplace=True)
    pivot_table_1 = pivot_table_1.merge(can_sum_others, on='Район', how='left')

    pivot_table_1.fillna(0, inplace=True)

    print(5)

    with pd.ExcelWriter(f'{files_dir}Итог/Реестр КП {datetime.now().strftime("%d.%m.%Y %H_%M")}.xlsx') as writer:
        cleaned_df_1.to_excel(writer, sheet_name='Реестр КГ', index=False)
        pivot_table_1.to_excel(writer, sheet_name='Сводная КГ', index=False)
        cleaned_df_2.to_excel(writer, sheet_name='Реестр КП КГО', index=False)
        pivot_table_2.to_excel(writer, sheet_name='Сводная КП КГО', index=False)

    # cleaned_df_1.to_excel(f'{files_dir}Итог/Реестр КП {datetime.now().strftime("%d.%m.%Y %H_%M")}.xlsx', index=False)


combine_files('files/disp/')
