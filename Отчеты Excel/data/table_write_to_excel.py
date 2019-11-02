import pandas as pd
import numpy as np

from openpyxl.styles import Alignment
from data import constants, table_style


def add_result(sheet, res_row, *args):
    """
    Добавление итоговой строки таблицы
    :param sheet: ссылка на страницу Excel
    :param res_row: номер строки, куда необходимо вставить результат
    :param args: параметры итоговой строки, type - кортеж кортежей
    :return: None
    """
    for column, value in args:
        sheet.cell(row=res_row, column=column).value = value


def create_month_report(table, obj_id, month, year):
    """
    Создание табеля для конкретного объекта, месяца и года
    :param table: ссылка на таблицу
    :param obj_id: id объекта
    :param month: номер месяца отчета
    :param year: год отчета
    :return: ссылка на табель конкретного объекта, месяца и года
    """
    table_obj = pd.pivot_table(table[(table['ObjectId'] == obj_id) &
                                     (table['month'] == month)],
                               index='FIO', values='WorkedHours',
                               columns=['Date'], fill_value=0, aggfunc=np.sum)

    cur_day = 1
    day = pd.Timestamp(year=year, month=month, day=cur_day)
    while day.month == month:
        if day not in list(table_obj.columns):
            table_obj[day] = pd.Series()
        day = day + pd.Timedelta(days=1)

    table_obj.fillna(0, inplace=True)
    table_obj.sort_index(axis=1, inplace=True)

    table_obj['ВСЕГО'] = table_obj.loc[:, :].apply(np.sum, axis=1)
    columns = ['ВСЕГО'] + list(table_obj.columns)[:-1]
    table_obj = table_obj[columns]

    return table_obj


def table_record(writer, sheet, table, table_res, table_name=None,
                 row_start=0, is_auto_dimension=False):
    """
    Печать в файл Excel таблицы
    :param writer: ссылка на writer
    :param sheet: ссылка на страницу Excel
    :param table: ссылка на таблицу
    :param table_res: параметры итоговой строки, type - кортеж кортежей
    :param table_name: имя таблицы. По умолчанию - None
    :param row_start: номер стартовой строки. По умолчанию - 0
    :param is_auto_dimension: необходима ли автоширина столбцов таблицы в файле Excel. По умолчанию - False
    :return: номер последней строки таблицы
    """
    if table.shape[0] != 0:
        if table_name:
            row_start += 2
            sheet.cell(row=row_start, column=1).value = table_name
            sheet.cell(row=row_start, column=1).font = constants.FONT_TITLE
        row_start += 1

        table.to_excel(writer, sheet.title, index=False, startrow=row_start - 1)
        if is_auto_dimension:
            table_style.auto_dimension(sheet=sheet, row_count=row_start - 1)

        row_start += table.shape[0] + 1
        add_result(sheet, row_start, *table_res)

        table_style.table_style(sheet=sheet, row_start=row_start - table.shape[0] - 1,
                                table_size=(table.shape[0] + 2, table.shape[1]),
                                is_result_style=[True, True])
    return row_start


def write_month_report(table_shift, table_done, obj_id, month,
                       work_book, writer):
    """
    Печать табеля в файл Excel
    :param table_shift: ссылка на таблицу с зарплатой
    :param table_done: ссылка на таблицу с выплатами
    :param obj_id: id объекта
    :param month: номер месяца
    :param work_book: ссылка на книгу Excel
    :param writer: ссылка на writer
    :return: None
    """
    if month == constants.MONTH_START:
        year = constants.YEAR_START
    else:
        year = constants.YEAR_END

    month_report_obj = create_month_report(table_shift, obj_id, month, year)

    group_by = ['Partner', 'PartnerName', 'ObjectId', 'ObjectName', 'ObjectAddress']
    shift_for_report_unique = pd.DataFrame({'WorkedHours': table_shift.groupby(group_by)['WorkedHours'].sum()}).reset_index()

    obj_name = shift_for_report_unique[shift_for_report_unique['ObjectId'] == obj_id]['ObjectName'].values[0]
    partner = shift_for_report_unique[shift_for_report_unique['ObjectId'] == obj_id]['PartnerName'].values[0]
    address = shift_for_report_unique[shift_for_report_unique['ObjectId'] == obj_id]['ObjectAddress'].values[0]
    cur_sheet_name = f'shift_report_{month}_{int(obj_id)}'

    work_book.create_sheet(cur_sheet_name)
    sheet = work_book[cur_sheet_name]
    writer.sheets = dict((ws.title, ws) for ws in work_book.worksheets)

    row = 1
    sheet.cell(row=row, column=1).value = f'Лист учета рабочего времени ' \
                                          f'{constants.MONTHS[month]} {year}'
    sheet.cell(row=row, column=1).font = constants.FONT_TITLE
    row += 1
    sheet.cell(row=row, column=1).value = f'Дата формирования отчета: '
    sheet.cell(row=row, column=2).value = f'{constants.TODAY:%d.%m.%Y}'
    row += 1
    sheet.cell(row=row, column=1).value = 'ПАРТНЕР'
    sheet.cell(row=row, column=1).font = constants.FONT_TITLE
    sheet.cell(row=row, column=2).value = partner
    row += 1
    sheet.cell(row=row, column=1).value = 'ОБЪЕКТ'
    sheet.cell(row=row, column=1).font = constants.FONT_TITLE
    sheet.cell(row=row, column=2).value = obj_name
    row += 1
    sheet.cell(row=row, column=1).value = 'АДРЕС ОБЪЕКТА'
    sheet.cell(row=row, column=1).font = constants.FONT_TITLE
    sheet.cell(row=row, column=2).value = address
    row += 1

    month_report_obj.to_excel(writer, cur_sheet_name, startrow=row)
    row += 1
    sheet.cell(row=row, column=1).value = 'ФИО сотрудника'

    sheet.insert_rows(row + 1)
    for i, cell in enumerate(sheet[row + 1]):
        if i < 2:
            continue
        weekday = int(f'{sheet.cell(row=row, column=i + 1).value:%w}')
        weekday = constants.WEEKDAYS[weekday]
        cell.value = weekday

    for i, _ in enumerate(sheet):
        sheet.cell(row=i + 1, column=1).alignment = Alignment(horizontal='left')

    res = [(1, 'ИТОГО ЧАСОВ')] + list(
        (i + 2, sum(month_report_obj[column])) for i, column in
        enumerate(month_report_obj.columns))
    add_result(sheet, row + month_report_obj.shape[0] + 2, *res)

    for i, cell in enumerate(sheet[row]):
        if i < 2:
            continue
        cell.value = f'{cell.value:%d.%m}'

    table_style.table_report_style(sheet=sheet, row_start=row,
                                   table_size=(
                                       month_report_obj.shape[0] + 3,
                                       month_report_obj.shape[1] + 1
                                   ),
                                   is_formatting=True,
                                   is_result_style=[True, True])

    margin = table_shift[table_shift['ObjectId'] == obj_id][['Date', 'month', 'Price_total']].groupby(['Date', 'month']).agg('sum').reset_index()
    done = table_done[table_done['ObjectId'] == obj_id][['Date', 'month', 'ResultOfShift_total']].groupby(['Date', 'month']).agg('sum').reset_index()

    res_margin = [(1, 'ИТОГО ВЫРУЧКА'), (2, sum(margin[margin['month'] == month]['Price_total']))]
    res_done = [(1, 'ВЫПЛАЧЕНО ЗАКАЗЧИКОМ'),
                (2, sum(done[done['month'] == month]['ResultOfShift_total']))]
    for i, date in enumerate(month_report_obj.columns[1:]):
        if date in list(margin['Date']):
            res_margin += [(i + 3, margin[margin['Date'] == date]['Price_total'].values[0])]
        else:
            res_margin += [(i + 3, 0)]
        if date in list(done['Date']):
            res_done += [(i + 3, done[done['Date'] == date]['ResultOfShift_total'].values[0])]
        else:
            res_done += [(i + 3, 0)]

    add_result(sheet, row + month_report_obj.shape[0] + 3, *res_margin)
    table_style.result_style(row=sheet[row + month_report_obj.shape[0] + 3],
                             column_count=month_report_obj.shape[1] + 1,
                             last_res=True)
    add_result(sheet, row + month_report_obj.shape[0] + 4, *res_done)
    table_style.result_style(row=sheet[row + month_report_obj.shape[0] + 4],
                             column_count=month_report_obj.shape[1] + 1,
                             last_res=True)
    table_style.auto_dimension(sheet, add_width=0,
                               row_count=row - 1)


def margin_res(sheet, row, table, writer, city=None):
    """
    Запись в Excel файл итогов таблицы по городам по валовой прибыли
    :param sheet: ссылка на страницу Excel
    :param row: номер стартовой строки
    :param table: ссылка на таблицу с валовой прибылью
    :param writer: ссылка на writer
    :param month: номер месяца
    :param city: (наименование города, краткое наименование города)
    :return: номер последней строки с результатом
    """
    if city is not None:
        city_name = city[0]
        short_city = city[1]

        row += 1
        res = ((1, f'ИТОГО {short_city}'),
               (2, sum(table[table['Город'] == city_name]['Себестоимость'])),
               (3, sum(table[table['Город'] == city_name]['Выручка'])),
               (4, sum(table[table['Город'] == city_name]['Валовая прибыль'])),
               (5, sum(table[table['Город'] == city_name]['Выплачено'])))
        add_result(sheet, row, *res)
        table_style.result_style(sheet[row], table.shape[1] - 1, last_res=True)

        row += 1
        diff = sum(table[table['Город'] == city_name]['Себестоимость']) - \
               sum(table[table['Город'] == city_name]['Выплачено'])
        if sum(table[table['Город'] == city_name]['Выручка']) == 0:
            percentage = 0
        else:
            percentage = sum(table[table['Город'] == city_name]['Валовая прибыль']) / \
                         sum(table[table['Город'] == city_name]['Выручка'])
        res = ((1, f'ИТОГО - ВЫДАНО {short_city}'),
               (2, diff),
               (4, percentage))
        add_result(sheet, row, *res)
        table_style.result_style(sheet[row], table.shape[1] - 1, last_res=True)
        sheet.cell(row=row, column=4).value = f'{sheet.cell(row=row, column=4).value * 100:.1f}%'
    else:
        columns_write = ['Заказчик', 'Себестоимость', 'Выручка',
                         'Валовая прибыль', 'Выплачено']
        res = ((1, 'ИТОГО'),
               (2, sum(table['Себестоимость'])),
               (3, sum(table['Выручка'])),
               (4, sum(table['Валовая прибыль'])),
               (5, sum(table['Выплачено'])))
        row = table_record(writer, sheet, table[columns_write], res,
                           row_start=row, is_auto_dimension=True)

    return row


def gross_record(table, sheet, file_name, work_book, month=None):
    """
    Запись в Excel файл информационных данных и таблицы по городам по валовой прибыли
    :param table: ссылка на таблицу с валовой прибылью
    :param sheet: ссылка на страницу Excel
    :param file_name: имя файла Excel
    :param work_book: ссылка на книгу Excel
    :param month: номер месяца
    :return: None
    """
    row = 1
    sheet.cell(row=row, column=1).value = f'ВАЛОВАЯ ПРИБЫЛЬ'
    if month is not None:
        sheet.cell(row=row, column=1).value += f'. {constants.MONTHS[month]}'
    sheet.cell(row=row, column=1).font = constants.FONT_TITLE
    row += 1
    sheet.cell(row=row, column=1).value = f'Дата формирования отчета: '
    sheet.cell(row=row, column=2).value = f'{constants.TODAY:%d.%m.%Y}'
    row += 1
    if month is None:
        sheet.cell(row=row, column=1).value = f'Отчетные даты: '
        sheet.cell(row=row,
                   column=2).value = f'{constants.START_LAST_WEEK:%d.%m.%Y} - ' \
                                     f'{constants.END_LAST_WEEK:%d.%m.%Y}'
        row += 1

    for i, _ in enumerate(sheet):
        sheet.cell(row=i + 1, column=1).alignment = Alignment(horizontal='left')

    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a',
                        datetime_format='DD/MM') as writer:
        writer.book = work_book
        writer.sheets = dict((ws.title, ws) for ws in work_book.worksheets)

        row = margin_res(sheet=sheet, row=row, table=table, writer=writer)

        row += 1
        city = 'Тюмень'
        row = margin_res(sheet=sheet, row=row, table=table, writer=writer,
                         city=(city, 'ТМН'))

        city = 'Екатеринбург'
        row = margin_res(sheet=sheet, row=row, table=table, writer=writer,
                         city=(city, 'ЕКБ'))
