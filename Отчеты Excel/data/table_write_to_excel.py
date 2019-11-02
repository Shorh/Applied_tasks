import pandas as pd
import openpyxl

from openpyxl.styles import Alignment
from data import constants, table_style
from data.table_create import ReportTables


class WriteTable:
    def __init__(self, file_name, first_sheet_name):
        self.file_name = file_name
        self.first_sheet_name = first_sheet_name
        self.work_book = None
        self.sheet = None
        self.__create_book()

        self.sheet_names = self.__get_sheets()
        self.row = 1

    def __create_book(self):
        self.work_book = openpyxl.Workbook()
        self.sheet = self.work_book.active
        self.sheet.title = self.sheet_names
        self.work_book.save(self.file_name)

    def __get_sheets(self):
        return dict((ws.title, ws) for ws in self.work_book.worksheets)

    def __get_writer(self):
        writer = pd.ExcelWriter(self.file_name, engine='openpyxl', mode='a',
                                datetime_format='DD.MM.YYYY')
        self.sheet_names = self.__get_sheets()
        return writer

    def __write_titles(self, type_table, **kwargs):
        """
        Запись в Excel описания таблицы
        :param type_table: тип описания
        1. Ведомость по ЗП. Первая страница
        2. Табель по объекту
        3. Валовая прибыль (за прошлый период или за месяц)
        :param sheet: ссылка на страницу
        :param kwargs: Если type_table == 1: city
        Если type_table == 2: month, year, partner, obj_name, address
        Если type_table == 3: None или month
        :return: None
        """
        if type_table == 1:
            city = kwargs['city']
            self.row = 1
            self.sheet.cell(row=self.row, column=1).value = f'ВЕДОМОСТЬ ПО ЗП'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.row += 1

            self.sheet.cell(row=self.row, column=1).value = f'ФИЛИАЛ'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.sheet.cell(row=self.row, column=2).value = city
            self.row += 1

            self.sheet.cell(row=self.row, column=1).value = f'Дата формирования отчета: '
            self.sheet.cell(row=self.row, column=2).value = f'{constants.TODAY:%d.%m.%Y}'
            self.row += 1

            self.sheet.cell(row=self.row, column=1).value = f'Отчетные даты: '
            self.sheet.cell(row=self.row, column=2).value = \
                f'{constants.START_LAST_WEEK:%d.%m.%Y} - ' \
                f'{constants.END_LAST_WEEK:%d.%m.%Y}'
            self.row += 1

        elif type_table == 2:
            month = kwargs['month']
            year = kwargs['year']
            partner = kwargs['partner']
            obj_name = kwargs['obj_name']
            address = kwargs['address']

            self.row = 1
            self.sheet.cell(row=self.row, column=1).value = \
                f'Лист учета рабочего времени {constants.MONTHS[month]} {year}'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE

            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = f'Дата формирования отчета: '
            self.sheet.cell(row=self.row, column=2).value = f'{constants.TODAY:%d.%m.%Y}'

            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = 'ПАРТНЕР'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.sheet.cell(row=self.row, column=2).value = partner

            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = 'ОБЪЕКТ'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.sheet.cell(row=self.row, column=2).value = obj_name

            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = 'АДРЕС ОБЪЕКТА'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.sheet.cell(row=self.row, column=2).value = address
            self.row += 1

        elif type_table == 3:
            self.row = 1
            self.sheet.cell(row=self.row, column=1).value = f'ВАЛОВАЯ ПРИБЫЛЬ'
            if 'month' in kwargs:
                self.sheet.cell(row=self.row, column=1).value += \
                    f'. {constants.MONTHS[kwargs["month"]]}'
            self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE

            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = f'Дата формирования отчета: '
            self.sheet.cell(row=self.row, column=2).value = f'{constants.TODAY:%d.%m.%Y}'

            self.row += 1
            if 'month' not in kwargs:
                self.sheet.cell(row=self.row, column=1).value = f'Отчетные даты: '
                self.sheet.cell(row=self.row, column=2).value = \
                    f'{constants.START_LAST_WEEK:%d.%m.%Y} - ' \
                    f'{constants.END_LAST_WEEK:%d.%m.%Y}'
                self.row += 1

            for i, _ in enumerate(self.sheet):
                self.sheet.cell(row=i + 1, column=1).alignment = Alignment(
                    horizontal='left')

    @staticmethod
    def __add_result(sheet, row, *args):
        """
        Добавление итоговой строки таблицы
        :param sheet: ссылка на страницу Excel
        :param row: номер строки, куда необходимо вставить результат
        :param args: параметры итоговой строки, type - кортеж кортежей
        :return: None
        """
        for column, value in args:
            sheet.cell(row=row, column=column).value = value

    def write_weekly_report(self, city, report_tables):
        salary_record = report_tables['record']
        salary_fine_detail = report_tables['fine']
        salary_done_detail = report_tables['done']

        self.work_book = openpyxl.load_workbook(self.file_name)
        self.sheet = self.work_book[self.sheet_names[0]]
        self.row = 1

        with self.__get_writer() as writer:
            writer.book = self.work_book
            writer.sheets = self.sheet_names
            self.__write_titles(type_table=1, city=city)

            res = ((1, 'ИТОГО'),
                   (4, sum(salary_record['Начислено'])),
                   (5, sum(salary_record['Штраф'])),
                   (6, sum(salary_record['Выплачено'])),
                   (7, sum(salary_record['Долг'])))
            self.__write_table(writer, salary_record, table_res=res,
                               is_auto_dimension=True)

            if salary_fine_detail.shape[0] != 0:
                res = ((1, 'ИТОГО'),
                       (5, sum(salary_fine_detail['Штраф'])))
                self.__write_table(writer, salary_fine_detail, table_res=res,
                                   table_name='ШТРАФЫ')

            if salary_done_detail.shape[0] != 0:
                res = ((1, 'ИТОГО'),
                       (4, sum(salary_done_detail['Сумма выплаты'])))
                self.__write_table(writer, salary_done_detail, table_res=res,
                                   table_name='ВЫПЛАТЫ')

        self.work_book.save(self.file_name)

    def __write_table(self, writer, table, **kwargs):
        """
        Печать в файл Excel таблиц зарплатной ведомости
        :param writer: контектсный менеджер pandas записи в Excel
        :param table: ссылка на таблицу
        :param kwargs: table_res: параметры итоговой строки, type - кортеж кортежей.
        table_name: имя таблицы. По умолчанию - None.
        is_auto_dimension: необходима ли автоширина столбцов таблицы в файле Excel
        :return: None
        """
        if table.shape[0] != 0:
            if 'table_name' in kwargs:
                self.row += 2
                self.sheet.cell(row=self.row, column=1).value = kwargs['table_name']
                self.sheet.cell(row=self.row, column=1).font = constants.FONT_TITLE
            self.row += 1

            table.to_excel(writer, self.sheet.title,
                           index=False, startrow=self.row - 1)
            if kwargs['is_auto_dimension']:
                table_style.auto_dimension(sheet=self.sheet,
                                           row_count=self.row - 1)

            if 'table_res' in kwargs:
                self.row += table.shape[0] + 1
                self.__add_result(self.sheet, self.row, *kwargs['table_res'])

            table_style.table_style(sheet=self.sheet,
                                    row_start=self.row - table.shape[0] - 1,
                                    table_size=(table.shape[0] + 2,
                                                table.shape[1]),
                                    is_result_style=[True, True])

    def write_month_report(self, tables_for_report, obj_id, month):
        """
        Печать табеля в файл Excel
        :param tables_for_report: словарь таблиц для табеля, содержащий
        shift_for_report и done_for_report
        :param obj_id: id объекта
        :param month: номер месяца
        :return: None
        """
        if month == constants.MONTH_START:
            year = constants.YEAR_START
        else:
            year = constants.YEAR_END

        shift_for_report = tables_for_report['shift_for_report']
        done_for_report = tables_for_report['done_for_report']

        table_shift = shift_for_report[(shift_for_report['ObjectId'] == obj_id) &
                                       (shift_for_report['month'] == month)]
        table_done = done_for_report[(done_for_report['ObjectId'] == obj_id) &
                                     (done_for_report['month'] == month)]

        month_report_obj = \
            ReportTables().get_month_report(table_shift, month, year)

        group_by = ['Partner', 'PartnerName', 'ObjectId',
                    'ObjectName', 'ObjectAddress']
        shift_for_report_unique = table_shift[group_by].drop_duplicates().reset_index()

        obj_name = \
            shift_for_report_unique[shift_for_report_unique['ObjectId'] ==
                                    obj_id]['ObjectName'].values[0]
        partner = \
            shift_for_report_unique[shift_for_report_unique['ObjectId'] ==
                                    obj_id]['PartnerName'].values[0]
        address = \
            shift_for_report_unique[shift_for_report_unique['ObjectId'] ==
                                    obj_id]['ObjectAddress'].values[0]
        sheet_name = f'shift_report_{month}_{int(obj_id)}'

        self.work_book = openpyxl.load_workbook(self.file_name)
        self.work_book.create_sheet(sheet_name)
        self.sheet = self.work_book[sheet_name]
        self.row = 1

        with self.__get_writer() as writer:
            writer.book = self.work_book
            writer.sheets = self.sheet_names
            self.__write_titles(type_table=2, month=month, year=year,
                                partner=partner, obj_name=obj_name,
                                address=address)

            month_report_obj.to_excel(writer, sheet_name, startrow=self.row)
            self.row += 1
            self.sheet.cell(row=self.row, column=1).value = 'ФИО сотрудника'

            self.sheet.insert_rows(self.row + 1)
            for i, cell in enumerate(self.sheet[self.row + 1]):
                if i < 2:
                    continue
                weekday = int(f'{self.sheet.cell(row=self.row, column=i + 1).value:%w}')
                weekday = constants.WEEKDAYS[weekday]
                cell.value = weekday

            for i, _ in enumerate(self.sheet):
                self.sheet.cell(row=i + 1, column=1).alignment = \
                    Alignment(horizontal='left')

            res = [(1, 'ИТОГО ЧАСОВ')] + list(
                (i + 2, sum(month_report_obj[column])) for i, column in
                enumerate(month_report_obj.columns))
            self.__add_result(self.sheet,
                              self.row + month_report_obj.shape[0] + 2,
                              *res)

            for i, cell in enumerate(self.sheet[self.row]):
                if i < 2:
                    continue
                cell.value = f'{cell.value:%d.%m}'

            table_style.table_report_style(sheet=self.sheet, row_start=self.row,
                                           table_size=(
                                               month_report_obj.shape[0] + 3,
                                               month_report_obj.shape[1] + 1
                                           ),
                                           is_formatting=True,
                                           is_result_style=[True, True])

            margin = table_shift[['Date', 'month', 'Price_total']].groupby(
                ['Date', 'month']).agg('sum').reset_index()
            done = table_done[['Date', 'month', 'ResultOfShift_total']].groupby(
                ['Date', 'month']).agg('sum').reset_index()

            res_margin = [(1, 'ИТОГО ВЫРУЧКА'),
                          (2, sum(margin['Price_total']))]
            res_done = [(1, 'ВЫПЛАЧЕНО ЗАКАЗЧИКОМ'),
                        (2, sum(done['ResultOfShift_total']))]
            for i, date in enumerate(month_report_obj.columns[1:]):
                if date in list(margin['Date']):
                    res_margin += [(i + 3, margin[
                        margin['Date'] == date]['Price_total'].values[0])]
                else:
                    res_margin += [(i + 3, 0)]

                if date in list(done['Date']):
                    res_done += [(i + 3, done[
                        done['Date'] == date]['ResultOfShift_total'].values[0])]
                else:
                    res_done += [(i + 3, 0)]

            self.__add_result(self.sheet, self.row + month_report_obj.shape[0] + 3, *res_margin)
            table_style.result_style(row=self.sheet[self.row + month_report_obj.shape[0] + 3],
                                     column_count=month_report_obj.shape[1] + 1,
                                     last_res=True)
            self.__add_result(self.sheet, self.row + month_report_obj.shape[0] + 4, *res_done)
            table_style.result_style(row=self.sheet[self.row + month_report_obj.shape[0] + 4],
                                     column_count=month_report_obj.shape[1] + 1,
                                     last_res=True)
            table_style.auto_dimension(self.sheet, add_width=0,
                                       row_count=self.row - 1)

        self.work_book.save(self.file_name)

    def write_gross_record(self, table, month=None):
        """
        Запись в Excel файл информационных данных и таблицы по городам по валовой прибыли
        :param table: ссылка на таблицу с валовой прибылью
        :param month: номер месяца
        :return: None
        """
        self.work_book = openpyxl.load_workbook(self.file_name)
        if month is not None:
            sheet_name = f'{self.first_sheet_name}_month_{month}'
            self.work_book.create_sheet(sheet_name)
        else:
            sheet_name = self.first_sheet_name

        self.sheet = self.work_book[sheet_name]
        self.row = 1

        with self.__get_writer() as writer:
            writer.book = self.work_book
            writer.sheets = self.sheet_names
            self.__write_titles(type_table=3, month=month)

            self.__write_margin_res(writer, table)

            self.row += 1
            city = 'Тюмень'
            self._write_margin_res(writer, table, city=(city, 'ТМН'))

            city = 'Екатеринбург'
            self._write_margin_res(writer, table, city=(city, 'ЕКБ'))

        self.work_book.save(self.file_name)

    def __write_margin_res(self, writer, table, city=None):
        """
        Запись в Excel файл итогов таблицы по городам по валовой прибыли
        или самой таблицы
        :param table: ссылка на таблицу с валовой прибылью
        :param writer: ссылка на writer
        :param city: (наименование города, краткое наименование города)
        :return: номер последней строки с результатом
        """
        if city is not None:
            city_name = city[0]
            short_city = city[1]

            self.row += 1
            res = ((1, f'ИТОГО {short_city}'),
                   (2, sum(table[table['Город'] == city_name]['Себестоимость'])),
                   (3, sum(table[table['Город'] == city_name]['Выручка'])),
                   (4, sum(table[table['Город'] == city_name]['Валовая прибыль'])),
                   (5, sum(table[table['Город'] == city_name]['Выплачено'])))
            self.__add_result(self.sheet, self.row, *res)
            table_style.result_style(self.sheet[self.row], table.shape[1] - 1,
                                     last_res=True)

            self.row += 1
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
            self.__add_result(self.sheet, self.row, *res)
            table_style.result_style(self.sheet[self.row], table.shape[1] - 1,
                                     last_res=True)
            self.sheet.cell(row=self.row, column=4).value = \
                f'{self.sheet.cell(row=self.row, column=4).value * 100:.1f}%'
        else:
            columns_write = ['Заказчик', 'Себестоимость', 'Выручка',
                             'Валовая прибыль', 'Выплачено']
            res = ((1, 'ИТОГО'),
                   (2, sum(table['Себестоимость'])),
                   (3, sum(table['Выручка'])),
                   (4, sum(table['Валовая прибыль'])),
                   (5, sum(table['Выплачено'])))
            self.__write_table(writer, table[columns_write], table_res=res,
                               is_auto_dimension=True)
