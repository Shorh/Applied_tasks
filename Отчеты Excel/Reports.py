from data import constants
from data.table_create import ReportTables
from data.table_write_to_excel import WriteMarginTables, WriteReportTables

FILE_NAMES = []
REPORT_TABLES = ReportTables()


def write_salary():
    # ЗП ведомость + Табель
    tables_for_report = REPORT_TABLES.get_tables_for_report()
    shift_for_report = tables_for_report['shift_for_report']
    first_sheet_name = 'salary_record'

    for city in shift_for_report['City'].unique():
        salary_table = REPORT_TABLES.get_salary_table(city)
        file_name = f'{constants.TODAY:%Y-%m-%d}. salary. {city}.xlsx'
        global FILE_NAMES
        FILE_NAMES += [file_name]
        table_of_objects = REPORT_TABLES.get_table_of_objects(city)

        write_report_tables = WriteReportTables(file_name=file_name,
                                                first_sheet_name=first_sheet_name)
        write_report_tables.write_weekly_report(city, salary_table)

        for month in shift_for_report['month'].unique():
            for obj_id in table_of_objects['ObjectId']:
                if shift_for_report[(shift_for_report['ObjectId'] == obj_id) & (
                        shift_for_report['month'] == month)].shape[0] != 0:
                    write_report_tables.write_month_report(
                        tables_for_report=tables_for_report, obj_id=obj_id,
                        month=month)


def write_margin():
    # Валовая прибыль
    file_name = f'{constants.TODAY:%Y-%m-%d}. margin.xlsx'
    global FILE_NAMES
    FILE_NAMES += [file_name]
    first_sheet_name = 'gross_margin'
    write_margin_tables = WriteMarginTables(file_name=file_name,
                                            first_sheet_name=first_sheet_name)

    gross_margin = REPORT_TABLES.get_margin()
    write_margin_tables.write_gross_record(gross_margin)

    gross_margin_month = REPORT_TABLES.get_margin(is_month=True)

    for month in gross_margin_month['month'].unique():
        write_margin_tables.write_gross_record(
            gross_margin_month[gross_margin_month['month'] == month], month=month)


write_salary()
write_margin()

print(f'Скрипт закончил работу. Сформированы файлы:')
for file in FILE_NAMES:
    print(file)
