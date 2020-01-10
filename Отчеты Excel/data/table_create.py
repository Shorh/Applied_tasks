import pandas as pd
import numpy as np

from data import constants
from data.tables import AmberTables


class ReportTables(AmberTables):
    def __init__(self, data_folder=False):
        super().__init__(data_folder)
        self.questionarys = self.get_questionarys()
        self.shift = self.get_shift()
        self.partners = self.get_partners()
        self.full_questionarys = self.__get_full_questionarys()
        self.shift_detail = self.__get_shift_detail()

    def __get_full_questionarys(self):
        """
        Создание полной таблицы анкет сотрудников
        :return: таблица
        """
        columns = ['Id', 'ShiftCount', 'NewNumber', 'Vozrast', 'Nomerkarty',
                   'Vladeleckarty', 'Telefon', 'FIO', 'CurrentReckoning',
                   'StartWorkDate', 'IsGiveSalary', 'Num', 'EmploymentStatus',
                   'EmployeeStatus', 'EmploymentType', 'Speciality', 'City']
        full_questionarys = self.questionarys.copy()[columns]
        employee_status = self.get_employee_status()
        employment_status = self.get_employment_status()
        employment_type = self.get_employment_type()
        specialty = self.get_specialty()

        full_questionarys['EmployeeStatus'] = [
            employee_status[employee_status['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in full_questionarys['EmployeeStatus']]

        full_questionarys['EmploymentStatus'] = [
            employment_status[
                employment_status['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in full_questionarys['EmploymentStatus']]

        full_questionarys['EmploymentType'] = [
            employment_type[employment_type['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in full_questionarys['EmploymentType']]

        full_questionarys['Speciality'] = [
            specialty[specialty['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in full_questionarys['Speciality']]

        return full_questionarys

    def __get_shift_detail(self):
        """
        Создание таблицы детализированных начислений и выплат сотрудникам
        :return: таблица начислений и выплат
        """
        def get_price(x):
            partner, specialty = x
            if pd.notna(partner) and pd.notna(specialty):
                return rate_for_partner[
                    (rate_for_partner['Klient'] == partner) &
                    (rate_for_partner['Specialty'] == specialty)
                    ]['Price'].values.item(0)
            else:
                return 0

        appointed_staff = self.get_appointed_staff()
        facility = self.get_facility()
        rate_for_partner = self.get_rate_for_partner()
        fine_reason = self.get_fine_reason()
        partners = self.partners
        contact = self.get_contact()
        base_city = self.get_base_city()

        shift_detail = self.shift.copy()

        merge_col = ['ObjectId', 'ObjectName', 'ObjectPartner', 'ObjectAddress']
        shift_detail = shift_detail.merge(
            facility[merge_col], left_on='Facility', right_on='ObjectId',
            how='left', validate='m:1')

        merge_col = ['City', 'Id', 'FIO', 'Nomerkarty', 'Vladeleckarty']
        shift_detail = shift_detail.merge(
            self.full_questionarys[merge_col], left_on='Questionary',
            right_on='Id', how='left', validate='m:1')

        shift_detail['FineReason'] = [
            fine_reason[fine_reason['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in shift_detail['FineReason']]

        shift_detail = shift_detail.rename(columns={'Id_x': 'Id'})
        shift_detail.index = shift_detail['Id']

        drop = ['IsAuto', 'Comment', 'Id', 'CustObjVersion', 'IdCreateBy',
                'CreateDate', 'IdUpdateBy', 'UpdateDate']
        shift_detail.drop(drop, axis=1, inplace=True)

        shift_detail['PaymentDate'] = [
            pd.Timestamp(date[:10])
            if date else pd.Timestamp('1900-01-01')
            for date in shift_detail['PaymentDate']]

        shift_detail['PartnerName'] = [
            partners[partners['Id'] == x]['LegalName'].values.item(0)
            if pd.notna(x) else 'Владелец'
            for x in shift_detail['Partner']]

        shift_detail['WhoPaidMoney'] = [
            contact[contact['Id'] == x]['Partner'].values.item(0)
            if pd.notna(x) else ''
            for x in shift_detail['WhoPaidMoney']]

        shift_detail['City'] = [
            base_city[base_city['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in shift_detail['City']]

        merge = ['Id', 'Specialty', 'Vydanonalichnymi']
        shift_detail = shift_detail.merge(
            appointed_staff[merge], left_on='AppointedStaff', right_on='Id',
            how='left', validate='m:1')
        shift_detail['Vydanonalichnymi'].fillna(0, inplace=True)
        for index, row in shift_detail[(
                shift_detail['Vydanonalichnymi'] != 0)].iterrows():
            row['WorkedHours'] = 0
            row['FineOrBonus'] = 0
            row['ResultOfShift'] = row['Vydanonalichnymi']
            row['IsShiftPaid'] = False
            row['IsComeToWork'] = None
            row['PaymentDate'] = row['Date']
            row['Type'] = 127
            row['WhoPaidMoney'] = row['Partner']
            shift_detail = shift_detail.append(row, ignore_index=True)

        shift_detail = shift_detail.assign(
            PriceForHour=shift_detail[
                ['Partner', 'Specialty']].apply(get_price, axis=1))
        shift_detail['WorkedHours'].fillna(0, inplace=True)
        shift_detail['InnerRate'].fillna(0, inplace=True)
        shift_detail['ResultOfShift_total'] = \
            shift_detail['WorkedHours'] * shift_detail['InnerRate'] + \
            shift_detail['FineOrBonus']
        shift_detail['ResultOfShift_total'] = [
            shift_detail.loc[index, 'ResultOfShift']
            if shift_detail.loc[index, 'Type'] != 126
            else shift_detail.loc[index, 'ResultOfShift_total']
            for index in shift_detail.index]

        shift_detail['Price'] = \
            shift_detail['WorkedHours'] * shift_detail['PriceForHour']

        shift_detail['Margin'] = \
            shift_detail['Price'] - shift_detail['ResultOfShift_total']
        shift_detail['Margin'] = [
            0 if shift_detail.loc[index, 'WorkedHours'] == 0
            else shift_detail.loc[index, 'Margin']
            for index in shift_detail.index]

        shift_detail['month'] = [date.month for date in shift_detail['Date']]

        drop = ['AppointedStaff', 'Id']
        shift_detail.drop(drop, axis=1, inplace=True)

        return shift_detail

    def get_salary_table(self, city):
        """
        Создание ведомости по ЗП, штрафам и выплатам по константам
        START_LAST_WEEK и END_LAST_WEEK
        :param city: Название города
        :return: Словарь из таблиц ведомости ЗП, штрафов и выплат
        """
        group_by = ['Questionary', 'City']
        merge_record = ['Id', 'FIO', 'Nomerkarty', 'Vladeleckarty']
        merge_other = ['Id', 'FIO']
        drop = ['Questionary', 'Id']
        salary = self.shift_detail[
            (self.shift_detail['Date'] >= constants.START_LAST_WEEK) &
            (self.shift_detail['City'] == city)]

        return {'record': self.__get_record_table(group_by, merge_record, drop, salary),
                'fine': self.__get_fine_table(group_by, merge_other, drop, salary),
                'done': self.__get_done_table(group_by, merge_other, drop, salary)}

    def __get_record_table(self, group_by, merge, drop, salary):
        """
        Создание таблицы ведомости ЗП по сотрудникам
        :param group_by: поля для группировки таблицы
        :param merge: поля для слияния таблиц
        :param drop: поля для удаления из таблицы
        :param salary: таблица с данными по выплатам и начислениям
        :return:
        """
        def duty(x):
            shift, done = x
            if shift > done:
                return shift - done
            else:
                return 0

        def payment(x):
            duty, fine = x
            if -fine < duty:
                return duty + fine
            else:
                return 0

        col_shift = group_by + ['ResultOfShift_total']
        col_fine = group_by + ['FineOrBonus']

        salary_done = salary[col_shift][
            (salary['Type'] != 126) &
            (salary['ResultOfShift_total'] != 0) &
            (salary['IsShiftPaid'] == False) &
            (salary['Date'] <= constants.END_LAST_WEEK)]
        salary_done = salary_done.groupby(
            ['Questionary']).agg('sum').reset_index()
        salary_done = salary_done.rename(
            columns={'ResultOfShift_total': 'Выплачено'})

        salary_fine = salary[col_fine][
            (salary['Type'] == 126) &
            (salary['FineOrBonus'] != 0) &
            (salary['Date'] >= constants.START_THIS_WEEK)]
        salary_fine = salary_fine.groupby(
            ['Questionary']).agg('sum').reset_index()
        salary_fine = salary_fine.rename(
            columns={'FineOrBonus': 'Штраф/премия текущей недели'})

        salary_record = salary[col_shift][
            (salary['Type'] == 126) &
            (salary['ResultOfShift_total'] != 0) &
            (salary['Date'] <= constants.END_LAST_WEEK)]
        salary_record = salary_record.groupby(group_by).agg(
            'sum').reset_index()
        salary_record = salary_record.rename(
            columns={'ResultOfShift_total': 'Начислено'})
        salary_record = salary_record.merge(
            salary_done, on='Questionary', how='left', validate='1:1')
        salary_record = salary_record.merge(
            salary_fine, on='Questionary', how='left', validate='1:1')

        salary_record.fillna(0, inplace=True)

        salary_record = salary_record.merge(
            self.full_questionarys[merge],
            left_on='Questionary', right_on='Id', how='left', validate='m:1')

        salary_record.drop(drop, axis=1, inplace=True)

        salary_record = salary_record.assign(
            duty=salary_record[['Начислено', 'Выплачено']].apply(duty, axis=1))
        salary_record = salary_record.rename(
            columns={'City': 'Город', 'FIO': 'ФИО', 'Nomerkarty': '№ карты',
                     'Vladeleckarty': 'Владелец карты', 'duty': 'Долг'})

        salary_record = salary_record.assign(payment=salary_record[
            ['Долг', 'Штраф/премия текущей недели']].apply(payment, axis=1))
        salary_record = salary_record.rename(columns={'payment': 'К выплате'})

        columns_record = ['Город', 'ФИО', '№ карты', 'Владелец карты',
                          'Начислено', 'Выплачено', 'Долг',
                          'Штраф/премия текущей недели', 'К выплате']
        salary_record = salary_record[columns_record]
        salary_record.sort_values(by=['К выплате'], ascending=False,
                                  inplace=True)

        return salary_record

    def __get_fine_table(self, group_by, merge, drop, salary):
        """
        Создание таблицы штрафов по сотрудникам
        :param group_by: поля для группировки таблицы
        :param merge: поля для слияния таблиц
        :param drop: поля для удаления из таблицы
        :param salary: таблица с данными по выплатам и начислениям
        :return:
        """
        group_by_fine = group_by + ['ObjectName', 'Date', 'FineReason']
        col_fine_detail = group_by_fine + ['FineOrBonus']

        salary_fine_detail = salary[col_fine_detail][
            (salary['FineOrBonus'] != 0) &
            (salary['Date'] >= constants.START_THIS_WEEK)]
        salary_fine_detail = salary_fine_detail.groupby(
            group_by_fine).agg('sum').reset_index()
        salary_fine_detail = salary_fine_detail.rename(
            columns={'FineOrBonus': 'Штраф',
                     'ObjectName': 'На каком объекте был штраф',
                     'Date': 'Дата штрафа', 'FineReason': 'Причина штрафа'})

        salary_fine_detail.fillna(0, inplace=True)

        salary_fine_detail = salary_fine_detail.merge(
            self.full_questionarys[merge], left_on='Questionary',
            right_on='Id', how='left', validate='m:1')
        salary_fine_detail.drop(drop, axis=1, inplace=True)

        salary_fine_detail = salary_fine_detail.rename(
            columns={'City': 'Город', 'FIO': 'ФИО'})
        columns_fine_detail = ['Город', 'ФИО', 'Причина штрафа',
                               'На каком объекте был штраф', 'Дата штрафа',
                               'Штраф']
        salary_fine_detail = salary_fine_detail[
            columns_fine_detail].sort_values(by=['Дата штрафа'])

        return salary_fine_detail

    def __get_done_table(self, group_by, merge, drop, salary):
        """
        Создание таблицы выплат по сотрудникам
        :param group_by: поля для группировки таблицы
        :param merge: поля для слияния таблиц
        :param drop: поля для удаления из таблицы
        :param salary: таблица с данными по выплатам и начислениям
        :return:
        """
        group_by_done = group_by + ['WhoPaidMoney', 'PaymentDate']
        col_done = group_by_done + ['ResultOfShift_total']

        salary_done_detail = salary[col_done][
            (salary['Type'] != 126) &
            (salary['ResultOfShift_total'] != 0) &
            (salary['IsShiftPaid'] == False) &
            (salary['Date'] <= constants.END_LAST_WEEK)]
        salary_done_detail['WhoPaidMoney'] = [
            self.partners[self.partners['Id'] == x][
                'LegalName'].values.item(0)
            if (pd.notna(x) and x != '') else 'Зарплата'
            for x in salary_done_detail['WhoPaidMoney']]
        salary_done_detail = salary_done_detail.groupby(
            group_by_done).agg('sum').reset_index()
        salary_done_detail = salary_done_detail.rename(
            columns={'ResultOfShift_total': 'Сумма выплаты',
                     'WhoPaidMoney': 'Кем выплачено',
                     'PaymentDate': 'Дата выплаты'})

        salary_done_detail.fillna(0, inplace=True)

        salary_done_detail = salary_done_detail.merge(
            self.full_questionarys[merge], left_on='Questionary',
            right_on='Id', how='left', validate='m:1')

        salary_done_detail.drop(drop, axis=1, inplace=True)
        salary_done_detail = salary_done_detail.rename(
            columns={'City': 'Город', 'FIO': 'ФИО'})
        columns_done_detail = ['Город', 'ФИО', 'Дата выплаты', 'Кем выплачено',
                           'Сумма выплаты']
        salary_done_detail = salary_done_detail[
            columns_done_detail].sort_values(by=['Дата выплаты'])

        return salary_done_detail

    def get_tables_for_report(self):
        """
        Создание таблиц по сменам и начислениям сотрудников для формирования
        табелей по объектам
        :return:
        """
        drop = ['FineOrBonus', 'ResultOfShift', 'IsShiftPaid', 'PaymentDate',
                'IsComeToWork', 'InstanceDate',
                'WhoPaidMoney', 'FineReason', 'Questionary',
                'ObjectPartner', 'InnerRate', 'Specialty', 'PriceForHour']
        shift_for_report = self.shift_detail.drop(drop, axis=1)

        shift_for_report.sort_values(by=['FIO'], inplace=True)
        shift_for_report = shift_for_report[
            (shift_for_report['month'] == constants.MONTH_START) |
            (shift_for_report['month'] == constants.MONTH_END)]

        done_for_report = shift_for_report[shift_for_report['Type'] != 126]
        shift_for_report = shift_for_report[shift_for_report['Type'] == 126]

        return {'shift_for_report': shift_for_report,
                'done_for_report': done_for_report}

    def get_table_of_objects(self, city):
        shift_for_report_unique = self.shift_detail[
            (self.shift_detail['City'] == city) &
            ((self.shift_detail['month'] == constants.MONTH_START) |
             (self.shift_detail['month'] == constants.MONTH_END))]

        group_by = ['Partner', 'PartnerName', 'ObjectId', 'ObjectName',
                    'ObjectAddress']
        shift_for_report_unique = shift_for_report_unique[
            group_by].drop_duplicates().reset_index()

        return shift_for_report_unique

    @staticmethod
    def get_month_report(df, month, year):
        """
        Создание табеля для конкретного объекта, месяца и года
        :param: df - таблица, из которой создается табель
        :param: month - месяц табеля
        :param: year - год табеля
        :return: ссылка на табель конкретного объекта, месяца и года
        """
        df_obj = pd.pivot_table(
            df,
            index='FIO', values='WorkedHours', columns=['Date'], fill_value=0,
            aggfunc=np.sum)

        cur_day = 1
        day = pd.Timestamp(year=year, month=month, day=cur_day)
        while day.month == month:
            if day not in list(df_obj.columns):
                df_obj[day] = pd.Series()
            day = day + pd.Timedelta(days=1)

        df_obj.fillna(0, inplace=True)
        df_obj.sort_index(axis=1, inplace=True)

        df_obj['ВСЕГО'] = df_obj.loc[:, :].apply(np.sum, axis=1)
        columns = ['ВСЕГО'] + list(df_obj.columns)[:-1]
        df_obj = df_obj[columns]

        for index in df_obj.index:
            if df_obj.loc[index, 'ВСЕГО'] == 0:
                df_obj.drop([index], inplace=True)

        return df_obj

    def get_margin(self, is_month=False):
        """
        Создание таблицы валовой прибыли либо за последние две недели, либо по месяцам
        :param is_month: таблица по месяцам или за две последние недели?
        :return:
        """
        group_by = ['PartnerName', 'Partner', 'City']
        if is_month:
            group_by += ['month']

        column = group_by + ['ResultOfShift_total', 'Price', 'Margin',
                             'Vydanonalichnymi']

        margin_table = self.shift_detail[self.shift_detail['Type'] == 126]

        if is_month:
            margin_table = margin_table[column]
        else:
            margin_table = \
                margin_table[(margin_table['Date'] >= constants.START_LAST_WEEK) &
                             (margin_table['Date'] <= constants.END_LAST_WEEK)][column]

        margin_table = margin_table.groupby(group_by).agg('sum').reset_index()

        drop = ['Partner']
        margin_table.drop(drop, axis=1, inplace=True)

        margin_table.sort_values(by=['Margin'], ascending=False, inplace=True)
        margin_table = margin_table.rename(
            columns={'PartnerName': 'Заказчик',
                     'ResultOfShift_total': 'Себестоимость',
                     'Price': 'Выручка', 'Margin': 'Валовая прибыль',
                     'Vydanonalichnymi': 'Выплачено', 'City': 'Город'})

        column = ['Город', 'Заказчик', 'Себестоимость', 'Выручка',
                  'Валовая прибыль', 'Выплачено']
        if is_month:
            column += ['month']

        return margin_table[column]


if __name__ == '__main__':
    pd.options.display.max_columns = 100
    report = ReportTables(data_folder=True)
    print(report.get_margin().head())
