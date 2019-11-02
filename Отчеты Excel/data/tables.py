import pandas as pd
from data.connection import connect


class AmberTables:
    def __init__(self, data_folder=False):
        current_connection = connect(data_folder)
        self.base = current_connection['base']
        self.connect_db = current_connection['connect_db']

    def __sql(self, table_name):
        return f'SELECT * FROM {self.base}.[{table_name}];'

    def get_questionarys(self):
        questionarys = pd.read_sql(self.__sql('Questionarys'), self.connect_db)

        questionarys['Referrer'].fillna('', inplace=True)
        questionarys['ShiftCount'].fillna(0, inplace=True)
        questionarys['Vozrast'].fillna('', inplace=True)
        questionarys['Nomerkarty'].fillna('', inplace=True)
        questionarys['Vladeleckarty'].fillna('', inplace=True)
        questionarys['Telefon'].fillna('', inplace=True)

        return questionarys

    def get_employee_status(self):
        return pd.read_sql(self.__sql('EmployeeStatus'), self.connect_db)

    def get_employment_status(self):
        return pd.read_sql(self.__sql('EmploymentStatus'), self.connect_db)

    def get_employment_type(self):
        return pd.read_sql(self.__sql('EmploymentType'), self.connect_db)

    def get_specialty(self):
        return pd.read_sql(self.__sql('Specialty'), self.connect_db)

    def get_appearance(self):
        return pd.read_sql(self.__sql('Appearance'), self.connect_db)

    def get_citizenship(self):
        return pd.read_sql(self.__sql('Citizenship'), self.connect_db)

    def get_interst_source(self):
        return pd.read_sql(self.__sql('InterstSource'), self.connect_db)

    def get_interst_source_details(self):
        return pd.read_sql(self.__sql('InterstSourceDetails'), self.connect_db)

    def get_phone_type(self):
        return pd.read_sql(self.__sql('PhoneType'), self.connect_db)

    def get_questionary_phones(self):
        questionary_phones = pd.read_sql(self.__sql('QuestionaryPhones'), self.connect_db)

        phone_type = self.get_phone_type()
        questionary_phones['CommunicationType'] = [
            phone_type[phone_type['Id'] == x]['Name'].values.item(0)
            if pd.notna(x) else ''
            for x in questionary_phones['CommunicationType']]

        return questionary_phones

    def get_facility(self):
        def add_object(s):
            return f'Object{s}'

        facility = pd.read_sql(self.__sql('Facility'), self.connect_db)
        facility.columns = list(map(add_object, facility.columns))

        return facility

    def get_fine_reason(self):
        return pd.read_sql(self.__sql('FineReason'), self.connect_db)

    def get_balance_type(self):
        return pd.read_sql(self.__sql('BalanceType'), self.connect_db)

    def get_shift(self):
        shift = pd.read_sql(self.__sql('Shift'), self.connect_db)

        shift['IsShiftPaid'].fillna(False, inplace=True)
        shift['FineOrBonus'].fillna(0, inplace=True)
        shift['ResultOfShift'].fillna(0, inplace=True)

        shift['Date'] = [pd.Timestamp(date[:10]) for date in shift['Date']]
        return shift

    def get_partners(self):
        return pd.read_sql(self.__sql('Partners'), self.connect_db)

    def get_contact(self):
        return pd.read_sql(self.__sql('Contact'), self.connect_db)

    def get_base_city(self):
        return pd.read_sql(self.__sql('BaseCity'), self.connect_db)

    def get_rate_for_partner(self):
        return pd.read_sql(self.__sql('RateForPartner'), self.connect_db)

    def get_interior_rates(self):
        return pd.read_sql(self.__sql('InteriorRates'), self.connect_db)

    def get_appointed_staff(self):
        return pd.read_sql(self.__sql('AppointedStaff'), self.connect_db)
