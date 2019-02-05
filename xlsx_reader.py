import pandas as pd
import xlrd

from fx_exchanger import FxExchanger


class XlsxReader:

    def __init__(self, annual_wage_file, min_wage_file, fx_file_name):
        self.annual_wage_file = annual_wage_file
        self.min_wage_file = min_wage_file
        self.row_namew = 'Current prices in NCU'
        self.current_price_in_nuc_by_country = {}
        self.fx = FxExchanger(fx_file_name)

        self.read_annual_wage_file()

    def read_annual_wage_file(self):
        xl_file = pd.ExcelFile(self.annual_wage_file)
        dx = xl_file.parse(xl_file.sheet_names[0], engine='xlrd').fillna(method='ffill')
        years = []
        for index, row in dx.iterrows():
            if row.values[0] == 'Time':
                years = row.values[4:]
            if row.values[1] == self.row_namew:
                currency_mark = self.convert_unit_name_to_location_mark(row.values[2])
                self.current_price_in_nuc_by_country[row.values[0]] = row.values[4:]
                self.current_price_in_nuc_by_country[row.values[0]] = [
                    self.fx.exchange_to_euro(currency_mark, years[i], x) for i, x in enumerate(row.values[4:])]

        print(self.current_price_in_nuc_by_country)
        return self.current_price_in_nuc_by_country

    def read_min_wage_file(self):
        xl_file = pd.ExcelFile(self.min_wage_file)
        dx = xl_file.parse(xl_file.sheet_names[0], engine='xlrd').fillna(method='ffill')

        pass

    def convert_unit_name_to_location_mark(self, unit_name):

        unit_names = [("Australian Dollar", "AUS"),
                      ("Canadian Dollar", "CAN"),
                      ("Chilean Peso", "CHL"),
                      ("Czech Koruna", "CZE"),
                      ("Danish Krone", "DNK"),
                      ("Euro", "Euro"),
                      ("Forint", "HUN"),
                      ("Iceland Krona", "ISL"),
                      ("Mexican Peso", "MEX"),
                      ("New Israeli Sheqel", "ISR"),
                      ("New Zealand Dollar", "NZL"),
                      ("Norwegian Krone", "NOR"),
                      ("Pound Sterling", "GBR"),
                      ("Swedish Krona", "SWE"),
                      ("Swiss Franc", "FRA"),
                      ("US Dollar", "USA"),
                      ("Won", "CHN"),
                      ("Yen", "JPN"),
                      ("Zloty", "POL")
                      ]
        mark = [item for item in unit_names if item[0] == unit_name]
        try:
            return mark[0][1]
        except Exception as e:
            print(e, unit_name)
            return 'Euro'
