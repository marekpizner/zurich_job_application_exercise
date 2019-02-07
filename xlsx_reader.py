import pandas as pd
import xlrd

from fx_exchanger import FxExchanger


class XlsxReader:

    def __init__(self, annual_wage_file, min_wage_file, fx_file_name):
        self.annual_wage_file = annual_wage_file
        self.min_wage_file = min_wage_file
        self.row_namew = 'Current prices in NCU'

        self.unit_names = [("Australian Dollar", "AUS"),
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
                           ("Won", "KOR"),
                           ("Yen", "JPN"),
                           ("Zloty", "POL"),
                           ("Turkish Lira", "TUR")
                           ]

        self.current_price_annual = {}
        self.current_price_min = {}

        self.fx = FxExchanger(fx_file_name)

        self.read_annual_wage_file()
        self.read_min_wage_file()

        self.first_table()
        self.second_table()

        self.write_to_file()

    def country_check(self, string):
        for x in self.unit_names:
            if string in x:
                return True
        return False

    def read_annual_wage_file(self):
        xl_file = pd.ExcelFile(self.annual_wage_file)
        dx = xl_file.parse(xl_file.sheet_names[0], engine='xlrd').fillna(method='ffill')
        self.annual_years = []
        for index, row in dx.iterrows():
            if row.values[0] == 'Time':
                self.annual_years = [int(x) for x in row.values[4:]]
            if row.values[1] == self.row_namew:
                currency_mark = self.convert_unit_name_to_location_mark(row.values[2])
                for i, x in enumerate(row.values[4:]):
                    if row.values[0] not in self.current_price_annual:
                        self.current_price_annual[row.values[0]] = {}
                    self.current_price_annual[row.values[0]][self.annual_years[i]] = self.fx.exchange_to_euro(
                        currency_mark,
                        self.annual_years[i], x)

        return self.current_price_annual

    def read_min_wage_file(self):
        xl_file = pd.ExcelFile(self.min_wage_file)
        dx = xl_file.parse(xl_file.sheet_names[0], engine='xlrd').fillna(method='ffill')
        self.min_years = []
        for index, row in dx.iterrows():
            if row.values[0] == 'Time':
                self.min_years = [int(x) for x in row.values[3:]]

            if index >= 4 and self.country_check(row.values[1]):
                currency_mark = self.convert_unit_name_to_location_mark(row.values[1])
                for i, x in enumerate(row.values[3:]):
                    if row.values[0] not in self.current_price_min:
                        self.current_price_min[row.values[0]] = {}
                    self.current_price_min[row.values[0]][self.min_years[i]] = \
                        self.fx.exchange_to_euro(currency_mark, self.min_years[i], x)

        return self.current_price_min

    def convert_unit_name_to_location_mark(self, unit_name):

        mark = [item for item in self.unit_names if item[0] == unit_name]
        try:
            return mark[0][1]
        except Exception as e:
            print(e, unit_name)
            return 'Euro'

    def first_table(self):
        min_wage = pd.DataFrame.from_dict(self.current_price_min, orient='index')
        ann_wage = pd.DataFrame.from_dict(self.current_price_annual, orient='index')
        years = list(min_wage)
        years_min = {}
        years_max = {}

        for key in years:
            if key in min_wage and key in ann_wage:
                min_for_year = min_wage[key]
                ann_for_year = ann_wage[key]

                years_min[key], years_max[key] = self.calculate_ratio(min_for_year, ann_for_year)

        tmp = {'maximum ratio for each year': years_max, 'minimum ratio for each year': years_min}

        tmp = pd.DataFrame.from_dict(tmp, orient='index')
        self.table_1 = min_wage
        self.table_1.append(tmp)

    def second_table(self):
        countries_ema = self.ema()
        countries_wages = self.current_price_annual
        ratio = {}
        for country, years in countries_ema.items():
            if country in countries_wages:
                for year, value_ema in years.items():
                    if year <= 2016:
                        if country not in ratio:
                            ratio[country] = {}
                        ratio[country][year] = value_ema / countries_wages[country][2017]
        self.table_2 = pd.DataFrame.from_dict(ratio).T
        pass

    def calculate_ratio(self, min_values, annual_values):
        min_ratio = 9999999
        max_ratio = 0
        for key, value in annual_values.items():
            if key in min_values:
                try:
                    diff = min_values[key] / value
                    if diff > max_ratio:
                        max_ratio = diff
                    if diff < min_ratio:
                        min_ratio = diff
                except Exception as e:
                    pass

        return min_ratio, max_ratio

    def ema(self):
        ann = pd.DataFrame.from_dict(self.current_price_annual, orient='index')
        countries_ema = {}
        for country, values in ann.iterrows():
            countries_ema[country] = self.ema_calculation(values.to_dict(), 2016)

        return countries_ema

    def ema_calculation(self, values, max_year):
        k = 2 / (7 + 1)
        prev_ema = 1
        new_values = {}
        for key, value in values.items():
            if key <= max_year:
                if prev_ema == 1:
                    new_values[key] = value
                else:
                    new_values[key] = value * k + prev_ema - prev_ema * k

                prev_ema = new_values[key]
            else:
                new_values[key] = value

        return new_values

    def write_to_file(self):
        p_annual = pd.DataFrame.from_dict(self.current_price_annual, orient='index')
        p_annual.to_csv('./data/annual_wages.csv')

        p_min = pd.DataFrame.from_dict(self.current_price_min, orient='index')
        p_min.to_csv('./data/min_wages.csv')

        self.table_1.to_csv('./data/table1.csv')
        self.table_2.to_csv('./data/table2.csv')
