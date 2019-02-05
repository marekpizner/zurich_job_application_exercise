from fx_exchanger import FxExchanger
from xlsx_reader import XlsxReader


def main(annual_wage_file_name, min_wage_file_name, fx_file_name):
    xlsx_reader = XlsxReader(annual_wage_file_name, min_wage_file_name, fx_file_name)


if __name__ == '__main__':
    fx_file_name = './data/FXrates.csv'
    annual_wage_file_name = './data/Dataset annual wages.xlsx'
    min_wage_file_name = './data/Dataset minimum wages.xlsx'

    main(annual_wage_file_name, min_wage_file_name, fx_file_name)
