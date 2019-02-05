import csv


class FxExchanger:

    def __init__(self, file_name):
        self.file = file_name
        self.exchange_value = {}

        self.read_file()

    def read_file(self):
        with open(self.file) as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                # row[0] - Location
                # row[5] - Year
                # row[6] - Value
                # Read header, skip
                if line_count == 0:
                    line_count += 1
                    continue

                if row[0] not in self.exchange_value:
                    self.exchange_value[row[0]] = {}

                self.exchange_value[row[0]][row[5]] = row[6]

    def exchange_to_euro(self, location, year, value):
        try:
            if location == 'Euro':
                return value
            else:
                return value / float(self.exchange_value[str(location)][str(int(year))])
        except Exception as e:
            print(e)
