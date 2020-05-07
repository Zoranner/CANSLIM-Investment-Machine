import os
import config
import pandas
from EmQuant.EmQuantAPI import *
from openpyxl.workbook import Workbook


def print_data(data):
    if (not isinstance(data, c.EmQuantData)):
        print(data)
    else:
        if (data.ErrorCode != 0):
            print("request css Error, ", data.ErrorMsg)
        else:
            print(data)
            for code in data.Codes:
                for i in range(0, len(data.Indicators)):
                    print(data.Data[code][i])


def season_day(year, season):
    day_switch = {1: '3-31', 2: '6-30', 3: '9-30', 4: '12-31'}
    return '{}-{}'.format(year, day_switch[season])


class DataHandler:
    data_list = []
    current_path = os.getcwd()
    excel_path = r'{}\\east_money.xlsx'.format(current_path)
    excel_writer = pandas.ExcelWriter(excel_path)

    def deal(self, year, season, data):
        # data.set_index(['CODES'])
        # data.reset_index(drop=False, inplace=True)
        season_str = '{}SE{}'.format(year, season)
        data.drop('DATES', axis=1, inplace=True)
        data.rename(columns={'INCOMESTATEMENTQ_83': '{}ISTQ83'.format(
            season_str), 'INCOMESTATEMENTQ_80': '{}ISTQ80'.format(season_str)}, inplace=True)
        self.data_list.append(data)
        print(season_str)

    def done(self):
        temp_data = pandas.DataFrame()
        for data in self.data_list:
            temp_data = pandas.concat([temp_data, data], axis=1)

        temp_data.to_excel(self.excel_writer,
                           sheet_name='MONEYDATA', index=True)
        self.excel_writer.save()


if __name__ == '__main__':
    data_handler = DataHandler()

    # loginresult为c.EmQuantData类型数据
    loginresult = c.start(
        options='TestLatency=0,ForceLogin=1,RecordLoginInfo=1')

    for year in range(2018, 2021):
        season_stop = 5
        if year == 2020:
            season_stop = 2
        for season in range(1, season_stop):
            season_data = c.css(config.codes, config.indicators,
                                'ReportDate={},Ispandas=1'.format(season_day(year, season)))
            data_handler.deal(year, season, season_data)

    data_handler.done()
