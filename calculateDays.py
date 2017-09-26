import pandas
from pandas.tseries.offsets import BDay
from pandas.tseries.offsets import CustomBusinessDay
import calendar

class next_business_day():

    def __init__(self):

        holidays = ['2017-12-25', '2017-12-26', '2018-01-01',
                    '2018-03-30', '2018-04-02', '2018-05-10',
                    '2018-12-24', '2018-12-25', '2018-12-26',
                    '2018-12-31', '2019-01-01']

        bday_test = CustomBusinessDay(holidays=holidays)
        self.today = (pandas.to_datetime('today'))
        self.tomorrow = (self.today + bday_test)

    def get_day_month(self):
        tomorrow = self.tomorrow.strftime("%Y%m%d")
        months = {k:v for k,v in enumerate(calendar.month_name)}
        today = self.today.strftime("%Y%m%d")[2:]
        day = tomorrow[2:]
        month = months[int(tomorrow[-3:-2])]
        year = tomorrow[:4]

        return today, day, month, year