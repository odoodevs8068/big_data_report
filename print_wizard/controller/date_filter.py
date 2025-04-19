
from odoo import models, fields, api, _
import calendar
from datetime import datetime, timedelta
from odoo.tools import date_utils, format_amount

class DateFilter():

    def get_last_yearly_dates(self):
        start_date = f"{datetime.now().year - 1}-01-01"
        end_date = f"{datetime.now().year - 1}-12-31"
        return start_date, end_date

    def get_yearly_dates(self):
        start_date = f"{datetime.now().year}-01-01"
        end_date = f"{datetime.now().year}-12-31"
        return start_date, end_date

    def get_this_week_date(self):
            today = datetime.now()
            start_of_week = today - timedelta(days=(today.weekday() + 1) % 7)
            end_of_week = start_of_week + timedelta(days=6)
            start_date = start_of_week.strftime('%Y-%m-%d')
            end_date = end_of_week.strftime('%Y-%m-%d')
            return start_date, end_date

    def get_last_week_date(self):
            today = datetime.now()
            start_of_this_week = today - timedelta(days=(today.weekday() + 1) % 7)
            start_date = (start_of_this_week - timedelta(weeks=1)).strftime('%Y-%m-%d')
            end_date = (start_of_this_week - timedelta(days=1)).strftime('%Y-%m-%d')
            return start_date, end_date

    def get_this_month(self):
            now = datetime.now()
            start_date = f"{now.year}-{now.month:02d}-01"
            end_date = f"{now.year}-{now.month:02d}-{calendar.monthrange(now.year, now.month)[1]}"
            return start_date, end_date

    def get_last_month(self):
            today = fields.date.today()
            previous_month = date_utils.subtract(today, months=1)
            start_date = date_utils.start_of(previous_month, "month")
            end_date = date_utils.end_of(previous_month, "month")
            return start_date, end_date

    def get_last_n_date_range(self, data):
            today = datetime.today()
            start_date = today - timedelta(days=int(data))
            end_date = today
            return start_date, end_date

    def get_date_query(self, data):
        if data == 'this_year':
            start_date, end_date = self.get_yearly_dates()
        elif data.strip() == 'last_year':
            start_date, end_date = self.get_last_yearly_dates()
        elif data == 'this_month':
            start_date, end_date = self.get_this_month()
        elif data == 'last_month':
            start_date, end_date = self.get_last_month()
        elif data == 'last_week':
            start_date, end_date = self.get_last_week_date()
        elif data == 'this_week':
            start_date, end_date = self.get_this_week_date()
        elif data in ('7','30', '60', '90', '120', '180', '365'):
            start_date, end_date = self.get_last_n_date_range(data)
        else:
            start_date = f"{datetime.now().year}-01-01"
            end_date = f"{datetime.now().year}-12-31"
        return start_date, end_date