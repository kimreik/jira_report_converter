import io
import re
import datetime

import xlrd as xlrd
from xlsxwriter import Workbook


class TimeTracking:

    def __init__(self, date, time, desctiption):
        self.date = date
        self.time = time
        desctiption = desctiption.strip()
        # Заменяем LF на CRLF
        match = re.search(r'[^\r]\n', desctiption)
        if match:
            desctiption = re.sub(r'\r', '', desctiption)
            desctiption = re.sub(r'\n', '\r\n', desctiption)
        self.description = desctiption


class ExcelConverter:

    def __init__(self, file):
        self.wb = xlrd.open_workbook(file_contents=file.read())

    def get_file_name(self):
        file_name = 'report_111.xls'
        try:
            date = self.get_some_date()
            name0 = self.get_developer().split()
            name = name0[1] + ' ' + name0[0][0:1] + '.'

            month = str(date.month)
            month = month if len(month) == 2 else '0' + month

            file_name = 'Отчёт о работе за ' + month + '.' + str(date.year) + ', ' + name + '.xlsx'
        except Exception as e:
            print(e)
        return file_name

    def convert(self):
        outBytes = io.BytesIO()
        book = Workbook(outBytes, {'in_memory': True})
        sheet = book.add_worksheet('Лист1')
        sheet.write(0, 0, 'Разработчик')
        sheet.write(0, 1, 'Дата')
        sheet.write(0, 2, 'Время, ч')
        sheet.write(0, 3, 'Комментарий')

        developer = self.get_developer()

        trackings = self.get_trackings()
        for i, x in enumerate(trackings):
            sheet.write(i+1, 0, developer)
            sheet.write(i+1, 1, x.date.strftime('%d.%m.%Y'))
            sheet.write(i+1, 2, x.time)
            sheet.write(i+1, 3, x.description)
        sheet.write(len(trackings)+1, 2, sum([x.time for x in trackings]))

        sheet.set_column(0, 0, 30)
        sheet.set_column(1, 2, 10)
        sheet.set_column(3, 3, 80)

        book.close()
        return outBytes

    def get_developer(self):
        ws = self.wb.sheet_by_index(0)
        return ws.cell(1, 5).value

    def get_some_date(self):
        ws = self.wb.sheet_by_index(0)
        return self.get_date(ws.cell(1, 3))

    def get_trackings(self):
        ws = self.wb.sheet_by_index(0)
        trackings = [x for x in ws.get_rows()][1:]
        trackings = [TimeTracking(self.get_date(x[3]), x[2].value, x[22].value) for x in trackings]
        trackings = sorted(trackings, key=lambda tt: tt.date)
        return trackings

    def get_date(self, cell):
        return datetime.datetime(*xlrd.xldate_as_tuple(cell.value, self.wb.datemode))
