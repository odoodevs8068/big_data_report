import xlsxwriter
import io

class WorkSheet():

    def workbook_worksheet(self):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:AZ', 20)
        return output, workbook, worksheet