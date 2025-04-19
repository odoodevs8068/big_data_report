import json
from odoo import http, _
from odoo.http import content_disposition, request, serialize_exception
from odoo.exceptions import ValidationError
from odoo.addons.print_wizard.controller.date_filter import DateFilter
from odoo.tools import html_escape
from odoo.http import content_disposition, dispatch_rpc, request, serialize_exception as _serialize_exception
date_filter = DateFilter()


class inventoryMovementReportController(http.Controller):

    @http.route([
        '/inventory_movement_report/excel_report/<model("print_wizard"):report_id>',
    ], type='http', auth="user", csrf=False)
    def get_sale_excel_report(self, report_id=None, **args):
        if report_id.date_filter != 'custom_date':
            start_date, end_date = date_filter.get_date_query(report_id.date_filter)
        else:
            start_date, end_date = report_id.date_from, report_id.date_to
        categ_ids = report_id.categ_id.ids
        categ = self.get_category(categ_ids)
        stock_valuation =  request.env['stock.valuation.layer']
        if report_id.report_type == 'valuation_report':
            report_name = 'Valuation Report'
            data = stock_valuation.get_inventory_valuation(start_date, end_date, categ)
        elif report_id.report_type in ('movement_report' , 'ov_movement_report'):
            report_name = 'Movement Report'
            data = stock_valuation.get_inventory_movement_data(start_date, end_date, categ, report_id.report_type)
        else:
            return
        report_bytes = stock_valuation.download_report(data, report_id.report_type)
        try:
            response = request.make_response(
                report_bytes,
                headers=[
                    ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                    ('Content-Disposition', content_disposition(f"{report_name}-{start_date}-{end_date}" + '.xlsx'))
                ]
            )
            response.set_cookie('fileToken', 'dummy-because-api-expects-one')
            return response
        except Exception as e:
            se = _serialize_exception(e)
            error = {
                'code': 200,
                'message': 'Odoo Server Error',
                'data': se
            }
            return request.make_response(html_escape(json.dumps(error)))

    def get_category(self, categ_ids):
        if len(categ_ids) > 1:
            categ = f"pt.categ_id IN {tuple(categ_ids)}"
        else:
            categ = f"pt.categ_id = {categ_ids[0]}"
        return categ


