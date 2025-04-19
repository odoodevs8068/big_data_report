import json
from odoo import http, _
from odoo.http import content_disposition, request, serialize_exception
from odoo.exceptions import ValidationError
from odoo.addons.print_wizard.controller.date_filter import DateFilter
from odoo.tools import html_escape
from odoo.http import content_disposition, dispatch_rpc, request, serialize_exception as _serialize_exception
date_filter = DateFilter()

class SalesReportController(http.Controller):

    @http.route(['/sales_po_report/excel_report/<model("print_wizard"):report_id>',], type='http', auth="user", csrf=False)
    def get_sale_excel_report(self, report_id=None, **args):
        if report_id.date_filter != 'custom_date':
            start_date, end_date = date_filter.get_date_query(report_id.date_filter)
        else :
            start_date, end_date = report_id.date_from , report_id.date_to
        sales = request.env['sale.order']
        report_bytes = None
        delivered_condition = f"sm.state = 'done' AND sm.sale_line_id IS NOT NULL AND DATE(sm.date) BETWEEN '{start_date}' AND '{end_date}' "
        pending_condition = f"sm.state NOT IN ('done', 'cancel') AND sm.sale_line_id IS NOT NULL AND DATE(sp.scheduled_date) BETWEEN '{start_date}' AND '{end_date}' "
        report_name = "Report"
        if report_id.report_type == 'sales_po':
            report_name = "Sales PO Report"
            domain_condition = f"DATE(so.date_order) BETWEEN '{start_date}' AND '{end_date}' AND so.state != 'cancel'"
            get_sales_po_data = sales.sales_po_query(domain_condition)
            report_bytes = sales.get_sale_po_xlsx_report(get_sales_po_data)
        if report_id.report_type == 'delivery_report':
            report_name = "Sales Delivered Report"
            get_delivery_report_data = sales.with_context(report_type='delivery_report').delivered_amt_report(delivered_condition)
            report_bytes = sales.with_context(report_type='delivery_report').generate_delivery_report(get_delivery_report_data)
        if report_id.report_type == 'pending':
            report_name = "Sales UnDelivered Report"
            get_pending_report_data =   sales.with_context(report_type='pending').delivered_amt_report(pending_condition)
            report_bytes = sales.with_context(report_type='pending').generate_delivery_report(get_pending_report_data)
        if report_id.report_type == 'sales_with_inv':
            report_name = "Sales With Delivery/Invoice Report"
            domain_condition = f"so.state in ('sale', 'done') AND DATE(so.date_order) BETWEEN '{start_date}' AND '{end_date}'"
            get_sales_with_inv_report_data = sales.get_sales_invoice_with_delivery(domain_condition)
            report_bytes = sales.prepare_report_delivery_with_invoice(get_sales_with_inv_report_data)
        if report_id.report_type == 'done_pending':
            report_name = "Sales Delivery/UnDelivered Report"
            delivered_report_data = sales.with_context(report_type='delivery_report').delivered_amt_report(delivered_condition)
            pending_report_data = sales.with_context(report_type='pending').delivered_amt_report(pending_condition)
            report_bytes = sales.generate_sales_delivered_undelivered_report(delivered_report_data, pending_report_data)
        # if report_id.report_type == 'detail_rep':
        #     sale_delivery_invoice_details = sales.sale_delivery_invoice_details(start_date, end_date)
        #     report_bytes = sales.prepare_report_sale_delivery_invoice(sale_delivery_invoice_details)
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

    @http.route(['/bom/excel_report/<model("print_wizard"):report_id>',], type='http', auth="user", csrf=False)
    def get_bom_excel_report(self, report_id=None, **args):
        bom = request.env['mrp.bom']
        report_bytes = bom.get_bom_excel_report()
        try:
            response = request.make_response(
                report_bytes,
                headers=[
                    ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                    ('Content-Disposition', content_disposition("BOM Structure & Cost Report" + '.xlsx'))
                ]
            )
            response.set_cookie('fileToken', 'dummy-because-api-expects-one')
            return response
        except Exception as e:
            se = _serialize_exception(e)
            error = {
                'code': 200,
                'message': f'There is Some Error While Generating Report \n Please Contact Your Administator',
                'data': se
            }
            return request.make_response(html_escape(json.dumps(error)))
