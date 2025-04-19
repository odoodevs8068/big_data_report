from odoo import api, models, fields, _


class PrintWizard(models.TransientModel):
    _name = "print_wizard"
    _description = "Sales Dynamic Report"

    report_type = fields.Selection(selection=lambda self: self._selection_report_type_values(), string='Report Type')
    date_filter = fields.Selection(
        [
            ('7', 'Last 7 days'),
            ('30', 'Last 30 days'),
            ('60', 'Last 60 days'),
            ('90', 'Last 90 days'),
            ('180', 'Last 180 days'),
            ('365', 'Last 365 days'),
            ('this_week', 'Current Week'),
            ('last_week', 'Past Week'),
            ('this_month', 'Current Month'),
            ('last_month', 'Past Month'),
            ('this_year', 'Current Year'),
            ('last_year', 'Last Year'),
            ('custom_date', 'Custom Date'),
        ], string='Date Filter', default='last_month'
    )
    date_from = fields.Date(string='From', default=fields.Datetime.now)
    date_to = fields.Date(string='To', default=fields.Datetime.now)
    categ_id = fields.Many2many('product.category',  string='Product Category')

    def _selection_report_type_values(self):
        values = [('none', 'none')]
        active_model = self.env.context.get('active_model')
        if active_model == 'sale.order':
            values = [
                ('delivery_report', 'Sales Delivered Report'),
                ('pending', 'Sales Undelivered Report'),
                ('done_pending', 'Sales Delivered/Undelivered Report'),
            ]
            if self.env.user.user_has_groups('print_wizard.group_sales_po_report'):
                values += [('sales_po', 'Sales PO Report'),]
            if self.env.user.user_has_groups('print_wizard.group_sales_with_inv_report'):
                values += [('sales_with_inv', 'Sales With Delivery/Invoice')]
        if active_model == 'account.move':
            values = [
                ('sales_with_inv', 'Sales With Delivery/Invoice '),
            ]
        if active_model == 'stock.valuation.layer':
            values = [
                ('valuation_report', 'Inventory Valuation Report'),
                ('ov_movement_report', 'Overall Movement Report'),
                ('movement_report', 'Detailed Inventory Movement Report'),
            ]
        if active_model == 'mrp.bom':
            values = [
                ('bom_report', 'BOM Report'),
            ]
        return values

    def button_print_report(self):
        if self.report_type in ('valuation_report', 'movement_report', 'ov_movement_report') :
            return {
                'type': 'ir.actions.act_url',
                'url': '/inventory_movement_report/excel_report/%s' % (self.id),
                'target': 'new',
            }
        elif self.report_type in ('sales_po', 'delivery_report', 'pending', 'sales_with_inv', 'done_pending') :
            return {
                'type': 'ir.actions.act_url',
                'url': '/sales_po_report/excel_report/%s' % (self.id),
                'target': 'new',
            }
        elif self.report_type == 'invoice_line' :
            return {
                'type': 'ir.actions.act_url',
                'url': '/invoice/excel_report/%s' % (self.id),
                'target': 'new',
            }
        elif self.report_type == 'bom_report' :
            return {
                'type': 'ir.actions.act_url',
                'url': '/bom/excel_report/%s' % (self.id),
                'target': 'new',
            }