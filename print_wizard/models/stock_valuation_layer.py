from odoo import api, fields, models, _
from odoo.addons.print_wizard.models.worksheet import WorkSheet
WS = WorkSheet()

class StockvaluationLayerInherit(models.Model):
    _inherit = 'stock.valuation.layer'

    def product_details(self):
        return f"""
            pt.name AS product_name,
            pt.default_code AS default_code,
            pc.name AS categ_name
        """

    def get_inventory_valuation(self, start_date, end_date, categ):
        query = f"""
                select 
    	            {self.product_details()},
                    {self.closing_balance(start_date, end_date)},
                    {self.closing_value(start_date, end_date)}
                    {self.condition_query(categ)}
            """
        self.env.cr.execute(query)
        stock_report = self.env.cr.dictfetchall()
        return stock_report

    def get_inventory_movement_data(self, start_date, end_date, categ, type):
        query = f"""
                select 
                    {self.product_details()},
                    {self.opening_balance_and_value(start_date)},
                    {self.stock_detailed_movement_data(start_date, end_date) if type == 'movement_report' else ''}
                    {self.stock_qty_value_in_out(start_date, end_date)},
                    {self.closing_balance(start_date, end_date)},
                    {self.closing_value(start_date, end_date)}
                    {self.condition_query(categ)}
        """
        self.env.cr.execute(query)
        stock_report = self.env.cr.dictfetchall()
        return stock_report

    def closing_balance(self, start_date, end_date):
        return f"""
                	COALESCE(SUM(CASE 
                        WHEN svl.quantity >= 0 
                             AND sm.state = 'done' 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) +
                    COALESCE(SUM(CASE 
                        WHEN svl.quantity <= 0 
                             AND sm.state = 'done' 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS closing_balance
            """

    def closing_value(self, start_date, end_date):
        return f"""
                    COALESCE(SUM(CASE 
                        WHEN svl.value >= 0 
                             AND sm.state = 'done' 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) +
                    COALESCE(SUM(CASE 
                        WHEN svl.value <= 0 
                             AND sm.state = 'done' 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS closing_value
            """

    def opening_balance_and_value(self, start_date):
        return f"""
               COALESCE(SUM(CASE 
                        WHEN svl.quantity >= 0 
                            AND sm.state = 'done' 
                            AND svl.create_date < '{start_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) + 
                    COALESCE(SUM(CASE 
                        WHEN svl.quantity <= 0 
                            AND sm.state = 'done' 
                            AND svl.create_date < '{start_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS opening_balance,    
                    COALESCE(SUM(CASE 
                        WHEN svl.value >= 0 
                            AND sm.state = 'done' 
                            AND svl.create_date < '{start_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) +
                    COALESCE(SUM(CASE 
                        WHEN svl.value <= 0 
                            AND sm.state = 'done' 
                            AND svl.create_date < '{start_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS opening_value     
        """

    def stock_qty_value_in_out(self, start_date, end_date):
        return f"""
            COALESCE(SUM(CASE 
                        WHEN svl.quantity >= 0 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS d_stock_quantity_in,

                    COALESCE(SUM(CASE 
                        WHEN svl.quantity <= 0 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS d_stock_quantity_out,

                    COALESCE(SUM(CASE	
                        WHEN svl.value >= 0 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS d_stock_value_in,

                    COALESCE(SUM(CASE 
                        WHEN svl.value <= 0 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                    ELSE 0 
                    END), 0) AS d_stock_value_out
        """

    def stock_detailed_movement_data(self, start_date, end_date):
        return f"""
                    {self.purchase_in_out(start_date, end_date)},
                    {self.get_mrp_in_out_scrap(start_date, end_date)},
                    {self.get_sales_in_out(start_date, end_date)},
                    {self.get_inventory_in_out(start_date, end_date)},
             """

    def purchase_in_out(self, start_date, end_date):
        return f"""
                    COALESCE(SUM(CASE
                        WHEN spt.code = 'incoming' AND sl_from.usage = 'supplier'
                            AND spt.name != 'Returns' 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS purchase_closing_balance,

                    COALESCE(SUM(CASE
                        WHEN sl_dest.usage = 'supplier'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS purchase_returns_closing_balance,

                    COALESCE(SUM(CASE
                        WHEN spt.code = 'incoming' AND sl_from.usage = 'supplier'
                            AND spt.name != 'Returns' 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS purchase_closing_value,

                    COALESCE(SUM(CASE
                        WHEN sl_dest.usage = 'supplier'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS purchase_returns_closing_value
        """

    def get_mrp_in_out_scrap(self, start_date, end_date):
        return f"""
                COALESCE(SUM(CASE
                        WHEN spt.code = 'mrp_operation' 
                            AND sm.production_id IS NOT NULL AND NOT sm.scrapped
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS mrp_closing_balance,                

                    COALESCE(SUM(CASE
                        WHEN spt.code = 'mrp_operation' 
                            AND sm.raw_material_production_id IS NOT NULL 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS mrp_rm_closing_balance,

                    COALESCE(SUM(CASE 
                        WHEN 
                             sm.state = 'done' 
                             AND sl_dest.usage = 'inventory' 
                             AND sm.scrapped 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS scrapped_cl_balance,

                    COALESCE(SUM(CASE
                        WHEN spt.code = 'mrp_operation' 
                            AND sm.production_id IS NOT NULL AND NOT sm.scrapped
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS mrp_closing_value,

                    COALESCE(SUM(CASE
                        WHEN spt.code = 'mrp_operation' 
                            AND sm.raw_material_production_id IS NOT NULL 
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS mrp_rm_closing_value,

                    COALESCE(SUM(CASE 
                        WHEN 
                             sm.state = 'done' 
                             AND sl_dest.usage = 'inventory' 
                             AND sm.scrapped 
                             AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS scrapped_cl_value

            """

    def get_sales_in_out(self, start_date, end_date):
        return f"""
                    COALESCE(SUM(CASE
                        WHEN sl_dest.usage = 'customer'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS sales_closing_balance,

                    COALESCE(SUM(CASE
                        WHEN sl_from.usage = 'customer'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0
                    END), 0) AS sales_returns_closing_balance,

                    COALESCE(SUM(CASE
                        WHEN sl_dest.usage = 'customer'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS sales_closing_value,

                    COALESCE(SUM(CASE
                        WHEN sl_from.usage = 'customer'
                            AND sm.state = 'done' 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0
                    END), 0) AS sales_returns_closing_value
        """

    def get_inventory_in_out(self, start_date, end_date):
        return f"""
                COALESCE(SUM(CASE 
                        WHEN svl.quantity >= 0 
                            AND sm.state = 'done' AND sl_from.usage = 'inventory' AND sm.is_inventory IS True 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS inv_incr_cl_balance,

                    COALESCE(SUM(CASE 
                        WHEN svl.quantity <= 0 
                            AND sm.state = 'done' AND sl_dest.usage = 'inventory' AND sm.is_inventory IS True 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.quantity
                        ELSE 0 
                    END), 0) AS inv_dec_cl_balance,

                    COALESCE(SUM(CASE 
                        WHEN svl.value >= 0 
                            AND sm.state = 'done' AND sl_from.usage = 'inventory' AND sm.is_inventory IS True 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS inv_incr_cl_value,

                    COALESCE(SUM(CASE 
                        WHEN svl.value <= 0 
                            AND sm.state = 'done' AND sl_dest.usage = 'inventory' AND sm.is_inventory IS True 
                            AND DATE(svl.create_date) BETWEEN '{start_date}' AND '{end_date}'
                        THEN svl.value
                        ELSE 0 
                    END), 0) AS inv_dec_cl_value
        """

    def condition_query(self, categ):
        cond = f"""
                FROM 
                   product_product pp 
                LEFT JOIN
                   product_template pt ON pt.id = pp.product_tmpl_id
                LEFT JOIN
                    product_category pc ON pt.categ_id = pc.id
                LEFT JOIN
                   stock_move sm ON sm.product_id = pp.id
                LEFT JOIN
                    stock_picking_type spt on spt.id = sm.picking_type_id
                LEFT JOIN
                    stock_location sl_from ON sl_from.id = sm.location_id
                LEFT JOIN
                    stock_location sl_dest ON sl_dest.id = sm.location_dest_id
                LEFT JOIN
                   stock_valuation_layer svl ON svl.stock_move_id = sm.id AND svl.product_id = pp.id
                WHERE
                    {categ}
                GROUP BY 
                   pt.name, pt.default_code, pc.name
                ORDER BY
                   pt.name ASC;
            """
        return cond

    def arrange_headers(self, type):
        if type == 'valuation_report':
            headers = ['Item code', 'Product', 'Product Category', 'Closing Balance', 'Closing Value']
        elif type == 'movement_report':
            headers = ['Item code', 'Product', 'Product Category',
                       'Opening Balance', 'Opening Value',
                       'Purchased Qty', 'Purchase Value', 'Purchase Returns Qty', 'Purchase Returns Value',
                       'MRP Qty', 'MRP Value', 'MRP RM Qty', 'MRP RM Value', 'Scrap Qty', 'Scrap Value',
                       'Sales Qty', 'Sales Value', 'Sales Returns Qty', 'Sales Returns Value',
                       'Inventory INC Qty', 'Inventory INC Value', 'Inventory DESC Qty', 'Inventory DESC Value',
                       'Total Qty IN', 'Total Value IN', 'Total Qty OUT', 'Total Value OUT',
                       'Closing Balance', 'Closing Value', 'Final Closing Balance', 'Final Closing Value'
                       ]
        else:
            headers = ['Item code', 'Product', 'Product Category', 'Opening Balance', 'Opening Value',
                       'Total Qty IN', 'Total Value IN', 'Total Qty OUT', 'Total Value OUT',
                       'Closing Balance', 'Closing Value', 'Final Closing Balance', 'Final Closing Value']
        return headers

    def download_report(self, data, type):
        output, workbook, worksheet = WS.workbook_worksheet()
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1})
        total_format = workbook.add_format({'bold': True, 'align': 'right', 'border': 1, 'bg_color': '#FDE9D9'})
        headers = self.arrange_headers(type)
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        row = 1
        total_op_value = 0
        total_cl_value = 0
        if type == 'movement_report':
            workbook, worksheet = self.arrange_detailed_movement_report_data(data, workbook, worksheet, row,
                                                                             total_op_value, total_cl_value,
                                                                             total_format)
        if type == 'valuation_report':
            workbook, worksheet = self.arrange_valuation_report_data(data, workbook, worksheet, row, total_cl_value,
                                                                     total_format)
        if type == 'ov_movement_report':
            workbook, worksheet = self.arrange_ov_movement_report_data(data, workbook, worksheet, row, total_op_value,
                                                                       total_cl_value, total_format)
        workbook.close()
        output.seek(0)
        return output.getvalue()

    def arrange_detailed_movement_report_data(self, data, workbook, worksheet, row, total_op_value, total_cl_value,
                                              total_format):
        total_tcl_value = 0
        for order in data:
            worksheet.write(row, 0, order['default_code'])
            worksheet.write(row, 1, order['product_name'])
            worksheet.write(row, 2, order['categ_name'])
            worksheet.write(row, 3, order['opening_balance'])
            worksheet.write(row, 4, order['opening_value'])
            worksheet.write(row, 5, order['purchase_closing_balance'])
            worksheet.write(row, 6, order['purchase_closing_value'])
            worksheet.write(row, 7, order['purchase_returns_closing_balance'])
            worksheet.write(row, 8, order['purchase_returns_closing_value'])
            worksheet.write(row, 9, order['mrp_closing_balance'])
            worksheet.write(row, 10, order['mrp_closing_value'])
            worksheet.write(row, 11, order['mrp_rm_closing_balance'])
            worksheet.write(row, 12, order['mrp_rm_closing_value'])
            worksheet.write(row, 13, order['scrapped_cl_balance'])
            worksheet.write(row, 14, order['scrapped_cl_value'])
            worksheet.write(row, 15, order['sales_closing_balance'])
            worksheet.write(row, 16, order['sales_closing_value'])
            worksheet.write(row, 17, order['sales_returns_closing_balance'])
            worksheet.write(row, 18, order['sales_returns_closing_value'])
            worksheet.write(row, 19, order['inv_incr_cl_balance'])
            worksheet.write(row, 20, order['inv_incr_cl_value'])
            worksheet.write(row, 21, order['inv_dec_cl_balance'])
            worksheet.write(row, 22, order['inv_dec_cl_value'])
            worksheet.write(row, 23, order['d_stock_quantity_in'])
            worksheet.write(row, 24, order['d_stock_value_in'])
            worksheet.write(row, 25, order['d_stock_quantity_out'])
            worksheet.write(row, 26, order['d_stock_value_out'])
            worksheet.write(row, 27, order['closing_balance'])
            worksheet.write(row, 28, order['closing_value'])
            worksheet.write(row, 29, order['opening_balance'] + order['closing_balance'])
            worksheet.write(row, 30, order['opening_value'] + order['closing_value'])
            total_op_value += order['opening_value']
            total_cl_value += order['closing_value']
            total_tcl_value += order['opening_value'] + order['closing_value']
            row += 1
        worksheet.write(row, 3, "Total OP Value", total_format)
        worksheet.write(row, 4, f"{total_op_value:,.2f} ", total_format)
        worksheet.write(row, 27, "Total CL Value", total_format)
        worksheet.write(row, 28, f"{total_cl_value:,.2f} ", total_format)
        worksheet.write(row, 30, f"{total_tcl_value:,.2f} ", total_format)
        return workbook, worksheet

    def arrange_ov_movement_report_data(self, data, workbook, worksheet, row, total_op_value, total_cl_value,
                                        total_format):
        total_tcl_value = 0
        for order in data:
            worksheet.write(row, 0, order['default_code'])
            worksheet.write(row, 1, order['product_name'])
            worksheet.write(row, 2, order['categ_name'])
            worksheet.write(row, 3, order['opening_balance'])
            worksheet.write(row, 4, order['opening_value'])
            worksheet.write(row, 5, order['d_stock_quantity_in'])
            worksheet.write(row, 6, order['d_stock_value_in'])
            worksheet.write(row, 7, order['d_stock_quantity_out'])
            worksheet.write(row, 8, order['d_stock_value_out'])
            worksheet.write(row, 9, order['closing_balance'])
            worksheet.write(row, 10, order['closing_value'])
            worksheet.write(row, 11, order['opening_balance'] + order['closing_balance'])
            worksheet.write(row, 12, order['opening_value'] + order['closing_value'])
            total_op_value += order['opening_value']
            total_cl_value += order['closing_value']
            total_tcl_value += order['opening_value'] + order['closing_value']
            row += 1
        worksheet.write(row, 3, "Total OP Value", total_format)
        worksheet.write(row, 4, f"{total_op_value:,.2f} ", total_format)
        worksheet.write(row, 9, "Total CL Value", total_format)
        worksheet.write(row, 10, f"{total_cl_value:,.2f} ", total_format)
        worksheet.write(row, 12, f"{total_tcl_value:,.2f} ", total_format)
        return workbook, worksheet

    def arrange_valuation_report_data(self, data, workbook, worksheet, row, total_cl_value, total_format):
        for order in data:
            worksheet.write(row, 0, order['default_code'])
            worksheet.write(row, 1, order['product_name'])
            worksheet.write(row, 2, order['categ_name'])
            worksheet.write(row, 3, order['closing_balance'])
            worksheet.write(row, 4, order['closing_value'])
            total_cl_value += order['closing_value']
            row += 1
        worksheet.write(row, 3, "Total CL Value", total_format)
        worksheet.write(row, 4, f"{total_cl_value:,.2f} ", total_format)
        return workbook, worksheet
