from odoo import api, models, fields, _
import json
from datetime import datetime, timedelta
from odoo.addons.print_wizard.models.worksheet import WorkSheet
WS = WorkSheet()


class SaleOrderInherit(models.Model):
    _inherit = 'sale.order'

    def sales_po_query(self, domain):
        query = f"""
                       WITH manufacturing_status_cte AS (
                                    SELECT 
                                        mp.product_id,
                                        CASE
                                            WHEN EXISTS (
                                                SELECT 1
                                                FROM mrp_production mp_sub
                                                WHERE mp_sub.product_id = mp.product_id AND mp_sub.state IN ('progress', 'confirmed')
                                            ) THEN 'In Production'
                                            WHEN EXISTS (
                                                SELECT 1
                                                FROM mrp_production mp_sub
                                                WHERE mp_sub.product_id = mp.product_id AND mp_sub.state = 'done'
                                            ) THEN 'Manufactured'
                                            WHEN EXISTS (
                                                SELECT 1
                                                FROM mrp_production mp_sub
                                                WHERE mp_sub.product_id = mp.product_id
                                            ) THEN 'Pending Production'
                                            ELSE 'Not Applicable'
                                        END AS manufacturing_status
                                    FROM 
                                        mrp_production mp
                                    GROUP BY 
                                        mp.product_id
                                )
                                SELECT 
                                    so.name AS order_name,
                                    rp.name AS partner_name,
                                    TO_CHAR(so.date_order, 'dd/mm/yyyy') AS create_date,
                                    so.commitment_date AS delivery_date,
                                   TRIM(
                                         COALESCE(dp.street, '') || ' ' ||
                                         COALESCE(dp.street2, '') || ' ' ||
                                         COALESCE(dp.city, '') || ' ' ||
                                         COALESCE(d_cs.name, '') || ' ' ||
                                         COALESCE(d_rc.name, '')
                                     ) AS delivery_address,

                                    so.client_order_ref AS order_type,
                                    ru_rp.name AS salesperson,
                                    so.amount_total AS amount_total,
                                    line.id AS order_line_id,
                                    pt.default_code AS default_code,
                                    pt.name AS product_name,
                                    line.product_uom_qty AS quantity,
                                    line.price_unit AS unit_price,
                                    um.name AS uom,
                                    line.product_uom_qty AS ordered_qty,
                                    line.qty_delivered AS qty_delivered,
                                    (line.qty_delivered * line.price_unit) AS delivered_amt,
                                    (line.product_uom_qty - line.qty_delivered) AS balance_qty_to_deliver,
                                    line.price_subtotal - (line.qty_delivered * line.price_unit) AS balance_amt_to_deliver,
                                    line.qty_invoiced AS qty_invoiced,
                                    (line.qty_invoiced * line.price_unit) AS invoiced_amt,
                                    (line.product_uom_qty - line.qty_invoiced) AS balance_qty_to_invoice,
                                    line.price_subtotal - (line.qty_invoiced * line.price_unit) AS balance_amt_to_invoice,
                                    line.price_subtotal AS subtotal,
                                    line.price_total AS line_total,
                                    cu.full_name AS currency_name,
                                    STRING_AGG(DISTINCT sm.reference, ', ') AS picking_names,
                                        CASE  so.invoice_status
                                        WHEN 'upselling' THEN 'Upselling Opportunity'
                                        WHEN 'invoiced' THEN 'Fully Invoiced'
                                        WHEN 'to invoice' THEN 'To Invoice'
                                        WHEN 'no' THEN 'Nothing to Invoice'
                                        ELSE so.invoice_status
                                    END AS invoice_status,
                                    pt.description_sale AS description_sale,
                                    ms_cte.manufacturing_status,
                                    STRING_AGG(DISTINCT crm.name, ', ') AS tags,
                                    JSON_AGG(
                                        DISTINCT (JSON_BUILD_OBJECT( 
                                            'name', sp.name,
                                            'quantity_demand', sm.product_uom_qty,
                                            'reserved_qty', COALESCE((
                                                SELECT SUM(sml.product_qty)
                                                FROM stock_move_line sml
                                                WHERE sml.move_id = sm.id AND sp.state NOT IN ('done', 'cancel') 
                                            ), 0),
                                            'done_qty', COALESCE((
                                                SELECT SUM(sml.qty_done)
                                                FROM stock_move_line sml
                                                WHERE sml.move_id = sm.id AND sp.state = 'done'
                                            ), 0),
                                            'remaining_qty_to_done', CASE WHEN sm.state NOT IN ('done', 'cancel') THEN sm.product_uom_qty ELSE 0 END,
                                            'state', CASE sp.state
                                                    WHEN 'done' THEN 'Delivered'
                                                    WHEN 'waiting' THEN 'Waiting Another Operation'
                                                    WHEN 'draft' THEN 'Draft'
                                                    WHEN 'confirmed' THEN 'Waiting'
                                                    WHEN 'assigned' THEN 'Ready'
                                                    WHEN 'cancel' THEN 'Cancelled'
                                                    ELSE sp.state
                                                END,
                                            'scheduled_date', TO_CHAR(sp.scheduled_date, 'dd/mm/yyyy'),
                                            'effective_date', TO_CHAR(sp.date_done, 'dd/mm/yyyy')
                                        ))::text 
                                    )::json AS picking_details
                                FROM 
                                    sale_order so
                                LEFT JOIN 
                                    res_partner rp ON rp.id = so.partner_id
                                LEFT JOIN 
                                    res_partner dp ON dp.id = so.partner_shipping_id
                                LEFT JOIN 
                                    sale_order_line line ON line.order_id = so.id AND line.display_type is NULL
                                LEFT JOIN 
                                    product_product pp ON pp.id = line.product_id
                                LEFT JOIN 
                                    product_template pt ON pt.id = pp.product_tmpl_id
                                LEFT JOIN 
                                    uom_uom um ON um.id = line.product_uom
                                LEFT JOIN 
                                    res_currency cu ON cu.id = so.currency_id
                                LEFT JOIN 
                                    res_country_state d_cs ON d_cs.id = dp.state_id
                                LEFT JOIN
                    	            res_country d_rc on d_rc.id = dp.country_id
                                LEFT JOIN 
                                    res_users ru ON ru.id = so.user_id
                                LEFT JOIN 
                                    res_partner ru_rp ON ru_rp.id = ru.partner_id
                                LEFT JOIN 
                                    stock_move sm ON sm.sale_line_id = line.id
                                LEFT JOIN 
                                    stock_picking sp ON sp.id = sm.picking_id
                                LEFT JOIN 
                                    manufacturing_status_cte ms_cte ON ms_cte.product_id = line.product_id
                                LEFT JOIN 
                                    sale_order_tag_rel rel ON rel.order_id = so.id 
                                LEFT JOIN 
                                    crm_tag crm ON crm.id = rel.tag_id
                                WHERE
                                    {domain}
                                GROUP BY 
                                    so.name, rp.name, line.id, so.date_order, pt.description_sale,delivery_address,
                                     so.client_order_ref, so.commitment_date,
                                    ru_rp.name, so.amount_total,
                                    pt.default_code, pt.name, line.product_uom_qty, line.price_unit, 
                                    um.name, cu.full_name, so.invoice_status, ms_cte.manufacturing_status
                                ORDER BY 
                                    so.name DESC;
                """
        self._cr.execute(query)
        sales_details = self.env.cr.dictfetchall()
        return sales_details

    def get_sale_po_xlsx_report(self, report_data):
        data = report_data
        output, workbook, worksheet = WS.workbook_worksheet()

        generated_by = "Generated by: {}".format(self.env.user.name)
        from pytz import timezone
        user_tz = self.env.user.tz or 'UTC'
        local_tz = timezone(user_tz)
        formatted_datetime = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        generated_on = f"Generated on: {formatted_datetime}"
        bold = workbook.add_format({'bold': True})

        worksheet.write(0, 0, generated_by, bold)
        worksheet.write(1, 0, generated_on, bold)

        bold = workbook.add_format({'bold': True})
        headers = ['Order Reference NO.',
                   'PO Customer NO',
                   'Create Date',
                   'Customer Name',
                   'Delivery Address',
                   '',
                   'SalesPerson',
                   'Customer PO Delivery Date',
                   'Tags',
                   'Description code',
                   'ITEM CODE#',
                   'Item Name',
                   'UOM',
                   'Unit Price',
                   'Total Item QTY',
                   'Total Item Amt',
                   'Currency',
                   'QTY Delivered',
                   'Delivered Amt',
                   'Balance Qty To Deliver',
                   'Balance Amt To Deliver',
                   'Qty Invoiced',
                   'Amt Invoiced',
                   'Balanace Qty To Invoice',
                   'Balance Amt To Invoice',
                   'Invoice Status',
                   'WH/Out Name',
                   'Qty Demand',
                   'Reserved Qty',
                   'Done Qty',
                   'Qty To Done',
                   'Scheduled Date',
                   'Delivered Date',
                   'WH/Out Status',
                   ]
        for col, header in enumerate(headers):
            worksheet.write(4, col, header, bold)

        row = 5
        for order in data:
            worksheet.write(row, 0, order['order_name'])
            worksheet.write(row, 1, order['order_type'])
            worksheet.write(row, 2, order['create_date'])
            worksheet.write(row, 3, order['partner_name'])
            worksheet.write(row, 4, order['delivery_address'])
            worksheet.write(row, 5, '')
            worksheet.write(row, 6, order['salesperson'])
            worksheet.write(row, 7, order['delivery_date'])
            worksheet.write(row, 8, order['tags'])
            worksheet.write(row, 9, order['description_sale'])
            worksheet.write(row, 10, order['default_code'])
            worksheet.write(row, 11, order['product_name'])
            worksheet.write(row, 12, order['uom'])
            worksheet.write(row, 13, order['unit_price'])
            worksheet.write(row, 14, order['ordered_qty'])
            worksheet.write(row, 15, order['subtotal'])
            worksheet.write(row, 16, order['currency_name'])
            worksheet.write(row, 17, order['qty_delivered'])
            worksheet.write(row, 18, order['delivered_amt'])
            worksheet.write(row, 19, order['balance_qty_to_deliver'])
            worksheet.write(row, 20, order['balance_amt_to_deliver'])
            worksheet.write(row, 21, order['qty_invoiced'])
            worksheet.write(row, 22, order['invoiced_amt'])
            worksheet.write(row, 23, order['balance_qty_to_invoice'])
            worksheet.write(row, 24, order['balance_amt_to_invoice'])
            worksheet.write(row, 25, order['invoice_status'])
            if 'picking_details' in order and order['picking_details']:
                for picking in order['picking_details']:
                    picking_data = json.loads(picking)
                    worksheet.write(row, 0, order['order_name'])
                    worksheet.write(row, 1, order['order_type'])
                    worksheet.write(row, 2, order['create_date'])
                    worksheet.write(row, 3, order['partner_name'])
                    worksheet.write(row, 4, order['delivery_address'])
                    worksheet.write(row, 5, '')
                    worksheet.write(row, 6, order['salesperson'])
                    worksheet.write(row, 7, order['delivery_date'])
                    worksheet.write(row, 8, order['tags'])
                    worksheet.write(row, 9, order['description_sale'])
                    worksheet.write(row, 10, order['default_code'])
                    worksheet.write(row, 11, order['product_name'])
                    worksheet.write(row, 12, order['uom'])
                    worksheet.write(row, 26, picking_data['name'])
                    worksheet.write(row, 27, picking_data['quantity_demand'])
                    worksheet.write(row, 28, picking_data['reserved_qty'])
                    worksheet.write(row, 29, picking_data['done_qty'])
                    worksheet.write(row, 30, picking_data['remaining_qty_to_done'])
                    worksheet.write(row, 31, picking_data.get('scheduled_date'))
                    worksheet.write(row, 32, picking_data.get('effective_date'))
                    worksheet.write(row, 33, picking_data['state'])
                    row += 1
            else:
                row += 1

        workbook.close()
        output.seek(0)
        return output.getvalue()

    #  SALE DELIVERY / UNDERVERED SECTION

    def get_qty_query(self):
        if 'report_type' in self.env.context :
            if self.env.context.get('report_type') == 'delivery_report':
                query = f"""
                    TO_CHAR(sm.date, 'YYYY-MM-DD') AS date,
                    COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id ), 0) AS done_qty,
                    (CASE 
                        WHEN sm.state = 'done' AND sm.sale_line_id IS NOT NULL THEN
                                        CASE 
                                            WHEN sol.currency_id != (SELECT currency_id FROM res_company WHERE id = sm.company_id) 
                                            THEN 
                                                CASE 
                                                    WHEN sol.product_uom IS NOT NULL AND sol.product_uom != sm.product_uom
                                                    THEN COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id ), 0) * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))  * (1 / COALESCE(rc.rate, 1))
                                                    ELSE COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id ), 0)  * sol.price_unit * (1 / COALESCE(rc.rate, 1))
                                                END
                                            ELSE 
                                                CASE 
                                                    WHEN sol.product_uom IS NOT NULL AND sol.product_uom != sm.product_uom
                                                    THEN COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id ), 0) * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))
                                                    ELSE COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id ), 0) * sol.price_unit
                                                END
                                        END
                                    ELSE 
                                        0.0 
                                END) * CASE WHEN sm.reference LIKE 'WH/RET%' THEN -1 ELSE 1 END AS exchanged_delivery_amt
                """
            else:
                query = f"""
                TO_CHAR(sp.scheduled_date, 'YYYY-MM-DD') AS date,
				sm.product_uom_qty AS balance_qty,
				(CASE
					WHEN sm.state NOT IN ('done', 'cancel') AND sm.sale_line_id IS NOT NULL THEN
						CASE 
							WHEN sol.currency_id != (SELECT currency_id FROM res_company WHERE id = sm.company_id)
							THEN
								CASE
									WHEN sol.product_uom != sm.product_uom
									THEN sm.product_uom_qty * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))  * (1 / COALESCE(rc.rate, 1))
									ELSE sm.product_uom_qty * (sol.price_unit * (1 / COALESCE(rc.rate, 1)))
								END
							ELSE
								CASE
									WHEN sol.product_uom != sm.product_uom
									THEN sm.product_uom_qty * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))  
									ELSE sm.product_uom_qty * sol.price_unit
								END
						END
                      ELSE 
                           0.0 
                END) * CASE WHEN sm.reference LIKE 'WH/RET%' THEN -1 ELSE 1 END AS exchanged_undelivery_amt
                """
            return query

    def delivered_amt_report(self,domain):
        query = f"""
            SELECT 
                sm.origin AS orgin,
                so.client_order_ref AS po_ref,
				CASE 
                    WHEN so.partner_id != so.partner_shipping_id  
                    THEN rp.name || ', ' || rs.name 
                    ELSE rs.name 
                END AS customer,
				(
					SELECT STRING_AGG(DISTINCT crm.name, ', ') 
					FROM sale_order_tag_rel rel
					LEFT JOIN crm_tag crm ON crm.id = rel.tag_id
					WHERE rel.order_id = so.id
				 ) AS tags,
                sm.reference,
                pt.default_code,
                pt.name AS product_name,
                p_uom.name AS product_uom,
                sol.price_unit,
                sol_cu.name AS currency,
                COALESCE((SELECT SUM(sml.product_qty)
                   FROM stock_move_line sml
                   WHERE sml.move_id = sm.id AND sm.state NOT IN ('done', 'cancel') 
                   ), 0) AS reserved_qty,
                CASE sp.state
                        WHEN 'done' THEN 'Delivered'
                        WHEN 'waiting' THEN 'Waiting Another Operation'
                        WHEN 'draft' THEN 'Draft'
                        WHEN 'confirmed' THEN 'Waiting'
                        WHEN 'assigned' THEN 'Ready'
                        WHEN 'cancel' THEN 'Cancelled'
                        ELSE sp.state
                END AS status,
                {self.get_qty_query()}
            FROM 
                stock_move sm
            LEFT JOIN 
                sale_order_line sol ON sm.sale_line_id = sol.id
            LEFT JOIN 
                sale_order so ON sol.order_id = so.id
            LEFT JOIN 
                res_partner rp ON so.partner_id = rp.id
			LEFT JOIN 
				res_partner rs ON so.partner_shipping_id = rs.id
            LEFT JOIN
                product_product pp on sm.product_id = pp.id
            LEFT JOIN 
                product_template pt on pt.id = pp.product_tmpl_id
            LEFT JOIN 
                uom_uom ac_uom on ac_uom.id = sol.product_uom
            LEFT JOIN 
                uom_uom p_uom ON p_uom.id = sm.product_uom
            LEFT JOIN 
                res_currency sol_cu ON sol_cu.id = sol.currency_id
            LEFT JOIN 
                res_currency_rate rc ON rc.currency_id = sol.currency_id
                AND rc.name = (SELECT MAX(name) FROM res_currency_rate WHERE currency_id = sol.currency_id)
            LEFT JOIN 
				stock_picking sp ON sp.id = sm.picking_id
			WHERE 
                {domain}
			ORDER BY sm.date DESC;
        """
        self._cr.execute(query)
        data = self._cr.dictfetchall()
        return data

    def hearders(self, worksheet, workbook, hearder_row):
        if 'report_type' in self.env.context and self.env.context.get('report_type') == 'delivery_report':
            head_1 = "WH/OUT QTY"
            head_2 = "DELIVERED AMT"
            head_3 = "WH/OUT DATE"
        else:
            head_1 = 'BALANCE QTY'
            head_2 = 'BALANCE AMT'
            head_3 = "CUSTOMER REQUESTED DATE"
        bold = workbook.add_format({'bold': True})
        headers = [
            'Order Reference NO.',
            'Po Reference NO.',
            'Customer',
            "Tags",
            'WH/OUT NAME',
            f'{head_3}',
            'DAMAD CODE',
            'PRODUCT NAME',
            'RESERVED QTY',
            f'{head_1}',
            'PRODUCT UOM',
            'UNIT PRICE',
            f'{head_2}',
            'WH/OUT STATUS'
        ]
        for col, header in enumerate(headers):
            worksheet.write(hearder_row, col, header, bold)
        return worksheet

    def generate_delivery_report(self, report_data):
        data = report_data
        output, workbook, worksheet =  WS.workbook_worksheet()
        worksheet = self.hearders(worksheet, workbook, 1)
        total_format = workbook.add_format({'bold': True, 'align': 'right', 'border': 1, 'bg_color': '#FDE9D9'})
        row = 2
        total_amt = 0
        worksheet, row, total_amt = self.arrange_data(data, worksheet, row, total_amt)
        worksheet.write(row, 11, "Total Amount", total_format)
        worksheet.write(row, 12, f"{total_amt:,.2f} ", total_format)
        workbook.close()
        output.seek(0)
        return output.getvalue()

    def arrange_data(self, data, worksheet, row, total_amt):
        for order in data:
            worksheet.write(row, 0, order['orgin'])
            worksheet.write(row, 1, order['po_ref'])
            worksheet.write(row, 2, order['customer'])
            worksheet.write(row, 3, order['tags'])
            worksheet.write(row, 4, order['reference'])
            worksheet.write(row, 5, order['date'])
            worksheet.write(row,6, order['default_code'])
            worksheet.write(row, 7, order['product_name'])
            worksheet.write(row, 8, order['reserved_qty'])
            worksheet.write(row, 10, order['product_uom'])
            worksheet.write(row, 11, order['price_unit'])
            # worksheet.write(row, 16, order['currency'])
            worksheet.write(row, 13, order['status'])
            if 'report_type' in self.env.context and self.env.context.get('report_type') == 'delivery_report':
                worksheet.write(row, 9, order['done_qty'])
                worksheet.write(row, 12, order['exchanged_delivery_amt'])
                total_amt += order['exchanged_delivery_amt'] if order['exchanged_delivery_amt'] else 0
            else:
                worksheet.write(row, 9, order['balance_qty'])
                worksheet.write(row, 12, order['exchanged_undelivery_amt'])
                total_amt += order['exchanged_undelivery_amt'] if order['exchanged_undelivery_amt'] else 0
            row += 1
        return worksheet, row, total_amt

    def generate_sales_delivered_undelivered_report(self, delivered_data, pending_data):
        output, workbook, worksheet =  WS.workbook_worksheet()
        worksheet = self.with_context(report_type='delivery_report').hearders(worksheet, workbook, 1)
        total_format = workbook.add_format({'bold': True, 'align': 'right', 'border': 1, 'bg_color': '#FDE9D9'})
        row = 2
        total_delivered_amt = 0
        worksheet, row, total_delivered_amt = self.with_context(report_type='delivery_report').arrange_data(delivered_data, worksheet, row, total_delivered_amt)
        worksheet.write(row, 11, "Total Amount", total_format)
        worksheet.write(row, 12, f"{total_delivered_amt:,.2f} ", total_format)

        row += 2
        total_pending_amt = 0
        worksheet = self.with_context(report_type='pending').hearders(worksheet, workbook, row)
        row +=1
        worksheet, row, total_pending_amt = self.with_context(report_type='pending').arrange_data(pending_data, worksheet, row, total_pending_amt)
        worksheet.write(row, 11, "Total Amount", total_format)
        worksheet.write(row, 12, f"{total_pending_amt:,.2f} ", total_format)

        workbook.close()
        output.seek(0)
        return output.getvalue()

#  SALES WITH INVOICE / DELIVERY
    def get_sales_invoice_with_delivery(self, domain):
        query = f"""
            SELECT 
                so.name AS origin,
                so.amount_total AS amount_total,
                COALESCE((
                    SELECT   
                        CASE 
                            WHEN MAX(so.currency_id) != res_co.currency_id  
                            THEN ((1.00 / rcr.rate) * MAX(so.amount_total)) 
                            ELSE so.amount_total
                        END
                    FROM res_company res_co
                    LEFT JOIN res_currency_rate rcr  
                        ON rcr.currency_id = MAX(so.currency_id)  
                        AND rcr.company_id = res_co.id
                    WHERE res_co.id = MAX(so.company_id)
                    ORDER BY rcr.name DESC 
                    LIMIT 1
                ), 0) AS sale_amt,
                TO_CHAR(so.date_order, 'YYYY-MM-DD') AS order_date,
                JSON_AGG(
                    JSONB_BUILD_OBJECT(
                        'ref', sm.reference,
                        'status', sm.state,
                        'date', TO_CHAR(sm.date, 'YYYY-MM-DD'),
                        'damad_code', pt.default_code,
                        'p_name', pt.name,
                        'done_qty', COALESCE(
                            (SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id), 0),
                        'curr', sol_cu.name,
                        'done_amt', 
                            (CASE 
                                WHEN sm.state = 'done' AND sm.sale_line_id IS NOT NULL THEN
                                    CASE 
                                        WHEN sol.currency_id != (SELECT currency_id FROM res_company WHERE id = sm.company_id) 
                                        THEN 
                                            CASE 
                                                WHEN sol.product_uom IS NOT NULL AND sol.product_uom != sm.product_uom
                                                THEN COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id), 0) 
                                                     * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))  
                                                     * (1 / COALESCE(rc.rate, 1))
                                                ELSE COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id), 0)  
                                                     * sol.price_unit 
                                                     * (1 / COALESCE(rc.rate, 1))
                                            END
                                        ELSE 
                                            CASE 
                                                WHEN sol.product_uom IS NOT NULL AND sol.product_uom != sm.product_uom
                                                THEN COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id), 0) 
                                                     * ((sol.product_uom_qty * sol.price_unit) / NULLIF(sm.product_uom_qty, 0))
                                                ELSE COALESCE((SELECT SUM(sml.qty_done) FROM stock_move_line sml WHERE sml.move_id = sm.id), 0) 
                                                     * sol.price_unit
                                            END
                                    END
                                ELSE 0.0 
                            END) * CASE WHEN sm.reference LIKE 'WH/RET%' THEN -1 ELSE 1 END 
                    )
                ) FILTER (WHERE sm.state = 'done') AS delivery_details, 
                (
                    SELECT JSON_AGG(JSONB_BUILD_OBJECT(
                        'inv_name', am.name,
                        'invoice_date', TO_CHAR(am.invoice_date, 'YYYY-MM-DD'),
                        'amount_total_signed', am.amount_total_signed
                    ))
                    FROM account_move am 
                    WHERE am.invoice_origin = so.name 
                    AND am.move_type IN ('out_invoice', 'out_refund') 
                    AND am.state = 'posted'
                ) AS invoice_details
            FROM sale_order so
            LEFT JOIN sale_order_line sol ON sol.order_id = so.id 
            LEFT JOIN stock_move sm ON sol.id = sm.sale_line_id 
            LEFT JOIN product_product pp ON sm.product_id = pp.id
            LEFT JOIN product_template pt ON pt.id = pp.product_tmpl_id
            LEFT JOIN uom_uom ac_uom ON ac_uom.id = sol.product_uom
            LEFT JOIN uom_uom p_uom ON p_uom.id = sm.product_uom
            LEFT JOIN res_currency sol_cu ON sol_cu.id = sol.currency_id
            LEFT JOIN res_currency_rate rc ON rc.currency_id = sol.currency_id
                AND rc.name = (SELECT MAX(name) FROM res_currency_rate WHERE currency_id = sol.currency_id)
            WHERE {domain}
            GROUP BY so.amount_total, so.name, so.date_order
            ORDER BY so.date_order DESC;
         """

        self._cr.execute(query)
        sales_delivery_invoice = self.env.cr.dictfetchall()
        return sales_delivery_invoice

    def prepare_report_delivery_with_invoice(self, data):
        output, workbook, worksheet =  WS.workbook_worksheet()
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9E1F2', 'border': 1})

        headers = ['Sales Reference', 'Order Date', 'Sales Amount', 'Exchanged AMT', 'WH/OUT name' , 'Delivered date', 'Item Code', 'Item Name', 'Done Amt', 'Invoice Reference', 'Invoice date', 'Invoiced Amt']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        row = 1
        for order in data:
            first_row = row
            worksheet.write(row , 0, order['origin'])
            worksheet.write(row , 1, order['order_date'])
            worksheet.write(row, 2, order['amount_total'])
            worksheet.write(row, 3, order['sale_amt'])

            delivery_details = order.get('delivery_details') or []
            invoice_details = order.get('invoice_details') or []

            delivery_count = len(delivery_details)
            invoice_count = len(invoice_details)
            max_rows = max(delivery_count, invoice_count, 1)
            for i, picking in enumerate(delivery_details):
                worksheet.write(row + i, 0, order['origin'])
                worksheet.write(row + i, 1, order['order_date'])
                worksheet.write(row + i, 4, picking['ref'])
                worksheet.write(row + i, 5, picking['date'])
                worksheet.write(row + i, 6, picking['damad_code'])
                worksheet.write(row + i, 7, picking['p_name'])
                worksheet.write(row + i, 8, picking['done_amt'])

            for i, invoice in enumerate(invoice_details):
                worksheet.write(row + i, 0, order['origin'])
                worksheet.write(row + i, 1, order['order_date'])
                worksheet.write(row + i, 9, invoice['inv_name'])
                worksheet.write(row + i, 10, invoice['invoice_date'])
                worksheet.write(row + i, 11, invoice['amount_total_signed'])

            row += max_rows

        workbook.close()
        output.seek(0)
        return output.getvalue()


    def invoice_query(self):
        query = """
                SELECT 
                sm.origin AS orgin,
                sm.reference,
                sm.date,
                pt.default_code,
                pt.name AS product_name,
                sol.product_uom_qty,
                ac_uom.name AS actual_uom,
                p_uom.name AS product_uom
			FROM 
                stock_move sm
            LEFT JOIN
                product_product pp on sm.product_id = pp.id
            LEFT JOIN 
                product_template pt on pt.id = pp.product_tmpl_id
            LEFT JOIN 
                uom_uom ac_uom on ac_uom.id = sol.product_uom
            LEFT JOIN 
                uom_uom p_uom ON p_uom.id = sm.product_uom
			ORDER BY sm.date DESC
        """
        self._cr.execute(query)
        top_product = self._cr.dictfetchall()
        return top_product

    def get_invoice_excel_report(self, report_data):
        data = report_data
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 20)
        worksheet.set_column('D:D', 20)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 20)
        worksheet.set_column('G:G', 20)
        worksheet.set_column('H:H', 20)
        worksheet.set_column('I:I', 20)
        worksheet.set_column('J:J', 20)
        worksheet.set_column('K:K', 20)
        worksheet.set_column('L:L', 20)
        worksheet.set_column('M:M', 20)
        worksheet.set_column('N:N', 20)
        worksheet.set_column('O:O', 20)
        worksheet.set_column('P:P', 20)
        worksheet.set_column('Q:Q', 20)
        worksheet.set_column('R:R', 20)


        headers = [
                   'Order Reference NO.',
                   'Customer',
                   'Region',
                   'WH/OUT NAME',
                   'DELIVERED Date',
                   'DAMAD CODE',
                   'PRODUCT NAME',
                   'ACTUAL QTY',
                   'ACTUAL UOM',
                   'PRODUCT UOM',
                   'UNIT PRICE',
                   'CURRENCY',
                   ]
        bold = workbook.add_format({'bold': True})
        for col, header in enumerate(headers):
            worksheet.write(1, col, header, bold)
        row = 2
        for order in data:
            worksheet.write(row, 0, order['orgin'])
            worksheet.write(row, 3, order['reference'])
            worksheet.write(row, 4, order['date'])
            worksheet.write(row, 5, order['default_code'])
            worksheet.write(row, 6, order['product_name'])
            worksheet.write(row, 7, order['actual_qty'])
            worksheet.write(row, 8, order['actual_uom'])
            worksheet.write(row, 10, order['product_uom'])
            row +=1
        workbook.close()
        output.seek(0)
        return output.getvalue()


class IrUiView(models.Model):
    _inherit = 'ir.ui.view'

    @api.model
    def get_view_name(self, view_id):
        view_name = self.sudo().env['ir.ui.view'].search([('id', '=', view_id)])
        return view_name.xml_id


class MrpBomInherit(models.Model):
    _inherit = "mrp.bom"

    def get_selling_price(self, product_variant_id):
        sale_price = []
        sale_lines = self.env['sale.order.line'].search([
            ('product_id', '=', self.env['product.product'].search([('product_tmpl_id', '=', product_variant_id)]).id)
        ])
        for line in sale_lines:
            price_list = {
                'order_name': line.order_id.name,
                'unit_price': line.price_unit,
                'product_uom': line.product_uom.name,
            }
            sale_price.append(price_list)
        return sale_price

    def prepare_bom_structure_and_cost(self):
        query = """
            SELECT
                bom.id AS bom_id,
                pp.id AS product_id,
                bom.product_qty AS quantity,
                tmpl.default_code AS default_code,
                tmpl.name AS product_name,
                tmpl.list_price AS actual_selling_price,
                uom.name AS uom
            FROM mrp_bom AS bom
            LEFT JOIN product_template  tmpl ON bom.product_tmpl_id = tmpl.id
            LEFT JOIN product_product pp ON pp.product_tmpl_id = tmpl.id
            LEFT JOIN uom_uom uom ON uom.id = bom.product_uom_id
        """
        self.env.cr.execute(query)
        results = self.env.cr.fetchall()
        data = []
        for row in results:
            bom_id, product_id, quantity,  default_code, product_name, actual_selling_price, uom= row
            doc = self.env['report.mrp.report_bom_structure']._get_pdf_line(
                bom_id, product_id=product_id, qty=quantity, unfolded=True
            )
            doc.update({
                'uom': uom,
                'default_code': default_code,
                'p_name': product_name,
                'actual_selling_price': actual_selling_price
            })

            data.append(doc)
        return data

    def get_bom_excel_report(self):
        report_data = self.prepare_bom_structure_and_cost()
        output, workbook, worksheet =  WS.workbook_worksheet()
        headers = ['Item Code', 'Product Name', 'Reference', 'UOM', 'Quantity', 'Product Cost', 'Bom Cost', 'Operation Cost' , 'Total Cost', 'Actual Selling price']
        bold = workbook.add_format({'bold': True})
        for col, header in enumerate(headers):
            worksheet.write(1, col, header, bold)
        row = 2
        for data in report_data:
            worksheet.write(row, 0, data['default_code'])
            worksheet.write(row, 1, data['p_name'])
            worksheet.write(row, 2, data['code'])
            worksheet.write(row, 3, data['uom'])
            worksheet.write(row, 4, data['bom_qty'])
            worksheet.write(row, 5, data['price'])
            worksheet.write(row, 6, data['bom_cost'])
            worksheet.write(row, 7, data['operations_cost'])
            worksheet.write(row, 8, data['total'])
            worksheet.write(row, 9, data['actual_selling_price'])
            # worksheet.write(row, 9, data['selling_price'])
            row +=1

        workbook.close()
        output.seek(0)
        return output.getvalue()
