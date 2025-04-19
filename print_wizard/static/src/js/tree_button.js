odoo.define('print_wizard.DownloadReportListController', function (require) {
    "use strict";

    var ListController = require('web.ListController');
    var rpc = require('web.rpc');
    var session = require('web.session');

    var DownloadReportListController = ListController.include({
        renderButtons: function ($node) {
            this._super.apply(this, arguments);

            var self = this;
            var models = [];

            var xmlPromise = this._rpc({
                model: "ir.ui.view",
                method: 'get_view_name',
                args: [this.viewId],
            }).then(function (result) {
                return result;
            });

            var poReportsPromise = session.user_has_group('print_wizard.group_sales_reports');
            var invReportsPromise = session.user_has_group('print_wizard.group_sales_with_inv_report');
            var stockReportsPromise = session.user_has_group('print_wizard.group_stock_report_damad');
            var bomReportsPromise = session.user_has_group('print_wizard.group_bom_report_damad');
            var poReportsPromise1 = session.user_has_group('print_wizard.group_sales_po_report');

            Promise.all([xmlPromise, poReportsPromise, invReportsPromise, stockReportsPromise, bomReportsPromise, poReportsPromise1]).then(function (values) {
                var xml_id = values[0];
                var hasPoReports = values[1] || values[5];
                var hasInvReports = values[2];
                var hasStockReports = values[3];
                var hasbomReports = values[4];

                if (hasPoReports && xml_id === 'sale.view_quotation_tree_with_onboarding') {
                    console.log("Adding sale.order", xml_id);
                    models.push('sale.order');
                }
                if (hasInvReports && xml_id === 'account.view_out_invoice_tree') {
                    console.log("Adding account.move", xml_id);
                    models.push('account.move');
                }
                if (hasStockReports && xml_id === 'stock_account.stock_valuation_layer_tree') {
                    console.log("Adding stock.valuation.layer", xml_id);
                    models.push('stock.valuation.layer');
                }
                if (hasbomReports && xml_id === 'mrp.mrp_bom_tree_view') {
                    models.push('mrp.bom');
                }
                if (!models.includes(self.modelName)) {
                    self.$buttons.find('.o_button_download_report').hide();
                }
            });

            if (this.$buttons) {
                this.$buttons.on('click', '.o_button_download_report', this._download_report.bind(this));
            }
        },

        _download_report: function (ev) {
            console.log("Downloading report for Model:", this.modelName);

            var context = {
                'active_model': this.modelName,
            };

            return this.do_action({
                name: "Print Report Wizard",
                type: "ir.actions.act_window",
                res_model: "print_wizard",
                view_mode: "form",
                views: [[false, "form"]],
                target: "new",
                context: context,
            });
        }
    });

    return DownloadReportListController;
});
