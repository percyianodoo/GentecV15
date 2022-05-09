"""Wizard Bank Reconciliation Report TransientModel."""

import xlsxwriter
import os
import base64
from datetime import datetime
# from calendar import monthrange
# from dateutil.relativedelta import relativedelta

from odoo import models, fields, api
from odoo.tools import ustr  # , DEFAULT_SERVER_DATE_FORMAT as DF

PAY_TYPE = {'outbound': 'Send Money', 'inbound': 'Receive Money',
            'transfer': 'Internal Transfer'}
PARTNER_TYPE = {'customer': 'Customer', 'supplier': 'Vendor'}


def _offset_format_timestamp2(src_tstamp_str, src_format, dst_format,
                              ignore_unparsable_time=True, context=None):
    if not src_tstamp_str:
        return False
    res = src_tstamp_str
    if src_format and dst_format:
        try:
            dt_value = datetime.strptime(src_tstamp_str, src_format)
            if context.get('tz', False):
                try:
                    import pytz
                    src_tz = pytz.timezone('UTC')
                    dst_tz = pytz.timezone(context['tz'])
                    src_dt = src_tz.localize(dt_value, is_dst=True)
                    dt_value = src_dt.astimezone(dst_tz)
                except Exception:
                    pass
            res = dt_value.strftime(dst_format)
        except Exception:
            if not ignore_unparsable_time:
                return False
            pass
    return res


class WizBankReconciliationReportExported(models.TransientModel):
    """Wizard Bank Reconciliation Report Exported TransientModel."""

    _name = 'wiz.bank.reconciliation.report.exported'
    _description = "Wizard Bank Reconciliation Report Exported"

    file = fields.Binary("Click On Download Link To Download Xlsx File",
                         readonly=True)
    name = fields.Char(string='File Name', size=32)


class WizBankReconciliationReport(models.TransientModel):
    """Wizard Bank Reconciliation Report TransientModel."""

    _name = 'wiz.bank.reconciliation.report'
    _description = "Wizard Bank Reconciliation Report"

    journal_id = fields.Many2one("account.journal",
                                 string="Bank Account/Journal")
    bnk_st_date = fields.Many2one('account.bank.statement',
                                  string="Date")
    company_id = fields.Many2one("res.company", string="Company",
                                 default=lambda self: self.env.user and
                                 self.env.user.company_id)

    @api.onchange('company_id')
    def onchange_company_id(self):
        if self.company_id:
            self.bnk_st_date = False
            self.journal_id = False

    def export_bank_reconciliation_report(self):
        """Method to export bank reconciliation report."""
        cr = self.env.args[0]
        uid = self.env.args[1]
        context = self.env.args[2]
        wiz_exported_obj = self.env['wiz.bank.reconciliation.report.exported']
        move_l_obj = self.env['account.move.line']
        bank_st_obj = self.env['account.bank.statement']
        bank_st_l_obj = self.env['account.bank.statement.line']
        # sheet Development
        file_path = 'Bank Reconciliation Report.xlsx'
        workbook = xlsxwriter.Workbook('/tmp/' + file_path)
        # num_format = workbook.add_format({'num_format': 'dd/mm/yy'})

        header_cell_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': 1,  # 'fg_color': '#96c5f4',
            'align': 'center',
            'border': 1,  # 'valign': 'vcenter'
            'text_wrap': True,
            'bg_color': '#d3d3d3'
        })
        header_cell_l_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': 1,  # 'fg_color': '#96c5f4',
            'align': 'left',
            # 'border': 1,  # 'valign': 'vcenter'
            'text_wrap': True
        })
        header_cell_r_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'bold': 1,  # 'fg_color': '#96c5f4',
            'align': 'right',
            'border': 1,  # 'valign': 'vcenter'
            'text_wrap': True,
            'bg_color': '#d3d3d3'
        })

        cell_l_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'left',  # 'valign': 'vcenter', 'text_wrap': True
            'text_wrap': True
        })

        cell_r_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'right',  # 'valign': 'vcenter'
            'text_wrap': True,
            'num_format': '#,##,###'
        })

        cell_r_bold_noborder = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'right',  # 'valign': 'vcenter'
            'text_wrap': True,
            'bold': 1
        })

        cell_c_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'center',  # 'valign': 'vcenter'
            'text_wrap': True
        })

        cell_c_head_fmat = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 14,
            'align': 'center',
            'bold': True,
            'border': 1,
            'text_wrap': True
        })

        company = self.company_id or False
        company_name = company and company.name or ''
        from_date = datetime.strftime(self.bnk_st_date.sudo().date, '%d/%m/%Y')
        # for journal in self.journal_ids:
        for journal in self.journal_id:
            currency_id = journal.sudo().currency_id or False
            currency_symbol = journal.sudo().currency_id and \
                journal.sudo().currency_id.symbol or \
                journal.sudo().company_id and \
                journal.sudo().company_id.currency_id and \
                journal.sudo().company_id.currency_id.symbol or ''
            currency_position = journal.sudo().currency_id and \
                journal.sudo().currency_id.position or \
                journal.sudo().company_id and \
                journal.sudo().company_id.currency_id and \
                journal.sudo().company_id.currency_id.position or ''

            bank_st_id = bank_st_obj.search([
                ('date', '=', self.bnk_st_date.sudo().date),
                ('journal_id', '=', journal.id),
                ('company_id', '=', company.id)], limit=1)


            last_bank_st_id = bank_st_obj.search([
                ('date', '<', self.bnk_st_date.sudo().date),
                ('journal_id', '=', journal.id),
                ('company_id', '=', company.id)], limit=1)
            last_st_balance = last_bank_st_id.balance_end

            last_reconcile_date_str = ''
            last_reconcile_date = ''
            # last_reconcile_amount = 0.0
            # curr_bal = bank_st_id and bank_st_id.balance_end or 0.0
            last_reconcile_bal = last_bank_st_id and \
                last_bank_st_id.balance_end or 0.0
            reconcile_cust_bnk_st_lines = False
            reconcile_vend_bnk_st_lines = False
            if last_bank_st_id:
                last_reconcile_date = last_bank_st_id.date
                last_reconcile_date_str = \
                    datetime.strftime(last_reconcile_date, '%d/%m/%Y')

                reconcile_cust_bnk_st_lines = bank_st_l_obj.search([
                    ('statement_id', '=', bank_st_id and bank_st_id.id or False),
                    ('statement_id.journal_id', '=', journal.id),
                    ('statement_id.company_id', '=', company.id),
                    ('move_line_ids', '!=', False),
                ])

                reconcile_vend_bnk_st_lines = bank_st_l_obj.search([
                    ('statement_id', '=', bank_st_id and bank_st_id.id or False),
                    ('statement_id.journal_id', '=', journal.id),
                    ('statement_id.company_id', '=', company.id),
                    ('move_line_ids', '!=', False),
                ])

            system_stf_fy_dt = self.bnk_st_date.sudo().date.\
                replace(year=2018, month=7, day=1)
            tot_virtual_gl_bal = 0.0
            account_ids = list(set([
                journal.default_account_id.id,
                journal.default_account_id.id]) - {False})
            lines_already_accounted = move_l_obj.search([
                ('account_id', 'in', account_ids),
                ('date', '<=', self.bnk_st_date.sudo().date),
                ('company_id', '=', company.id)])

            odoo_balance = sum(lines_already_accounted.mapped('balance'))
            # Bank statement lines not reconciled with a payment
            bank_st_positiove_l = bank_st_l_obj.search([
                ('statement_id.journal_id', '=', journal.id),
                ('date', '<=', self.bnk_st_date.sudo().date),
                ('amount', '>', 0),
                ('company_id', '=', company.id)])

            outstanding_plus_tot = sum(bank_st_positiove_l.mapped('amount'))

            bank_st_minus_l = bank_st_l_obj.search([
                ('statement_id.journal_id', '=', journal.id),
                ('date', '<=', self.bnk_st_date.sudo().date),
                ('amount', '<', 0),
                ('company_id', '=', company.id)])

            outstanding_minus_tot = sum(bank_st_minus_l.mapped('amount'))
            unreconcile_checks_payments = move_l_obj.search([
                '|', '&',
                ('move_id.journal_id.type', 'in', ['cash', 'bank']),
                ('move_id.journal_id', '=', journal.id),
                '&',
                ('move_id.journal_id.type', 'not in', ['cash', 'bank']),
                ('move_id.journal_id', '=', journal.id),
                '|',
                ('statement_line_id', '=', False),
                ('statement_line_id.date', '>', self.bnk_st_date.sudo().date),
                ('full_reconcile_id', '=', False),
                ('date', '<=', self.bnk_st_date.sudo().date),
                '&',
                ('company_id', '=', company.id),
                ('date', '>=', system_stf_fy_dt)
            ])
            unrec_tot = sum(unreconcile_checks_payments.mapped('balance'))

            tot_virtual_gl_bal = odoo_balance + outstanding_plus_tot + \
                outstanding_minus_tot + unrec_tot
            difference = tot_virtual_gl_bal - last_st_balance

            worksheet = workbook.add_worksheet(journal.name)
            worksheet.set_column(0, 0, 5)
            worksheet.set_column(1, 1, 13)
            worksheet.set_column(2, 2, 10)
            worksheet.set_column(3, 3, 35)
            worksheet.set_column(4, 4, 35)
            worksheet.set_column(5, 5, 20)
            worksheet.set_column(6, 6, 15)
            worksheet.set_row(1, 20)
            worksheet.merge_range(
                1, 0, 1, 5, company_name, cell_c_head_fmat)
            worksheet.merge_range(
                2, 0, 2, 5,
                'Reconciliation Details - ' + journal.name, cell_c_head_fmat)
            worksheet.merge_range(
                3, 0, 3, 5,
                'As of ' + ustr(from_date),
                cell_c_head_fmat)
            row = 5
            col = 0
            worksheet.write(row, col, 'ID', header_cell_fmat)
            col += 1
            worksheet.write(row, col, 'Transaction Type', header_cell_fmat)
            col += 1
            worksheet.write(row, col, 'Date', header_cell_fmat)
            col += 1
            worksheet.write(row, col, 'Customer/Partner Name',
                            header_cell_fmat)
            col += 1
            worksheet.write(row, col, 'Lable/Memo', header_cell_fmat)
            col += 1
            worksheet.write(row, col, 'Balance', header_cell_r_fmat)
            row += 1
            worksheet.merge_range(row, 0, row, 1, 'Reconciled',
                                  header_cell_l_fmat)
            row += 1
            worksheet.merge_range(row, 1, row, 4,
                                  'Cleared Deposits and Other Credits',
                                  header_cell_l_fmat)
            col = 0
            row += 1
            tot_cust_payment = 0.0
            if reconcile_cust_bnk_st_lines:
                for cust_pay_line in reconcile_cust_bnk_st_lines:
                    account_ids = list(set([
                        journal.default_account_id.id,
                        journal.default_account_id.id]) - {False})
                    move_lines = move_l_obj.search([
                        ('statement_line_id', '=', cust_pay_line.id),
                        ('payment_id', '!=', False),
                        ('payment_id.payment_type', 'in', ['inbound', 'transfer']),
                        ('balance', '>=', 0.0),
                        ('account_id', 'in', account_ids)
                    ])
                    for move_l in move_lines:
                        balance = move_l and move_l.balance or 0.0
                        if currency_id:
                            balance = move_l and move_l.amount_currency or 0.0
                            if balance in [0.0, -0.0]:
                                balance = move_l and move_l.balance or 0.0

                        tot_cust_payment = tot_cust_payment + balance or 0.0
                        payment_date = ''
                        if move_l and move_l.date:
                            payment_date = datetime.strftime(
                                move_l.date, '%d-%m-%Y')

                        payment = move_l and move_l.payment_id or False
                        name = payment and payment.partner_id and \
                            payment.partner_id.name or ''
                        pay_no = payment and payment.name or ''
                        pay_reference = payment and payment.payment_reference or ''
                        pay_memo = payment and payment.communication or ''
                        batch_pay_no = payment and payment.batch_payment_id and \
                            payment.batch_payment_id.name or ''
                        pay_method = payment and payment.batch_payment_id and \
                            payment.batch_payment_id.payment_method_id and \
                            payment.batch_payment_id.payment_method_id.name or ''
                        jou_entry_ref = move_l and move_l.move_id and \
                            move_l.move_id.name or ''
                        partner = "Partner Name : " + name + '\n'
                        partner += "Batch Payment Number : " + batch_pay_no + '\n'
                        partner += "Payment Method : " + pay_method + '\n'
                        partner += "Payment Number : " + pay_no + '\n'
                        partner += "Payment Reference : " + pay_reference + '\n'
                        partner += "Journal Entry Number : " + jou_entry_ref + '\n'
                        partner += "Invoice Reference : " + pay_memo + '\n'

                        cust_pay_memo = cust_pay_line.name or ''

                        worksheet.write(row, col, ' ', cell_c_fmat)
                        col += 1
                        worksheet.write(row, col, 'Payment', cell_c_fmat)
                        col += 1
                        worksheet.write(row, col, payment_date, cell_c_fmat)
                        col += 1
                        worksheet.set_row(row, 90)
                        worksheet.write(row, col, partner, cell_l_fmat)
                        col += 1
                        worksheet.write(row, col, cust_pay_memo or '', cell_l_fmat)
                        col += 1
                        bal_str = balance
                        if currency_position == 'after':
                            bal_str = ustr(bal_str) + ustr(currency_symbol)
                        else:
                            bal_str = ustr(currency_symbol) + ustr(bal_str)
                        worksheet.write(row, col, bal_str or ustr(0.0),
                                        cell_r_fmat)
                        col = 0
                        row += 1

            worksheet.set_row(row, 40)
            row += 1
            worksheet.set_row(row, 40)

            tot_cust_payment_str = round(tot_cust_payment, 2)
            if currency_position == 'after':
                tot_cust_payment_str = ustr(
                    tot_cust_payment_str) + ustr(currency_symbol)
            else:
                tot_cust_payment_str = ustr(
                    currency_symbol) + ustr(tot_cust_payment_str)

            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Cleared Deposits and Other Credits',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, tot_cust_payment_str or 0.0,
                            cell_r_bold_noborder)
            row += 1
            worksheet.set_row(row, 40)
            worksheet.merge_range(row, 1, row, 4,
                                  'Cleared Checks and Payments',
                                  header_cell_l_fmat)

            col = 0
            row += 1
            tot_vend_payment = 0.0
            if reconcile_vend_bnk_st_lines:
                for vend_pay_line in reconcile_vend_bnk_st_lines:
                    account_ids = list(set([
                        journal.default_account_id.id,
                        journal.default_account_id.id]) - {False})
                    move_lines = move_l_obj.search([
                        ('statement_line_id', '=', vend_pay_line.id),
                        ('payment_id', '!=', False),
                        ('payment_id.payment_type', 'in',
                            ['outbound', 'transfer']),
                        ('balance', '<=', 0.0),
                        ('account_id', 'in', account_ids)
                    ])
                    for move_l in move_lines:
                        balance = move_l and move_l.balance or 0.0
                        if currency_id:
                            balance = move_l and move_l.amount_currency or 0.0
                            if balance in [0.0, -0.0]:
                                balance = move_l and move_l.balance or 0.0

                        tot_vend_payment = tot_vend_payment + balance or 0.0
                        payment_date = ''
                        if move_l and move_l.date:
                            payment_date = datetime.strftime(
                                move_l.date, '%d-%m-%Y')
                        payment = move_l and move_l.payment_id or False
                        name = payment and payment.partner_id and \
                            payment.partner_id.name or ''
                        pay_no = payment and payment.name or ''
                        pay_reference = payment and payment.payment_reference or ''
                        pay_memo = payment and payment.communication or ''
                        batch_pay_no = payment and payment.batch_payment_id and \
                            payment.batch_payment_id.name or ''
                        pay_method = payment and payment.batch_payment_id and \
                            payment.batch_payment_id.payment_method_id and \
                            payment.batch_payment_id.payment_method_id.name or ''
                        jou_entry_ref = move_l and move_l.move_id and \
                            move_l.move_id.name or ''
                        partner = "Partner Name : " + name + '\n'
                        partner += "Batch Payment Number : " + batch_pay_no + '\n'
                        partner += "Payment Method : " + pay_method + '\n'
                        partner += "Payment Number : " + pay_no + '\n'
                        partner += "Payment Reference : " + pay_reference + '\n'
                        partner += "Journal Entry Number : " + jou_entry_ref + '\n'
                        partner += "Invoice Reference : " + pay_memo + '\n'

                        vend_pay_memo = move_l.name or ''

                        worksheet.write(row, col, ' ', cell_c_fmat)
                        col += 1
                        worksheet.write(row, col, 'Bill Payment', cell_c_fmat)
                        col += 1
                        worksheet.write(row, col, payment_date, cell_c_fmat)
                        col += 1
                        worksheet.set_row(row, 90)
                        worksheet.write(row, col, partner, cell_l_fmat)
                        col += 1
                        worksheet.write(row, col, vend_pay_memo or '', cell_l_fmat)
                        col += 1
                        bal_str = balance
                        if currency_position == 'after':
                            bal_str = ustr(bal_str) + ustr(currency_symbol)
                        else:
                            bal_str = ustr(currency_symbol) + ustr(bal_str)
                        worksheet.write(row, col, bal_str or ustr(0.0),
                                        cell_r_fmat)
                        col = 0
                        row += 1

            row += 1

            tot_vend_pay_str = round(tot_vend_payment, 2)
            if currency_position == 'after':
                tot_vend_pay_str = ustr(tot_vend_pay_str) + \
                    ustr(currency_symbol)
            else:
                tot_vend_pay_str = ustr(currency_symbol) + \
                    ustr(tot_vend_pay_str)

            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Cleared Checks and Payments',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, tot_vend_pay_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Total - Reconciled', header_cell_l_fmat)

            filter_bal = tot_cust_payment + tot_vend_payment

            filter_bal_str = round(filter_bal, 2)
            if currency_position == 'after':
                filter_bal_str = ustr(filter_bal_str) + ustr(currency_symbol)
            else:
                filter_bal_str = ustr(currency_symbol) + ustr(filter_bal_str)

            worksheet.write(row, 5, filter_bal_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(
                row, 0, row, 3,
                'Last Reconciled Statement Balance - ' +
                ustr(last_reconcile_date_str),
                header_cell_l_fmat)

            last_recon_bal_str = round(last_reconcile_bal, 2)
            if currency_position == 'after':
                last_recon_bal_str = ustr(
                    last_recon_bal_str) + ustr(currency_symbol)
            else:
                last_recon_bal_str = ustr(
                    currency_symbol) + ustr(last_recon_bal_str)

            worksheet.write(row, 5, last_recon_bal_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Current Reconciled Balance',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, filter_bal_str, cell_r_bold_noborder)
            row += 1

            worksheet.merge_range(
                row, 0, row, 3,
                'Reconcile Statement Balance - ' + ustr(from_date),
                header_cell_l_fmat)
            re_st_bal_tot = filter_bal + last_reconcile_bal

            re_st_bal_tot_str = round(re_st_bal_tot, 2)
            unrec_tot_str = round(unrec_tot, 2)
            difference_str = round(0.0, 2)
            if currency_position == 'after':
                re_st_bal_tot_str = ustr(re_st_bal_tot_str) + \
                    ustr(currency_symbol)
                unrec_tot_str = ustr(unrec_tot_str) + ustr(currency_symbol)
                difference_str = ustr(difference_str) + ustr(currency_symbol)
            else:
                re_st_bal_tot_str = ustr(currency_symbol) + \
                    ustr(re_st_bal_tot_str)
                unrec_tot_str = ustr(currency_symbol) + ustr(unrec_tot_str)
                difference_str = ustr(currency_symbol) + ustr(difference_str)

            worksheet.write(row, 5, re_st_bal_tot_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(
                row, 0, row, 3, 'Difference', header_cell_l_fmat)
            worksheet.write(row, 5, difference_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3, 'Unreconciled',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, unrec_tot_str, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Uncleared  Checks and Payments',
                                  header_cell_l_fmat)

            col = 0
            row += 1
            tot_unreconcile_cust_payment = 0.0
            for cust_unrecon_l in unreconcile_checks_payments:
                trns_type = 'Payment'
                if cust_unrecon_l.payment_id and \
                        cust_unrecon_l.payment_id.payment_type:
                    if cust_unrecon_l.payment_id.payment_type in \
                            ['outbound', 'transfer']:
                        trns_type = 'Bill Payment'

                cust_balance = cust_unrecon_l and cust_unrecon_l.balance or 0.0
                if currency_id:
                    cust_balance = cust_unrecon_l and \
                        cust_unrecon_l.amount_currency or 0.0
                    if cust_balance in [0.0, -0.0]:
                        cust_balance = cust_unrecon_l and \
                            cust_unrecon_l.balance or 0.0

                tot_unreconcile_cust_payment = tot_unreconcile_cust_payment + \
                    cust_balance or 0.0

                payment_date = ''
                if cust_unrecon_l.date:
                    payment_date = datetime.strftime(
                        cust_unrecon_l.date, '%d-%m-%Y')

                partner = cust_unrecon_l.partner_id and \
                    cust_unrecon_l.partner_id.name or ''
                cust_unrecon_pay_memo = cust_unrecon_l.name or ''

                worksheet.write(row, col, ' ', cell_c_fmat)
                col += 1
                worksheet.write(row, col, trns_type, cell_c_fmat)
                col += 1
                worksheet.write(row, col, payment_date, cell_c_fmat)
                col += 1
                worksheet.set_row(row, 40)
                worksheet.write(row, col, partner, cell_l_fmat)
                col += 1
                worksheet.write(row, col, cust_unrecon_pay_memo or '',
                                cell_l_fmat)
                col += 1

                cust_balance = cust_unrecon_l and cust_unrecon_l.balance or 0.0
                if currency_id:
                    cust_balance = cust_unrecon_l and \
                        cust_unrecon_l.amount_currency or 0.0
                    if cust_balance in [0.0, -0.0]:
                        cust_balance = cust_unrecon_l and \
                            cust_unrecon_l.balance or 0.0

                cust_bal_str = round(cust_balance, 2)
                if currency_position == 'after':
                    cust_bal_str = ustr(cust_bal_str) + ustr(currency_symbol)
                else:
                    cust_bal_str = ustr(currency_symbol) + ustr(cust_bal_str)

                worksheet.write(row, col, cust_bal_str, cell_r_fmat)
                col = 0
                row += 1
            row += 1

            tot_unrec_cust_pay_str = round(tot_unreconcile_cust_payment, 2)
            if currency_position == 'after':
                tot_unrec_cust_pay_str = ustr(
                    tot_unrec_cust_pay_str) + ustr(currency_symbol)
            else:
                tot_unrec_cust_pay_str = ustr(
                    currency_symbol) + ustr(tot_unrec_cust_pay_str)

            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Uncleared Checks and Payments',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, tot_unrec_cust_pay_str,
                            cell_r_bold_noborder)
            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Unreconciled',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, tot_unrec_cust_pay_str,
                            cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Total as of ' + ustr(from_date),
                                  header_cell_l_fmat)
            worksheet.write(row, 5, re_st_bal_tot_str,
                            cell_r_bold_noborder)

        workbook.close()
        buf = base64.encodestring(open('/tmp/' + file_path, 'rb').read())
        try:
            if buf:
                os.remove(file_path + '.xlsx')
        except OSError:
            pass
        wiz_rec = wiz_exported_obj.create({
            'file': buf,
            'name': 'Bank Reconciliation Report.xlsx'
        })
        form_view = self.env.ref(
            'account_reports_extended.wiz_bank_reconcil_rep_exported_form')
        if wiz_rec and form_view:
            return {
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_id': wiz_rec.id,
                'res_model': 'wiz.bank.reconciliation.report.exported',
                'views': [(form_view.id, 'form')],
                'view_id': form_view.id,
                'target': 'new',
            }
        else:
            return {}
