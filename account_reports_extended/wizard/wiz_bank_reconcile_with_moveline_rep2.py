"""Wizard Bank Reconcilition Report TransientModel."""

import xlsxwriter
import os
import base64
from datetime import datetime
from calendar import monthrange
# from dateutil.relativedelta import relativedelta

from odoo import models, fields, api
from odoo.tools import ustr, DEFAULT_SERVER_DATE_FORMAT as DF

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
    """Wizard Bank Reconcilition Report Exported TransientModel."""

    _name = 'wiz.bank.reconciliation.report.exported'
    _description = "Wizard Bank Reconcilition Report Exported"

    file = fields.Binary("Click On Download Link To Download Xlsx File",
                         readonly=True)
    name = fields.Char(string='File Name', size=32)


class WizBankReconciliationReport(models.TransientModel):
    """Wizard Bank Reconcilition Report TransientModel."""

    _name = 'wiz.bank.reconciliation.report'
    _description = "Wizard Bank Reconcilition Report"

    @api.model
    def default_get(self, fields=[]):
        """Method to update start and end date."""
        res = super(WizBankReconciliationReport, self).default_get(fields)
        curr_dt = datetime.today()
        tot_days = monthrange(curr_dt.year, curr_dt.month)[1]
        st_dt = datetime.today().replace(day=1).date()
        end_dt = datetime.today().replace(day=int(tot_days)).date()
        res.update({'date_from': st_dt, 'date_to': end_dt})
        return res

    date_from = fields.Date(string='Start Date')
    date_to = fields.Date(string='End Date')
    company_id = fields.Many2one("res.company", string="Company",
                                 default=lambda self: self.env.user and
                                 self.env.user.company_id)
    journal_ids = fields.Many2many("account.journal",
                                   "wiz_bank_recon_journal_rel",
                                   "wiz_bankrecon_id", "journal_id",
                                   string="Bank Accounts")

    @api.onchange('company_id')
    def onchange_company_id(self):
        """Onchange company to set journals."""
        journal_obj = self.env['account.journal']
        self.journal_ids = [(6, 0, [])]
        if self.company_id:
            journals = journal_obj.search([
                ('type', '=', 'bank'),
                ('company_id', '=', self.company_id.id)])
            self.journal_ids = [(6, 0, journals.ids)]

    def export_bank_reconciliation_report(self):
        """Method to export bank reconciliation report."""
        cr, uid, context = self.env.args
        wiz_exported_obj = self.env['wiz.bank.reconciliation.report.exported']
        move_l_obj = self.env['account.move.line']
        bank_st_obj = self.env['account.bank.statement']
        bank_st_l_obj = self.env['account.bank.statement.line']
        # sheet Development
        file_path = 'Bank Reconcilition Report.xlsx'
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
            'text_wrap': True
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
        to_date = ''
        from_date = ''
        prev_year_from_date = self.date_from
        prev_year_to_date = self.date_to
        if self.date_from:
            from_date = datetime.strftime(self.date_from, '%d/%m/%Y')
            from_dt = self.date_from
            from_year = self.date_from.year
            f_dt = self.date_from
            prev_year_from_date = \
                from_dt.replace(day=1, month=1, year=from_year - 1)
            prev_year_to_date = f_dt.replace(day=31,
                                             month=12, year=from_year - 1)

        if self.date_to:
            to_date = datetime.strftime(self.date_to, '%d/%m/%Y')
            f_dt = self.date_to

        company = self.company_id or False
        company_name = company and company.name or ''
        for journal in self.journal_ids:

            pre_reconcile_cust_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', prev_year_from_date),
                ('date', '<=', prev_year_to_date),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '!=', False),
                ('amount', '>', 0.0)
            ])
            tot_pre_reconcile_cust_lines = \
                sum(pre_reconcile_cust_bnk_st_lines.mapped('amount'))

            pre_reconcile_vend_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', prev_year_from_date),
                ('date', '<=', prev_year_to_date),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '!=', False),
                ('amount', '<', 0.0)
            ])
            tot_pre_reconcile_vend_lines = \
                sum(pre_reconcile_vend_bnk_st_lines.mapped('amount'))

            pre_unreconcile_cust_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', prev_year_from_date),
                ('date', '<=', prev_year_to_date),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '=', False),
                ('amount', '>', 0.0)
            ])
            tot_pre_unreconcile_cust_lines = \
                sum(pre_unreconcile_cust_bnk_st_lines.mapped('amount'))

            pre_unreconcile_vend_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', prev_year_from_date),
                ('date', '<=', prev_year_to_date),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '=', False),
                ('amount', '<', 0.0)
            ])
            tot_pre_unreconcile_vend_lines = \
                sum(pre_unreconcile_vend_bnk_st_lines.mapped('amount'))

            # --------------------------------------------------------------

            reconcile_cust_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', self.date_from),
                ('date', '<=', self.date_to),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '!=', False),
                ('amount', '>', 0.0)
            ])
            tot_reconcile_cust_lines = \
                sum(reconcile_cust_bnk_st_lines.mapped('amount'))

            reconcile_vend_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', self.date_from),
                ('date', '<=', self.date_to),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '!=', False),
                ('amount', '<', 0.0)
            ])
            tot_reconcile_vend_lines = \
                sum(reconcile_vend_bnk_st_lines.mapped('amount'))

            unreconcile_cust_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', self.date_from),
                ('date', '<=', self.date_to),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '=', False),
                ('amount', '>', 0.0)
            ])
            tot_unreconcile_cust_lines = \
                sum(unreconcile_cust_bnk_st_lines.mapped('amount'))

            unreconcile_vend_bnk_st_lines = bank_st_l_obj.search([
                ('date', '>=', self.date_from),
                ('date', '<=', self.date_to),
                ('statement_id.journal_id', '=', journal.id),
                ('statement_id.company_id', '=', company.id),
                ('journal_entry_ids', '=', False),
                ('amount', '<', 0.0)
            ])
            tot_unreconcile_vend_lines = \
                sum(unreconcile_vend_bnk_st_lines.mapped('amount'))

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
                'As of ' + ustr(from_date) + ' To ' + ustr(to_date),
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
            for cust_pay_line in reconcile_cust_bnk_st_lines:
                tot_cust_payment = tot_cust_payment + \
                    cust_pay_line.amount or 0.0
                payment_date = ''
                if cust_pay_line.date:
                    payment_date = \
                        datetime.strftime(cust_pay_line.date, '%d-%m-%Y')

                cust_pay_name = cust_pay_line.name or ''
                partner = cust_pay_line.partner_id and \
                    cust_pay_line.partner_id.name or ''
                cust_pay_memo = cust_pay_line.name or ''

                worksheet.write(row, col, ' ', cell_c_fmat)
                col += 1
                worksheet.write(row, col, 'Payment', cell_c_fmat)
                col += 1
                worksheet.write(row, col, payment_date, cell_c_fmat)
                col += 1
                worksheet.write(row, col, partner, cell_l_fmat)
                col += 1
                worksheet.write(row, col, cust_pay_memo or '', cell_l_fmat)
                col += 1
                worksheet.write(row, col, cust_pay_line.amount or 0.0,
                                cell_r_fmat)
                col = 0
                row += 1

            row += 1
            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Cleared Deposits and Other Credits',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, tot_cust_payment or 0.0,
                            cell_r_bold_noborder)

            row += 1
            worksheet.merge_range(row, 1, row, 4,
                                  'Cleared Checks and Payments',
                                  header_cell_l_fmat)

            col = 0
            row += 1
            tot_vend_payment = 0.0
            for vend_pay_line in reconcile_vend_bnk_st_lines:
                tot_vend_payment = tot_vend_payment + vend_pay_line.amount
                payment_date = ''
                if vend_pay_line.date:
                    payment_date = \
                        datetime.strftime(vend_pay_line.date, '%d-%m-%Y')

                vend_pay_name = vend_pay_line.name or ''
                partner = vend_pay_line.partner_id and \
                    vend_pay_line.partner_id.name or ''
                vend_pay_memo = vend_pay_line.name or ''

                worksheet.write(row, col, ' ', cell_c_fmat)
                col += 1
                worksheet.write(row, col, 'Bill Payment', cell_c_fmat)
                col += 1
                worksheet.write(row, col, payment_date, cell_c_fmat)
                col += 1
                worksheet.write(row, col, partner, cell_l_fmat)
                col += 1
                worksheet.write(row, col, vend_pay_memo or '', cell_l_fmat)
                col += 1
                worksheet.write(row, col, vend_pay_line.amount or '',
                                cell_r_fmat)
                col = 0
                row += 1

            row += 1
            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Cleared Checks and Payments',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, round(tot_vend_payment, 2) or 0.0,
                            cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Total - Reconciled', header_cell_l_fmat)
            filter_bal = tot_cust_payment + tot_vend_payment
            worksheet.write(row, 5, round(filter_bal, 2) or 0.0,
                            cell_r_bold_noborder)
            row += 1
            prev_year_to_date_str = datetime.strftime(
                prev_year_to_date, '%d/%m/%Y')
            worksheet.merge_range(
                row, 0, row, 3,
                'Last Reconciled Statement Balance - ' +
                ustr(prev_year_to_date_str),
                header_cell_l_fmat)
            prev_bal = tot_pre_reconcile_vend_lines + \
                tot_pre_reconcile_cust_lines

            worksheet.write(row, 5, round(prev_bal, 2), cell_r_bold_noborder)
            row += 1
            curr_bal = filter_bal + prev_bal
            worksheet.merge_range(row, 0, row, 3,
                                  'Current Reconciled Balance',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, round(curr_bal, 2)
                            or 0.0, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(
                row, 0, row, 3,
                'Reconcile Statement Balance - ' + ustr(to_date),
                header_cell_l_fmat)
            worksheet.write(row, 5, round(curr_bal, 2), cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(
                row, 0, row, 3, 'Difference', header_cell_l_fmat)
            worksheet.write(row, 5, 0.0, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3, 'Unreconciled',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, 0.0, cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Uncleared  Checks and Payments',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, 0.0, cell_r_bold_noborder)

            col = 0
            row += 1
            tot_unreconcile_cust_payment = 0.0
            for cust_unrecon_l in unreconcile_cust_bnk_st_lines:
                tot_unreconcile_cust_payment = tot_unreconcile_cust_payment + \
                    cust_unrecon_l.amount or 0.0
                payment_date = ''
                if cust_unrecon_l.date:
                    payment_date = \
                        datetime.strftime(cust_unrecon_l.date, '%d-%m-%Y')

                cust_unrecon_pay_name = cust_unrecon_l.name or ''
                partner = cust_unrecon_l.partner_id and \
                    cust_unrecon_l.partner_id.name or ''
                cust_unrecon_pay_memo = cust_unrecon_l.name or ''

                worksheet.write(row, col, ' ', cell_c_fmat)
                col += 1
                worksheet.write(row, col, 'Payment', cell_c_fmat)
                col += 1
                worksheet.write(row, col, payment_date, cell_c_fmat)
                col += 1
                worksheet.write(row, col, partner, cell_l_fmat)
                col += 1
                worksheet.write(row, col, cust_unrecon_pay_memo or '',
                                cell_l_fmat)
                col += 1
                worksheet.write(row, col, cust_unrecon_l.amount or 0.0,
                                cell_r_fmat)
                col = 0
                row += 1
            row += 1
            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Uncleared Checks and Payments',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, round(tot_unreconcile_cust_payment, 2) or 0.0,
                            cell_r_bold_noborder)
            worksheet.merge_range(row, 1, row, 4,
                                  'Total - Unreconciled',
                                  header_cell_l_fmat)
            worksheet.write(row, 5, round(tot_unreconcile_cust_payment, 2) or 0.0,
                            cell_r_bold_noborder)
            row += 1
            worksheet.merge_range(row, 0, row, 3,
                                  'Total as of ' + ustr(to_date),
                                  header_cell_l_fmat)
            worksheet.write(row, 5, 0.0, cell_r_bold_noborder)

        workbook.close()
        buf = base64.encodestring(open('/tmp/' + file_path, 'rb').read())
        try:
            if buf:
                os.remove(file_path + '.xlsx')
        except OSError:
            pass
        wiz_rec = wiz_exported_obj.create({
            'file': buf,
            'name': 'Bank Reconcilition Report.xlsx'
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
