# -*- coding: utf-8 -*-
import copy
import ast
import io
try:
    from odoo.tools.misc import xlsxwriter
except ImportError:
    import xlsxwriter

from odoo import models, fields, api, _
from odoo.tools.misc import format_date, formatLang
from datetime import datetime, timedelta
from odoo.addons.web.controllers.main import clean_action
from odoo.tools import float_is_zero
from odoo.tools.safe_eval import safe_eval
from datetime import datetime
import datetime
from odoo.osv import expression
from dateutil.relativedelta import relativedelta
from odoo.exceptions import UserError, ValidationError
import pytz
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT as DTF


class report_account_general_ledger(models.AbstractModel):
    _inherit = "account.general.ledger"
    _description = "General Ledger Report"

    def _group_by_account_id(self, options, line_id):
        accounts = {}
        results = self._do_query_group_by_account(options, line_id)
        initial_bal_date_to = fields.Date.from_string(
            self.env.context['date_from_aml']) + timedelta(days=-1)
        initial_bal_results = self.with_context(
            date_to=initial_bal_date_to.strftime(
                '%Y-%m-%d'))._do_query_group_by_account(options, line_id)

        context = self.env.context

        last_day_previous_fy = self.env.user.company_id.\
            compute_fiscalyear_dates(fields.Date.from_string(
                self.env.context['date_from_aml']))['date_from'] + \
            timedelta(days=-1)
        unaffected_earnings_per_company = {}
        for cid in context.get('company_ids', []):
            company = self.env['res.company'].browse(cid)
            unaffected_earnings_per_company[company] = \
                self.with_context(date_to=last_day_previous_fy.strftime(
                    '%Y-%m-%d'), date_from=False).\
                _do_query_unaffected_earnings(options, line_id, company)

        unaff_earnings_treated_companies = set()
        unaffected_earnings_type = self.env.ref(
            'account.data_unaffected_earnings')
        for account_id, result in results.items():
            account = self.env['account.account'].browse(account_id)
            accounts[account] = result
            accounts[account]['initial_bal'] = initial_bal_results.get(
                account.id,
                {'balance': 0, 'amount_currency': 0, 'debit': 0, 'credit': 0})
            if account.user_type_id == unaffected_earnings_type and \
                    account.company_id not in unaff_earnings_treated_companies:
                # add the benefit/loss of previous fiscal year to unaffected
                # earnings accounts
                unaffected_earnings_results = unaffected_earnings_per_company[
                    account.company_id]
                for field in ['balance', 'debit', 'credit']:
                    accounts[account]['initial_bal'][
                        field] += unaffected_earnings_results[field]
                    accounts[account][
                        field] += unaffected_earnings_results[field]
                unaff_earnings_treated_companies.add(account.company_id)
            # use query_get + with statement instead of a search in order to
            # work in cash basis too
            aml_ctx = {}
            if context.get('date_from_aml'):
                aml_ctx = {
                    'strict_range': True,
                    'date_from': context['date_from_aml'],
                }
            aml_ids = self.with_context(**aml_ctx).\
                _do_query(options, account_id, group_by_account=False)
            aml_ids = [x[0] for x in aml_ids]

            accounts[account]['total_lines'] = len(aml_ids)
            offset = int(options.get('lines_offset', 0))
            # We Removed Below three lines to Load All Lines At once.
            # if self.MAX_LINES:
            #     stop = offset + self.MAX_LINES
            # else:
            stop = None
            aml_ids = aml_ids[offset:stop]

            accounts[account]['lines'] = self.env[
                'account.move.line'].browse(aml_ids)

        # For each company, if the unaffected earnings account wasn't in the
        # selection yet: add it manually
        user_currency = self.env.user.company_id.currency_id
        for cid in context.get('company_ids', []):
            company = self.env['res.company'].browse(cid)
            if company not in unaff_earnings_treated_companies and \
                    not float_is_zero(
                        unaffected_earnings_per_company[company]['balance'],
                        precision_digits=user_currency.decimal_places):
                unaffected_earnings_account = \
                    self.env['account.account'].search([
                        ('user_type_id', '=', unaffected_earnings_type.id),
                        ('company_id', '=', company.id)
                    ], limit=1)
                if unaffected_earnings_account and \
                        (not line_id or
                         unaffected_earnings_account.id == line_id):
                    accounts[unaffected_earnings_account[0]
                             ] = unaffected_earnings_per_company[company]
                    accounts[unaffected_earnings_account[0]][
                        'initial_bal'] = \
                        unaffected_earnings_per_company[company]
                    accounts[unaffected_earnings_account[0]]['lines'] = []
                    accounts[unaffected_earnings_account[0]]['total_lines'] = 0
        return accounts


class AccountReport(models.AbstractModel):
    _inherit = 'account.report'

    def get_html(self, options, line_id=None, additional_context=None):
        '''
        return the html value of report, or html value of unfolded line
        * if line_id is set, the template used will be the line_template
        otherwise it uses the main_template. Reason is for efficiency, when unfolding a line in the report
        we don't want to reload all lines, just get the one we unfolded.
        '''

        # Prevent inconsistency between options and context.
        self = self.with_context(self._set_context(options))

        templates = self._get_templates()
        report_manager = self._get_report_manager(options)
        report = {'name': self._get_report_name(),
                  'summary': report_manager.summary,
                  'company_name': self.env.user.company_id.name, }
        render_values = self._get_html_render_values(options, report_manager)
        if additional_context:
            render_values.update(additional_context)

        is_profit = False
        profit_losse_action = self.env.ref(
            'account_reports.account_financial_report_profitandloss0')
        if profit_losse_action:
            is_profit = True
            options['is_profit'] = is_profit
            lines = self._get_lines(options, line_id=line_id)
        else:
            lines = self._get_lines(options, line_id=line_id)
        if options.get('hierarchy'):
            lines = self._create_hierarchy(lines)
        footnotes_to_render = []
        if self.env.context.get('print_mode', False):
            # we are in print mode, so compute footnote number and include them in lines values, otherwise, let the js compute the number correctly as
            # we don't know all the visible lines.
            footnotes = dict([(str(f.line), f)
                              for f in report_manager.footnotes_ids])
            number = 0
            for line in lines:
                f = footnotes.get(str(line.get('id')))
                if f:
                    number += 1
                    line['footnote'] = str(number)
                    footnotes_to_render.append(
                        {'id': f.id, 'number': number, 'text': f.text})
        ctx_rec = dict(self.env.context)
        ctx_rec.update({
            'is_profit': is_profit
        })
        rcontext = {'report': report,
                    'lines': {'columns_header': self.with_context(ctx_rec).get_header(options), 'lines': lines},
                    'options': options,
                    'context': self.env.context,
                    'model': self,
                    }

        if profit_losse_action:
            rcontext.update({
                'is_total': True
            })
        if additional_context and type(additional_context) == dict:
            rcontext.update(additional_context)
        if self.env.context.get('analytic_account_ids'):
            rcontext['options']['analytic_account_ids'] = [
                {'id': acc.id, 'name': acc.name} for acc in self.env.context['analytic_account_ids']
            ]
        render_template = templates.get(
            'main_template', 'account_reports.main_template')
        if line_id is not None:
            render_template = templates.get(
                'line_template', 'account_reports.line_template')
        html = self.env['ir.ui.view']._render_template(
            render_template,
            values=dict(rcontext),
        )
        if self.env.context.get('print_mode', False):
            for k, v in self._replace_class().items():
                html = html.replace(k, v)
            # append footnote as well
            html = html.replace(b'<div class="js_account_report_footnotes"></div>',
                                self.get_html_footnotes(footnotes_to_render))
        return html

    @api.model
    def get_currency(self, value):
        currency_id = self.env.user.company_id.currency_id
        return formatLang(self.env, value, currency_obj=currency_id)

    def get_xlsx(self, options, response):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet(self._get_report_name()[:31])

        date_default_col1_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666', 'indent': 2, 'num_format': 'yyyy-mm-dd'})
        date_default_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666', 'num_format': 'yyyy-mm-dd'})
        default_col1_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666', 'indent': 2})
        default_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666'})
        title_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'bottom': 2})
        super_col_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'align': 'center'})
        level_0_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 13, 'bottom': 6, 'font_color': '#666666'})
        level_1_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 13, 'bottom': 1, 'font_color': '#666666'})
        level_2_col1_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 12, 'font_color': '#666666', 'indent': 1})
        level_2_col1_total_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 12, 'font_color': '#666666'})
        level_2_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 12, 'font_color': '#666666'})
        level_3_col1_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666', 'indent': 2})
        level_3_col1_total_style = workbook.add_format(
            {'font_name': 'Arial', 'bold': True, 'font_size': 12, 'font_color': '#666666', 'indent': 1})
        level_3_style = workbook.add_format(
            {'font_name': 'Arial', 'font_size': 12, 'font_color': '#666666'})

        # Set the first column width to 50
        sheet.set_column(0, 0, 50)
        super_columns = self._get_super_columns(options)
        y_offset = bool(super_columns.get('columns')) and 1 or 0

        sheet.write(y_offset, 0, '', title_style)
        is_profit = False
        profit_losse_action = self.env.ref(
            'account_reports.account_financial_report_profitandloss0')
        if profit_losse_action and profit_losse_action.id == self.id:
            is_profit = True
        # Todo in master: Try to put this logic elsewhere
        x = super_columns.get('x_offset', 0)
        for super_col in super_columns.get('columns', []):
            cell_content = super_col.get('string', '').replace(
                '<br/>', ' ').replace('&nbsp;', ' ')
            x_merge = super_columns.get('merge')
            if x_merge and x_merge > 1:
                sheet.merge_range(0, x, 0, x + (x_merge - 1),
                                  cell_content, super_col_style)
                x += x_merge
            else:
                sheet.write(0, x, cell_content, super_col_style)
                x += 1
        ctx = self._set_context(options)
        ctx.update({'is_profit': is_profit})
        header = self.get_header(options)
        if len(header) > 0:
            # Below Code to print the On-Screen Report Header Portion
            merge_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
            })
            cell_left_bold_fmt = workbook.add_format({
                'font_name': 'Arial',
                'bold': True,
                'align': 'left'
            })
            cell_left_fmt = workbook.add_format({
                'font_name': 'Arial',
                'align': 'left'
            })
            comapny_name = self.env.user.company_id and \
                self.env.user.company_id.name or ''
            sheet.merge_range(0, 0, 1, len(header[0]) - 1,
                              comapny_name, merge_format)
            sheet.merge_range(2, 0, 2, len(header[0]) - 1,
                              self._get_report_name(), merge_format)
            current_dt = datetime.datetime.now().strftime(DTF)
            tz = pytz.timezone(self._context.get('tz', 'Asia/Calcutta'))
            current_dt = pytz.timezone('UTC').localize(
                datetime.datetime.strptime(current_dt, DTF)).\
                astimezone(tz).strftime(DTF)
            sheet.set_column(3, 2, 12)
            sheet.set_column(3, 3, 12)
            sheet.write(3, 3, "Printing Date:", cell_left_bold_fmt)
            sheet.set_column(3, 4, 15)
            sheet.write(3, 4, current_dt, cell_left_fmt)
            # sheet.set_row(0, 25)
            # sheet.set_row(1, 25)
            # sheet.set_row(2, 25)
        for row in self.with_context({
                'is_profit': is_profit}).get_header(options):
            y_offset = y_offset + 5
            x = 0
            for column in row:
                colspan = column.get('colspan', 1)
                header_label = column.get('name', '').replace(
                    '<br/>', ' ').replace('&nbsp;', ' ')
                if colspan == 1:
                    sheet.write(y_offset, x, header_label, title_style)
                else:
                    sheet.merge_range(y_offset, x, y_offset,
                                      x + colspan - 1, header_label,
                                      title_style)
                x += colspan
            y_offset += 1
        options['is_profit'] = is_profit
        ctx = self._set_context(options)
        ctx.update({'no_format': True, 'print_mode': True,
                    'is_profit': is_profit})
        lines = self.with_context(ctx)._get_lines(options)

        if options.get('hierarchy'):
            lines = self._create_hierarchy(lines)

        # write all data rows
        for y in range(0, len(lines)):
            level = lines[y].get('level')
            if lines[y].get('caret_options'):
                style = level_3_style
                col1_style = level_3_col1_style
            elif level == 0:
                y_offset += 1
                style = level_0_style
                col1_style = style
            elif level == 1:
                style = level_1_style
                col1_style = style
            elif level == 2:
                style = level_2_style
                col1_style = 'total' in lines[y].get('class', '').split(
                    ' ') and level_2_col1_total_style or level_2_col1_style
            elif level == 3:
                style = level_3_style
                col1_style = 'total' in lines[y].get('class', '').split(
                    ' ') and level_3_col1_total_style or level_3_col1_style
            else:
                style = default_style
                col1_style = default_col1_style

            if 'date' in lines[y].get('class', ''):
                # write the dates with a specific format to avoid them being casted as floats in the XLSX
                if isinstance(lines[y]['name'], (datetime.date, datetime.datetime)):
                    sheet.write_datetime(
                        y + y_offset, 0, lines[y]['name'], date_default_col1_style)
                else:
                    sheet.write(y + y_offset, 0,
                                lines[y]['name'], date_default_col1_style)
            else:
                # write the first column, with a specific style to manage the indentation
                sheet.write(y + y_offset, 0, lines[y]['name'], col1_style)

            # write all the remaining cells
            for x in range(1, len(lines[y]['columns']) + 1):
                this_cell_style = style
                if 'date' in lines[y]['columns'][x - 1].get('class', ''):
                    # write the dates with a specific format to avoid them being casted as floats in the XLSX
                    this_cell_style = date_default_style
                    if isinstance(lines[y]['columns'][x - 1].get('name', ''), (datetime.date, datetime.datetime)):
                        sheet.write_datetime(y + y_offset, x + lines[y].get(
                            'colspan', 1) - 1, lines[y]['columns'][x - 1].get('name', ''), this_cell_style)
                    else:
                        sheet.write(y + y_offset, x + lines[y].get(
                            'colspan', 1) - 1, lines[y]['columns'][x - 1].get('name', ''), this_cell_style)
                else:
                    sheet.write(y + y_offset, x + lines[y].get(
                        'colspan', 1) - 1, lines[y]['columns'][x - 1].get('name', ''), this_cell_style)

        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()


class ReportAccountFinancialReport(models.Model):

    _inherit = "account.financial.html.report"

    def _get_columns_name(self, options):
        columns = [{'name': ''}]
        if self.debit_credit and not options.get('comparison', {}).get('periods', False):
            columns += [{'name': _('Debit'), 'class': 'number'},
                        {'name': _('Credit'), 'class': 'number'}]
        columns += [{'name': self.format_date(options), 'class': 'number'}]
        if options.get('comparison') and options['comparison'].get('periods'):
            for period in options['comparison']['periods']:
                columns += [{'name': period.get('string'), 'class': 'number'}]
            if options['comparison'].get('number_period') == 1 and not options.get('groups'):
                columns += [{'name': '%', 'class': 'number'}]
        if self._context.get('is_profit'):
            # columns += [{'name': '%', 'class': 'number'}]
            columns += [{'name': 'Total', 'class': 'number'}]

        if options.get('groups', {}).get('ids'):
            columns_for_groups = []
            for column in columns[1:]:
                for ids in options['groups'].get('ids'):
                    group_column_name = ''
                    for index, id in enumerate(ids):
                        column_name = self._get_column_name(
                            id, options['groups']['fields'][index])
                        group_column_name += ' ' + column_name
                    columns_for_groups.append(
                        {'name': column.get('name') + group_column_name, 'class': 'number'})
            columns = columns[:1] + columns_for_groups
        return columns


class AccountFinancialReportLine(models.Model):

    _inherit = "account.financial.html.report.line"

    def _get_lines(self, financial_report, currency_table, options, linesDicts):
        final_result_table = []
        comparison_table = [options.get('date')]
        comparison_table += options.get(
            'comparison') and options['comparison'].get('periods', []) or []
        currency_precision = self.env.user.company_id.currency_id.rounding

        # build comparison table
        is_profit = False
        opinic_total = 0
        if self._context.get('is_profit') or options.get('is_profit'):
            is_profit = True
        if (is_profit):
            if self._context and self._context.get('opinic_total'):
                opinic_total = float(self._context.get('opinic_total'))
            j = 0
            opnic_id = self.search([('code', '=', 'OPINC')])
            op_res = []
            op_domain_ids = {'line'}
            debit_credit = len(comparison_table) == 1
            for period in comparison_table:
                date_from = period.get('date_from', False)
                date_to = period.get(
                    'date_to', False) or period.get('date', False)
                date_from, date_to, strict_range = opnic_id.with_context(
                    date_from=date_from, date_to=date_to)._compute_date_range()

                r = opnic_id.with_context(
                    date_from=date_from,
                    date_to=date_to,
                    strict_range=strict_range)._eval_formula(
                        financial_report,
                        debit_credit,
                        currency_table,
                        linesDicts[j],
                        groups=options.get('groups'))
                debit_credit = False
                op_res.extend(r)
                for column in r:
                    op_domain_ids.update(column)
                j += 1
            op_res = opnic_id._put_columns_together(op_res, op_domain_ids)
            rec_total = 0
            for rec in op_res['line']:
                rec_total += rec
            opinic_total = rec_total
        for line in self:
            res = []
            debit_credit = len(comparison_table) == 1
            domain_ids = {'line'}
            k = 0
            for period in comparison_table:
                date_from = period.get('date_from', False)
                date_to = period.get(
                    'date_to', False) or period.get('date', False)
                date_from, date_to, strict_range = line.with_context(
                    date_from=date_from, date_to=date_to)._compute_date_range()
                r = line.with_context(date_from=date_from,
                                      date_to=date_to,
                                      strict_range=strict_range)._eval_formula(financial_report,
                                                                               debit_credit,
                                                                               currency_table,
                                                                               linesDicts[k],
                                                                               groups=options.get('groups'))
                debit_credit = False
                res.extend(r)
                for column in r:
                    domain_ids.update(column)
                k += 1
            res = line._put_columns_together(res, domain_ids)
            if line.hide_if_zero and all([float_is_zero(k, precision_rounding=currency_precision) for k in res['line']]):
                continue

            total = 0
            if is_profit and line.code not in ['INTP', 'OTP', 'NTP']:
                for rec in res['line']:
                    total += rec
                res['line'].append(total)
            elif is_profit:
                if line.code == 'INTP':
                    per_id = self.search([('code', '=', 'GRP')])
                if line.code == 'OTP':
                    per_id = self.search([('code', '=', 'TOP')])
                if line.code == 'NTP':
                    per_id = self.search([('code', '=', 'NEP')])
                if per_id:
                    i = 0
                    opnic_res = []
                    domain_ids_line = {'line'}
                    debit_credit = len(comparison_table) == 1
                    for period in comparison_table:
                        date_from = period.get('date_from', False)
                        date_to = period.get(
                            'date_to', False) or period.get('date', False)
                        date_from, date_to, strict_range = per_id.with_context(
                            date_from=date_from,
                            date_to=date_to)._compute_date_range()

                        r = per_id.with_context(
                            date_from=date_from,
                            date_to=date_to,
                            strict_range=strict_range)._eval_formula(
                                financial_report,
                                debit_credit,
                                currency_table,
                                linesDicts[i],
                                groups=options.get('groups'))
                        debit_credit = False
                        opnic_res.extend(r)
                        for column in r:
                            domain_ids_line.update(column)
                        i += 1
                    opnic_res = per_id._put_columns_together(
                        opnic_res, domain_ids_line)
                    total = 0
                    for rec in opnic_res['line']:
                        total += rec
                    res['line'].append(
                        line._build_cmp_percentage(total, opinic_total))
            vals = {
                'id': line.id,
                'name': line.name,
                'level': line.level,
                'class': 'o_account_reports_totals_below_sections' if self.env.user.company_id.totals_below_sections else '',
                'columns': [{'name': l} for l in res['line']],
                'unfoldable': len(domain_ids) > 1 and line.show_domain != 'always',
                'unfolded': line.id in options.get('unfolded_lines', []) or line.show_domain == 'always',
                'page_break': line.print_on_new_page,
            }
            if financial_report.tax_report and line.domain and not line.action_id:
                vals['caret_options'] = 'tax.report.line'

            if line.action_id:
                vals['action_id'] = line.action_id.id
            domain_ids.remove('line')
            lines = [vals]
            groupby = line.groupby or 'aml'
            if line.id in options.get('unfolded_lines', []) or line.show_domain == 'always':
                if line.groupby:
                    domain_ids = sorted(
                        list(domain_ids), key=lambda k: line._get_gb_name(k))
                for domain_id in domain_ids:
                    name = line._get_gb_name(domain_id)
                    if not self.env.context.get('print_mode') or not self.env.context.get('no_format'):
                        name = name[:40] + \
                            '...' if name and len(name) >= 45 else name
                    vals = {
                        'id': domain_id,
                        'name': name,
                        'level': line.level,
                        'parent_id': line.id,
                        'columns': [{'name': l} for l in res[domain_id]],
                        'caret_options': groupby == 'account_id' and 'account.account' or groupby,
                    }
                    if line.financial_report_id.name == 'Aged Receivable':
                        vals['trust'] = self.env['res.partner'].browse(
                            [domain_id]).trust
                    lines.append(vals)
                if domain_ids and self.env.user.company_id.totals_below_sections:
                    lines.append({
                        'id': 'total_' + str(line.id),
                        'name': _('Total') + ' ' + line.name,
                        'level': line.level,
                        'class': 'o_account_reports_domain_total',
                        'parent_id': line.id,
                        'columns': copy.deepcopy(lines[0]['columns']),
                    })
            for vals in lines:
                if len(comparison_table) == 2 and not options.get('groups'):
                    if is_profit and line.code not in ['INTP', 'OTP', 'NTP'] and len(vals['columns']) > 2:
                        vals['columns'].insert(-1, line._build_cmp(
                            vals['columns'][0]['name'], vals['columns'][1]['name']))
                    elif line.code not in ['INTP', 'OTP', 'NTP']:
                        vals['columns'].append(line._build_cmp(
                            vals['columns'][0]['name'], vals['columns'][1]['name']))
                    elif is_profit and line.code in ['INTP', 'OTP', 'NTP']:
                        vals['columns'].insert(-1, line._build_cmp(0, 0))

                    if line.code not in ['INTP', 'OTP', 'NTP']:

                        for i in [0, 1]:
                            vals['columns'][i] = line._format(
                                vals['columns'][i])
                        if is_profit and len(vals['columns']) > 3:
                            vals['columns'][-1] = line._format(
                                vals['columns'][-1])
                    # elif is_profit:
                    elif is_profit and line.code in ['INTP', 'OTP', 'NTP']:
                        for i in [0, 1]:
                            vals['columns'][i] = line._build_percentage_total(
                                vals['columns'][i].get('name'))
                        if is_profit:
                            vals['columns'][-1] = line._build_percentage_total(
                                vals['columns'][-1].get('name'))
                else:
                    if is_profit and line.code in ['INTP', 'OTP', 'NTP']:
                        vals['columns'] = [line._build_percentage_total(
                            v.get('name')) for v in vals['columns']]
                    else:
                        vals['columns'] = [line._format(
                            v) for v in vals['columns']]
                if not line.formulas:
                    vals['columns'] = [{'name': ''} for k in vals['columns']]
            if len(lines) == 1:
                new_lines = line.children_ids.with_context(opinic_total=opinic_total, is_profit=is_profit)._get_lines(
                    financial_report, currency_table, options, linesDicts)
                if new_lines and line.formulas:
                    if self.env.user.company_id.totals_below_sections:
                        divided_lines = self.with_context(
                            opinic_total=opinic_total, is_profit=is_profit)._divide_line(lines[0])
                        if is_profit and line.code == 'INC':
                            result = [divided_lines[0]] + new_lines
                        else:
                            result = [divided_lines[0]] + \
                                new_lines + [divided_lines[-1]]
                    else:
                        result = [lines[0]] + new_lines
                else:
                    result = lines + new_lines
            else:
                result = lines
            final_result_table += result

        return final_result_table

    def _build_cmp_percentage(self, balance, comp, is_true=False):
        if comp != 0:
            res = round(balance / comp * 100, 1)
            return res
            # In case the comparison is made on a negative figure, the color should be the other
            # way around. For example:
            #                       2018         2017           %
            # Product Sales      1000.00     -1000.00     -200.0%
            #
            # The percentage is negative, which is mathematically correct, but my sales increased
            # => it should be green, not red!
        #     if (res > 0) != (self.green_on_positive and comp > 0):
        #         return {'name': str(res) + '%', 'class': 'number color-red'}
        #     else:
        #         return {'name': str(res) + '%', 'class': 'number color-green'}
        # else:
        #     return {'name': _('n/a')}

    def _build_percentage_total(self, balance):
        if balance and balance != 0:
            # In case the comparison is made on a negative figure, the color should be the other
            # way around. For example:
            #                       2018         2017           %
            # Product Sales      1000.00     -1000.00     -200.0%
            #
            # The percentage is negative, which is mathematically correct, but my sales increased
            # => it should be green, not red!
            balance = round(balance, 1)
            if (balance > 0) != (self.green_on_positive):
                return {'name': str(balance) + '%', 'class': 'number color-red'}
            else:
                return {'name': str(balance) + '%', 'class': 'number color-green'}
        else:
            return {'name': _('n/a')}
