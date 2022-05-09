# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import models, _


class AccountFollowupReport(models.AbstractModel):
    _inherit = "account.followup.report"
    _description = "Follow-up Report"

    def _get_columns_name(self, options):
        headers = [{},
                   {'name': _('Date'), 'class': 'date',
                    'style': 'text-align:center; white-space:nowrap;'},
                   {'name': _('Due Date'), 'class': 'date',
                    'style': 'text-align:center; white-space:nowrap;'},
                   {'name': _('Source Document'),
                       'style': 'text-align:center; white-space:nowrap;'},
                   {'name': _('Communication'),
                       'style': 'text-align:right; white-space:normal;'},
                   {'name': _('Expected Date'), 'class': 'date',
                       'style': 'white-space:nowrap;'},
                   {'name': _('Excluded'), 'class': 'date',
                       'style': 'white-space:nowrap;'},
                   {'name': _('Total Due'), 'class': 'number o_price_total',
                       'style': 'text-align:right; white-space:nowrap;'}
                   ]
        if self.env.context.get('print_mode'):
            headers = headers[:5] + headers[7:]
        return headers
