"""Account Bank Statement Model."""

from odoo import models, api


class AccountBankStatement(models.Model):
    _inherit = "account.bank.statement"


    def name_get(self):
        bnk_st_ctx = self.env.context.get('bank_st_as_date', False)
        res = []
        if bnk_st_ctx:
            for bk_st in self:
                res.append((bk_st.id, bk_st.date))
        else:
            return super(AccountBankStatement, self).name_get()
        return res
