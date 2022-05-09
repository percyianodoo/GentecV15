# -*- coding: utf-8 -*-

{
    # Module information
    'name': 'Accounting Reports',
    'version': '12.0.1.0.0',
    'category': 'Accounting',
    'sequence': 1,
    'summary': """Accounting Reports.""",
    'description': """Accounting Reports.""",

    # Author
    'author': 'Serpent Consulting Services Pvt. Ltd.',
    'website': 'http://www.serpentcs.com',

    # Dependencies
    'depends': ['account_reports', 'account_followup'],

    # Views
    'data': [
        'security/ir.model.access.csv',
        'data/account_financial_report_data.xml',
        'views/account_view.xml',
        'wizard/wiz_bank_reconciliation_report_view.xml',
        # 'views/followup_view.xml',
    ],

    # Techical
    'installable': True,
    'auto_install': False
}
