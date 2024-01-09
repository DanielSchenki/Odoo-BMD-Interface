{
    'name': 'BMD-Export',
    'version': '1.0',
    'author': 'it-fact',
    'website': 'https://it-fact.com',
    'category': 'Accounting',
    'summary': 'Export von Daten für BMD',
    'application': True,
    'description': """
        Export von Daten für BMD
    """,
    'sequence': '1',
    'depends': ['base', 'account'],
    'data': [
        'wizard/bmd_export.xml',
        'security/ir.model.access.csv'
    ]
}
