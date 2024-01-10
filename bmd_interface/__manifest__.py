{
    'name': 'BMD-Export',
    'version': '1.0',
    'author': 'it-fact GmbH',
    'website': 'https://it-fact.com',
    'category': 'Accounting',
    'summary': 'Export von Daten für BMD',
    'application': True,
    'description': """
        Export von Daten für BMD
    """,
    'depends': ['base','account'],
    'data': [
        'wizard/bmd_export.xml',
        'security/ir.model.access.csv'
    ]
}