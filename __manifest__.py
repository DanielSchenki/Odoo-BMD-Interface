{
    'name': 'BMD-Export',
    'version': '1.0',
    'author': 'It-fact',
    'website': 'https://it-fact.com',
    'category': 'Accounting',
    'summary': 'Export von Daten für BMD',
    'application': True,
    'description': """
        Export von Daten für BMD
    """,
    'sequence': '1',
    'depends': ['base','account'],
    'data': [
        'wizard/bmd_export.xml',
        'views/error_popup_template.xml'
    ]
}