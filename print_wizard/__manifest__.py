{
    'name': 'Big Data Reports',
    'version': '1.2',
    'sequence': 100,
    'author': "JD DEVS",
    'depends': ['base', 'mrp', 'sale', 'purchase', 'stock', 'sale_stock', 'account'],
    'data': [
        'security/ir.model.access.csv',
        'security/groups.xml',
        'wizards/print_wizard.xml',
    ],
    'assets': {
        'web.assets_backend': [
            "print_wizard/static/src/js/tree_button.js",
        ],
        'web.assets_qweb': [
            "print_wizard/static/legacy/xml/base.xml",
        ],
    },
    'price': 150,
    'currency': 'USD',
    'installable': True,
    'application': True,
    'auto_install': False,
    'license': 'AGPL-3',
    'images': ['static/description/assets/screenshots/banner.png'],
}
