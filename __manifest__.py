
{
    "name": "Export Payment Transactions to Excel (HTTP)",
    "version": "18.0",
    "depends": ["payment", "mail"],
    "data": [
        "security/ir.model.access.csv",
        "views/payment_transaction_export_views.xml",
        "views/payment_transaction_views.xml"
    ],
    "installable": True,
    "license": "LGPL-3",
}
