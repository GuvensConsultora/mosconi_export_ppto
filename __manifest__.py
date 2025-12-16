
{
    "name": "Export Payment Transactions to Excel (HTTP)",
    "version": "18.0",
    "depends": ["payment"],
    "data": [
        "security/ir.model.access.csv",
        "views/payment_transaction_export_views.xml"
    ],
    "installable": True,
}
