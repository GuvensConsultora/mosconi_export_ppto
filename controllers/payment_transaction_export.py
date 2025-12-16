
from odoo import http
from odoo.http import request
import base64

class PaymentTransactionExportController(http.Controller):

    @http.route(
        "/mosconi/export/payment_transactions",
        type="http",
        auth="user"
    )
    def export_payment_transactions(self, filename, data, **kw):
        filecontent = base64.b64decode(data)

        return request.make_response(
            filecontent,
            headers=[
                ("Content-Type",
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ("Content-Disposition", f'attachment; filename="{filename}"')
            ],
        )
