
from odoo import http
from odoo.http import request


class PaymentTransactionExportController(http.Controller):

    @http.route(
        "/mosconi/export/payment_transactions",
        type="http",
        auth="user"
    )
    def export_payment_transactions(self, filename, **kw):
        filecontent = request.session.pop("mosconi_xlsx", None)
        #filecontent = base64.b64decode(data)

        if not filecontent:
            return request.not_found()
        
        return request.make_response(
            filecontent,
            headers=[
                ("Content-Type",
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                ("Content-Disposition", f'attachment; filename="{filename}"')
            ],
        )
