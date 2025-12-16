
from odoo import models
import io
import base64
import xlsxwriter
from datetime import datetime

class PaymentTransactionExportWizard(models.TransientModel):
    _name = "payment.transaction.export.wizard"
    _description = "Export Payment Transactions to Excel"

    def action_export(self):
        transactions = self.env["payment.transaction"].browse(
            self.env.context.get("active_ids", [])
        )

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Transacciones")
        header = workbook.add_format({"bold": True})

        headers = [
            "Referencia", "Creado el", "Método",
            "Proveedor", "Cliente", "Email",
            "Importe", "Estado", "Empresa"
        ]

        for col, h in enumerate(headers):
            sheet.write(0, col, h, header)

        row = 1
        for t in transactions:
            sheet.write(row, 0, t.reference or "")
            sheet.write(row, 1, str(t.create_date or ""))
            sheet.write(row, 2, t.payment_method_id.name or "")
            sheet.write(row, 3, t.provider_id.name or "")
            sheet.write(row, 4, t.partner_id.name or "")
            sheet.write(row, 5, t.partner_id.email or "")
            sheet.write(row, 6, t.amount or 0.0)
            sheet.write(row, 7, t.state or "")
            sheet.write(row, 8, t.company_id.name or "")
            row += 1

        workbook.close()
        output.seek(0)

        # Guardamos en sesión (NO en URL)
        request.session["mosconi_xlsx"] = output.read()
        #data = base64.b64encode(output.read()).decode()
        filename = f"payment_transactions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        return {
            "type": "ir.actions.act_url",
            "url": f"/mosconi/export/payment_transactions?filename={filename}&data={data}",
            "target": "self",
        }
