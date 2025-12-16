from odoo import models, fields
import io
import base64
import xlsxwriter
from datetime import datetime
from odoo.exceptions import UserError

class PaymentTransactionExportWizard(models.TransientModel):
    _name = "payment.transaction.export.wizard"
    _description = "Export Payment Transactions to Excel"

    file_data = fields.Binary(readonly=True)
    file_name = fields.Char(readonly=True)

    def action_export(self):
        transactions = self.env["payment.transaction"].browse(
            self.env.context.get("active_ids", [])
        )

        # 1️⃣ Buffer en memoria
        output = io.BytesIO()

        # 2️⃣ Workbook correctamente inicializado
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Transacciones")

        header = workbook.add_format({"bold": True})

        headers = [
            "Importar",
            "Fecha Pedido",
            "Fecha Entrega",
            "Cod, Hora",
            "Referencia",
            "Apellido y Nombre",
            "Entrega",
            "Proveedor",
            "Cliente",
            "Email",
            "Importe",
            "Estado",
            "Empresa",
        ]

        for col, h in enumerate(headers):
            sheet.write(0, col, h, header)

        row = 1
        for t in transactions:
            texto = ""
            for sale_order in t.sale_order_ids:
                for lines in sale_order.order_line:
                    or_line = 1
                    for line in lines:
                        if or_line == "1":
                            row_flete = row
                        if line.product_template_id.l10n_ar_ncm_code == "9999":
                            sheet.write(row_flete, 9, line.price_unit or "")
                            texto += (f"Nro {line.display_name} \n Cantidad {line.product_uom_qty} \n Precio {line.price_unit} \n ")
                    #raise UserError(f"Referencia {t.reference} \n  Ppto {t.sale_order_ids} \n Lineas de pptos {sale_order.order_line} \n {texto}" )
                        else:
                            sheet.write(row, 4, t.reference or "")
                            sheet.write(row, 2, str(t.date_order or ""))
                            sheet.write(row, 3, str(t.date_order or ""))
                            sheet.write(row, 5, t.partner_id.name or "")
                            sheet.write(row, 6, t.partner_id.partner_shipping_id or "")
                            sheet.write(row, 5, t.partner_id.email or "")
                            sheet.write(row, 6, float(t.amount or 0.0))
                            sheet.write(row, 7, t.state or "")
                            sheet.write(row, 8, line.display_name or "")
                            row += 1

        # 3️⃣ Cerrar workbook (CRÍTICO)
        workbook.close()

        # 4️⃣ Volver al inicio del buffer (CRÍTICO)
        output.seek(0)

        # 5️⃣ Base64 UNA sola vez
        self.file_data = base64.b64encode(output.read())
        self.file_name = (
            f"payment_transactions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        # 6️⃣ Descargar correctamente
        return {
            "type": "ir.actions.act_url",
            "url": (
                f"/web/content/?model={self._name}"
                f"&id={self.id}"
                f"&field=file_data"
                f"&filename={self.file_name}"
                f"&download=true"
            ),
            "target": "self",
        }
