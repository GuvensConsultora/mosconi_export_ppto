from odoo import models, fields
import io
import base64
import xlsxwriter
from datetime import datetime

class PaymentTransactionExportWizard(models.TransientModel):
    _name = "payment.transaction.export.wizard"
    _description = "Export Payment Transactions to Excel"

    file_data = fields.Binary(readonly=True)
    file_name = fields.Char(readonly=True)

    def action_export(self):
        transactions = self.env["payment.transaction"].browse(
            self.env.context.get("active_ids", [])
        )

        # Por qué: Recopilar datos del periodo para el informe del chatter
        export_datetime = datetime.now()
        total_sale_orders = len(transactions.sale_order_ids)
        total_lines = sum(len(so.order_line.filtered(
            lambda l: l.product_id.l10n_ar_ncm_code != "9999"
        )) for so in transactions.sale_order_ids)

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
            "NRO. OC ID VENTA",
            "Apellido y Nombre",
            "Tipo Doc",
            "Nro",
            "Resp frente IVA",
            "Domicilio",
            "Pais",
            "Provincia",
            "Ciudad",
            "Cod. Postal",
            "Email",
            "Teléfono",
            "Vacio",
            "Status",
            "SKU",
            "Nombre Producto",
            "Cantidad",
            "Precio Unitario",
            "Pcio Tot con Imp",
            "Id Cliente",
            "Valor Flete",
            
        ]

        for col, h in enumerate(headers):
            sheet.write(0, col, h, header)

        row = 1
        for t in transactions:
            for sale_order in t.sale_order_ids:
                # Por qué: Necesitamos la primera fila del pedido para escribir el flete
                first_row_of_order = row

                # Por qué: Iteramos directamente order_line (ya es un recordset)
                for line in sale_order.order_line:
                    # Por qué: Productos con NCM "9999" son flete (se escriben en col 24 de la primera línea)
                    if line.product_id.l10n_ar_ncm_code == "9999":
                        sheet.write(first_row_of_order, 24, line.price_unit or "")
                    else:
                        # Por qué: Concatenamos street y street2 para dirección completa
                        street = sale_order.partner_shipping_id.street or ""
                        street2 = sale_order.partner_shipping_id.street2 or ""
                        domicilio = f"{street} {street2}".strip()

                        sheet.write(row, 1, str(sale_order.date_order or ""))
                        sheet.write(row, 2, str(sale_order.date_order or ""))
                        sheet.write(row, 4, t.reference or "")
                        sheet.write(row, 5, t.partner_id.name or "")
                        sheet.write(row, 6, t.partner_id.l10n_latam_identification_type_id.name or "")
                        sheet.write(row, 7, t.partner_id.vat or "")
                        sheet.write(row, 8, t.partner_id.l10n_ar_afip_responsibility_type_id.name or "")
                        sheet.write(row, 9, domicilio)
                        sheet.write(row, 10, sale_order.partner_shipping_id.country_id.name or "")
                        sheet.write(row, 11, sale_order.partner_shipping_id.state_id.name or "")
                        sheet.write(row, 12, sale_order.partner_shipping_id.city or "")
                        sheet.write(row, 13, t.partner_id.zip or "")
                        sheet.write(row, 14, t.partner_id.email or "")
                        sheet.write(row, 15, sale_order.partner_shipping_id.phone or sale_order.partner_invoice_id.phone)
                        sheet.write(row, 16, "")
                        sheet.write(row, 17, t.state or "")
                        sheet.write(row, 18, line.product_id.default_code or "")
                        sheet.write(row, 19, line.product_id.name or "")
                        sheet.write(row, 20, line.product_uom_qty or "")
                        sheet.write(row, 21, line.price_unit or "")
                        sheet.write(row, 22, line.price_total or "")
                        sheet.write(row, 23, sale_order.partner_id.id or "")
                        # Tip: Dejar col 24 vacía, se llena con el producto flete al final
                        row += 1

        # 3️⃣ Cerrar workbook (CRÍTICO)
        workbook.close()

        # 4️⃣ Volver al inicio del buffer (CRÍTICO)
        output.seek(0)

        # 5️⃣ Base64 UNA sola vez
        self.file_data = base64.b64encode(output.read())
        self.file_name = (
            f"payment_transactions_{export_datetime.strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        # Por qué: Registrar en el chatter de cada transacción para auditoría
        # Patrón: message_post() en modelos mail.thread
        export_info = f"""
        <p><strong>📊 Exportación a Excel realizada</strong></p>
        <ul>
            <li><strong>Fecha/Hora:</strong> {export_datetime.strftime('%d/%m/%Y %H:%M:%S')}</li>
            <li><strong>Usuario:</strong> {self.env.user.name}</li>
            <li><strong>Archivo:</strong> {self.file_name}</li>
            <li><strong>Órdenes exportadas:</strong> {total_sale_orders}</li>
            <li><strong>Líneas de productos:</strong> {total_lines}</li>
        </ul>
        """

        for transaction in transactions:
            transaction.message_post(
                body=export_info,
                subject="Exportación a Excel",
                message_type="notification",
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
