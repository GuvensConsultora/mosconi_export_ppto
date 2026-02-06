from odoo import models


class PaymentTransaction(models.Model):
    # Por qué: Extendemos payment.transaction agregando mail.thread para chatter
    # Patrón: Herencia múltiple en Odoo - _name + _inherit con lista de modelos
    # Tip: _name es necesario para evitar conflictos en campos Many2many
    _name = "payment.transaction"
    _inherit = ["payment.transaction", "mail.thread"]
