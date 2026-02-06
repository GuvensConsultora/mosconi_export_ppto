from odoo import models


class PaymentTransaction(models.Model):
    # Por qué: Heredamos payment.transaction para agregar funcionalidad de chatter
    # Patrón: Herencia por extensión en Odoo (_inherit)
    _inherit = "payment.transaction"
    # Por qué: Agregamos mail.thread para habilitar message_post() y el chatter
    _inherit = ["payment.transaction", "mail.thread"]
