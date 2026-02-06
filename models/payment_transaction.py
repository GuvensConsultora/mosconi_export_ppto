from odoo import models


class PaymentTransaction(models.Model):
    # Por qué: Heredamos payment.transaction para agregar funcionalidad de chatter
    # Patrón: Herencia múltiple en Odoo - usar lista cuando se hereda más de un modelo
    # Tip: Solo UNA declaración de _inherit, con lista de modelos a heredar
    _inherit = ["payment.transaction", "mail.thread"]
