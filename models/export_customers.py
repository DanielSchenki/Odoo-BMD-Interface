from odoo import models, api, fields
import csv
import os


class CustomerExport(models.Model):
    _name = 'customer.export'
    _description = "Exportieren von Kunden"

    name = fields.Char("Name")
    reference = fields.Char("Reference")



    def export_bmd(self):
        print("==============> Generating csv Files for BMD export")
        pass

