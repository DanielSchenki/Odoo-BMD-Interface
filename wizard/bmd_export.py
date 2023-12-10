import csv
import re
import datetime
import tkinter as tk
from tkinter import filedialog
import os

from odoo import models, fields, api


class AccountBmdExport(models.TransientModel):
    _name = 'account.bmd'
    _description = 'BMD Export'

    period_date_from = fields.Date(string="Von:", required=True)
    period_date_to = fields.Date(string="Bis:", required=True)

    path = fields.Char(string="Pfad:", required=False)


    @api.model
    def export_account(self):
        accounts = self.env['account.account'].search([])

        result_data = []
        for acc in accounts:
            if not acc.tax_ids:
                kontoart_mapping = {
                    'asset': 1,
                    'equity': 2,
                    'liability': 2,
                    'expense': 3,
                    'income': 4
                }
                kontoart = kontoart_mapping.get(acc.internal_group, '')
                result_data.append({
                    'Konto-Nr': acc.code,
                    'Bezeichnung': acc.name,
                    'Ustcode': '',
                    'USTPz': '',
                    'Kontoart': kontoart
                })
            else:
                for tax in acc.tax_ids:
                    kontoart_mapping = {
                        'asset': 1,
                        'equity': 2,
                        'liability': 2,
                        'expense': 3,
                        'income': 4
                    }
                    kontoart = kontoart_mapping.get(acc.internal_group, '')
                    result_data.append({
                        'Konto-Nr': acc.code,
                        'Bezeichnung': acc.name,
                        'Ustcode': tax.tax_group_id.id if tax.tax_group_id else '',
                        'USTPz': tax.amount,
                        'Kontoart': kontoart
                    })

        # if len(accounts) != len(result_data):
        #     raise Warning('Steuerklassen sind nicht für alle Sachkonten gepflegt')
        #

        # Create a Tkinter window
        window = tk.Tk()


        #Angepasst von mir an die Allgemeine Directory!
        save_path = self.path + '/Sachkonten.csv'
        directory = os.path.dirname(save_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Write the data to the CSV file
        with open(save_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Konto-Nr', 'Bezeichnung', 'Ustcode', 'USTPz', 'Kontoart']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')

            writer.writeheader()
            for row in result_data:
                cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                               row.items()}
                writer.writerow(cleaned_row)

        window.destroy()

        return True

    @api.model
    def selectPath(self):
        window = tk.Tk()
        print("selectPath")
        self.path = filedialog.askdirectory()
        window.destroy()


    @api.model
    @api.model
    def export_customers(self):

        customers = self.env['res.partner'].search([])

        path1 = self.path + '/PersonenkontenAllAccountsBothAccounts.csv'
        directory = os.path.dirname(path1)
        if not os.path.exists(directory):
            os.makedirs(directory)

        # Write to the CSV file
        with open(path1, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Konto-Nr', 'Name', 'Straße', 'PLZ', 'Ort', 'Land', 'UID-Nummer', 'E-Mail', 'Webseite',
                          'Phone', 'IBAN', 'Zahlungsziel', 'Skonto', 'Skontotage']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')

            writer.writeheader()
            for customer in customers:
                # Write row for receivable account
                writer.writerow({
                    'Konto-Nr': customer.property_account_receivable_id.code if customer.property_account_receivable_id else '',
                    'Name': customer.name if customer.name else '',
                    'E-Mail': customer.email if customer.email else '',
                    'Phone': customer.phone if customer.phone else '',
                    'Ort': customer.city if customer.city else '',
                    'Straße': customer.street if customer.street else '',
                    'PLZ': customer.zip if customer.zip else '',
                    'Webseite': customer.website if customer.website else '',
                    'UID-Nummer': customer.vat if customer.vat else '',
                    'Land': customer.state_id.code if customer.state_id else '',
                })
                # Write row for payable account
                writer.writerow({
                    'Konto-Nr': customer.property_account_payable_id.code if customer.property_account_payable_id else '',
                    'Name': customer.name if customer.name else '',
                    'E-Mail': customer.email if customer.email else '',
                    'Phone': customer.phone if customer.phone else '',
                    'Ort': customer.city if customer.city else '',
                    'Straße': customer.street if customer.street else '',
                    'PLZ': customer.zip if customer.zip else '',
                    'Webseite': customer.website if customer.website else '',
                    'UID-Nummer': customer.vat if customer.vat else '',
                    'Land': customer.state_id.code if customer.state_id else '',
                })

        return True

    def export_buchungszeilen(self):
        print("==============> Generating csv Files for BMD export")
        gkonto = ""

        # date formatter from yyyy-mm-dd to dd.mm.yyyy
        def date_formatter(date):
            return date.strftime('%d.%m.%Y')

        buchsymbol_mapping = {
            'sale': 'AR',
            'purchase': 'ER',
            'general': 'SO',
            'cash': 'KA',
            'bank': 'BK'
        }
        journal_items = self.env['account.move.line'].search([])
        result_data = []
        for line in journal_items:
            belegdatum = line.date
            if self.period_date_from > belegdatum or belegdatum > self.period_date_to:
                continue
            belegdatum = date_formatter(belegdatum)
            if line.account_id.code == False:
                 continue
            else:
                konto = line.account_id.code[:10]
            if line.tax_ids.amount == False:
                prozent = 0
            else:
                prozent = round(line.tax_ids.amount,3)

            if line.price_total == False or line.price_subtotal == False:
                steuer = 0
            else:
                steuer = round((line.price_total - line.price_subtotal),2)
            if line.move_id.name == False:
                belegnr = ""
            else:
                belegnr = line.move_id.name[:20]
            if line.name == False:
                text = ""
            else:
                text = line.name[:255]
            if steuer != 0 and line.tax_ids.name != False:
                steuercode_before_cut = line.tax_ids.name
                # Test String
                # steuercode_before_cut = "UST_056 Tax invoiced accepted (§ 11 Abs. 12 und 14, § 16 Abs. 2 sowie gemäß Art. 7 Abs. 4) BMDSC043"
                pattern = r"BMDSC\d{3}$"
                if re.search(pattern, steuercode_before_cut):
                    steuercode = int(steuercode_before_cut[-3:])
                else:
                    steuercode = "002" #Default
            else:
                steuercode = "002" #Default

            if line.debit > 0:
                buchcode = 1
                habenBuchung = False
            else:
                buchcode = 2
                habenBuchung = True

            buchsymbol = buchsymbol_mapping.get(line.journal_id.type, '')


            if habenBuchung:
                betrag = round(-line.credit,2)  # Haben Buchungen müssen negativ sein
            else:
                betrag = round(line.debit,2)


            if buchsymbol == 'ER':
                gkonto = line.partner_id.property_account_payable_id.code
            elif buchsymbol == 'AR':
                gkonto = line.partner_id.property_account_receivable_id.code
            else:
                for invoice_line in line.move_id.invoice_line_ids:  # runs through all invoice lines to find the right "Gegenkonto"
                    if invoice_line.credit > 0:
                        gkonto = invoice_line.account_id.code

            #switch Soll Haben bei Ausgangsrechnungen
            if buchsymbol == 'AR':
                temp_gkonto = gkonto
                gkonto = konto
                konto = temp_gkonto

            buchungszeile = line.move_id

            # TODO: Add the correct values for the following fields
            satzart = 0

            result_data.append({
                'satzart': satzart,
                'konto': konto,
                'gKonto': gkonto,
                'belegnr': belegnr,
                'belegdatum': belegdatum,
                'steuercode': steuercode,
                'buchcode': buchcode,
                'betrag': betrag,
                'prozent': prozent,
                'steuer': steuer,
                'text': text,
                'buchsymbol': buchsymbol,
                'buchungszeile': buchungszeile,
            })

        # Remove tax lines and haben buchung

        for data in result_data:
            for check_data in result_data:
                if (data['buchsymbol'] == 'ER' or data['buchsymbol'] == 'AR') and data['buchungszeile'] == check_data['buchungszeile'] and data['prozent'] and not check_data['prozent']:
                    result_data.remove(check_data)


        save_path = self.path + '/Buchungszeilen.csv'
        directory = os.path.dirname(save_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        with open(save_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['satzart', 'konto', 'gKonto', 'belegnr', 'belegdatum', 'steuercode', 'buchcode', 'betrag',
                          'prozent', 'steuer', 'text', 'buchsymbol']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')

            writer.writeheader()
            for row in result_data:
                cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                               row.items()}
                del cleaned_row['buchungszeile']
                writer.writerow(cleaned_row)

        print("==============> Done")


    def execute(self):
        self.selectPath()
        self.export_account()
        self.export_customers()
        self.export_buchungszeilen()
        return True
