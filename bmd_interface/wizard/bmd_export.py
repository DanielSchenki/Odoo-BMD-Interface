import base64
import csv
import io
import re
import zipfile
import math

from odoo.exceptions import ValidationError
from odoo.http import request

from odoo import models, fields, api, http


class CsvDownloadController(http.Controller):
    @http.route('/download', type='http', auth="user")
    def download_csv(self):
        zip_content = request.env['account.bmd'].combine_to_zip()
        date_form = request.env['account.bmd'].search([])[-1]
        formatted_date_from = date_form.period_date_from.strftime('%y%m%d')
        formatted_date_to = date_form.period_date_to.strftime('%y%m%d')
        formatted_company = date_form.company.name.replace(' ', '_')
        # Return the ZIP file
        response = http.request.make_response(zip_content.getvalue(), headers=[('Content-Type', 'application/zip'), (
            'Content-Disposition',
            f'attachment; filename="BMD_Export_{formatted_company}_{formatted_date_from}_{formatted_date_to}.zip"')])
        return response


def commercial_round_3_digits(number):
    # Schritt 1: Multipliziere mit 1000
    number *= 1000

    # Extrahiere die dritte und vierte Dezimalstelle
    dritte_dezimal = int(number) % 10
    vierte_dezimal = int(number * 10) % 10

    # Schritt 2: Kaufmännisches Runden
    if vierte_dezimal < 5:
        number = math.floor(number)
    elif vierte_dezimal > 5:
        number = math.ceil(number)
    else:
        # Schritt 3: Runde zur nächsten geraden Zahl, wenn vierte Dezimalstelle genau 5 ist
        if dritte_dezimal % 2 == 0:  # Gerade Zahl
            number = math.floor(number)
        else:  # Ungerade Zahl
            number = math.ceil(number)

    # Teile durch 1000, um das Ergebnis zu normalisieren
    return number / 1000

def testCommercialRound():
    testzahlen = [
        123.45674,  # Vierte Dezimalstelle < 5
        123.45676,  # Vierte Dezimalstelle > 5
        123.45650,  # Vierte Dezimalstelle = 5, dritte Dezimalstelle gerade
        123.45550,  # Vierte Dezimalstelle = 5, dritte Dezimalstelle ungerade
        100,  # Keine Dezimalstellen
        123.45,  # Weniger als drei Dezimalstellen
        -123.45675  # Negative Zahl
    ]
    ergebnisse = [commercial_round_3_digits(zahl) for zahl in testzahlen]
    print(ergebnisse)


class AccountBmdExport(models.TransientModel):
    _name = 'account.bmd'
    _description = 'BMD Export'

    period_date_from = fields.Date(string="Von:", required=True)
    period_date_to = fields.Date(string="Bis:", required=True)
    company = fields.Many2one('res.company', string="Mandant:", required=True)

    # Checks if the date is valid
    @api.constrains('period_date_from', 'period_date_to')
    def _check_date(self):
        for _ in self:
            if self.period_date_from > self.period_date_to:
                raise ValidationError('Das Startdatum muss vor dem Enddatum liegen!')

    # Returns the accounts
    @api.model
    def export_accounts(self):
        accounts = self.env['account.account'].search([])
        date_form = self.env['account.bmd'].search([])[-1]
        result_data = []

        for acc in accounts:
            if acc.company_id.id != date_form.company.id:
                continue
            kontoart_mapping = {'asset': 1, 'equity': 2, 'liability': 2, 'expense': 3, 'income': 4}
            kontoart = kontoart_mapping.get(acc.internal_group, '')
            if not acc.tax_ids:
                result_data.append({
                    'Konto-Nr': acc.code,
                    'Bezeichnung': acc.name,
                    'Ustcode': '',
                    'USTPz': '',
                    'Kontoart': kontoart
                })
            else:
                for tax in acc.tax_ids:
                    result_data.append({
                        'Konto-Nr': acc.code,
                        'Bezeichnung': acc.name,
                        'Ustcode': tax.tax_group_id.id if tax.tax_group_id else '',
                        'USTPz': tax.amount,
                        'Kontoart': kontoart
                    })

        # Write the data to the CSV file
        account_buffer = io.StringIO()
        fieldnames = ['Konto-Nr', 'Bezeichnung', 'Ustcode', 'USTPz', 'Kontoart']
        writer = csv.DictWriter(account_buffer, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in result_data:
            cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                           row.items()}
            writer.writerow(cleaned_row)

        return account_buffer.getvalue()

    # Returns the customers
    @api.model
    def export_customers(self):
        customers = self.env['res.partner'].search([])
        date_form = self.env['account.bmd'].search([])[-1]
        # Write to the CSV file
        customer_buffer = io.StringIO()
        fieldnames = ['Konto-Nr', 'Name', 'Straße', 'PLZ', 'Ort', 'Land', 'UID-Nummer', 'E-Mail', 'Webseite', 'Phone',
                      'IBAN', 'Zahlungsziel', 'Skonto', 'Skontotage']
        writer = csv.DictWriter(customer_buffer, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        for customer in customers:
            if customer.company_id.id != date_form.company.id:
                continue
            # Write row for receivable account
            writer.writerow({
                'Konto-Nr': customer.property_account_receivable_id.code if customer.property_account_receivable_id else '',
                'Name': customer.name if customer.name else '', 'E-Mail': customer.email if customer.email else '',
                'Phone': customer.phone if customer.phone else '', 'Ort': customer.city if customer.city else '',
                'Straße': customer.street if customer.street else '', 'PLZ': customer.zip if customer.zip else '',
                'Webseite': customer.website if customer.website else '',
                'UID-Nummer': customer.vat if customer.vat else '',
                'Land': customer.state_id.code if customer.state_id else '', })
            # Write row for payable account
            writer.writerow({
                'Konto-Nr': customer.property_account_payable_id.code if customer.property_account_payable_id else '',
                'Name': customer.name if customer.name else '', 'E-Mail': customer.email if customer.email else '',
                'Phone': customer.phone if customer.phone else '', 'Ort': customer.city if customer.city else '',
                'Straße': customer.street if customer.street else '', 'PLZ': customer.zip if customer.zip else '',
                'Webseite': customer.website if customer.website else '',
                'UID-Nummer': customer.vat if customer.vat else '',
                'Land': customer.state_id.code if customer.state_id else '', })

            return customer_buffer.getvalue()

    # Returns the booking lines
    def get_account_movements(self):
        gkonto = ""

        # date formatter from yyyy-mm-dd to dd.mm.yyyy
        def date_formatter(date):
            return date.strftime('%d.%m.%Y')

        buchsymbol_mapping = {'sale': 'AR', 'purchase': 'ER', 'general': 'SO', 'cash': 'KA', 'bank': 'BK'}
        journal_items = self.env['account.move.line'].search([])
        date_form = self.env['account.bmd'].search([])[-1]
        result_data = []
        docs = []

        journal_items = sorted(journal_items, key=lambda x: x.move_id.id)

        for line in journal_items:

            if line.company_id.id != date_form.company.id:
                continue

            belegdatum = line.date
            if date_form.period_date_from > belegdatum or belegdatum > date_form.period_date_to:
                continue

            belegdatum = date_formatter(belegdatum)
            konto = line.account_id.code
            prozent = line.tax_ids.amount
            steuer = line.price_total - line.price_subtotal
            belegnr = line.move_id.name
            text = line.name
            if (steuer != 0) and (line.tax_ids.name is not False):
                steuercode_before_cut = line.tax_ids.name
                pattern = r"BMDSC\d{3}$"
                if re.search(pattern, steuercode_before_cut):
                    steuercode = int(steuercode_before_cut[-3:])
                else:
                    steuercode = "2"
            else:
                steuercode = "2"

            if line.debit > 0:
                buchcode = 1
                haben_buchung = False
            else:
                buchcode = 2
                haben_buchung = True

            buchsymbol = buchsymbol_mapping.get(line.journal_id.type, '')

            if haben_buchung:
                betrag = -line.credit  # Haben Buchungen müssen negativ sein
            else:
                betrag = line.debit

            if buchsymbol == 'ER':
                gkonto = line.partner_id.property_account_payable_id.code
            elif buchsymbol == 'AR':
                gkonto = line.partner_id.property_account_receivable_id.code
            else:
                # runs through all invoice lines to find the right "Gegenkonto"
                for invoice_line in line.move_id.invoice_line_ids:
                    if invoice_line.credit > 0:
                        gkonto = invoice_line.account_id.code

            # switch Soll Haben bei Ausgangsrechnungen
            if buchsymbol == 'AR':
                temp_gkonto = gkonto
                gkonto = konto
                konto = temp_gkonto

            buchungszeile = line.move_id

            move_id = line.move_id.id

            satzart = 0
            # Rounding
            betrag = commercial_round_3_digits(betrag)
            steuer = commercial_round_3_digits(steuer)


            dokument = ''

            additional_documents = []

            attachments = self.env['ir.attachment'].search([])
            for att in attachments:
                if move_id == att.res_id and dokument == '' and att.id not in docs:
                    dokument = att.name
                    docs.append(att.id)
                elif move_id == att.res_id and att.id not in docs:
                    additional_documents.append({'document': att.name})
                    docs.append(att.id)

            result_data.append(
                {'satzart': satzart, 'konto': konto, 'gKonto': gkonto, 'belegnr': belegnr, 'belegdatum': belegdatum,
                 'steuercode': steuercode, 'buchcode': buchcode, 'betrag': betrag, 'prozent': prozent,
                 'steuer': steuer, 'text': text, 'buchsymbol': buchsymbol, 'buchungszeile': buchungszeile,
                 'move_id': move_id, 'dokument': dokument})

            if additional_documents:
                for doc in additional_documents:
                    result_data.append({
                        'satzart': '5', 'konto': '', 'gKonto': '', 'belegnr': '', 'belegdatum': '', 'steuercode': '',
                        'buchcode': '', 'betrag': '', 'prozent': '', 'steuer': '', 'text': '', 'buchsymbol': '',
                        'buchungszeile': '', 'move_id': '', 'dokument': doc['document']})

        # Remove tax lines and haben buchung
        for data in result_data:
            for check_data in result_data:
                if (data['buchsymbol'] == 'ER' or data['buchsymbol'] == 'AR') and data['buchungszeile'] == check_data[
                    'buchungszeile'] and data['prozent'] and not check_data['prozent']:
                    result_data.remove(check_data)

        return result_data

    # Exports the documents
    def export_attachments(self):
        attachments = self.env['ir.attachment'].search([])
        return_data = []
        docs = []

        for att in attachments:
            if att.res_model == 'account.move':
                for data in self.get_account_movements():
                    if data['move_id'] == att.res_id and att.id not in docs:
                        return_data.append(att)
                        docs.append(att.id)

        return return_data

    # Exports the booking lines
    def export_account_movements(self):
        csvBuffer = io.StringIO()
        fieldnames = ['satzart', 'konto', 'gKonto', 'belegnr', 'belegdatum', 'steuercode', 'buchcode', 'betrag',
                      'prozent', 'steuer', 'text', 'buchsymbol', 'dokument']
        writer = csv.DictWriter(csvBuffer, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in self.get_account_movements():
            del row['move_id']
            del row['buchungszeile']
            cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                           row.items()}
            writer.writerow(cleaned_row)

        return csvBuffer.getvalue()

    # Executes the export
    def execute(self):
        action = {'type': 'ir.actions.act_url', 'url': '/download', 'target': 'self', }
        return action

    # Combines all files to one zip file
    def combine_to_zip(self):
        #testCommercialRound()
        zip_buffer = io.BytesIO()
        date_form = self.env['account.bmd'].search([])[-1]
        formatted_date_from = date_form.period_date_from.strftime('%y%m%d')
        formatted_date_to = date_form.period_date_to.strftime('%y%m%d')
        formatted_company = date_form.company.name.replace(' ', '_')

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            accountContent = self.export_accounts()
            zip_file.writestr(f'Sachkonten_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv',
                              accountContent)
            customerContent = self.export_customers()
            zip_file.writestr(f'Personenkonten_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv',
                              customerContent)
            entryContent = self.export_account_movements()
            zip_file.writestr(f'Buchungszeilen_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv',
                              entryContent)
            for att in self.export_attachments():
                zip_file.writestr(att.name, base64.b64decode(att.datas))
        zip_buffer.seek(0)

        return zip_buffer
