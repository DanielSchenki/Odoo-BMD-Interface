import base64
import csv
import re
import io
import zipfile

from odoo import models, fields, api, http
from odoo.http import request


class CsvDownloadController(http.Controller):
    @http.route('/download', type='http', auth="user")
    def download_csv(self):
        zip_content = request.env['account.bmd'].combine_to_zip()

        date_form = request.env['account.bmd'].search([])[-1]
        formated_date_from = date_form.period_date_from.strftime('%y%m%d')
        formated_date_to = date_form.period_date_to.strftime('%y%m%d')
        formated_company = date_form.company.name.replace(' ', '_')
        # Return the ZIP file
        response = http.request.make_response(zip_content.getvalue(),
                                              headers=[
                                                  ('Content-Type', 'application/zip'),
                                                  ('Content-Disposition',
                                                   'attachment; filename="BMD_Export_'
                                                   + formated_company + '_'
                                                   + formated_date_from + '_'
                                                   + formated_date_to + '.zip"')
                                              ])
        return response


class AccountBmdExport(models.TransientModel):
    _name = 'account.bmd'
    _description = 'BMD Export'

    period_date_from = fields.Date(string="Von:", required=True)
    period_date_to = fields.Date(string="Bis:", required=True)
    company = fields.Many2one('res.company', string="Mandant:", required=True)

    path = fields.Char(string="Pfad:", required=False)

    # Returns the accounts
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

        # Write the data to the CSV file
        account_buffer = io.StringIO()
        fieldnames = ['Konto-Nr', 'Bezeichnung', 'Ustcode', 'USTPz', 'Kontoart']
        writer = csv.DictWriter(account_buffer, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in result_data:
            cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                           row.items()}
            writer.writerow(cleaned_row)

        # window.destroy()

        return account_buffer.getvalue()

    # Returns the customers
    @api.model
    def export_customers(self):

        customers = self.env['res.partner'].search([])

        # Write to the CSV file
        customer_buffer = io.StringIO()
        fieldnames = ['Konto-Nr', 'Name', 'Straße', 'PLZ', 'Ort', 'Land', 'UID-Nummer', 'E-Mail', 'Webseite',
                      'Phone', 'IBAN', 'Zahlungsziel', 'Skonto', 'Skontotage']
        writer = csv.DictWriter(customer_buffer, fieldnames=fieldnames, delimiter=';')

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

        return customer_buffer.getvalue()

    # Returns the booking lines
    def get_buchungszeilen(self):
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
        date_form = self.env['account.bmd'].search([])[-1]
        result_data = []
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
            if (steuer != 0) and (line.tax_ids.name != False):
                steuercode_before_cut = line.tax_ids.name
                pattern = r"BMDSC\d{3}$"
                if re.search(pattern, steuercode_before_cut):
                    steuercode = int(steuercode_before_cut[-3:])
                else:
                    steuercode = "002"
            else:
                steuercode = "002"

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
                'move_id': move_id,
                'dokument': ''
            })

        # Remove tax lines and haben buchung
        for data in result_data:
            for check_data in result_data:
                if ((data['buchsymbol'] == 'ER' or data['buchsymbol'] == 'AR') and
                        data['buchungszeile'] == check_data['buchungszeile'] and
                        data['prozent'] and not check_data['prozent']):
                    result_data.remove(check_data)

        return result_data

    # Exports the documents
    def export_documents(self):
        attachments = self.env['ir.attachment'].search([])
        return_data = []
        unique_move_ids = set()

        for att in attachments:
            if att.res_model == 'account.move':
                for data in self.get_buchungszeilen():
                    if data['move_id'] == att.res_id:
                        if data['move_id'] not in unique_move_ids:
                            return_data.append(att)
                            unique_move_ids.add(data['move_id'])

        return return_data

    # Adds the documents to the booking lines
    def add_documents_to_booking_lines(self):
        attachments = self.export_documents()
        booking_lines = self.get_buchungszeilen()
        for att in attachments:
            for line in booking_lines:
                if line['move_id'] == att.res_id:
                    line['dokument'] = '/export/' + att.name
                    break

        booking_lines = [{key: value for key, value in line.items() if key not in ['buchungszeile', 'move_id']} for line
                         in booking_lines]
        return booking_lines

    # Exports the booking lines
    def export_buchungszeilen(self):
        csvBuffer = io.StringIO()
        fieldnames = ['satzart', 'konto', 'gKonto', 'belegnr', 'belegdatum', 'steuercode', 'buchcode', 'betrag',
                      'prozent', 'steuer', 'text', 'buchsymbol', 'dokument']
        writer = csv.DictWriter(csvBuffer, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in self.add_documents_to_booking_lines():
            cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                           row.items()}
            writer.writerow(cleaned_row)

        return csvBuffer.getvalue()

    # Executes the export
    def execute(self):
        action = {
            'type': 'ir.actions.act_url',
            'url': '/download',
            'target': 'self',
        }
        return action

    # Combines all files to one zip file
    def combine_to_zip(self):
        zip_buffer = io.BytesIO()
        date_form = self.env['account.bmd'].search([])[-1]
        formated_date_from = date_form.period_date_from.strftime('%y%m%d')
        formated_date_to = date_form.period_date_to.strftime('%y%m%d')
        formated_company = date_form.company.name.replace(' ', '_')

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            accountContent = self.export_account()
            zip_file.writestr(
                f'Sachkonten_' + formated_company + '_' + formated_date_from + '_' + formated_date_to + '.csv',
                accountContent)
            customerContent = self.export_customers()
            zip_file.writestr(
                f'Personenkonten_' + formated_company + '_' + formated_date_from + '_' + formated_date_to + '.csv',
                customerContent)
            entryContent = self.export_buchungszeilen()
            zip_file.writestr(
                f'Buchungszeilen_' + formated_company + '_' + formated_date_from + '_' + formated_date_to + '.csv',
                entryContent)
            for att in self.export_documents():
                zip_file.writestr(att.name, base64.b64decode(att.datas))
        zip_buffer.seek(0)

        return zip_buffer
