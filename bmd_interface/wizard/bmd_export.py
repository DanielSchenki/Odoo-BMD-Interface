import base64
import csv
import io
import re
import zipfile
import math
import time

from odoo.exceptions import ValidationError
from odoo.http import request
from datetime import datetime

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
            f'attachment; filename="BMD_Export_{sanitize_filename(formatted_company)}_{formatted_date_from}_{formatted_date_to}.zip"')])
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

def sanitize_filename(filename):
    if '.' in filename:
        name, ext = filename.rsplit('.', 1)
        name = name.replace('.', '_')
        sanitized_name = f"{name}.{ext}"
    else:
        # Handle filenames without an extension
        sanitized_name = filename.replace('.', '_')
    return re.sub(r'[<>:" /\\|?*]', '_', sanitized_name)

def replace_dot_with_comma(value):
    return str(value).replace('.',',')





class AccountBmdExport(models.TransientModel):
    _name = 'account.bmd'
    _description = 'BMD Export'

    period_date_from = fields.Date(string="Von:", required=True)
    period_date_to = fields.Date(string="Bis:", required=True)
    company = fields.Many2one('res.company', string="Mandant:", required=True)
    documents = fields.Boolean(string="Dokumente ausgeben?", default=True)

    start_time = fields.Float(string="Startzeit", default=time.time())
    checkpointNr = fields.Integer(string="Checkpoint", default=0)




    # Prints a checkpoint
    @api.model
    def checkpoint(self, text):
        record = self.env['account.bmd'].search([], limit=1, order='id desc')
        if not record:
            return
        end_time = datetime.now()
        elapsed_time = (end_time - datetime.fromtimestamp(record.start_time)).total_seconds()
        minutes, seconds = divmod(elapsed_time, 60)
        seconds, milliseconds = divmod(seconds, 1)
        print(f"CheckpointNr: {record.checkpointNr} {text}, Elapsed time: {int(minutes)} minutes, {int(seconds)} seconds and {int(milliseconds * 1000)} milliseconds")
        record.checkpointNr += 1
        record.write({'checkpointNr': record.checkpointNr, 'start_time': record.start_time})



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
        pattern = r"BMDSC\d{3}$"

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
                    if re.search(pattern, tax.name):
                        steuercode = int(tax.name[-3:])
                    else:
                        steuercode = "2"

                    result_data.append({
                        'Konto-Nr': acc.code,
                        'Bezeichnung': acc.name,
                        'Ustcode': steuercode,
                        'USTPz': int(tax.amount),
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

        journal_items = self.env['account.move.line'].search([])
        date_form = self.env['account.bmd'].search([])[-1]
        account_codes = []

        journal_items = sorted(journal_items, key=lambda x: x.move_id.id)

        for line in journal_items:

            if line.company_id.id != date_form.company.id:
                continue

            belegdatum = line.date
            if date_form.period_date_from > belegdatum or belegdatum > date_form.period_date_to:
                continue

            if len(str(line.account_id.code))>5:
                account_codes.append(line.account_id.code)



        #customers = self.env['res.partner'].search([])
        customers = self.env['res.partner'].search([
            '|',
            ('property_account_receivable_id.code', 'in', account_codes),
            ('property_account_payable_id.code', 'in', account_codes)
        ])
        date_form = self.env['account.bmd'].search([])[-1]

        # Write to the CSV file
        customer_buffer = io.StringIO()
        fieldnames = ['Konto-Nr', 'Name', 'Straße', 'PLZ', 'Ort', 'Land', 'UID-Nummer', 'E-Mail', 'Webseite', 'Phone',
                      'IBAN', 'Zahlungsziel', 'Skonto', 'Skontotage']
        writer = csv.DictWriter(customer_buffer, fieldnames=fieldnames, delimiter=';')
        writer.writeheader()
        for customer in customers:
            ###if customer.company_id.id != date_form.company.id:
            ###    continue
            if customer.property_account_receivable_id.code in account_codes:
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
            if customer.property_account_payable_id.code in account_codes:
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
        buchsymbol_mapping_AR_ER = {'out_invoice': 'AR', 'in_invoice': 'ER', 'out_refund': 'GU', 'in_refund': 'EG'}
        journal_items = self.env['account.move.line'].search([])
        date_form = self.env['account.bmd'].search([])[-1]
        result_data = []
        docs = []

        journal_items = sorted(journal_items, key=lambda x: x.move_id.id)
        move_ids = set()

        #self.checkpoint("Start journal_items loop")
        for line in journal_items:

            if line.company_id.id != date_form.company.id:
                continue

            belegdatum = line.date
            if date_form.period_date_from > belegdatum or belegdatum > date_form.period_date_to:
                continue

            #general fields
            belegdatum = date_formatter(belegdatum)
            konto = line.account_id.code
            prozent = line.tax_ids.amount
            steuer = line.price_total - line.price_subtotal
            belegnr = line.move_id.name
            text = line.name
            buchsymbol = buchsymbol_mapping.get(line.journal_id.type, '')
            if (steuer != 0) and (line.tax_ids.name is not False):
                steuercode_before_cut = line.tax_ids.name
                pattern = r"BMDSC\d{3}$"
                if re.search(pattern, steuercode_before_cut):
                    steuercode = int(steuercode_before_cut[-3:])
                else:
                    steuercode = "2"
            else:
                steuercode = "2"

            satzart = 0
            dokument = ''
            fwbetrag = ''
            fwsteuer = ''
            waehrung = ''
            move_id = line.move_id.id
            additional_documents = []

            attachments = self.env['ir.attachment'].search([])
            for att in attachments:
                if move_id == att.res_id and dokument == '' and att.id not in docs:
                    dokument = att.name
                    docs.append(att.id)
                elif move_id == att.res_id and att.id not in docs:
                    additional_documents.append({'document': att.name})
                    docs.append(att.id)

            #special logic for ARs
            if buchsymbol == 'AR':
                buchsymbol = buchsymbol_mapping_AR_ER.get(line.move_id.move_type, '')
                if line.display_type != 'product':
                    continue
                gkonto = line.partner_id.property_account_receivable_id.code
                temp_gkonto = gkonto
                gkonto = konto
                konto = temp_gkonto
                buchcode = '1'
                betrag = line.price_total
                if buchsymbol == 'GU':
                    betrag = -betrag
                else:
                    steuer = -steuer

            #special logic for ERs
            if buchsymbol == 'ER':
                buchsymbol = buchsymbol_mapping_AR_ER.get(line.move_id.move_type, '')
                if line.display_type != 'product':
                    continue
                gkonto = line.partner_id.property_account_payable_id.code
                temp_gkonto = gkonto
                gkonto = konto
                konto = temp_gkonto
                buchcode = '2'
                betrag = line.price_total
                if buchsymbol == 'EG':
                    steuer = -steuer
                else:
                    betrag = -betrag

            #special logic for foreign currency ERs
            if buchsymbol == 'ER' and line.move_id.currency_id != line.company_id.currency_id:
                fwbetrag = -line.amount_currency
                betrag = line.move_id.amount_total_signed
                steuer = -line.move_id.amount_tax_signed
                fwsteuer = line.move_id.amount_tax
                waehrung = line.move_id.currency_id.name

            #special logic for BKs and KAs
            if buchsymbol == 'BK' or buchsymbol == 'KA':
                if line.matching_number == False:
                    continue
                else:
                    konto = line.account_id.code
                    gkonto = line.payment_id.outstanding_account_id.code

            #special logic for SOs
            if buchsymbol == 'SO':
                if line.move_id.id in move_ids:
                    continue
                if line.debit > 0:
                    buchcode = 1
                    haben_buchung = False
                else:
                    buchcode = 2
                    haben_buchung = True

                if haben_buchung:
                    betrag = -line.credit  # Haben Buchungen müssen negativ sein
                else:
                    betrag = line.debit

                for invoice_line in line.move_id.invoice_line_ids:
                    if line.move_id.id == invoice_line.move_id.id and konto != invoice_line.account_id.code:
                        gkonto = invoice_line.account_id.code

                move_ids.add(line.move_id.id)

            # Rounding
            betrag = commercial_round_3_digits(betrag)
            steuer = commercial_round_3_digits(steuer)

            result_data.append(
                {'satzart': satzart, 'konto': konto, 'gKonto': gkonto, 'belegnr': belegnr, 'belegdatum': belegdatum,
                 'steuercode': steuercode, 'buchcode': buchcode, 'betrag': replace_dot_with_comma(betrag),
                 'prozent': replace_dot_with_comma(prozent),
                 'steuer': replace_dot_with_comma(steuer), 'fwbetrag': replace_dot_with_comma(fwbetrag),
                 'fwsteuer': replace_dot_with_comma(fwsteuer), 'waehrung': waehrung, 'text': text, 'buchsymbol': buchsymbol,
                 'move_id': move_id, 'dokument': sanitize_filename(dokument)})

            if additional_documents:
                for doc in additional_documents:
                    result_data.append({
                        'satzart': '5', 'konto': '', 'gKonto': '', 'belegnr': '', 'belegdatum': '', 'steuercode': '',
                        'buchcode': '', 'betrag': '', 'prozent': '', 'steuer': '', 'fwbetrag': '', 'fwsteuer': '', 'waehrung': '', 'text': '', 'buchsymbol': '',
                        'move_id': '', 'dokument': sanitize_filename(doc['document'])})

        #self.checkpoint("Finished journal_items loop")
        # Remove tax lines and haben buchung
        #return_data = []
        #seen_move_ids = set()
        #for data in result_data:
        #    if data['move_id'] not in seen_move_ids and (
        #            data['buchsymbol'] in ('ER', 'AR', 'GU')):
        #        return_data.append(data)
        #        seen_move_ids.add(data['move_id'])
        #    elif data['move_id'] == '':
        #        return_data.append(data)

        return_data = result_data
        return return_data

    # Exports the documents
    def export_attachments(self):
        date_form = self.env['account.bmd'].search([])[-1]
        journal_items = self.env['account.move.line'].search([
            ('company_id.id', '=', date_form.company.id),
            ('date', '>=', date_form.period_date_from),
            ('date', '<=', date_form.period_date_to)
        ])
        move_ids = [item.move_id.id for item in journal_items]
        attachments = self.env['ir.attachment'].search([('res_id', 'in', move_ids), ('res_model', '=', 'account.move')])
        return attachments

    # Exports the booking lines
    def export_account_movements(self):
        csvBuffer = io.StringIO()
        fieldnames = ['satzart', 'konto', 'gKonto', 'belegnr', 'belegdatum', 'steuercode', 'buchcode', 'betrag',
                      'prozent', 'steuer', 'fwbetrag', 'fwsteuer', 'waehrung', 'text', 'buchsymbol', 'dokument']
        writer = csv.DictWriter(csvBuffer, fieldnames=fieldnames, delimiter=';')

        writer.writeheader()
        for row in self.get_account_movements():
            del row['move_id']
            cleaned_row = {key: value.replace('\n', ' ') if isinstance(value, str) else value for key, value in
                           row.items()}
            writer.writerow(cleaned_row)

        return csvBuffer.getvalue()

    # Executes the export
    def execute(self):
        #reset checkpoint logger
        record = self.env['account.bmd'].search([], limit=1, order='id desc')
        record.start_time = time.time()
        record.checkpointNr = 0
        record.write({'checkpointNr': record.checkpointNr, 'start_time': record.start_time})
        self.checkpoint("Start export")

        action = {'type': 'ir.actions.act_url', 'url': '/download', 'target': 'self', }
        return action

    # Combines all files to one zip file
    def combine_to_zip(self):
        self.checkpoint("Start combining to zip")
        zip_buffer = io.BytesIO()
        date_form = self.env['account.bmd'].search([])[-1]
        formatted_date_from = date_form.period_date_from.strftime('%y%m%d')
        formatted_date_to = date_form.period_date_to.strftime('%y%m%d')
        formatted_company = date_form.company.name.replace(' ', '_')

        print("Company: " + formatted_company)
        print("Date from: " + formatted_date_from)
        print("Date to: " + formatted_date_to)

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            self.checkpoint("Start Sachkonten")
            accountContent = self.export_accounts()
            if accountContent is None:
                accountContent = ""
                print("Keine Sachkonten vorhanden")
            else:
                print("Sachkonten: " + accountContent)
            self.checkpoint("Finished Sachkonten")
            zip_file.writestr(sanitize_filename(f'Sachkonten_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv'),
                              accountContent)
            self.checkpoint("Sachkonten written to zip, start Personenkonten")


            customerContent = self.export_customers()
            if customerContent is None:
                customerContent = ""
                print("Keine Personenkonten vorhanden")
            else:
                print("Personenkonten: " + customerContent)
            self.checkpoint("Finished Personenkonten")

            zip_file.writestr(sanitize_filename(f'Personenkonten_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv'),
                              customerContent)
            self.checkpoint("Personenkonten written to zip, start Buchungszeilen")


            entryContent = self.export_account_movements()
            self.checkpoint("Finished Buchungszeilen")
            zip_file.writestr(sanitize_filename(f'Buchungszeilen_{formatted_company}_{formatted_date_from}_{formatted_date_to}.csv'),
                              entryContent)
            self.checkpoint("Buchungszeilen written to zip, start Dokumente")
            if date_form.documents is True:
                for att in self.export_attachments():
                    zip_file.writestr(sanitize_filename(att.name), base64.b64decode(att.datas))
            self.checkpoint("Dokumente written to zip, start closing")
        zip_buffer.seek(0)
        self.checkpoint("Finished combining to zip")
        return zip_buffer
