from odoo.tests.common import TransactionCase

class TestCustomerExport(TransactionCase):

    def test_export_customers(self):
        self.assertEqual(1 + 1, 2, "1 + 1 sollte 2 ergeben")
        print("Test erfolgreich")
        #Partner = self.env['res.partner']
        #Partner.create({'name': 'Test Kunde', 'email': 'test@example.com', 'phone': '1234567890', 'customer_rank': 1})
        #print("Partner created")

        #self.env['customer.export'].create({}).export_customers()


