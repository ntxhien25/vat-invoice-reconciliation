import unittest
from privacy_check import unsafe_text
class PublicPrivacyTests(unittest.TestCase):
    def test_rejects_invoice_url(self):
        self.assertTrue(unsafe_text('https://example.invalid/invoice?code=FICTIONAL'))
    def test_rejects_local_path(self):
        self.assertTrue(unsafe_text('C:/Users/fictional/Documents/source.xlsx'))
    def test_filename_provenance_is_allowed(self):
        self.assertFalse(unsafe_text('Purchases_coded_2025-04.xlsx'))
if __name__=='__main__':unittest.main()
