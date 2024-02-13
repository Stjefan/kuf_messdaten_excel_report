import unittest
from src.kuf_messdaten_excel_report import create_monatsbericht_immendingen
from dotenv import load_dotenv
import os
load_dotenv()
class TestAddFunction(unittest.TestCase):
    def test_create(self):
        create_monatsbericht_immendingen(2024, 1)
        self.assertEqual(1+2, 3)


if __name__ == '__main__':
    unittest.main()