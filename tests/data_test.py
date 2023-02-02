import unittest
from src.correct_data import *
from src.utils import options

class TestDataFuctions(unittest.TestCase):
    def test_correct_legacy_id(self):
        expected = correct_legacy_id(12345678901)
        id_with_w = correct_legacy_id("W1234567890")
        short_id = correct_legacy_id(12345678)
        self.assertEqual(len(expected), 10)
        self.assertEqual(len(id_with_w), 9)
        self.assertEqual(len(short_id), 8)
        self.assertIsInstance(expected, str)


    def test_correct_name(self):
        expected = correct_name("$John")
        self.assertTrue(expected, "John")

    
    def test_clean_phone(self):
        expected = clean_phone("(812)555-5647")
        self.assertTrue(expected, "8125555647")

    
    def test_correct_race(self):
        expected = correct_race("American Indian/Alaskan Native")
        self.assertTrue(expected, options['race']['a'])


    def test_correct_ethnicity(self):
        expected = correct_ethnicity('a. hispanic')
        self.assertTrue(expected, options['ethnicity']['a'])

    
    def test_correct_eng_prof(self):
        expected = correct_eng_prof('a. Household is Limited English Proficient')
        self.assertTrue(expected, options['englishProficiency']['a'])


    def test_correct_case_type(self):
        expected = correct_case_type('Home Purchase')
        self.assertTrue(expected, options['caseType']['a'])


    def test_correct_rural(self):
        expected = correct_rural('Household lives in a rural area')
        self.assertTrue(expected, options['RuralAreaStatus']['a'])


    def test_correct_household(self):
        expected = correct_household('Female-Single Parent')
        self.assertTrue(expected, options['HouseholdType']['b'])

if __name__ == 'main':
    unittest.main()