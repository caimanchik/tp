import unittest
from program import Salary, Vacancy, DataSet, HelpMethods


class SalaryTests(unittest.TestCase):

    def setUp(self) -> None:
        self.salary = Salary([10.0, 20.0, 'RUR'])

    def test_salary_type(self):
        self.assertEqual(type(self.salary).__name__, 'Salary')

    def test_salary_from(self):
        self.assertEqual(self.salary._Salary__salary_from, 10.0)

    def test_salary_to(self):
        self.assertEqual(self.salary._Salary__salary_to, 20.0)

    def test_salary_currency(self):
        self.assertEqual(self.salary._Salary__salary_currency, 'RUR')

    def test_int_to_float(self):
        self.assertEqual(float(Salary([10, 20, 'RUR'])), 15.0)

    def test_float_salary_from_to_flaot(self):
        self.assertEqual(float(Salary([10.0, 20, 'RUR'])), 15.0)

    def test_to_float_from_eur(self):
        self.assertEqual(float(Salary([10.0, 30.0, 'EUR'])), 1198.0)


class VacancyTests(unittest.TestCase):

    def setUp(self) -> None:
        self.vacancy = Vacancy(['Руководитель', '<strong>Обязанности:</strong>', 'Организаторские', 'between3And6', 'FALSE', 'ПМЦ Авангард', '80000', '100000', 'FALSE', 'RUR', 'Санкт-Петербург', '2022-07-17T18:23:06+0300'], ['name', 'description', 'key_skills', 'experience_id', 'premium', 'employer_name', 'salary_from', 'salary_to', 'salary_gross', 'salary_currency', 'area_name', 'published_at'])

    def test_vacancy_type(self):
        self.assertEqual(type(self.vacancy).__name__, 'Vacancy')

    def test_vacancy_area(self):
        self.assertEqual(self.vacancy.get_area(), 'Санкт-Петербург')

    def test_salary_from_vacancy(self):
        self.assertEqual(self.vacancy.get_salary(), 90000.0)

    def test_vacancy_year(self):
        self.assertEqual(self.vacancy.get_date(), 2022)

    def test_vacancy_is_suitible(self):
        self.assertTrue(self.vacancy.is_suitible('Руководитель'))

    def test_vacancyis_not_suitible(self):
        self.assertFalse(self.vacancy.is_suitible('Not suitible'))


class TestDataSet(unittest.TestCase):

    def setUp(self) -> None:
        self.dataset = DataSet('test.csv')

    def test_dataset_type(self):
        self.assertEqual(type(self.dataset).__name__, 'DataSet')

    def test_dataset_length(self):
        self.assertEqual(self.dataset._DataSet__len, 1)


class TestHelpMethods(unittest.TestCase):

    def test_delete_rubbish_when_normal(self):
        self.assertEqual(HelpMethods.delete_rubbish('Test string'), 'Test string')

    def test_delete_rubbish_when_tag_only(self):
        self.assertEqual(HelpMethods.delete_rubbish('<strong>Test string</strong>'), 'Test string')

    def test_delete_rubbish_when_tag_and_extra_spaces(self):
        self.assertEqual(HelpMethods.delete_rubbish('     <p>Test    string   </p>'), 'Test string')

    def test_refadctor_label_when_normal(self):
        self.assertEqual(HelpMethods.refactor_label('Питер'), 'Питер')

    def test_refactor_label_when_has_space(self):
        self.assertEqual(HelpMethods.refactor_label('Санкт Петербург'), 'Санкт\nПетербург')

    def test_refactor_label_when_has_defis(self):
        self.assertEqual(HelpMethods.refactor_label('Санкт-Петербург'), 'Санкт-\nПетербург')


if __name__ == '__main__':
    unittest.main()