import unittest
import io
import project_for_test
# import unittest.mock


class TestProject(unittest.TestCase):
    def test_max_val(self):
        self.assertEqual(5, project_for_test.max_val([1, 2, 3, 4, 5]))
        self.assertIsNot(4, project_for_test.max_val([1, 2, 3, 4, 5]))

    def test_check_validation_of_product_code(self):
        self.assertIs(False, project_for_test.check_validation_of_product_code(1234))
        self.assertIs(False, project_for_test.check_validation_of_product_code(666))
        self.assertIs(True, project_for_test.check_validation_of_product_code(1))

    def test_build_one_shift(self):
        result = project_for_test.build_one_shift(1, 1)
        self.assertIn('tair', result)
        result = project_for_test.build_one_shift(1, 2)
        self.assertIn('tair', result)
        result = project_for_test.build_one_shift(2, 7)
        self.assertIn('asaf', result)

    def test_make_shifts_for_shift_manager(self):
        self.assertIn('michal', project_for_test.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])
        self.assertIn('emilia', project_for_test.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])

    def test_Daily_Money_amount(self):
        self.assertNotEqual(260, project_for_test.Daily_Money_amount('2018-12-27'))
        self.assertEqual(1216.89, project_for_test.Daily_Money_amount('2019-01-04'))

    def test_GetPrice(self):
        self.assertEqual(119.9, project_for_test.GetPrice(2))
        self.assertEqual(99.9, project_for_test.GetPrice(1))

    def test_GetName(self):
        self.assertEqual('T-shirts', project_for_test.GetName(1))
        self.assertIsNot('hat', project_for_test.GetName(1))

    def test_check_recipect_number_validation(self):
        self.assertIs(True, project_for_test.check_recipect_number_validation(1))

    def test_get_total_price_of_recipect(self):
        self.assertEqual(320.76, project_for_test.get_total_price_of_recipect(4))

    def test_The_number_of_next_recipct(self):
        self.assertEqual(8, project_for_test.The_number_of_next_recipct())
        self.assertNotEqual(9, project_for_test.The_number_of_next_recipct())

    def test_get_recipect_date(self):
        self.assertEqual('2019-01-02', project_for_test.get_recipect_date(1))
        self.assertEqual('2019-01-02', project_for_test.get_recipect_date(2))
        self.assertNotEqual('2019-01-31', project_for_test.get_recipect_date(1))

    def test_make_shift_by_random(self):
        self.assertEqual(['no one can', 'no one can'], project_for_test.make_shift_by_random([]))
        self.assertEqual(['yoni', 'no one can'], project_for_test.make_shift_by_random(['yoni']))
        self.assertIn('tair', project_for_test.make_shift_by_random(['tair', 'asaf']))
        self.assertNotEqual(['no one can', 'no one can'], project_for_test.make_shift_by_random(['yoni']))

    def test_count_shift_for_worker(self):
        self.assertEqual({'adir': 4, 'stav': 3, 'tair': 5, 'yoni': 2, 'asaf': 4, 'rotem': 6},
                         project_for_test.count_shift_for_worker())
        self.assertIsNot({'adir': 4, 'stav': 3, 'tair': 5, 'yoni': 2},
                         project_for_test.count_shift_for_worker())

    def test_find_custumer(self):
        self.assertEqual(True, project_for_test.find_custumer(123321123))
        self.assertEqual(False, project_for_test.find_custumer(12331123))
        self.assertNotEqual(True, project_for_test.find_custumer(3321123))

    def test_check_if_customer_is_member_club(self):
        self.assertEqual(True, project_for_test.check_if_customer_is_member_club('242532654'))
        self.assertEqual(True, project_for_test.check_if_customer_is_member_club('123456789'))
        self.assertNotEqual(True, project_for_test.check_if_customer_is_member_club('42532654'))
        self.assertNotEqual(True, project_for_test.check_if_customer_is_member_club(''))


if __name__ == '__test_project__':
    unittest.main()

