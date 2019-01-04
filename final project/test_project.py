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

    # def test_build_one_shift(self):
    #     result = almost_final_project.build_one_shift(1, 1)
    #     self.assertIn('tair', result)

    def test_make_shifts_for_shift_manager(self):
        # result = project.make_shifts_for_shift_manager([['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])
        self.assertIn('michal', project_for_test.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])
        self.assertIn('emilia', project_for_test.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])

# check this test again!
    def test_Daily_Money_amount(self):
        self.assertNotEqual(260, project_for_test.Daily_Money_amount(2018, 12, 27))
       # self.assertEqual(219.8, project_for_test.Daily_Money_amount(2019, 1, 3))

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


if __name__ == '__test_project__':
    unittest.main()

