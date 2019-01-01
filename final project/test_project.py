import unittest
import io
import almost_final_project
# import unittest.mock


class TestProject(unittest.TestCase):
    def test_max_val(self):
        self.assertEqual(5, project.max_val([1, 2, 3, 4, 5]))
        self.assertIsNot(4, project.max_val([1, 2, 3, 4, 5]))

    def test_check_validation_of_product_code(self):
        self.assertIs(False, project.check_validation_of_product_code(1234))
        self.assertIs(False, project.check_validation_of_product_code(666))
        self.assertIs(True, project.check_validation_of_product_code(1))

    def test_build_one_shift(self):
        result = project.build_one_shift(1, 1)
        self.assertIn('tair', result)

    def test_make_shifts_for_shift_manager(self):
        # result = project.make_shifts_for_shift_manager([['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])
        self.assertIn('michal', project.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])
        self.assertIn('emilia', project.make_shifts_for_shift_manager(
            [['michal', [1, 1], [2, 6]], ['emilia', [1, 6], [2, 1]]])[0])

    def test_Daily_Money_amount(self):
        self.assertEqual(260, project.Daily_Money_amount(2018, 12, 20))
        self.assertEqual(430, project.Daily_Money_amount(2017, 12, 20))

    def test_GetPrice(self):
        self.assertEqual(200, project.GetPrice(2))
        self.assertEqual(89.9, project.GetPrice(1))

    def test_GetName(self):
        self.assertEqual('shirts', project.GetName(1))
        self.assertIsNot('hat', project.GetName(1))

# not working
    def test_check_recipect_number_validation(self):
        self.assertIs(True, project.check_recipect_number_validation(1))
    #
    # def test_get_recipect_date(self):
    #


if __name__ == '__test_project__':
    unittest.main()

