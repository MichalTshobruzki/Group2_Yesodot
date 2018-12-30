import unittest
import io
import storeProject
import unittest.mock


class TestProject(unittest.TestCase):

    def test_max_val(self):
        result = storeProject.max_val('walk')
        TestDict = {'ID': 1, 'TodoItem': 'walk', 'isDone': False}
        self.assertIn(TestDict, result)
        todolist.RemoveItem(1)



if __name__=='__main__':
     unittest.main()
