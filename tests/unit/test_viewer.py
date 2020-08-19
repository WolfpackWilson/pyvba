import unittest
from pyvba import viewer as v

viewer = v.Viewer('CATIA.Application')


class TestViewer(unittest.TestCase):
    def test_instance(self):
        self.assertIsInstance(viewer, v.Viewer)

    def test_com(self):
        self.assertIsNotNone(viewer.com)

    def test_name(self):
        self.assertEqual(viewer.name, "CATIA.Application")

    def test_parent(self):
        self.assertIsNone(viewer.parent)

    def test_objects(self):
        self.assertTrue(len(viewer.objects) > 0)

    def test_methods(self):
        self.assertTrue(len(viewer.methods) > 0)

    def test_variables(self):
        self.assertTrue(len(viewer.variables.keys()) > 0)
        self.assertIsInstance(viewer.variables, dict)

    def test_errors(self):
        self.assertTrue(len(viewer.errors) >= 0)

    def test_getattr(self):
        self.assertIsInstance(getattr(viewer, "ActiveDocument"), v.Viewer)
        self.assertIsInstance(viewer.getattr("ActiveDocument"), v.Viewer)
        self.assertIsInstance(viewer.ActiveDocument, v.Viewer)

        # note that these objects don't refer to the same object
        self.assertIsNot(viewer.ActiveDocument, getattr(viewer, "ActiveDocument"))

    def test_func(self):
        viewer2 = viewer.ActiveDocument.Part.Bodies
        self.assertIsNotNone(viewer2.func("Item", 1))
        self.assertIsNotNone(viewer2.Item(1))

    def test_view(self):
        self.assertIsInstance(viewer.view("ActiveDocument"), v.Viewer)


class TestFunctionViewer(unittest.TestCase):
    def test_instance(self):
        pass

    def test_func(self):
        pass

    def test_name(self):
        pass

    def test_fullargspec(self):
        pass

    def test_args(self):
        pass

    def test_call(self):
        pass


class TestIterableFunctionViewer(unittest.TestCase):
    def test_instance(self):
        pass

    def test_iter(self):
        pass

    def test_count(self):
        pass

    def test_items(self):
        pass

    def test_item(self):
        pass


if __name__ == '__main__':
    unittest.main()
