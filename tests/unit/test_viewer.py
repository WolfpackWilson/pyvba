import unittest
from pyvba import viewer as v

APP = "CATIA.Application"


class TestViewer(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.viewer = v.Viewer(APP)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.viewer = None

    def test_instance(self):
        self.assertIsInstance(self.viewer, v.Viewer)

    def test_com(self):
        self.assertIsNotNone(self.viewer.com)

    def test_name(self):
        self.assertEqual(self.viewer.name, APP)

    def test_parent(self):
        self.assertIsNone(self.viewer.parent)

    def test_objects(self):
        self.assertTrue(len(self.viewer.objects) > 0)

    def test_methods(self):
        self.assertTrue(len(self.viewer.methods) > 0)

    def test_variables(self):
        self.assertTrue(len(self.viewer.variables.keys()) > 0)
        self.assertIsInstance(self.viewer.variables, dict)

    def test_errors(self):
        self.assertTrue(len(self.viewer.errors) >= 0)

    def test_getattr(self):
        self.assertIsInstance(getattr(self.viewer, "ActiveDocument"), v.Viewer)
        self.assertIsInstance(self.viewer.getattr("ActiveDocument"), v.Viewer)
        self.assertIsInstance(self.viewer.ActiveDocument, v.Viewer)

    def test_func(self):
        viewer2 = self.viewer.ActiveDocument.Part.Bodies
        self.assertIsNotNone(viewer2.func("Item", 1))
        self.assertIsNotNone(viewer2.Item(1))

    def test_view(self):
        self.assertIsInstance(self.viewer.view("ActiveDocument"), v.Viewer)


class TestFunctionViewer(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.f_viewer = v.Viewer(APP).ActiveDocument.Save

    @classmethod
    def tearDownClass(cls) -> None:
        cls.f_viewer = None

    def test_instance(self):
        self.assertIsInstance(self.f_viewer, v.FunctionViewer)

    def test_func(self):
        self.assertIsNone(self.f_viewer())

    def test_name(self):
        self.assertEqual(self.f_viewer.name, "Save")

    def test_fullargspec(self):
        self.assertEqual(
            str(self.f_viewer.fullargspec),
            "FullArgSpec(args=['self'], varargs=None, varkw=None, defaults=None, kwonlyargs=[], kwonlydefaults=None, "
            "annotations={})"
        )

    def test_args(self):
        self.assertIsInstance(self.f_viewer.args, list)

    def test_call(self):
        self.assertIsNone(self.f_viewer.call())


class TestIterableFunctionViewer(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.i_viewer = v.Viewer(APP).ActiveDocument.Part.Bodies.Item

    @classmethod
    def tearDownClass(cls) -> None:
        cls.i_viewer = None

    def test_instance(self):
        self.assertIsInstance(self.i_viewer, v.IterableFunctionViewer)

    def test_iter(self):
        items = [item for item in self.i_viewer]
        self.assertTrue(len(items) > 0)

    def test_count(self):
        self.assertEqual(len(self.i_viewer.items), self.i_viewer.count)

    def test_items(self):
        self.assertTrue(len(self.i_viewer.items) > 0)
        self.assertIsInstance(self.i_viewer.items, list)

    def test_item(self):
        self.assertIsNotNone(self.i_viewer.item(0))


if __name__ == '__main__':
    unittest.main()
