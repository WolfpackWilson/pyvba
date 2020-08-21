import unittest
from pyvba import browser as b

APP = 'CATIA.Application'


class TestBrowser(unittest.TestCase):
    # generation will be checked in integration tests

    @classmethod
    def setUpClass(cls) -> None:
        cls.browser = b.Browser(APP)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.browser = None

    def test_instance(self):
        self.assertIsInstance(self.browser, b.Browser)

    def test_skip(self):
        self.browser.skip("Parent")
        self.assertEqual(len(self.browser._skip), 2)

        self.browser.skip("ActiveDocument")
        self.assertTrue("ActiveDocument" in self.browser._skip)

        self.browser.rm_skip("ActiveDocument")
        self.assertTrue("ActiveDocument" not in self.browser._skip)

    def test_reset(self):
        self.browser.reset()
        self.assertEqual(self.browser._checked, {})
        self.assertEqual(self.browser._all, {})

    def test_reset_all(self):
        self.browser.skip("ActiveDocument")
        self.browser.reset_all()

        self.assertEqual(self.browser._checked, {})
        self.assertEqual(self.browser._all,{})
        self.assertEqual(len(self.browser._skip), 0)

        self.browser.skip("Parent")
        self.browser.skip("Application")


class IterableFunctionBrowser(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.i_browser = b.Browser(APP).ActiveDocument.Part.Bodies.Item

    @classmethod
    def tearDownClass(cls) -> None:
        cls.i_browser = None

    def test_instance(self):
        self.assertIsInstance(self.i_browser, b.IterableFunctionBrowser)

    def test_all(self):
        self.assertTrue(len(self.i_browser.all) > 0)


if __name__ == '__main__':
    unittest.main()
