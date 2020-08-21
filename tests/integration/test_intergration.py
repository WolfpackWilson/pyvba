import io
import os
import unittest
from contextlib import redirect_stdout

import pyvba

APP = "CATIA.Application"
PRINT = False


def redirect_print(func, *args, **kwargs):
    out = io.StringIO()
    with redirect_stdout(out):
        func(*args, **kwargs)

    if PRINT:
        print(out.getvalue())
    return out.getvalue()


class TestBrowser(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.b = pyvba.Browser(APP)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.b = None

    def test_all(self):
        self.assertTrue(len(self.b.all) > 0)
        self.assertIsInstance(self.b.all, dict)

    def test_print_browser(self):
        output = redirect_print(self.b.print_browser)
        self.assertTrue(len(output) > 100)

    def test_search_and_goto(self):
        results = self.b.search("Part", True)
        self.assertTrue(len(results) > 0)

        part = self.b.goto(results[0])
        self.assertIsInstance(part, pyvba.Browser)

    def test_regen(self):
        self.b.regen()
        self.assertNotEqual(self.b._checked, {})
        self.assertNotEqual(self.b._all, {})


class TestXMLExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.b = pyvba.Browser(APP).ActiveDocument
        cls.e = pyvba.XMLExport(cls.b)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.e = None
        cls.b = None

    def test_string(self):
        self.assertTrue(len(self.e.string) > 100)

    def test_min(self):
        self.assertTrue(len(self.e.min) > 100)
        self.assertEqual(self.e.min.count("\n"), 0)
        self.assertEqual(self.e.min.count("\t"), 0)

    def test_print(self):
        output = redirect_print(self.e.print)
        self.assertTrue(len(output) > 100)

    def test_save(self):
        if os.path.exists(r".\out\xml_test.xml"):
            os.remove(r".\out\xml_test.xml")

        self.e.save("xml_test", r".\out")

        with open(r".\out\xml_test.xml", "r") as f:
            self.assertTrue(len(f.read()) > 100)


class TestJSONExport(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.b = pyvba.Browser(APP).ActiveDocument
        cls.e = pyvba.JSONExport(cls.b)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.e = None
        cls.b = None

    def test_string(self):
        self.assertTrue(len(self.e.string) > 100)

    def test_min(self):
        self.assertTrue(len(self.e.min) > 100)
        self.assertEqual(self.e.min.count("\n"), 0)
        self.assertEqual(self.e.min.count("\t"), 0)

    def test_print(self):
        output = redirect_print(self.e.print)
        self.assertTrue(len(output) > 100)

    def test_save(self):
        if os.path.exists(r".\out\json_test.json"):
            os.remove(r".\out\json_test.json")

        self.e.save("json_test", r".\out")

        with open(r".\out\json_test.json", "r") as f:
            self.assertTrue(len(f.read()) > 100)


if __name__ == '__main__':
    unittest.main()
