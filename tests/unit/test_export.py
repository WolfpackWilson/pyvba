import os
import unittest

from pyvba import export as e


class TestXMLExport(unittest.TestCase):
    # generation will be checked in integration tests

    def test_xml_encode(self):
        self.assertEqual(e.XMLExport.xml_encode("Ed & Eddy"), "Ed &amp; Eddy")


class TestTag(unittest.TestCase):
    def setUp(self) -> None:
        self.tag = e.XMLExport.Tag("Example", lang="English")

    def tearDown(self) -> None:
        self.tag = None

    def test_name(self):
        self.assertEqual(self.tag.name, "Example")

    def test_attrs(self):
        self.assertEqual(len(self.tag.attrs), 1)
        self.assertIsInstance(self.tag.attrs, dict)

    def test_open_tag(self):
        self.assertEqual(self.tag.open_tag, '<Example lang="English">')

    def test_close_tag(self):
        self.assertEqual(self.tag.close_tag, '</Example>')

    def test_format_name(self):
        self.assertEqual(e.XMLExport.Tag.format_name("XML:Example.Name"), "Example.Name")
        self.assertEqual(self.tag.format_name(self.tag.name), "Example")

    def test_enclose(self):
        self.assertEqual(
            self.tag.enclose("This is an example.", 0),
            '<Example lang="English">This is an example.</Example>\n'
        )

    def test_add_attr(self):
        self.tag.add_attr("Count", 1)
        self.assertEqual(len(self.tag.attrs), 2)

    def rm_attr(self):
        self.tag.rm_attr("lang")
        self.assertEqual(len(self.tag.attrs), 0)


class TestJSONExport(unittest.TestCase):
    # generation will be checked in integration tests

    def test_json_encode(self):
        self.assertEqual(e.XMLExport.xml_encode(r"Ed\Eddy"), "Ed\\Eddy")


class TestExport(unittest.TestCase):
    def test_save_as(self):
        if os.path.exists(r".\out\save_test.txt"):
            os.remove(r".\out\save_test.txt")

        e.save_as(
            "This is some text.\n",
            "save_test",
            ".txt",
            r".\out"
        )

        with open(r".\out\save_test.txt", "r") as f:
            self.assertEqual(f.read(), "This is some text.\n")


if __name__ == '__main__':
    unittest.main()
