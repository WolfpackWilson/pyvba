import unittest
from pyvba import export


class TestXMLExport(unittest.TestCase):
    # generation will be checked in integration tests

    def test_xml_encode(self):
        pass

class TestTag(unittest.TestCase):
    def test_name(self):
        pass

    def test_attrs(self):
        pass

    def test_open_tag(self):
        pass

    def test_close_tag(self):
        pass

    def test_format_name(self):
        pass

    def test_enclose(self):
        pass

    def test_add_attr(self):
        pass

    def rm_attr(self):
        pass


class TestJSONExport(unittest.TestCase):
    def test_json_encode(self):
        pass


class TestExport(unittest.TestCase):
    def test_save_as(self):
        pass


if __name__ == '__main__':
    unittest.main()
