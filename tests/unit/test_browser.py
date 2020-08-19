import unittest
from pyvba import browser as b

browser = b.Browser('CATIA.Application')


class TestBrowser(unittest.TestCase):
    # generation will be checked in integration tests

    def test_instance(self):
        pass

    def test_skip(self):
        pass

    def test_rm_skip(self):
        pass

    def test_reset(self):
        pass

    def test_reset_all(self):
        pass


class IterableFunctionBrowser(unittest.TestCase):
    def test_all(self):
        pass


if __name__ == '__main__':
    unittest.main()
