import unittest
import pyvba


class TestProgram(unittest.TestCase):
    def test_program1(self):
        catia = pyvba.Browser("CATIA.Application")
        shapes = pyvba.Browser(catia.ActiveDocument.Part.Bodies.Item(1)).Shapes
        items = shapes.Item

        print("PartBody structure:")
        for value in items.all.values():
            print(value.Name)


if __name__ == '__main__':
    unittest.main()
