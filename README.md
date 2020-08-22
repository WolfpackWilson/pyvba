# pyvba
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/pyvba)
[![PyPI](https://img.shields.io/pypi/v/pyvba)](https://pypi.org/project/pyvba/)
[![GitHub](https://img.shields.io/github/license/TheEric960/pyvba)](https://github.com/TheEric960/pyvba)

The pyvba package was designed to gather data from VBA-based applications (e.g. Microsoft Excel, CATIA, etc.). It
may also be used to assist programming VBA macro scripts in a more sensical language. 

## Getting Started
Install the Python Package:
```cmd
pip install pyvba
```

To export data from a VBA program:
```python
import pyvba

catia = pyvba.Browser("CATIA.Application")
active_document = catia.ActiveDocument

exporter = pyvba.XMLExport(active_document)
exporter.save("output", r"C:\Documents")
```

The current supported output types are XML and JSON formats.

## Developer Notes
This package is still in alpha. Hence, there are still some problematic bugs and issues that cause errors in certain
applications. Contributors are welcome! The project is [hosted on GitHub](https://github.com/TheEric960/pyvba). Report 
any issues at [the issue tracker](https://github.com/TheEric960/pyvba/issues), but please check to see if the issue 
already exists!
