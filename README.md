# pyvba
![PyPI - Python Version](https://img.shields.io/pypi/pyversions/pyvba)
[![PyPI](https://img.shields.io/pypi/v/pyvba)](https://pypi.org/project/pyvba/)
[![GitHub](https://img.shields.io/github/license/TheEric960/pyvba)](https://github.com/TheEric960/pyvba)

The pyvba package was designed to gather data from VBA-based applications (e.g. Microsoft Excel, CATIA, etc.). It may also be used to assist programming VBA macro scripts using the Python language. 

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

The current supported output types are XML and JSON formats. Both support a form the imitates the VBA object tree as well as a dictionary form where each unique object is in the outermost layer.

Example Output:
> Note: `BrowserObject` denotes an object defined elsewhere in the output.
```JSON
{ "MainBody": [
	{ "Pad": [
		{ "Shapes": [
			{ "DirectionOrientation": 0 },
			{ "DirectionType": 0 },
			{ "FirstLimit": "BrowserObject" },
			{ "IsSymmetric": false },
			{ "IsThin": false },
			{ "MergeEnd": false },
			{ "Name": "Pad.1" },
			{ "NeutralFiber": false },
			{ "SecondLimit": "BrowserObject" },
			{ "Sketch": "BrowserObject" }
		]},
		{ "Shapes": [
			{ "DirectionOrientation": 0 },
			{ "DirectionType": 0 },
			{ "FirstLimit": "BrowserObject" },
			{ "IsSymmetric": false },
			{ "IsThin": false },
			{ "MergeEnd": false },
			{ "Name": "Pad.1" },
			{ "NeutralFiber": false },
			{ "SecondLimit": "BrowserObject" },
			{ "Sketch": "BrowserObject" }
		]},
		{ "Shapes": [
			{ "DirectionOrientation": 0 },
			{ "DirectionType": 0 },
			{ "FirstLimit": "BrowserObject" },
			{ "IsSymmetric": false },
			{ "IsThin": false },
			{ "MergeEnd": false },
			{ "Name": "Pad.3" },
			{ "NeutralFiber": false },
			{ "SecondLimit": "BrowserObject" },
			{ "Sketch": "BrowserObject" }
		]}
	]}
]}
```


## Developer Notes
This package is in beta. Therefore, there are still some problematic bugs and issues that cause errors in certain applications. Contributors are welcome! The project is [hosted on GitHub](https://github.com/TheEric960/pyvba). Report any issues at [the issue tracker](https://github.com/TheEric960/pyvba/issues), but please check to see if the issue already exists!
