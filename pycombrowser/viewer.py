from win32com.client.gencache import EnsureDispatch
from inspect import getfullargspec


class COMViewer:
    def __init__(self, app, **kwargs):
        """Create a viewer from an application string or win32com object.

        The COMViewer object used to observe and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """

        self._com = EnsureDispatch(app)
        self._name = kwargs.get('name', app)
        self._parent = kwargs.get('parent', None)
        self._kwargs = kwargs
        self._objects = [key for key in getattr(self._com, '_prop_map_get_').keys()]
        self._methods = [
            i for i in dir(self._com)
            if '_' not in i and i not in ['CLSID', 'coclass_clsid']
        ]
        self._errors = {}

    def __getattr__(self, item):
        """Return the attribute or an error.

        Parameters
        ----------
        item
            The attribute to search for.

        Notes
        -----
        If an issue occurred trying to find the attribute, an error is returned.
        """

        try:
            return getattr(self._com, item)
        except BaseException as e:
            self._errors[item] = e.args
            return e

    def __iter__(self):
        """Return iteration of the combined names of the objects and methods."""
        return (str(obj) for obj in self._objects + self._methods)

    def __str__(self):
        return "<class 'COMViewer'>: " + self._name

    @property
    def com(self):
        """Return the COM object."""
        return self._com

    @property
    def name(self):
        """Return the name of the COM object."""
        return self._name

    @property
    def parent(self):
        """Return the parent COM object."""
        return self._parent

    @property
    def objects(self) -> list:
        """Return a list of the objects."""
        return self._objects

    @property
    def methods(self) -> list:
        """Return a list of the methods"""
        return self._methods

    @property
    def variables(self) -> dict:
        """Return a dictionary in format {name: value}"""
        variables = {}

        for key in self._objects:
            if not isinstance(self.view(key), (COMViewer, BaseException)):
                variables[key] = self.getattr(key)
        return variables

    @property
    def errors(self):
        """Return a dictionary in format {obj: Error}"""
        return self._errors

    def getattr(self, item):
        """Return a variable, object, or method."""
        return getattr(self, item)

    def func(self, name, *args):
        """Runs a function based on arguments given and returns the result."""
        return getattr(self, name)(*args)

    def view(self, attr):
        """Return a variable, FunctionViewer, or COMBrowser object."""
        obj = getattr(self, attr)

        if '<bound method' in str(obj):
            return FunctionViewer(obj, attr)
        elif 'win32com' in str(obj) or 'COMObject' in str(obj):
            return COMViewer(obj, parent=self._com, name=attr)
        else:
            return obj


class FunctionViewer:
    def __init__(self, func, name: str = None):
        """Create a viewer from a stored function.

        A viewer object used to observe and run functions extracted from the COMViewer.

        Parameters
        ----------
        func
            A bound method.
        name : str, optional
            The name of the bound method.
        """

        self._func = func
        self._name = name
        self._fullargspec = getfullargspec(func)
        self._args = self._fullargspec.args

    def __call__(self, *args, **kwargs):
        """Calls the function and returns the function output."""
        return self._func(*args, **kwargs)

    def __str__(self):
        """Returns a string of the class and how to use the function.

        Notes
        -----
        The `self` argument is typically included as the first parameter and should be ignored when
        calling the function.
        """
        name = "func_name" if self._name is None else self._name
        args = ""

        for arg in self._args:
            args += arg + ', '
        return "<class 'FunctionViewer'>: {}({})".format(name, args[:-2])

    @property
    def func(self):
        """Return the function instance."""
        return self._func

    @property
    def name(self) -> str:
        """Return the function name."""
        return self._name

    @property
    def fullargspec(self):
        """Return the inspect.fullargspec object."""
        return self._fullargspec

    @property
    def args(self):
        """Return the function arguments."""
        return self._args

    def call(self, *args, **kwargs):
        """Alternative function call."""
        return self.__call__(*args, **kwargs)
