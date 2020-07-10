from win32com.client.gencache import EnsureDispatch
from inspect import getfullargspec


class COMViewer:
    def __init__(self, app, **kwargs):
        """Initialize the class from an application string or win32com object.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """

        self._com = EnsureDispatch(app)
        self._parent = kwargs.get('parent', None)
        self._kwargs = kwargs
        self._objects = [key for key in getattr(self._com, '_prop_map_get_').keys()]
        self._methods = [
            i for i in dir(self._com)
            if '_' not in i and i not in ['CLSID', 'coclass_clsid']
        ]
        self._errors = {}

    def __getattr__(self, item):
        try:
            return getattr(self._com, item)
        except BaseException as e:
            self._errors[item] = e.args
            return e

    def __iter__(self):
        return (str(obj) for obj in self._objects + self._methods)

    @property
    def com(self):
        """Return the COM object."""
        return self._com

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
            return COMViewer(obj, parent=self._com)
        else:
            return obj


class FunctionViewer:
    def __init__(self, func, name: str = None):
        self._func = func
        self._name = name
        self._fullargspec = getfullargspec(func)
        self._args = self._fullargspec.args

    def __call__(self, *args, **kwargs):
        return self._func(*args, **kwargs)

    def __str__(self):
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
