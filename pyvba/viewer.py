from inspect import getfullargspec

from win32com.client.gencache import EnsureDispatch


class Viewer:
    def __init__(self, app, **kwargs):
        """Create a viewer from an application string or win32com object.

        The Viewer object used to observe and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """

        self._com = EnsureDispatch(app)
        self._name = kwargs.get('name', repr(app)[1:-1])
        self._parent = kwargs.get('parent', None)
        self._kwargs = kwargs
        self._objects = [key for key in getattr(self._com, '_prop_map_get_').keys()]
        self._methods = [
            i for i in dir(self._com)
            if '_' not in i and i not in ['CLSID', 'coclass_clsid']
        ]
        self._errors = {}

    def __getattr__(self, item):
        """Return a variable, FunctionViewer, or Browser object."""
        try:
            obj = getattr(self._com, item)
        except BaseException as e:
            self._errors[item] = e.args
            return e

        if '<bound method' in repr(obj):
            if "Item" in item:
                try:
                    count = getattr(self._com, 'Count')
                    return IterableFunctionViewer(obj, item, count)
                except AttributeError:
                    return FunctionViewer(obj, item)
            else:
                return FunctionViewer(obj, item)
        elif 'win32com' in repr(obj) or 'COMObject' in repr(obj):
            return Viewer(obj, parent=self._com, name=item)
        else:
            return obj

    def __iter__(self):
        """Return iteration of the combined names of the objects and methods."""
        return (str(obj) for obj in self._objects + self._methods)

    def __str__(self):
        return "<class 'Viewer'>: " + self._name

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
        """Return a dictionary in format {name: value}."""
        variables = {}

        for key in self._objects:
            if not isinstance(self.view(key), (Viewer, BaseException)):
                variables[key] = self.getattr(key)
        return variables

    @property
    def errors(self):
        """Return a dictionary in format {obj: Error}."""
        return self._errors

    def getattr(self, item):
        """Alternative to `Viewer.Attribute`. Return a variable, object, or method."""
        return self.view(item)

    def func(self, name, *args):
        """Alternative to `Viewer.Attribute(*args)`. Runs a function based on arguments given and returns
        the result."""
        return getattr(self, name)(*args)

    def view(self, attr):
        """Alternative to `Viewer.Attribute`. Return a variable, FunctionViewer, or Browser object."""
        return getattr(self, attr)


class FunctionViewer:
    def __init__(self, func, name: str = None, **kwargs):
        """Create a viewer from a stored function.

        A viewer object used to observe and run functions extracted from the Viewer.

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
        """Return a string of the class and how to use the function."""
        name = "func_name" if self._name is None else self._name
        args = ""

        for arg in self._args[1:]:
            args += arg + ', '
        return f"<class 'FunctionViewer'>: {name}({args[:-2]})"

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
    def args(self) -> list:
        """Return the function arguments."""
        return self._args[1:]

    def call(self, *args, **kwargs):
        """Alternative function call."""
        return self(*args, **kwargs)


class IterableFunctionViewer(FunctionViewer):
    def __init__(self, func, name, count, **kwargs):
        """Create a viewer from a stored iterable function.

        Parameters
        ----------
        func
            A bound method.
        name : str, optional
            The name of the bound method.
        count : int
            The number of items held.
        """
        super().__init__(func, name, **kwargs)
        self._count = count
        self._items = [
            Viewer(func(i), name=str(i), **kwargs)
            for i in range(1, count + 1)
        ]

    def __str__(self):
        """Return a string of the class and how to use the function."""
        return super().__str__().replace("FunctionViewer", "IterableFunctionViewer")

    def __iter__(self):
        """Iterate through each item in the iterable function."""
        return (item for item in self._items)

    @property
    def count(self) -> int:
        """Return the count of the iterable function."""
        return self._count

    @property
    def items(self) -> list:
        """Return the items of the iterable function."""
        return self._items

    def item(self, i: int):
        """Return a specific item of the iterable function."""
        return self._items[i]
