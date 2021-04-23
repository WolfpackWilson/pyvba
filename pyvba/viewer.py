import os
import re
import shutil
import sys
from inspect import getfullargspec

from win32com.client.gencache import EnsureDispatch

# define regular expressions
class_re = re.compile(r"(?<=\.)[^.]+?(?='>)")


class Viewer:
    def __init__(self, app, name: str = None, parent: object = None):
        """Create a viewer from an application string or win32com object.

        The Viewer object used to observe and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application"), Viewer, or win32com object.
        name: str
            The name of the object. It will generate automatically unless string is given.
        parent: object
            The parent object, if applicable.
        """

        self._com = self.ensure_dispatch(app) if not isinstance(app, Viewer) else app
        self._name = name if name else self._com.Name
        self._type = class_re.findall(str(self._com.__class__))[0]
        self._parent = parent

        self._objects = [key for key in getattr(self._com, '_prop_map_get_').keys()]
        self._methods = [
            FunctionViewer(getattr(self._com, i), i)
            for i in dir(self._com)
            if '_' not in i and i not in ['CLSID', 'Item']
        ]

        self._errors = {}

    def __getattr__(self, item):
        return self.getattr(item)

    def __str__(self):
        return "<class 'Viewer'>: " + self._name

    @staticmethod
    def ensure_dispatch(com):
        """Ensures the COM object is generated and retrieved.

        Sometimes the cache needs to be cleared. In this case, an attribute error is thrown and caught.
        """
        try:
            app = EnsureDispatch(com)
        except (AttributeError, TypeError):
            # Remove cache and try again.
            module_list = [m.__name__ for m in sys.modules.values()]
            for module in module_list:
                if re.match(r'win32com\.gen_py\..+', module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
            app = EnsureDispatch(com)
        return app

    @staticmethod
    def gettype(obj, item: str = None, parent: object = None):
        """Return the appropriate variable or Viewer instance."""
        if '<bound method' in repr(obj):
            return FunctionViewer(obj, item)
        elif 'win32com' in repr(obj) or 'COMObject' in repr(obj):
            try:
                _ = len(obj)
                return CollectionViewer(obj, item, parent)
            except (TypeError, AttributeError):
                return Viewer(obj, item, parent)
        return obj

    def getattr(self, item):
        """Return a variable, FunctionViewer, or Viewer object."""
        try:
            obj = getattr(self._com, item)
        except (AttributeError, KeyboardInterrupt):
            raise
        except BaseException as e:
            self._errors[item] = e
            return e

        return self.gettype(obj, item)

    def cf(self, other) -> bool:
        """Comparison alternative to __eq__.

        The comparison avoids checking any Viewer instances and compares the values of the standard objects within.
        """
        if type(self) != type(other):
            return False
        return self._type == other.type and self._name == other.name and self._objects == other.objects

    @property
    def com(self):
        """Return the COM object."""
        return self._com

    @property
    def name(self) -> str:
        """Return the name of the COM object."""
        return self._name

    @property
    def type(self) -> str:
        """Return the type of the object within the COM object."""
        return self._type

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
    def errors(self) -> dict:
        """Return a dictionary in format {obj: Error}."""
        return self._errors

    def view(self, attr):
        """Alternative to `Viewer.Attribute`. Return a variable, FunctionViewer, or Browser object."""
        return self.getattr(attr)


class FunctionViewer:
    def __init__(self, func, name: str):
        """Create a viewer from a stored function.

        A viewer object used to observe and run functions extracted from the Viewer.

        Parameters
        ----------
        func
            A bound method.
        """

        self._func = func
        self._name = name
        self._fullargspec = getfullargspec(func)
        self._args = self._fullargspec.args

    def __call__(self, *args, **kwargs):
        """Calls the function and returns the function output."""
        return Viewer.gettype(self._func(*args, **kwargs))

    def __str__(self):
        """Return a string of the class and how to use the function."""
        args = ""

        for arg in self._args[1:]:
            args += arg + ', '
        return f"<class 'FunctionViewer'>: {self._name}({args[:-2]})"

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


class CollectionViewer(Viewer):
    def __init__(self, obj, name: str = None, parent: object = None):
        super().__init__(obj, name, parent)

        self._count = len(self._com)
        self._items = [
            Viewer.gettype(i, name, self)
            for i in self._com
        ]

    def __str__(self):
        return super().__str__().replace('Viewer', 'CollectionViewer')

    def __len__(self):
        return self._count

    def __getitem__(self, item):
        return self._items[item]

    @property
    def count(self) -> int:
        """Return the number of items in the collection."""
        return self._count

    @property
    def items(self) -> list:
        """Return the items in the collection."""
        return self._items

    def item(self, index):
        """Return one of the items."""
        return self[index]
