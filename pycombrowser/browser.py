from win32com.client.gencache import EnsureDispatch


class COMBrowser:
    # TODO: implement __iter__, find(), and error handling
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

    def __getattr__(self, item):
        return getattr(self._com, item)

    def __iter__(self):
        pass

    @property
    def com(self):
        return self._com

    @property
    def parent(self):
        return self._parent

    @property
    def objects(self) -> list:
        return self._objects

    @property
    def methods(self) -> list:
        return self._methods

    @property
    def variables(self) -> dict:
        """Return a dictionary in format {name: value}"""
        variables = {}

        for key in self._objects:
            if not isinstance(self.browse(key), COMBrowser):
                variables[key] = self.getattr(key)
        return variables

    def getattr(self, item):
        """Return a variable, object, or method."""
        return getattr(self._com, item)

    def func(self, name, *args):
        """Runs a function based on arguments given and returns the result."""
        return getattr(self._com, name)(*args)

    def browse(self, attr):
        """Return a variable, method, or COMBrowser object."""
        obj = getattr(self._com, attr)
        if '<bound method' in str(obj):
            return obj
        elif 'win32com' in str(obj) or 'COMObject' in str(obj):
            return COMBrowser(obj, parent=self._com)
        else:
            return obj



