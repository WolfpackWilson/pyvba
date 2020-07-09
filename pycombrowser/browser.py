from win32com.client.gencache import EnsureDispatch


class COMBrowser:
    def __init__(self, app: str, **kwargs):
        """Initialize the class from an application string.

        Parameters
        ----------
        app : str
            The application string (e.g. "Excel.Application")
        """

        self._app = EnsureDispatch(app)

        self._methods = [
            i for i in dir(self._app)
            if '_' not in i and i not in ['CLSID', 'coclass_clsid']
        ]

        self._variables = [key for key in getattr(self._app, '_prop_map_put_').keys()]

        self._objects = [
            key for key in getattr(self._app, '_prop_map_get_').keys()
            if key not in self._variables
        ]

        self._obj_filter = ["Application", "Parent"] if kwargs.get('defaults', False) else []

    def __getattr__(self, item):
        return getattr(self._app, item)

