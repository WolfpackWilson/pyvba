from pyvba.viewer import Viewer, FunctionViewer, CollectionViewer

# used to skip predetermined objects by exact name
# fixme: remove extra parameters
skip = ['Application', 'Parent', 'Units', 'Parameters']


# TODO: fix infinite recursion issue
class Browser(Viewer):
    def __init__(self, app, name: str, parent: Viewer = None):
        """Create a browser from an application string or win32com object.

        The Browser object used to iterate through and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """
        super().__init__(app, name, parent)
        self._all = {}

    def __str__(self):
        return super().__str__().replace('Viewer', 'Browser')

    def __getattr__(self, item):
        if self._all == {}:
            self._generate()

        obj = super().getattr(item)
        return obj if isinstance(obj, FunctionViewer) or not isinstance(obj, Viewer) else self.from_viewer(obj)

    @staticmethod
    def from_viewer(viewer, parent=None):
        """Turn a Viewer object into a Browser object."""
        return CollectionBrowser(viewer) if isinstance(viewer, CollectionViewer) \
            else Browser(viewer.com, viewer.name, viewer.parent if parent is None else parent)

    @staticmethod
    def skip(item: str):
        """Adds a keyword to the skip list."""
        global skip
        if item not in skip:
            skip.append(item)

    @staticmethod
    def rm_skip(item: str):
        """Remove a keyword from the skip list."""
        global skip
        if item in skip:
            skip.remove(item)

    @staticmethod
    def clr_skip():
        """Resets the skip list to its original state."""
        global skip
        skip = ['Application', 'Parent']

    @property
    def all(self) -> dict:
        """Return a dict of objects in the form `{name: item}`."""
        if self._all == {}:
            self._generate()
        return self._all

    def _generate(self):
        """Iterates through all objects when called upon."""
        global skip

        for name in self._objects + [i.name for i in self._methods]:
            if name in skip:
                continue

            try:
                obj = super().getattr(name)

                if isinstance(obj, Viewer):
                    self._all[name] = self.from_viewer(obj, self)
                else:
                    self._all[name] = obj
            except BaseException as e:
                self._errors[name] = e.args
                continue

    def search(self, name: str, exact: bool = False):
        """Return a dictionary in format {path: item} matching the name.

        Search for all instances of a method or object containing a name.

        Parameters
        ----------
        exact: bool
            A flag that searches for exact matches.
        name: str
            The name of the attribute to search for.

        Returns
        -------
        dict
            The results of the search in format {path: item}.
        """
        ...

    def goto(self, path: str):
        """Retrieve an object at a given location.

        Parameters
        ----------
        path: str
            The location of the item delimited by '/'.

        Examples
        --------
        `goto('Bodies/Item/1/HybridShapes/GetItem')` yields the 'GetItem' function.

        """
        ...

    def regen(self):
        """Regenerate the `all` property."""
        self._all = {}
        self._generate()


class CollectionBrowser(Browser, CollectionViewer):
    def __init__(self, obj: CollectionViewer):
        super().__init__(obj.com, obj.name, obj.parent)
        self._items = [self.from_viewer(i) for i in self._items]

    def __str__(self):
        return "<class 'CollectionBrowser'>: " + self._name

    def _generate(self):
        super()._generate()
        self._all['Item'] = self._items
