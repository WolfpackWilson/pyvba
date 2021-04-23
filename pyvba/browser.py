from pyvba.viewer import Viewer, FunctionViewer, CollectionViewer
from collections import OrderedDict

# used to skip predetermined objects by exact name
skip = ['Application', 'Parent']

# store a dictionary of the discovered items
visited = OrderedDict()


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
    def clr_found():
        """Clears the stored dictionary of items browsed."""
        global visited
        visited = OrderedDict()

    @staticmethod
    def skip(*item: str):
        """Add one or more keywords to the skip list."""
        global skip
        for i in item:
            if i not in skip:
                skip.append(i)

    @staticmethod
    def rm_skip(*item: str):
        """Remove one or more keywords from the skip list."""
        global skip
        for i in item:
            if i not in skip:
                skip.remove(i)

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
        global skip, visited

        # iterate through items
        for name in self._objects + [i.name for i in self._methods]:
            if name in skip:
                continue

            try:
                obj = super().getattr(name)

                if isinstance(obj, Viewer):
                    self._all[name] = self.from_viewer(obj, self)
                else:
                    self._all[name] = obj
            except KeyboardInterrupt:
                raise
            except BaseException as e:
                self._errors[name] = e.args
                continue

        # add items to the visited dictionary
        for name, value in self._all.items():
            if isinstance(value, Viewer):
                if value.type not in visited:
                    visited[value.type] = []

                if not any(map(lambda item: value.cf(item), visited[value.type])):
                    visited[value.type].append(value)

    def browse_all(self):
        """Populate the browser and all descendents of the browser."""
        if self._all == {}:
            self._generate()

        # populate child browsers if not already visited
        for name, value in self._all.items():
            if isinstance(value, Browser) and not any(map(lambda item: value.cf(item), visited[value.type])):
                name.browse_all()

    def cf(self, other) -> bool:
        """Comparison alternative to __eq__.

        The comparison avoids checking any Viewer instances and compares the values of the standard objects within.
        """
        return super().cf(other) and all([
            a == b
            for a, b in zip(self._all.values(), other.all.values())
            if not isinstance(a, (Viewer, FunctionViewer, list)) and not isinstance(b, (Viewer, FunctionViewer, list))
        ])

    def regen(self):
        """Regenerate the `all` property."""
        self._all = {}
        self._generate()


class CollectionBrowser(Browser, CollectionViewer):
    def __init__(self, obj):
        super().__init__(obj.com, obj.name, obj.parent)
        self._items = [self.from_viewer(item) for item in self._items]

    def __str__(self):
        return "<class 'CollectionBrowser'>: " + self._name

    def _generate(self):
        global visited
        super()._generate()
        self._all['Item'] = self._items

        # generate the Item list
        if len(self._items) > 0 and isinstance(self._items[0], Browser):
            _ = [item.all for item in self._items]

        # add items to the visited dictionary
        for value in self._all['Item']:
            if isinstance(value, Viewer):
                if value.type not in visited:
                    visited[value.type] = []

                if value not in visited[value.type]:
                    visited[value.type].append(value)

