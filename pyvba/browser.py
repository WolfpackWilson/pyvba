from pyvba.viewer import Viewer, FunctionViewer, IterableFunctionViewer


class Browser(Viewer):
    def __init__(self, app, **kwargs):
        """Create a browser from an application string or win32com object.

        The Browser object used to iterate through and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """
        super().__init__(app, **kwargs)

        # define filters
        self._skip = kwargs.get('skip', ["Application", "Parent"])      # user-defined skips
        self._checked = kwargs.get('checked', {})                       # checked items
        self._max_checks = kwargs.get('max_checks', 1)                  # maximum allowable instances of an object

        self._all = {}

    def __str__(self):
        return super().__str__().replace("Viewer", "Browser")

    def __getattr__(self, item):
        if self._all == {}:
            self._generate()

        try:
            obj = self._all[item]
        except AttributeError:
            obj = getattr(self._com, item)
        return obj

    @property
    def all(self) -> dict:
        """Return a dict of objects in the form `{name: item}`."""
        if self._all == {}:
            self._generate()
        return self._all

    def _generate(self):
        """Iterates through all instances of COM objects when called upon.

        See Also
        --------
        Viewer.__getattr___
        """

        for attr in self:
            # skip if checked or user opts to skip it
            if attr in self._skip or self._checked.get(attr, self._max_checks) <= 0:
                continue

            # attempt to collect and observe the attribute
            try:
                obj = getattr(self._com, attr)

                # sort by type
                if '<bound method' in repr(obj):
                    # check for Item array
                    if "Item" == attr:
                        # create the IterableFunctionBrowser
                        # TODO infinite item issue
                        try:
                            count = getattr(self._com, 'Count')
                            self._all[attr] = IterableFunctionBrowser(
                                obj, attr, count,
                                parent_name=self._name,
                                parent=self,
                                skip=self._skip,
                                checked=self._checked,
                                max_checks=self._max_checks
                            )
                        except Exception:
                            self._all[attr] = FunctionViewer(obj, attr)
                    else:
                        self._all[attr] = FunctionViewer(obj, attr)
                elif 'win32com' in repr(obj) or 'COMObject' in repr(obj):
                    self._checked[attr] = self._checked.get(attr, self._max_checks) - 1
                    self._all[attr] = Browser(
                        obj,
                        parent=self,
                        name=attr,
                        skip=self._skip,
                        checked=self._checked,
                        max_checks=self._max_checks
                    )
                else:
                    self._all[attr] = obj
            except BaseException as e:
                self._errors[attr] = e.args
                self._all[attr] = e
                continue

    def print_browser(self, **kwargs):
        """Prints out `all` in a readable way."""
        item = kwargs.get('item', self)
        tabs = kwargs.get('tabs', 0)
        name = ''

        # attach a name if not self describing
        if 'class' not in str(item):
            name = kwargs.get('name') + ': '

        print("  " * tabs + name + str(item))

        # iterate through the corresponding browser
        if isinstance(item, (Browser, IterableFunctionBrowser)):
            for i in list(item.all.keys()):
                self.print_browser(item=item.all[i], name=i, tabs=(tabs + 1))

    def skip(self, item: str):
        """Adds a keyword to the skip list."""
        if item not in self._skip:
            self._skip.append(item)

    def rm_skip(self, item: str):
        """Remove a keyword from the skip list."""
        if item in self._skip:
            self._skip.remove(item)

    def search(self, name: str, exact: bool = False, **kwargs):
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

        item = kwargs.get('item', self)
        paths = []
        path = kwargs.get('path', self._name)

        # add to list if found
        if type(item) == FunctionViewer and name in item.name:
            if not exact or (exact and name == item.name):
                paths.append(path + '/' + item.name)

        # search
        if isinstance(item, (Browser, IterableFunctionBrowser)):
            for i in item.all:
                if name in i:
                    if not exact or (exact and i == name):
                        paths.append(path + '/' + i)

                value = item.all[i]
                if isinstance(value, (Browser, IterableFunctionBrowser)):
                    paths += self.search(name, exact, item=value, path=(path + '/' + i))

        return paths

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
        item = self
        path = path.split('/')

        for loc in path[1:]:
            if isinstance(item, IterableFunctionBrowser):
                item = item(int(loc))
            else:
                item = getattr(item, loc)
        return item

    def reset(self):
        """Clear the `all` property and the checked list."""
        self._checked = {}
        self._all = {}

    def reset_all(self):
        """Clear the `all` property and all empties the skip list."""
        self._skip = []
        self.reset()

    def regen(self):
        """Regenerate the `all` property."""
        self.reset()
        self._generate()


class IterableFunctionBrowser(IterableFunctionViewer):
    def __init__(self, func, name, count, **kwargs):
        """Create a browser for an iterable function to view its components."""
        super().__init__(func, name, count, **kwargs)

        self._all = {
            str(i): Browser(func(i), name=(kwargs.get('parent_name')), **kwargs)
            for i in range(1, count + 1)
        }

    def __str__(self):
        return super().__str__().replace('FunctionViewer', "FunctionBrowser")

    def __getattr__(self, item):
        if self._all == {}:
            self._generate()

        try:
            obj = self._all[item]
        except AttributeError:
            obj = getattr(self._com, item)
        return obj

    @property
    def all(self):
        """Return a dict of objects in the form `{name: item}`."""
        return self._all
