from pycombrowser.viewer import COMViewer, FunctionViewer, IterableFunctionViewer


class COMBrowser(COMViewer):
    def __init__(self, app, **kwargs):
        """Create a browser from an application string or win32com object.

        The COMBrowser object used to iterate through and explore COM objects from external
        applications.

        Parameters
        ----------
        app
            The application string (e.g. "Excel.Application") or win32com object.
        """
        super().__init__(app, **kwargs)

        # define filters
        self._skip = kwargs.get('skip', [])
        self._checked = kwargs.get('checked', [])

        self._all = {}

    def __str__(self):
        return super().__str__().replace("COMViewer", "COMBrowser")

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
        COMViewer.__getattr___
        """

        for attr in self:
            # skip if checked or user opts to skip it
            if attr in self._skip + self._checked:
                continue

            # attempt to collect the attribute
            try:
                obj = getattr(self._com, attr)
            except BaseException as e:
                self._errors[attr] = e.args
                self._all[attr] = e
                return

            # sort by type
            if '<bound method' in str(obj):
                # check for Item array
                if "Item" == attr:
                    try:
                        count = getattr(self._com, 'Count')
                        self._all[attr] = IterableFunctionBrowser(
                            obj, attr, count,
                            parent_name=self._name,
                            parent=self,
                            skip=self._skip,
                            checked=self._checked
                        )
                    except (AttributeError, NameError):
                        self._all[attr] = FunctionViewer(obj, attr)
                else:
                    self._all[attr] = FunctionViewer(obj, attr)
            elif 'win32com' in str(obj) or 'COMObject' in str(obj):
                self._checked.append(attr)
                self._all[attr] = COMBrowser(
                    obj,
                    parent=self,
                    name=attr,
                    skip=self._skip,
                    checked=self._checked,
                )
            else:
                self._all[attr] = obj

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
        if isinstance(item, (COMBrowser, IterableFunctionBrowser)):
            for i in list(item.all.keys()):
                self.print_browser(item=item.all[i], name=i, tabs=(tabs + 1))

    def skip(self, item: str):
        """Adds a keyword to the skip list."""
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
        if isinstance(item, (COMBrowser, IterableFunctionBrowser)):
            for i in item.all:
                if name in i:
                    if not exact or (exact and i == name):
                        paths.append(path + '/' + i)

                value = item.all[i]
                if isinstance(value, (COMBrowser, IterableFunctionBrowser)):
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
        self._checked = []
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
            str(i): COMBrowser(func(i), name=(kwargs.get('parent_name') + '_' + str(i)), **kwargs)
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
