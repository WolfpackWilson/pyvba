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
        # TODO: implement ignore filter?

        self._all = {}

    def __str__(self):
        return "<class 'COMBrowser'>: " + self._name

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
        """Return a dict"""
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
            if attr in self._skip + self._checked:
                continue

            try:
                obj = getattr(self._com, attr)
            except BaseException as e:
                self._errors[attr] = e.args
                self._all[attr] = e
                return

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

    def skip(self, item: str):
        """Adds a keyword to the skip list."""
        self._skip.append(item)

    def search(self, name):
        """Search for all instances of a method or object."""
        pass

    def reset(self):
        """Clear the `all` property."""
        pass

    def reset_all(self):
        """Clear the `all` property and all filters."""
        pass

    def regenerate(self):
        """Regenerate the `all` property."""
        pass

    def print_all(self):
        """Prints out `all` in a readable way"""
        pass


class IterableFunctionBrowser(IterableFunctionViewer):
    def __init__(self, func, name, count, **kwargs):
        super().__init__(func, name, count, **kwargs)
        self._all = [
            COMBrowser(func(i), name=(kwargs.get('parent_name') + str(i)), **kwargs)
            for i in range(1, count + 1)
        ]

    def __str__(self):
        return super().__str__().replace('FunctionViewer', "FunctionBrowser")

    @property
    def all(self):
        return self._all
