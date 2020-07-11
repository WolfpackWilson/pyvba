import pycombrowser.viewer as viewer


class COMBrowser(viewer.COMViewer):
    def __init__(self, app, **kwargs):
        super().__init__(app, **kwargs)

        # define filters
        self._skip = kwargs.get('skip', [])
        self._checked = kwargs.get('checked', [])
        self._funcs = kwargs.get('funcs', [])   # TODO: implement function exploration

        self._all = {}
        self._generate()    # TODO: do later and add __iter__?

    def __str__(self):
        return "<class 'COMBrowser'>: " + self._name

    @property
    def all(self):
        return self._all

    def _generate(self):
        for attr in self:
            if attr in self._checked:
                continue

            obj = getattr(self, attr)

            if '<bound method' in str(obj):
                self._all[attr] = viewer.FunctionViewer(obj, attr)
            elif 'win32com' in str(obj) or 'COMObject' in str(obj):
                self._checked.append(attr)
                self._all[attr] = COMBrowser(obj, parent=self._com, name=attr, checked=self._checked)
            else:
                self._all[attr] = obj

    def search(self):
        pass
