import os
import re
import copy

from win32com.universal import com_error

from pyvba.browser import Browser, visited
from pyvba.viewer import FunctionViewer


class ExportStr:
    def __init__(self, browser: Browser, skip_func: bool = False, skip_err: bool = False, vba_form: bool = False):
        """The base class for exporting

        Parameters
        ----------
        browser: Browser
            The object used to gather all variables.
        skip_func: bool
            Skips reporting any FunctionViewer instant.
        skip_err: bool
            Skips reporting any error.
        vba_form: bool
            A flag that determines if the output mimics the VBA tree structure or a more list-like view.
        """
        self._browser = browser
        self._data = None

        self._skip_func = skip_func
        self._skip_err = skip_err
        self._vba_form = vba_form

    @property
    def data_str(self) -> str:
        """Return the data in string format."""
        self._check()
        return self._data

    @property
    def data_min(self) -> str:
        """Return the data in a minimized string format.

        The minimized version removes all newlines and tabs.
        """
        self._check()
        return re.sub(r'\n*\t*', '', self._data)

    def _check(self):
        """Check if the string needs to be generated."""
        if self._data is None:
            self._generate_vba() if self._vba_form else self._generate_dict()

    def _generate_vba(self, *args, **kwargs):
        """Begin generating the string based on the VBA tree."""
        pass

    def _generate_dict(self, *args, **kwargs):
        """Begin generating the string based on the browser.visited dictionary."""
        pass

    def save_as(self, name: str, ext: str, path: str = '.\\', minimize: bool = False):
        """Save a string object to a specified name and location.

        Parameters
        ----------
        name: str
            The name of the file.
        ext: str
            The file extension (e.g. .xml, .json, etc.)
        path: str
            The save location.
        minimize: bool
            A flag that determines if the data is returned in a minimized string format.
        """
        os.makedirs(path, exist_ok=True)
        with open(os.path.join(path, name + ext), "w") as file:
            file.write(self.data_str if not minimize else self.data_min)
            file.close()

    def print(self, minimize: bool = False):
        """Print the string in the normal or minimized version.

        Parameters
        ----------
        minimize: bool
            A flag that determines if the data is returned in a minimized string format.
        """
        self._check()
        print(self.data_str if not minimize else self.data_min)


class XMLExport(ExportStr):
    XML_ESCAPE_CHARS = {
        "&": "&amp;",
        '"': "&quot;",
        "'": "&apos;",
        ">": "&gt;",
        "<": "&lt;",
    }

    def __init__(self, browser: Browser, version=1.0, encoding: str = "UTF-8", skip_func: bool = False,
                 skip_err: bool = False, vba_form: bool = False):
        """Create a well-formed XML string for export.

        Parameters
        ----------
        browser: Browser
            The object used to gather all variables.
        version
            The current version of the XML.
        encoding: str
            The encoding type (default is UTF-8).
        """
        super().__init__(browser, skip_func, skip_err, vba_form)

        self._xml_head = f'<?xml version="{str(version)}" encoding="{encoding}"?>\n'

    @staticmethod
    def xml_encode(text: str) -> str:
        """Map special XML characters to their encoded form in a given string."""
        return "".join(XMLExport.XML_ESCAPE_CHARS.get(c, c) for c in str(text))

    def _generate_vba(self):
        """Begin generating the XML string based on the VBA tree."""
        self._data = self._xml_head + self._generate_tag(self._browser)

        # convert empty elements to a single tag
        self._data = re.sub(r'></.*?>', ' />', self._data)

    def _generate_tag(self, elem, tabs: int = 0, **kwargs) -> str:
        """Recursively generate each element into a string.

        Parameters
        ----------
        elem
            The element to convert into an XML string.
        tabs: int
            The indentation level of the current element.

        Returns
        -------
        str
            The XML string of the element and sub-elements.
        """

        xml = ''
        stack = kwargs.get('stack', [])

        if isinstance(elem, Browser):
            tag = XMLExport.Tag(elem.name)

            # check if in stack already
            if any(map(lambda obj: elem.cf(obj), stack)):
                return tag.enclose('BrowserObject: See ancestors', tabs)
            else:
                stack.append(elem)

            # setup the tag attributes
            attrs = ["Name", "Count"]
            [
                tag.add_attr(attr, value)
                for attr, value in elem.all.items()
                if attr in attrs
            ]

            # add the element and start adding the sub-elements
            xml += '\t' * tabs + tag.open_tag + '\n'
            for item, value in elem.all.items():
                if isinstance(value, list):
                    item_tag = XMLExport.Tag("Item")

                    xml += '\t' * (tabs + 1) + item_tag.open_tag + '\n'
                    for i in value:
                        xml += self._generate_tag(i, tabs + 2, stack=stack)
                    xml += '\t' * (tabs + 1) + item_tag.close_tag + '\n'

                elif item not in attrs:
                    # overlook objects that point to themselves
                    if item == elem.name:
                        continue
                    else:
                        xml += self._generate_tag(value, tabs + 1, name=item, stack=stack)

            xml += '\t' * tabs + tag.close_tag + '\n'

        elif isinstance(elem, FunctionViewer):
            if not self._skip_func:
                # display the function and its properties
                tag = XMLExport.Tag("Function", name=elem.name, args=len(elem.args))
                xml += tag.enclose(str(elem)[26:], tabs)

        elif isinstance(elem, com_error):
            # display the error location and method
            if not self._skip_err:
                tag = XMLExport.Tag("Error")
                xml += tag.enclose(self.xml_encode(str(elem)), tabs)

        else:
            # display the variable and value
            tag = XMLExport.Tag(kwargs.get('name', 'Unknown'))
            xml += tag.enclose(self.xml_encode(str(elem)), tabs)
        return xml

    def _generate_dict(self):
        """Begin generating the XML string based on the visited dictionary."""
        # populate browser and copy visited
        self._browser.browse_all()
        visited2 = copy.copy(visited)

        tag = XMLExport.Tag(self._browser.name, count=len(visited2))
        xml = self._xml_head + tag.open_tag + "\n"

        # iterate through dictionary
        for var, value in visited2.items():
            tag1 = XMLExport.Tag(var, count=len(value))
            xml += "\t" + tag1.open_tag + "\n"

            # iterate through each list
            for item in value:
                tag2 = XMLExport.Tag(item.name)
                xml += "\t" * 2 + tag2.open_tag + "\n"

                # add name attribute
                if 'Name' in item.all:
                    tag2.add_attr('Name', item.Name)

                # iterate through each browser in the list
                for var2, value2 in item.all.items():
                    tag3 = XMLExport.Tag(var2)

                    # add name attribute
                    if isinstance(value2, Browser) and 'Name' in value2.all:
                        tag3.add_attr('Name', value2.Name)

                    # check for a collection object
                    if isinstance(value2, list):
                        tag3.add_attr('count', len(value2))
                        xml += "\t" * 3 + tag3.open_tag + "\n"

                        # iterate through the browser's collection
                        for item2 in value2:
                            tag4 = XMLExport.Tag(item2.name if isinstance(item2, Browser) else item2)

                            # add name attribute
                            if isinstance(item2, Browser) and 'Name' in item2.all:
                                tag4.add_attr('Name', item2.Name)

                            xml += tag4.enclose(item2.name if isinstance(item2, Browser) else item2, 4)

                        xml += "\t" * 3 + tag3.close_tag + "\n"
                    else:
                        if isinstance(value2, Browser):
                            output = 'BrowserObject'
                        elif isinstance(value2, com_error):
                            if self._skip_err:
                                continue
                            output = self.xml_encode(repr(value2))
                        elif isinstance(value2, FunctionViewer):
                            if self._skip_func:
                                continue
                            tag3 = XMLExport.Tag("Function", name=value2.name, args=len(value2.args))
                            output = str(value2)[26:]
                        else:
                            output = self.xml_encode(value2)
                        xml += tag3.enclose(output, 3)

                xml += "\t" * 2 + tag1.close_tag + "\n"

            xml += "\t" + tag1.close_tag + "\n"

        self._data = xml + tag.close_tag

    def save(self, name: str, path: str = '.\\', minimize: bool = False):
        """Save to a file."""
        super().save_as(name, '.xml', path, minimize)

    class Tag:
        NAME_RE = re.compile(r'(^xml)|(^[0-9]*)', re.IGNORECASE)

        def __init__(self, tag_name: str, **attrs):
            """Create and store XML tag information in the proper formatting.

            Parameters
            ----------
            tag_name: str
                The name that will be displayed.
            attrs
                The attributes to add.
            """
            self._name = self.format_name(tag_name)
            self._attrs = {
                self.format_name(key): value
                for key, value in attrs.items()
            }

        @property
        def name(self) -> str:
            """Return the name of the tag."""
            return self._name

        @property
        def attrs(self) -> dict:
            """Return a dictionary of the tag attributes in the form {attr: value}."""
            return self._attrs

        @property
        def open_tag(self) -> str:
            """Return the formatted opening tag."""
            tag = "<" + self._name
            if len(self._attrs) > 0:
                # add the attributes
                tag += " " + " ".join(
                    f'{key}="{XMLExport.xml_encode(value)}"'
                    for key, value in self._attrs.items()
                    if not isinstance(value, com_error)
                )
            return tag + ">"

        @property
        def close_tag(self) -> str:
            """Return the formatted closing tag."""
            return f"</{self._name}>"

        @staticmethod
        def format_name(text: str) -> str:
            """Return a string formatted to XML tag naming conventions."""
            text = XMLExport.Tag.NAME_RE.sub('', text)
            return text.strip('!"#$%&\'()*+,-./:;<=>?@[\\]^_`{|}~')

        def enclose(self, text: str, tabs: int):
            """Return a string enclosed with the tag."""
            return "\t" * tabs + self.open_tag + text + self.close_tag + "\n"

        def add_attr(self, attr: str, value):
            """Add an attribute to the tag."""
            attr = self.format_name(attr)
            self._attrs[attr] = value

        def rm_attr(self, attr):
            """Remove and return a tag attribute."""
            attr = self.format_name(attr)
            return self._attrs.pop(attr)


class JSONExport(ExportStr):
    JSON_ESCAPE_CHARS = ["\b", "\f", "\n", "\r", "\t", "\"", "\\"]

    def __init__(self, browser: Browser, skip_func: bool = False, skip_err: bool = False, vba_form: bool = False):
        super(JSONExport, self).__init__(browser, skip_func, skip_err, vba_form)

    @staticmethod
    def json_encode(text: str) -> str:
        """Map special JSON characters to their encoded form in a given string."""
        return "".join(
            "\\" + c
            if c in JSONExport.JSON_ESCAPE_CHARS else c
            for c in str(text)
        )

    def _check(self):
        """Check if the JSON string needs to be generated."""
        if self._data is None:
            self._data = self._generate_vba(self._browser) if self._vba_form else self._generate_dict()
            self._data = re.sub(r',(?!\s*?[{\[\"\'\w])', '', self._data)

    def _generate_vba(self, elem, tabs: int = 0, **kwargs) -> str:
        """Recursively generate each element into a string.

        Parameters
        ----------
        elem
            The element to convert into an JSON string.
        tabs: int
            The indentation level of the current element.

        Returns
        -------
        str
            The JSON string of the element and sub-elements.
        """

        json = ''
        stack = kwargs.get('stack', [])

        if isinstance(elem, Browser):
            # check if in stack already
            if any(map(lambda obj: elem.cf(obj), stack)):
                return "\t" * tabs + f"{{ \"{self.json_encode(elem.name)}\": \"BrowserObject: See ancestors\" }},\n"
            else:
                stack.append(elem)

            # display the browser and its children
            json += "\t" * tabs + f"{{ \"{self.json_encode(elem.name)}\": [\n"

            for item, value in elem.all.items():
                if type(value) is list and len(value) > 0:
                    json += "\t" * (tabs + 1) + "{ \"Item\": [\n"
                    for i in value:
                        json += self._generate_vba(i, tabs + 2, stack=stack)
                    json += "\t" * (tabs + 1) + "]},\n"
                else:
                    json += self._generate_vba(value, tabs + 1, name=item, stack=stack)

            json += "\t" * tabs + "]},\n"
        elif isinstance(elem, FunctionViewer):
            if not self._skip_func:
                # display the function and its properties
                json += "\t" * tabs + f"{{ \"{self.json_encode(elem.name)}\": [\n"
                json += "\t" * (tabs + 1) + f"{{ \"name\": \"{self.json_encode(elem.name)}\" }},\n"
                json += "\t" * (tabs + 1) + f"{{ \"args\": {self.json_encode(str(len(elem.args)))} }},\n"
                json += "\t" * (tabs + 1) + f"{{ \"use\": \"{self.json_encode(str(elem)[26:])}\" }}\n"
                json += "\t" * tabs + "]},\n"
        elif isinstance(elem, com_error):
            # display the error location and method
            if not self._skip_err:
                json += "\t" * tabs + "{ \"Error\": [\n"
                try:
                    json += "\t" * (tabs + 1) + f"{{ \"on\": \"{self.json_encode(str(elem.args[2][1]))}\" }},\n"
                    json += "\t" * (tabs + 1) + f"{{ \"message\": \"{self.json_encode(str(elem.args[2][2]))}\" }}\n"
                except TypeError:
                    json += "\t" * (tabs + 1) + f"{{ \"message\": \"{self.json_encode(str(elem.args[2]))}\" }}\n"
                except IndexError:
                    json += "\t" * (tabs + 1) + f"{{ \"message\": \"{self.json_encode(str(elem))}\" }}\n"
                json += "\t" * tabs + "]},\n"
        else:
            # display the variable and value
            name = self.json_encode(kwargs.get('name', 'Unknown'))

            if isinstance(elem, bool):
                elem = str(elem).lower()
            elif not isinstance(elem, (int, float, complex)):
                elem = f"\"{self.json_encode(str(elem))}\""

            json += "\t" * tabs + f"{{ \"{name}\": {elem} }},\n"
        return json

    def _generate_dict(self) -> str:
        """Begin generating the JSON string based on the visited dictionary."""

        # populate browser and copy visited
        self._browser.browse_all()
        visited2 = copy.copy(visited)
        json = f'{{ "{self._browser.name}": [\n'

        # iterate through dictionary items
        for var, value in visited2.items():
            json += f'\t{{ "{var}": [\n'

            # iterate through each list
            for item in value:
                json += f'\t\t{{ "{item.name}": [\n'

                # iterate through each browser in the list
                for var2, value2 in item.all.items():
                    # check for a collection object
                    if isinstance(value2, list):
                        json += f'\t\t\t{{ "{var2}": [\n'

                        # iterate through the browser's collection
                        for item2 in value2:
                            output = item2.name if isinstance(item2, Browser) else item2
                            json += f'\t\t\t\t{{ "{output}": "BrowserObject" }},\n'

                        json += '\t\t\t]},\n'
                    else:
                        if isinstance(value2, Browser):
                            json += f'\t\t\t{{ "{value2.name}": "BrowserObject" }},\n'

                        elif isinstance(value2, com_error):
                            if self._skip_err:
                                continue

                            json += "\t\t\t{ \"Error\": [\n"
                            try:
                                json += f"\t\t\t\t{{ \"on\": \"{self.json_encode(str(value2.args[2][1]))}\" }},\n"
                                json += f"\t\t\t\t{{ \"message\": \"{self.json_encode(str(value2.args[2][2]))}\" }}\n"
                            except TypeError:
                                json += f"\t\t\t\t{{ \"message\": \"{self.json_encode(str(value2.args[2]))}\" }}\n"
                            except IndexError:
                                json += f"\t\t\t\t{{ \"message\": \"{self.json_encode(str(value2))}\" }}\n"
                            json += "\t\t\t]},\n"

                        elif isinstance(value2, FunctionViewer):
                            if self._skip_func:
                                continue
                            # display the function and its properties
                            json += f"\t\t\t{{ \"{self.json_encode(value2.name)}\": [\n"
                            json += f"\t\t\t\t{{ \"name\": \"{self.json_encode(value2.name)}\" }},\n"
                            json += f"\t\t\t\t{{ \"args\": {self.json_encode(str(len(value2.args)))} }},\n"
                            json += f"\t\t\t\t{{ \"use\": \"{self.json_encode(str(value2)[26:])}\" }}\n"
                            json += "\t\t\t]},\n"

                        else:
                            # display the variable and value
                            name = self.json_encode(var2)
                            elem = value2

                            if isinstance(value2, bool):
                                elem = str(value2).lower()
                            elif not isinstance(value2, (int, float, complex)):
                                elem = f"\"{self.json_encode(str(value2))}\""

                            json += f"\t\t\t{{ \"{name}\": {elem} }},\n"

                json += '\t\t]},\n'

            json += '\t]},\n'

        return json + ']}\n'
