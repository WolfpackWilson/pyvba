import os
import re

from pyvba.browser import Browser, IterableFunctionBrowser
from pyvba.viewer import FunctionViewer


class XMLExport:
    XML_ESCAPE_CHARS = {
        "&": "&amp;",
        '"': "&quot;",
        "'": "&apos;",
        ">": "&gt;",
        "<": "&lt;",
    }

    def __init__(self, browser: Browser, version=1.0, encoding: str = "UTF-8", skip_func: bool = False,
                 skip_err: bool = False):
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
        self._browser = browser
        self._xml_head = f'<?xml version="{str(version)}" encoding="{encoding}"?>\n'
        self._xml_str = None
        self._attrs = ['Name', 'Count']
        self._skip_func = skip_func
        self._skip_err = skip_err

    @property
    def string(self) -> str:
        """Return the well-formed XML in string format."""
        self._check()
        return self._xml_str

    @property
    def min(self) -> str:
        """Return the minimized XML in string format.

        The minimized version removes all newlines and tabs.
        """
        self._check()
        return re.sub(r'\n*\t*', '', self._xml_str)

    @staticmethod
    def xml_encode(text: str) -> str:
        """Map special XML characters to their encoded form in a given string."""
        return "".join(XMLExport.XML_ESCAPE_CHARS.get(c, c) for c in str(text))

    def _check(self):
        """Check if the xml_string needs to be generated."""
        if self._xml_str is None:
            self._generate()

    def _generate(self):
        """Begin generating the XML string."""
        self._xml_str = self._xml_head + self._generate_tag(self._browser)

        # convert empty elements to a single tag
        self._xml_str = re.sub(r'></.*?>', ' />', self._xml_str)

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
        if isinstance(elem, (Browser, IterableFunctionBrowser)):
            # display the browser and its children

            # setup the tag and attributes
            attrs = ["Name", "Count"]
            tag = XMLExport.Tag(elem.name)
            [
                tag.add_attr(attr, value)
                for attr, value in elem.all.items()
                if attr in attrs
            ]

            # add the element and start adding the sub-elements
            xml += '\t' * tabs + tag.open_tag + '\n'
            for item, value in elem.all.items():
                if item not in attrs:
                    xml += self._generate_tag(value, tabs + 1, name=item)
            xml += '\t' * tabs + tag.close_tag + '\n'
        elif isinstance(elem, FunctionViewer):
            if not self._skip_func:
                # display the function and its properties
                tag = XMLExport.Tag("Function", name=elem.name, args=len(elem.args))
                xml += tag.enclose(str(elem)[26:], tabs)
        elif isinstance(elem, BaseException):
            # display the error location and method
            if not self._skip_err:
                try:
                    tag = XMLExport.Tag("Error", on=str(elem.args[2][1]))
                    xml += tag.enclose(self.xml_encode(str(elem.args[2][2])), tabs)
                except (TypeError, IndexError):
                    tag = XMLExport.Tag("Error")
                    xml += tag.enclose(self.xml_encode(str(elem.args[2])), tabs)
                except IndexError:
                    tag = XMLExport.Tag("Error")
                    xml += tag.enclose(self.xml_encode(str(elem)), tabs)
        else:
            # display the variable and value
            tag = XMLExport.Tag(kwargs.get('name', 'Unknown'))
            xml += tag.enclose(self.xml_encode(str(elem)), tabs)
        return xml

    def print(self, minimize: bool = False):
        """Print the XML string in the normal or minimized version."""
        self._check()
        print(self._xml_str) if not minimize else print(self.min)

    def save(self, name: str, path: str = '.\\', minimize: bool = False):
        """Save to an XML file to a specified name and location.

        Parameters
        ----------
        name: str
            The file name.
        path: str
            The file path.
        minimize: bool
            A flag that determines the format of the final output.
        """
        save_as(
            self.string if not minimize else self.min,
            name, ".xml", path
        )

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


class JSONExport:
    JSON_ESCAPE_CHARS = ["\b", "\f", "\n", "\r", "\t", "\"", "\\"]

    def __init__(self, browser: Browser, skip_func: bool = False, skip_err: bool = False):
        self._browser = browser
        self._skip_func = skip_func
        self._skip_err = skip_err
        self._json_str = None

    @property
    def string(self) -> str:
        """Return the JSON in string format."""
        self._check()
        return self._json_str

    @property
    def min(self) -> str:
        """Return the minimized JSON in string format.

        The minimized version removes all newlines and tabs.
        """
        self._check()
        return re.sub(r'\n*\t*', '', self._json_str)

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
        if self._json_str is None:
            self._json_str = "{\n" + self._generate(self._browser, 1) + "}\n"
            self._json_str = re.sub(r',(?!\s*?[{\[\"\'\w])', '', self._json_str)

    def _generate(self, elem, tabs: int = 0, **kwargs) -> str:
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
        if isinstance(elem, Browser):
            # display the browser and its children
            json += "\t" * tabs + f"\"{self.json_encode(elem.name)}\": {{\n"

            for item, value in elem.all.items():
                json += self._generate(value, tabs + 1, name=item)

            json += "\t" * tabs + "},\n"
        elif isinstance(elem, IterableFunctionBrowser):
            # display the function browser and its children
            json += "\t" * tabs + f"\"{elem.name}\": {{\n"
            json += "\t" * (tabs + 1) + f"\"Name\": \"{self.json_encode(elem.name)}\",\n"
            json += "\t" * (tabs + 1) + f"\"Count\": {elem.count},\n"
            json += "\t" * (tabs + 1) + f"\"Items\": [\n"

            for item, value in elem.all.items():
                json += "\t" * (tabs + 2) + "{\n"
                json += self._generate(value, tabs + 3, name=item)
                json += "\t" * (tabs + 2) + "},\n"

            json += "\t" * (tabs + 1) + "]\n"
            json += "\t" * tabs + "},\n"
        elif isinstance(elem, FunctionViewer):
            if not self._skip_func:
                # display the function and its properties
                json += "\t" * tabs + f"\"{self.json_encode(elem.name)}\": {{\n"
                json += "\t" * (tabs + 1) + f"\"Name\": \"{self.json_encode(elem.name)}\",\n"
                json += "\t" * (tabs + 1) + f"\"args\": {self.json_encode(str(len(elem.args)))},\n"
                json += "\t" * (tabs + 1) + f"\"use\": \"{self.json_encode(str(elem)[26:])}\"\n"
                json += "\t" * tabs + "},\n"
        elif isinstance(elem, BaseException):
            # display the error location and method
            if not self._skip_err:
                json += "\t" * tabs + "\"Error\": {\n"
                try:
                    json += "\t" * (tabs + 1) + f"\"on\": \"{self.json_encode(str(elem.args[2][1]))}\",\n"
                    json += "\t" * (tabs + 1) + f"\"message\": \"{self.json_encode(str(elem.args[2][2]))}\"\n"
                except TypeError:
                    json += "\t" * (tabs + 1) + f"\"message\": \"{self.json_encode(str(elem.args[2]))}\"\n"
                except IndexError:
                    json += "\t" * (tabs + 1) + f"\"message\": \"{self.json_encode(str(elem))}\"\n"
                json += "\t" * tabs + "},\n"
        else:
            # display the variable and value
            name = self.json_encode(kwargs.get('name', 'Unknown'))

            if isinstance(elem, bool):
                elem = str(elem).lower()
            elif not isinstance(elem, (int, float, complex)):
                elem = f"\"{self.json_encode(str(elem))}\""

            json += "\t" * tabs + f"\"{name}\": {elem},\n"
        return json

    def print(self, minimize: bool = False):
        """Print the XML string in the normal or minimized version."""
        self._check()
        print(self._json_str) if not minimize else print(self.min)

    def save(self, name: str, path: str = '.\\', minimize: bool = False):
        """Save to an JSON file to a specified name and location.

        Parameters
        ----------
        name: str
            The file name.
        path: str
            The file path.
        minimize: bool
            A flag that determines the format of the final output.
        """
        save_as(
            self.string if not minimize else self.min,
            name, ".json", path
        )


def save_as(data: str, name: str, ext: str, path: str = '.\\'):
    """Save a string object to a specified name and location.

    Parameters
    ----------
    data: str
        The string data to be printed (e.g. XML, JSON, etc.).
    name: str
        The name of the file.
    ext: str
        The file extension (e.g. .xml, .json, etc.)
    path: str
        The save location.
    """
    os.makedirs(path, exist_ok=True)
    with open(os.path.join(path, name + ext), "w") as file:
        file.write(data)
        file.close()
