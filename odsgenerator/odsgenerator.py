# Copyright 2021-2024 Jérôme Dumonteil
# Licence: MIT
# Authors: jerome.dumonteil@gmail.com
"""Generate an OpenDocument Format .ods file from json or yaml file.

When used as a script, odsgenerator parses a JSON or YAML description of
tables and generates an ODF document using the odfdo library.

When used as a library, odsgenerator parses a python description of tables
and returns the ODF content as bytes.

    -  description can be minimalist: a list of lists of lists,
    -  description can be complex, allowing styles at row or cell level.

See also https://github.com/jdum/odsparsator which is doing the reverse
operation, ods => json.


Usage
-----

   odsgenerator [-h] [--version] input_file output_file


Arguments
---------

input_file: input file containing data in json or yaml format

output_file: output file, .ods file generated from input

Use `odsgenerator --help` for more details about input file parameters
and look at examples in the tests folder.


From python code
----------------

    import odsgenerator
    raw = odsgenerator.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
    with open("sample1.ods", "wb") as f:
        f.write(raw)

Another example with more parameters:

    raw = odsgenerator.ods_bytes(
        [
            {
                "name": "first tab",
                "style": "cell_decimal2",
                "table": [
                    {
                        "row": ["a", "b", "c"],
                        "style": "bold_center_bg_gray_grid_06pt",
                    },
                    [10, 20, 30],
                ],
            }
        ]
    )
    with open("sample2.ods", "wb") as f:
        f.write(raw)


Principle
---------

-  a document is a list or dict containing tabs,
-  a tab is a list or dict containing rows,
-  a row is a list or dict containing cells.


A cell can be:
    - int, float or str
    - a dict, with the following keys (only the 'value' key is mandatory):
        - value: int, float or str,
        - style: str or list of str, a style name or a list of style names,
        - text: str, a string representation of the value (for ODF readers
          who use it),
        - formula: str, content of the 'table:formula' attribute, some "of:"
          OpenFormula string,
        - colspanned: int, the number of spanned columns,
        - rowspanned: int, the number of spanned rows.

A row can be:
    - a list of cells,
    - a dict, with the following keys (only the 'row' key is mandatory):
        - row: a list of cells, see above,
        - style: str or list of str, a style name or a list of style names.

A tab can be:
    - a list of rows,
    - a dict, with the following keys (only the 'table' key is mandatory):
        - table: a list of rows,
        - width: a list containing the width of each column of the table,
        - name: str, the name of the tab,
        - style: str or list of str, a style name or a list of style names.

A tab may have some post transformation:
    - a list of span areas, cell coordinates are defined in the tab after
      its creation using odfo method Table.set_span(), with either
      coordiante system: "A1:B3" or [0, 0, 2, 1].

A document can be:
    - a list of tabs,
    - a dict, with the following keys (only the 'body' key is mandatory):
        - body: a list of tabs,
        - styles: a list of dict of styles definitions,
        - defaults: a dict, for the defaults styles.

A style definition is a dict with 2 items:
    - name: str, the name of the style (optional, if not present the
      attribute style:name of the definition is used),
    - an XML definition of the ODF style, see list below.

The styles provided for a row or a table can be of family table-row or
table-cell, they apply to row and below cells. A style defined at a
lower level (cell for instance) has priority over the style defined above
(row for instance).

In short, if you don't need custom styles, this is a valid document
description:
    [ [ ["a", "b", "c" ] ] ]

This list will create a document with only one tab (name will be "Tab 1"
by default), containing one row of 3 values "a", "b", "c".


Styles
------

Styles are XML strings of OpenDocument styles. They can be extracted from the
content.xml part of an existing .ods document.

    - The DEFAULT_STYLES constant defines styles always available, they can be
      called by their name for cells or rows.
    - To add a custom style, use the "styles" category of the document dict.
      A style is a dict with 2 keys, "definition" and "name".

List of provided styles:

    - 'grid_06pt' means that the cell is surrounded by a black border of 0.6
      point,
    - 'gray' means that the cell has a gray background.
    - The file doc/styles.ods displays all the styles provided.

Row styles:
    - default_table_row
    - table_row_1cm
Cell styles:
    - bold
    - bold_center
    - left
    - right
    - center
    - cell_decimal1
    - cell_decimal2
    - cell_decimal3
    - cell_decimal4
    - cell_decimal6
    - grid_06pt
    - bold_left_bg_gray_grid_06pt
    - bold_right_bg_gray_grid_06pt
    - bold_center_bg_gray_grid_06pt
    - bold_left_grid_06pt
    - bold_right_grid_06pt
    - bold_center_grid_06pt
    - left_grid_06pt
    - right_grid_06pt
    - center_grid_06pt
    - integer_grid_06pt
    - integer_no_zero_grid_06pt
    - center_integer_no_zero_grid_06pt
    - decimal1_grid_06pt
    - decimal2_grid_06pt
    - decimal3_grid_06pt
    - decimal4_grid_06pt
    - decimal6_grid_06pt
"""

import io
import sys

try:  # noqa: SIM105
    import yaml
except ModuleNotFoundError:  # pragma: no cover
    pass
import json

from odfdo import Cell, Document, Element, Row, Table

__version__ = "1.9.0"

DEFAULT_STYLES = [
    {
        "name": "default_table_row",
        "definition": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="4.52mm"
            fo:break-before="auto" style:use-optimal-row-height="true"/>
            </style:style>
        """,
    },
    {
        "name": "table_row_1cm",
        "definition": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="1cm"
            fo:break-before="auto"/>
            </style:style>
        """,
    },
    {
        "name": "bold",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:text-properties fo:font-weight="bold"
            style:font-weight-asian="bold" style:font-weight-complex="bold"/>
            <style:table-cell-properties style:text-align-source="value-type"/>
            <style:paragraph-properties
            fo:margin-right="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "bold_center",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:text-properties fo:font-weight="bold"
            style:font-weight-asian="bold" style:font-weight-complex="bold"/>
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="center"/>
            </style:style>
        """,
    },
    {
        "name": "left",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="start"
            fo:margin-left="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "right",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="end"
            fo:margin-right="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "center",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="center"/>
            </style:style>
        """,
    },
    {
        "name": "decimal1",
        "definition": """
            <number:number-style><number:number number:decimal-places="1"
            loext:min-decimal-places="1" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "cell_decimal1",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal1">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal2",
        "definition": """
            <number:number-style><number:number number:decimal-places="2"
            loext:min-decimal-places="2" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "cell_decimal2",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal2">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal3",
        "definition": """
            <number:number-style><number:number number:decimal-places="3"
            loext:min-decimal-places="3" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "cell_decimal3",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal3">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal4",
        "definition": """
            <number:number-style><number:number number:decimal-places="4"
            loext:min-decimal-places="4" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "cell_decimal4",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal4">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal6",
        "definition": """
            <number:number-style><number:number number:decimal-places="6"
            loext:min-decimal-places="6" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "cell_decimal6",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal6">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "integer",
        "definition": """
            <number:number-style><number:number number:decimal-places="0"
            loext:min-decimal-places="0" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "integer_no_zero",
        "definition": """
            <number:number-style><number:number number:decimal-places="0"
            loext:min-decimal-places="0" number:min-integer-digits="0"
            number:grouping="false"/>
            </number:number-style>
        """,
    },
    {
        "name": "grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "bold_left_bg_gray_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:background-color="#dddddd" fo:border="0.06pt solid #000000"
             style:text-align-source="fix"/>
             <style:paragraph-properties fo:text-align="start"
             fo:margin-left="1.2mm"/>
             <style:text-properties fo:font-weight="bold"
             style:font-weight-asian="bold" style:font-weight-complex="bold"/>
             </style:style>
         """,
    },
    {
        "name": "bold_right_bg_gray_grid_06pt",
        "definition": """
              <style:style style:family="table-cell"
              style:parent-style-name="Default">
              <style:table-cell-properties
              fo:background-color="#dddddd" fo:border="0.06pt solid #000000"
              style:text-align-source="fix"/>
              <style:paragraph-properties fo:text-align="end"
              fo:margin-right="1.2mm"/>
              <style:text-properties fo:font-weight="bold"
              style:font-weight-asian="bold" style:font-weight-complex="bold"/>
              </style:style>
          """,
    },
    {
        "name": "bold_center_bg_gray_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:background-color="#dddddd" fo:border="0.06pt solid #000000"
             style:text-align-source="fix"/>
             <style:paragraph-properties fo:text-align="center"/>
             <style:text-properties fo:font-weight="bold"
             style:font-weight-asian="bold" style:font-weight-complex="bold"/>
             </style:style>
         """,
    },
    {
        "name": "bold_left_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"
             style:text-align-source="fix"/>
             <style:paragraph-properties fo:text-align="start"
             fo:margin-left="1.2mm"/>
             <style:text-properties fo:font-weight="bold"
             style:font-weight-asian="bold" style:font-weight-complex="bold"/>
             </style:style>
         """,
    },
    {
        "name": "bold_right_grid_06pt",
        "definition": """
              <style:style style:family="table-cell"
              style:parent-style-name="Default">
              <style:table-cell-properties
              fo:border="0.06pt solid #000000"
              style:text-align-source="fix"/>
              <style:paragraph-properties fo:text-align="end"
              fo:margin-right="1.2mm"/>
              <style:text-properties fo:font-weight="bold"
              style:font-weight-asian="bold" style:font-weight-complex="bold"/>
              </style:style>
          """,
    },
    {
        "name": "bold_center_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"
             style:text-align-source="fix"/>
             <style:paragraph-properties fo:text-align="center"/>
             <style:text-properties fo:font-weight="bold"
             style:font-weight-asian="bold" style:font-weight-complex="bold"/>
             </style:style>
         """,
    },
    {
        "name": "left_grid_06pt",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"
            fo:border="0.06pt solid #000000"/>
            <style:paragraph-properties
            fo:margin-left="1.2mm" fo:text-align="start"/>
            </style:style>
        """,
    },
    {
        "name": "right_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties style:text-align-source="fix"
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-right="1.2mm" fo:text-align="end"/>
             </style:style>
         """,
    },
    {
        "name": "center_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties style:text-align-source="fix"
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties fo:text-align="center"/>
             </style:style>
         """,
    },
    {
        "name": "integer_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="integer">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "integer_no_zero_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="integer_no_zero">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             </style:style>
         """,
    },
    {
        "name": "center_integer_no_zero_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="integer_no_zero">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties fo:text-align="center"/>
             </style:style>
         """,
    },
    {
        "name": "decimal1_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal1">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal2_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal2">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal3_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal3">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal4_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal4">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "decimal6_grid_06pt",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default"
             style:data-style-name="decimal6">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
]

DEFAULTS_DICT = {
    "style_table_row": "default_table_row",
    "style_table_cell": "",
    "style_str": "left",
    "style_int": "right",
    "style_float": "right",
    "style_other": "left",
}

BODY = "body"
TABLE = "table"
ROW = "row"
VALUE = "value"
FORMULA = "formula"
COLSPAN = "colspanned"
ROWSPAN = "rowspanned"
TEXT = "text"
NAME = "name"
DEFINITION = "definition"
WIDTH = "width"
SPAN = "span"
STYLE = "style"
STYLES = "styles"
DEFAULTS = "defaults"
DEFAULT_TAB_PREFIX = "Tab"


class ODSGenerator:
    """Core class of odsgenerator.

    The class parses the input description and generate an ODF document.
    Use ODSGenerator via the front-end functions ods_bytes() or
    content_to_ods().

    Args:
        content (list or dict): Description of tables.
    """

    def __init__(self, content):
        self.doc = Document("spreadsheet")
        self.doc.body.clear()
        self.tab_counter = 0
        self.defaults = DEFAULTS_DICT
        self.styles_elements = {}
        self.used_styles = set()
        self.spanned_cells = []
        self.parse(content)

    def save(self, path):
        """Save the resulting ODF document.

        Args:
            path (str or Path or BytesIO): Path of the ODF output file.
        """
        self.doc.save(path)

    def parse_styles(self, styles, insert=False):
        """Load available styles, either from default and the input description.

        Args:
            styles (list): List of styles definitions.
            insert (bool): Force inseerti in document.
        """
        if not styles:
            return
        for style_item in styles:
            name = style_item.get(NAME)
            definition = style_item.get(DEFINITION)
            style = Element.from_tag(definition)
            if name:
                style.name = name
            else:
                name = style.name
            self.styles_elements[name] = style
            if insert:
                self.insert_style(name)

    def insert_style(self, name, automatic=True):
        """Insert the named style into the ODF document."""
        if name and name not in self.used_styles and name in self.styles_elements:
            style = self.styles_elements[name]
            self.doc.insert_style(style, automatic=automatic)
            self.used_styles.add(name)
            # add style dependacies
            for key, value in style.attributes.items():
                if key.endswith("style-name"):
                    self.insert_style(value)

    def guess_style(self, opt, family, default):
        """Guess which style to apply.

        Search list of styles under the "style" key, check against family of
        style, apply default if none found.

        Args:
            opt (dict): Part of input description.
            family (str): ODF family style.
            default (str): Default style name.

        Returns:
            str or None: Name of he style to apply.
        """
        style_list = opt.get(STYLE, [])
        if not isinstance(style_list, list):
            style_list = [style_list]
        for style_name in style_list:
            if style_name:
                style = self.styles_elements.get(style_name)
                if style and style.family == family:
                    return style_name
        if default:
            style = self.styles_elements.get(default)
            if style and style.family == family:
                return default
        return None

    @staticmethod
    def split(item, key):
        """Extract the value of the key if item is a dict.

        If item is a dict, pop the value from the key, else consider that item
        is already the response.

        Args:
            item (any type): Part of input description.
            key (str): Key to extract.

        Returns:
            tuple: extracted content, remaining dict.
        """
        if isinstance(item, dict):
            inner = item.pop(key, [])
            return (inner, item)
        # item can be list or value, or None
        return (item, {})

    def parse(self, content):
        """Parse the top level of the input description."""
        body, opt = self.split(content, BODY)
        self.defaults.update(opt.get(DEFAULTS, {}))
        self.parse_styles(DEFAULT_STYLES)
        self.parse_styles(opt.get(STYLES), insert=True)
        for table_content in body:
            self.parse_table(table_content)

    def parse_table(self, table_content):
        """Parse a table level from the input description."""
        rows, opt = self.split(table_content, TABLE)
        self.tab_counter += 1
        table = Table(opt.get(NAME, f"{DEFAULT_TAB_PREFIX} {self.tab_counter}"))
        style_table_row = self.guess_style(
            opt, "table-row", self.defaults["style_table_row"]
        )
        style_table_cell = self.guess_style(
            opt, "table-cell", self.defaults["style_table_cell"]
        )
        self.spanned_cells = []  # spanned_cells is relative to this table
        for row_content in rows:
            self.parse_row(table, row_content, style_table_row, style_table_cell)
        self.parse_width(table, opt)
        self.parse_spanned(table, opt)
        self.doc.body.append(table)

    def parse_row(self, table, row_content, style_table_row, style_table_cell):
        """Parse a row level from the input description."""
        cells, opt = self.split(row_content, ROW)
        style_table_row = self.guess_style(opt, "table-row", style_table_row)
        self.insert_style(style_table_row)
        row = Row(style=style_table_row)
        style_table_cell = self.guess_style(opt, "table-cell", style_table_cell)
        row = table.append_row(row)
        for cell_content in cells:
            self.parse_cell(row, cell_content, style_table_cell)

    def parse_cell(self, row, cell_content, style_table_cell):
        """Parse a cell level from the input description."""
        value, opt = self.split(cell_content, VALUE)
        if style_table_cell:
            default = style_table_cell
        elif value is None or isinstance(value, str):
            default = self.defaults["style_str"]
        elif isinstance(value, int):
            default = self.defaults["style_int"]
        elif isinstance(value, float):
            default = self.defaults["style_float"]
        else:
            default = self.defaults["style_other"]
        style = self.guess_style(opt, "table-cell", default)
        self.insert_style(style)
        cell = Cell(
            value=value, style=style, text=opt.get(TEXT), formula=opt.get(FORMULA)
        )
        attr = opt.get("attr")
        if attr:
            for key, value in attr.items():
                cell.set_attribute(key, value)
        cell = row.append(cell)
        if COLSPAN in opt or ROWSPAN in opt:
            self.store_spanned_cell(cell, opt)

    def column_width_style(self, width):
        """Generate an ODF style for a column width.

        Args:
            width (str): The required width, any ODF format like "10.5mm".

        Returns:
            str: The XML string of the style.
        """
        return self.doc.insert_style(
            Element.from_tag(
                f"""
                    <style:style style:family="table-column">
                    <style:table-column-properties fo:break-before="auto"
                    style:column-width="{width}"/>
                    </style:style>
                """
            ),
            automatic=True,
        )

    def parse_width(self, table, opt):
        """Parse the width tag of the input description."""
        width_opt = opt.get(WIDTH)
        if not width_opt:
            return
        if isinstance(width_opt, list):
            for position, width in enumerate(width_opt):
                if width:
                    column = table.get_column(position)
                    column.style = self.column_width_style(width)
                    table.set_column(position, column)
            return
        for position, column in enumerate(table.get_columns()):
            column.style = self.column_width_style(width_opt)
            table.set_column(position, column)

    def parse_spanned(self, table, opt):
        """Parse the span tag of the input description."""
        span_opt = opt.get(SPAN)
        if span_opt:
            if not isinstance(span_opt, list):
                span_opt = [span_opt]
            span_opt.extend(self.spanned_cells)
        else:
            span_opt = self.spanned_cells
        for area in span_opt:
            table.set_span(area)

    def store_spanned_cell(self, cell, opt):
        colspan = max(1, int(opt.get(COLSPAN, 1)))
        rowspan = max(1, int(opt.get(ROWSPAN, 1)))
        if colspan < 2 and rowspan < 2:
            return
        self.spanned_cells.append(
            [cell.x, cell.y, cell.x + colspan - 1, cell.y + rowspan - 1]
        )


def content_to_ods(content, output_path):
    """Parse document description and save resulting ODF to file.

    Args:
        content (list or dict): Input description of tables.
        output_path (str or Path or BytesIO): Path of the ODF output file.
    """
    document = ODSGenerator(content)
    document.save(output_path)


def file_to_ods(input_path, output_path):
    """Parse the input file and save resulting ODF to file.

    The input file can be JSON or other YAML format.

    Args:
        input_path (str or Path): Path of the file with desription to parse.
        output_path (str or Path or BytesIO): Path of the ODF output file.
    """
    if "yaml" in sys.modules:
        with open(input_path, encoding="utf8") as file:
            content = yaml.load(file, yaml.SafeLoader)
    else:  # fall back to json
        with open(input_path, encoding="utf8") as file:
            content = json.load(file)
    content_to_ods(content, output_path)


def ods_bytes(content):
    """Parse the document description and generate an ODF document as bytes.

    This is the recommended front-end when odsgenerator is used as a library.

    Example:
    >>> raw = odsgenerator.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
    >>> with open("sample1.ods", "wb") as f:
    ...     f.write(raw)

    ...
    7266

    Args:
        content (list or dict): Input description of tables.

    Returns:
        bytes: Zipped OpenDocument format.
    """
    with io.BytesIO() as iobytes:
        document = ODSGenerator(content)
        document.save(iobytes)
        return iobytes.getvalue()
