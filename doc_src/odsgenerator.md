# odsgenerator, an .ods generator.

Generate an OpenDocument Format `.ods` file from a `.json` or `.yaml` file.


When used as a script, `odsgenerator` parses a JSON or YAML description of
tables and generates an ODF document using the `odfdo` library.

When used as a library, `odsgenerator` parses a Python description of tables
and returns the ODF content as bytes (ready to be saved as a valid ODF document).

-  The content description can be minimalist: a list of lists of lists,
-  or description can be complex, allowing styles at row or cell level.

See also https://github.com/jdum/odsparsator which is doing the reverse
operation, `.osd` to `.json`.

`odsgenerator` is a `Python3` package, using the [odfdo](https://github.com/jdum/odfdo) library. Current version requires Python >= 3.9, see prior versions for older environments.

Project:
    https://github.com/jdum/odsgenerator

Author:
    jerome.dumonteil@gmail.com

License:
    MIT


## Installation

Installation from Pypi (recommended):

```python
pip install odsgenerator
```

Installation from sources (requiring setuptools):

```python
pip install .
```


## CLI usage

```
odsgenerator [-h] [--version] input_file output_file
```

### arguments


``input_file``: input file containing data in JSON or YAML format

``output_file``: output file, `.ods` file generated from the input

Use ``odsgenerator --help`` for more details about input file parameters
and look at examples in the tests folder.


## Usage from python code

```python
from odsgenerator import odsgenerator

content = odsgenerator.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
with open("sample.ods", "wb") as file:
    file.write(content)
```

The resulting `.ods` file loaded in a spreadsheet:

![spreadsheet screnshot](https://raw.githubusercontent.com/jdum/odsgenerator/main/doc/sample1_ods.png)

Another example with more parameters:

```python
import odsgenerator

content = odsgenerator.ods_bytes(
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
with open("sample2.ods", "wb") as file:
    file.write(content)
```

The `.ods` file loaded in a spreadsheet, with gray background on first line:

![spreadsheet screnshot](https://raw.githubusercontent.com/jdum/odsgenerator/main/doc/sample2_ods.png)


## Tutorial example

The doc folder contains:

- A tutorial model, see `tutorial.json` or  `tutorial.yml` and resulting `tutorial.ods`,
- a showcase of the default styles: : `styles.json` and resulting `styles.ods`.


## Principle

-  A document is a list or dict containing tabs,
-  a tab is a list or dict containing rows,
-  a row is a list or dict containing cells.


A **cell** can be:

- int, float or str,
- a dict, with the following keys (only the 'value' key is mandatory):
    - value: int, float or str,
    - style: str or list of str, a style name or a list of style names,
    - text: str, a string representation of the value (for ODF readers
      who use it),
    - formula: str, content of the 'table:formula' attribute, some "of:"
      OpenFormula string,
    - colspanned: int, the number of spanned columns,
    - rowspanned: int, the number of spanned rows.

A **row** can be:

- a list of cells,
- a dict, with the following keys (only the 'row' key is mandatory):
    - row: a list of cells, see above,
    - style: str or list of str, a style name or a list of style names.

A **tab** can be:

- a list of rows,
- a dict, with the following keys (only the 'table' key is mandatory):
    - table: a list of rows,
    - width: a list containing the width of each column of the table
    - name: str, the name of the tab,
    - style: str or list of str, a style name or a list of style names.

A tab may have some post transformation:

- a list of span areas, cell coordinates are defined in the tab after
  its creation using odfo method Table.set_span(), with either
  coordiante system: "A1:B3" or [0, 0, 2, 1].

A **document** can be:

- a list of tabs,
- a dict, with the following keys (only the 'body' key is mandatory):
    - body: a list of tabs,
    - styles: a list of dict of styles definitions,
    - defaults: a dict, for the defaults styles.

A **style** definition is a dict with 2 items:

- the name of the style (optional, if not present the attribute
  style:name of the definition is used),
- an XML definition of the ODF style, see list below.

The styles provided for a row or a table can be of family table-row or
table-cell, they apply to row and below cells. A style defined at a
lower level (cell for instance) has priority over the style defined above
(row for instance).

In short, if you don't need custom styles, this is a valid document
description:

 ``[ [ ["a", "b", "c" ] ] ]``

 This list will create a document with only one tab (name will be "Tab 1"
 by default), containing one row of 3 values "a", "b", "c".


## Styles

Styles are XML strings of OpenDocument styles. They can be extracted from the
content.xml part of an existing `.ods` document.

- The `DEFAULT_STYLES` constant defines styles always available, they can be
  called by their name for cells or rows.
- To add a custom style, use the "styles" category of the document dict. A
  style is a dict with 2 keys, "definition" and "name".

List of provided styles
-----------------------
- `grid_06pt` means that the cell is surrounded by a black border of 0.6
  point,
- `gray` means that the cell has a gray background.
- The file `doc/styles.ods` displays all the provided styles.

**Row styles:**

- default_table_row
- table_row_1cm

**Cell styles:**

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
