.. _odsgenerator-an-ods-generator:


odsgenerator, an .ods generator.
================================

Generate an OpenDocument Format .ods file from json or yaml file.


When used as a script, odsgenerator parses a JSON or YAML description of
tables and generates an ODF document using the odfdo library.

When used as a library, odsgenerator parses a python description of tables
and returns the ODF content as bytes.

-  description can be minimalist: a list of lists of lists,
-  description can be complex, allowing styles at row or cell level.


installation
------------

.. code-block:: bash

    $ pip install odsgenerator


usage
-----

::

   odsgenerator [-h] [--version] input_file output_file


arguments
---------

``input_file``: input file containing data in json or yaml format

``output_file``: output file, .ods file generated from input

Use ``odsgenerator --help`` for more details about input file parameters
and look at examples in the tests folder.


from python code
----------------

.. code-block:: python

    import odsgenerator
    raw = odsgenerator.ods_bytes([[["a", "b", "c"], [10, 20, 30]]])
    with open("sample1.ods", "wb") as f:
        f.write(raw)


The .ods file loaded in a spreadsheet:

.. figure:: https://raw.githubusercontent.com/jdum/odsgenerator/main/doc/sample1_ods.png

Another example with more parameters:

.. code-block:: python

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

The .ods file loaded in a spreadsheet:

.. figure:: https://raw.githubusercontent.com/jdum/odsgenerator/main/doc/sample2_ods.png


principle
---------

-  a document is a list or dict containing tabs,
-  a tab is a list or dict containing rows,
-  a row is a list or dict containing cells.


documentation
-------------

See ``tutorial.json`` or ``tutorial.yml`` and ``tutorial.ods`` in the doc folder.


license
-------

This project is licensed under the MIT License (see the
``LICENSE`` file for details).
