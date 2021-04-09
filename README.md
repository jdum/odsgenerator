## odsgenerator

odsgenerator generates an OpenDocument Format .ods file from json or yaml file.

The script parses a JSON or YAML description of tables and generates an
ODF document using the odfdo library.

  * description can be minimalist: a list of lists of lists,
  * description can be complex, allowing styles at row or cell level.

**principle:**

  * a document is a list of tabs,
  * a tab is a list of rows,
  * a row is a list of cells.

**usage:**

    odsgenerator [-h] input_file output_file   

**positional arguments:**

  `input_file`:   input file containing data in json or yaml format

  `output_file`:  output file, .ods file generated from input


Use `odsgenerator --help` for more details about input file parameters and
look at examples in the test folder.
