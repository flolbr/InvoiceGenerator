# Invoice Generator

This project is made to generate formal invoice-like documents.

It works by generating a $\LaTeX$ document using [Jinja2](https://palletsprojects.com/p/jinja/) templates.

An Excel file is used as the data source.

## Installation

```shell
pip install -r requirements.txt
```

## Development

### Data collection

#### Named values

Data can be extracted from the Excel file using named cells.
These can then be recovered using `get_named_value`.
It will fetch the value from the cell in the worksheet and return it in its best format,
inferred from the cell attribues.

#### Tables

Tables can be extracted from the Excel file using the `Table` class,
passing the sheet they're in as the first argument and their name as the second.
Each row will be extracted in the `Table.item` attribute, as `Item` instances.

Values from the columns are available as attributes of `Item`.

These classes can be inherited from to add data from the defined values.

### Templates

The templates are stored in the [/templates](/templates) directory.

The main template for the document is [/templates/invoice.jinja.tex](/templates/invoice.jinja.tex)

These work separately from $\LaTeX$ includes and rely on Jinja2 instead.

The bracket/escape format from Jinja templates had to be modified in order not to mess with the $\LaTeX$ documents.

The new format is as follows:

- Block `<+ ... +>`
- Variable `<@ ... @>`
- Comment `<- ... ->`

## Usage

```shell
usage: main.py [-h] [-n NUMBER] [-o OUTPUT] [--compile] input_file

positional arguments:
  input_file            File to use as input for the generated invoice.

options:
  -h, --help            show this help message and exit
  -n NUMBER, --number NUMBER
                        File number
  -o OUTPUT, --output OUTPUT
                        Output file
  --compile             Compile the LaTeX document
```

By default the $\LaTeX$ document will be compiled.

## Examples

See [/examples](/examples)