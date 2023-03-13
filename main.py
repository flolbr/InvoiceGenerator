import argparse
import os
import shutil
import subprocess
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import load_workbook
from openpyxl.workbook import Workbook

from contact import Client, Me
from document import Document
from table import Nomenclature

os.system('')  # Needed to enable VT100 support
# sys.stdout.write("\x1b[8;{rows};{cols}t".format(rows=32, cols=999))


file_path = Path(os.path.realpath(__file__)).parent

env = Environment(
    loader=FileSystemLoader(file_path / 'templates'),
    autoescape=select_autoescape(),
    block_start_string='<+',
    block_end_string='+>',
    variable_start_string='<@',
    variable_end_string='@>',
    comment_start_string='<-',
    comment_end_string='->',
    trim_blocks=True,
    lstrip_blocks=True,
)

data_wb: Workbook


def app():
    global data_wb

    # Parse the arguments from the command line
    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", type=Path, help="File to use as input for the generated invoice.")
    parser.add_argument("-n", "--number", type=str, help="File number")
    parser.add_argument("-o", "--output", type=Path, help="Output file")
    parser.add_argument("--compile", action="store_false", help="Compile the LaTeX document")
    args = parser.parse_args()

    # Figure out the paths
    input_file = args.input_file
    output = args.output

    path = input_file.resolve().absolute().parent

    print(f'Input file: "{input_file.absolute()}"')

    data_wb = load_workbook(filename=input_file, data_only=True)

    # Extract the data from the Excel file
    client = Client(data_wb['Client'])
    me = Me(data_wb['Self'])

    nom = Nomenclature(data_wb['Nomenclature'])

    document = Document(data_wb['Info'])

    # Render the template
    template = env.get_template('invoice.jinja.tex')
    rendered = template.render(client=client, me=me, document=document, nomenclature=nom)
    # print(rendered)

    # Output the generated document
    if output:
        output_file = (output if output.is_absolute() else path / output).with_suffix('.tex')
    else:
        output_file = path / f'{document.number} - {input_file.stem}.tex'

    print(f'Output file: "{output_file.absolute()}"')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(rendered)

    # Compile the document if required
    if args.compile:
        # Temp file directory
        aux_dir = output_file.parent / '.texgen'

        command = ' '.join([
            'xelatex',
            '-c-style-errors',
            '-file-line-error',
            '-interaction=nonstopmode',
            f'-output-directory="{output_file.parent}"',
            f'-aux-directory="{aux_dir}"',
            f'"{output_file.name}"',
        ])

        # Command call
        xelatex = subprocess.run(
            command,
            capture_output=False,
            shell=False,
            check=True
        )

        # Remove the temp files on success only
        if xelatex.returncode == 0:
            shutil.rmtree(aux_dir)

    return


if __name__ == '__main__':
    app()
