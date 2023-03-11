import argparse
import os
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

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


def format_cell(cell):
    if cell.style == 'Monétaire':
        return f'{cell.value :.02f}\\,€'.replace('.', ',')
    elif cell.is_date:
        return cell.value.strftime('%d/%m/%Y')
    else:
        return cell.value


def get_named_value(worksheet: Worksheet, cell_name: str):
    cell: Cell = worksheet[worksheet.defined_names[cell_name].value.split('!')[1]]
    return format_cell(cell)


class Contact:
    class Info:

        def __init__(self, worksheet: Worksheet, cell_name: str, title: str, template_name: str) -> None:
            self.cell_name = cell_name
            self.title = title.replace('_', '\\_')
            self.template_name = template_name
            try:
                self.value = get_named_value(worksheet, cell_name).replace('_', '\\_')
            except AttributeError:
                pass

        def __str__(self) -> str:
            return self.value

        def __repr__(self) -> str:
            return f'<{self.title}: {self.value}>'

    def __init__(self, worksheet: Worksheet) -> None:
        self.fields: dict[str, Contact.Info] = {
            'name': Contact.Info(worksheet, 'Name', 'Nom', 'CONTACT-NAME'),
            'address': Contact.Info(worksheet, 'Address', 'Adresse', 'CONTACT-ADDRESS'),
            'address_2': Contact.Info(worksheet, 'Address2', '', 'CONTACT-ADDRESS-2'),
            'phone': Contact.Info(worksheet, 'Phone', 'Téléphone', 'CONTACT-PHONE'),
            'mail': Contact.Info(worksheet, 'Mail', 'Email', 'CONTACT-MAIL'),
            'siret': Contact.Info(worksheet, 'SIRET', 'SIREN', 'CONTACT-SIRET')
        }
        self.is_client = False

    def __getattr__(self, item):
        get = self.fields.get(item)
        if get is None:
            raise AttributeError(f'Attribute "{item}" does not exist.')
        return get


class Client(Contact):

    def __init__(self, worksheet: Worksheet) -> None:
        super().__init__(worksheet)
        self.is_client = True


class Me(Contact):

    def __init__(self, worksheet: Worksheet) -> None:
        super().__init__(worksheet)
        self.is_client = False


class Document:

    def __init__(self, sheet: Worksheet, number: str = None) -> None:
        self.number = number or get_named_value(sheet, 'Number')
        self.type = get_named_value(sheet, 'Type')
        self.date = get_named_value(sheet, 'DocDate') or datetime.now().strftime('%d/%m/%Y')
        self.info = {
            'start_date': get_named_value(sheet, 'StartDate'),
            'delivery_date': get_named_value(sheet, 'DeliveryDate'),
        }


class Nomenclature:
    class Item:
        def __repr__(self) -> str:
            return self.item

    def __init__(self, name='Nomenclature') -> None:
        sheet: Worksheet = data_wb[name]
        self.columns, self.items = Nomenclature.read_excel_table(sheet, name)
        self.total = get_named_value(sheet, 'Total')

    def __repr__(self) -> str:
        return str(self.columns)

    @staticmethod
    def read_excel_table(sheet: Worksheet, table_name: str) -> [[str], [Item]]:
        """
        This function will read an Excel table
        and return a tuple of columns and data

        This function assumes that tables have column headers
        :param sheet: the sheet
        :param table_name: the name of the table
        :return: columns (list) and items (list[dict])
        """
        table_range = sheet.tables[table_name].ref

        table_head = sheet[table_range][0]
        table_data = sheet[table_range][1:]

        columns = [column.value for column in table_head]

        items = []

        for row in table_data:
            row_val = []
            for cell in row:
                value = format_cell(cell)
                row_val.append(value)
            item = Nomenclature.Item()
            for key, val in zip(columns, row_val):
                item.__setattr__(key, val)
            items.append(item)

        return columns, items


def app():
    global data_wb

    parser = argparse.ArgumentParser()
    parser.add_argument("input_file", type=Path, help="File to use as input for the generated invoice.")
    parser.add_argument("-n", "--number", type=str, help="File number")
    parser.add_argument("-o", "--output", type=Path, help="Output file")
    args = parser.parse_args()

    input_file = args.input_file
    output = args.output

    path = input_file.resolve().absolute().parent

    print(f'Input file: "{input_file.absolute()}"')

    data_wb = load_workbook(filename=input_file, data_only=True)

    client = Client(data_wb['Client'])
    me = Me(data_wb['Self'])

    nom = Nomenclature()

    document = Document(data_wb['Info'])
    template = env.get_template('invoice.jinja.tex')
    rendered = template.render(client=client, me=me, document=document, nomenclature=nom)

    # print(rendered)

    if output:
        output_file = (output if output.is_absolute() else path / output).with_suffix('.tex')
    else:
        output_file = path / f'{document.number} - {input_file.stem}.tex'

    print(f'Output file: "{output_file.absolute()}"')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(rendered)

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

    xelatex = subprocess.run(
        command,
        capture_output=False,
        shell=False,
        check=True
    )

    if xelatex.returncode == 0:
        shutil.rmtree(aux_dir)

    return


if __name__ == '__main__':
    app()
