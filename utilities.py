from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def format_cell(cell: Cell) -> any:
    """
    Converts a cell into a string using the cell attributes to infer the format.

    ⚠️Some of the formats detection are using style names which are language specific.
    These might have to be modified to support your language.
    :param cell: Input cell
    :return: Formatted cell value
    """
    if cell.style == 'Monétaire':
        return f'{cell.value :.02f}\\,€'.replace('.', ',')
    elif cell.is_date:
        return cell.value.strftime('%d/%m/%Y')
    else:
        return cell.value


def get_named_value(worksheet: Worksheet, cell_name: str) -> any:
    """
    Fetches a cell value from a worksheet and returns its formatted value.
    :param worksheet: Worksheet to extract the cell from
    :param cell_name: Cell defined name
    :return: Formatted cell value
    """
    cell: Cell = worksheet[worksheet.defined_names[cell_name].value.split('!')[1]]
    return format_cell(cell)
