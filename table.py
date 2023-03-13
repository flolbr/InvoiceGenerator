from openpyxl.worksheet.worksheet import Worksheet

from utilities import format_cell, get_named_value


class Table:
    class Item:
        pass

    def __init__(self, sheet: Worksheet, table_name: str) -> None:
        self.sheet = sheet
        self.columns, self.items = Table.read_excel_table(self.sheet, table_name)

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


class Nomenclature(Table):
    class Item:
        def __repr__(self) -> str:
            return self.item

    def __init__(self, sheet, table_name='Nomenclature') -> None:
        super().__init__(sheet, table_name)
        self.total = get_named_value(self.sheet, 'Total')
