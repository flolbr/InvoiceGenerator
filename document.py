from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from utilities import get_named_value


class Document:

    def __init__(self, sheet: Worksheet, number: str = None) -> None:
        self.number = number or get_named_value(sheet, 'Number')
        self.type = get_named_value(sheet, 'Type')
        self.date = get_named_value(sheet, 'DocDate') or datetime.now().strftime('%d/%m/%Y')
        self.info = {
            'start_date': get_named_value(sheet, 'StartDate'),
            'delivery_date': get_named_value(sheet, 'DeliveryDate'),
        }
