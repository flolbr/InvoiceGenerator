from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from utilities import get_named_value


class Document:
    data: {}

    def __init__(self, sheet: Worksheet, template: str) -> None:
        self.template = template
        self.sheet = sheet


class Invoice(Document):

    def __init__(self, sheet: Worksheet, template: str = 'invoice.jinja.tex', number: str = None) -> None:
        super().__init__(sheet, template)
        self.number = number or get_named_value(self.sheet, 'Number')
        self.type = get_named_value(self.sheet, 'Type')
        self.date = get_named_value(self.sheet, 'DocDate') or datetime.now().strftime('%d/%m/%Y')
        self.info = {
            'start_date': get_named_value(self.sheet, 'StartDate'),
            'delivery_date': get_named_value(self.sheet, 'DeliveryDate'),
            'title': get_named_value(self.sheet, 'Title'),
        }
