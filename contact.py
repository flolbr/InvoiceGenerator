from openpyxl.worksheet.worksheet import Worksheet

from utilities import get_named_value


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
