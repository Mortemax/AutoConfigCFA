import sys

from abc import ABC, abstractmethod
from typing import Iterable, List, Tuple
from xml.etree import ElementTree
from xml.etree.ElementTree import Element

from openpyxl import load_workbook
from openpyxl.cell import Cell

from ipaddress import IPv4Address, IPv4Network, ip_network

__author__ = "Alejandro Cano Bermúdez"

#################
## Base classes
#################

class RuleOption(ABC):
    """Rule option.
    """

    @abstractmethod
    def can_apply(self, value: str) -> bool:
        """Checks if this option should be applied for the given value.

        Parameters
        ----------
        value : str
            The column value

        Returns
        -------
        bool
            True if should be applied
        """
        ...

    @abstractmethod
    def apply(self, xml: Element, value: str) -> None:
        """Applies this option to the given XML document.

        Parameters
        ----------
        xml : Element
            The XML element
        value : str
            The option value
        """
        ...


class Rule(ABC):
    """Rule.
    """

    @abstractmethod
    def can_apply(self, title: str) -> bool:
        """Checks if this rule should be applied in the given column.

        Parameters
        ----------
        title : str
            The column title

        Returns
        -------
        bool
            True if should be applied
        """
        ...

    @abstractmethod
    def get_options(self) -> Iterable[RuleOption]:
        """Gets the rule options.

        Returns
        -------
        Iterable[RuleOption]
            The options
        """
        ...

    def apply(self, xml: Element, value: str) -> None:
        """Applies this rule to the given XML document.

        Parameters
        ----------
        xml : Element
            The XML element
        value : str
            The column value
        """
        for option in self.get_options():
            if option.can_apply(value):
                print(f'\t\tApplying {option.__class__.__name__}')
                option.apply(xml, value)


####################
## Utility methods
####################

def update_node_value(xml: Element, xpath: str, value: str) -> None:
    node = xml.find(xpath)
    assert node is not None
    node.text = value

def update_rule_type(xml: Element, tracker: str, new_type: str) -> None:
    update_node_value(xml, f'filter/rule/tracker[.="{tracker}"]..type', new_type)


######################
## Rules definitions
######################

class ProtocolsRule(Rule):


    class OptionFTP(RuleOption):

        def can_apply(self, value: str) -> bool:
            return 'ftp' in value.lower()

        def apply(self, xml: Element, value: str) -> None:
            update_rule_type(xml, '1629481790', 'pass') # WAN
            update_rule_type(xml, '1629485503', 'pass') # LAN


    class OptionSMB(RuleOption):

        def can_apply(self, value: str) -> bool:
            return 'smb' in value.lower()

        def apply(self, xml: Element, value: str) -> None:
            update_rule_type(xml, '1629485258', 'pass') # WAN
            update_rule_type(xml, '1629485518', 'pass') # LAN


    class OptionSSH(RuleOption):

        def can_apply(self, value: str) -> bool:
            return 'ssh' in value.lower()

        def apply(self, xml: Element, value: str) -> None:
            update_rule_type(xml, '1629479843', 'pass') # WAN
            update_rule_type(xml, '1629479704', 'pass') # LAN


    def get_options(self) -> Iterable[RuleOption]:
        return (
            self.OptionFTP(),
            self.OptionSMB(),
            self.OptionSSH(),
        )

    def can_apply(self, title: str) -> bool:
        return 'estos protocolos utiliza' in title.lower()


class TeleworkingRule(Rule):


    class OptionNo(RuleOption):

        def can_apply(self, value: str) -> bool:
            return 'no' in value.lower()

        def apply(self, xml: Element, value: str) -> None:
            update_rule_type(xml, '1614115961', 'block') # OpenVPN


    def get_options(self) -> Iterable[RuleOption]:
        return (
            self.OptionNo(),
        )

    def can_apply(self, title: str) -> bool:
        return 'realizar teletrabajo' in title.lower()


class BlockBadTrafficRule(Rule):


    class OptionNo(RuleOption):

        def can_apply(self, value: str) -> bool:
            return 'no quiero bloquear' in value.lower()

        def apply(self, xml: Element, value: str) -> None:
            update_node_value(xml, 'installedpackages/squidguarddefault/config/dest', 'all')

            nodes = xml.findall('installedpackages/pfblockerngblacklist/item/selected')
            for node in nodes:
                node.text = None


    def get_options(self) -> Iterable[RuleOption]:
        return (
            self.OptionNo(),
        )

    def can_apply(self, title: str) -> bool:
        return 'sitios ociosos' in title.lower()


class AdminEmailRule(Rule):


    class Option(RuleOption):

        def can_apply(self, value: str) -> bool:
            return len(value) > 0

        def apply(self, xml: Element, value: str) -> None:
            update_node_value(xml, 'installedpackages/squid/config/admin_email', value)


    def get_options(self) -> Iterable[RuleOption]:
        return (
            self.Option(),
        )

    def can_apply(self, title: str) -> bool:
        return 'correo electrónico' in title.lower()

class NetworkRule(Rule):


    class Option(RuleOption):

        def can_apply(self, value: str) -> bool:        
            return IPv4Network('192.168.100.1/24',strict=False).overlaps(IPv4Network(value,strict=False))

        def apply(self, xml: Element, value: str) -> None:
            update_node_value(xml, 'interfaces/lan/ipaddr', '10.0.0.1')
        
    def get_options(self) -> Iterable[RuleOption]:
        return (
            self.Option(),
        )

    def can_apply(self, title: str) -> bool:
        return ' ruter de salida o gateway' in title.lower()


RULES: Iterable[Rule] = (
    ProtocolsRule(),
    TeleworkingRule(),
    BlockBadTrafficRule(),
    AdminEmailRule(),
    NetworkRule(),
)


#################
## Main program
#################

def get_excel_columns(filename: str, row_id: int) -> Iterable[Tuple[str, str]]:
    """Gets the excel columns and their values for the given row ID.

    Parameters
    ----------
    filename : str
        The XLSX filename
    row_id : int
        The wanted row ID

    Returns
    -------
    Iterable[Tuple[str, str]]
        The iterable of tuples (title, value) for the found row
    """
    # Open workbook
    wb = load_workbook(filename)
    sheet = wb.active

    # Get first row (columns titles)
    first_row: List[Cell] = next(sheet.rows)
    assert first_row is not None

    # Find wanted row
    row: List[Cell] = []

    for row in sheet.rows:
        if row[0].value == str(row_id):
            break

    if not row:
        print(f'Row with ID {row_id} was not found!')
        exit(3)

    return ((str(title.value), str(cell.value)) for title, cell in zip(first_row, row))

if __name__ == '__main__':
    if len(sys.argv) != 5:
        print('Usage: main.py <xlsx-path> <base-xml-path> <output-xml-path> <xlsx-row-id>')
        exit(1)

    input_xlsx_file = sys.argv[1]
    input_xml_file = sys.argv[2]
    output_file = sys.argv[3]
    row_id_str = sys.argv[4]

    try:
        row_id = int(row_id_str)
    except ValueError:
        print('The row ID should be a number!')
        exit(2)

    print('Reading XLSX file...')

    columns = get_excel_columns(input_xlsx_file, row_id)

    print('Reading base configuration...')

    try:
        xml_tree = ElementTree.parse(input_xml_file)
    except OSError:
        print('Could not open base configuration file!')
        exit(4)

    xml_root = xml_tree.getroot()

    print('Applying rules...')

    for title, value in columns:
        for rule in RULES:
            if rule.can_apply(title):
                print(f'\tApplying rule {rule.__class__.__name__}')
                rule.apply(xml_root, value)

    print('Writing output...')

    xml_tree.write(output_file)

    print('Bye!')
