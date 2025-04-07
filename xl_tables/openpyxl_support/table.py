from ..fields import get_row_text, get_table_text, save_table, text_to_table, parse_table
from .workbook import Workbook


__all__ = ['Table']


class Table(Workbook):
    get_row_text = staticmethod(get_row_text)
    get_table_text = staticmethod(get_table_text)
    save_table = staticmethod(save_table)
    text_to_table = staticmethod(text_to_table)
    parse_table = staticmethod(parse_table)
