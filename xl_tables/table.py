import os
from .prop_utils import HashDict
from .excel_attributes import Workbook
from .fields import is_iterable


__all__ = ['get_row_text', 'get_table_text', 'save_table', 'text_to_table', 'parse_table', 'Table']


def get_row_text(values, delimiter='\t', serializer=None):
    """Return a text row for the given list of values

    Args:
        values (list/tuple/object): List of table values. Example: (a1, b1, c1)
        delimiter (str)['\t']: Delimiter to separate column values by.
        serializer (callable/function)[None/str]: Function to convert the cell value to a string.
            This may be important for saving and loading files (json.dumps may be desired).

    Returns:
        text (str): Text row. Example: "a1\tb1\tc1"
    """
    if serializer is None:
        serializer = str

    if is_iterable(values):
        return delimiter.join(serializer(c) for c in values)
    else:
        return serializer(values)


def get_table_text(values, head=None, header=None, delimiter='\t', head_delimiter=' = ', row_delimiter='\n',
                   serializer=None, head_serializer=None):
    """Return a text table for the given list of values

    Args:
        values (list/tuple/object): List of table values. Example: [(a1, b1, c1), (a2, b2, c2)].
        head (dict)[None]: Dictionary of Name Value pairs to save before the table.
        header (list/str)[None]: String table column headers.
        delimiter (str)['\t']: Delimiter to separate column values by.
        head_delimiter (str)[' = ']: Delimiter to separate the head Name Value pairs saved before the table.
        row_delimiter (str)['\n']: Delimiter to separate rows by.
        serializer (callable/function)[None/str]: Function to convert the cell value to a string.
            This may be important for saving and loading files (json.dumps may be desired).
        head_serializer (callable/function)[None/str]: Function to convert the head values to a string.
            This may be important for saving and loading files (json.dumps may be desired).

    Returns:
        text (str): Text table. Example: "a1    b1    c1\na2    b2    c2"
    """
    if serializer is None:
        serializer = str
    if head_serializer is None:
        head_serializer = str

    # Lines to convert to text
    lines = []

    # Save Head values (Name Value pairs before the table).
    if isinstance(head, dict):
        head = ['{}{}{}'.format(str(k), head_delimiter, head_serializer(v))
                for k, v in head.items()]
        if len(head) > 0:
            head.append('')  # Add an extra space between the head values and the table.
        lines.extend(head)

    # Table header
    if isinstance(header, (list, tuple)):
        header = get_row_text(header, delimiter=delimiter)
    if isinstance(header, str):
        lines.append(header)

    # Save array table as lines
    lines.extend((get_row_text(row, delimiter=delimiter, serializer=serializer) for row in values))

    # Return the lines as text separated by \n
    return row_delimiter.join(lines)


def save_table(filename, values, head=None, header=None, delimiter=',', head_delimiter=' = ', row_delimiter='\n',
             serializer=None, head_serializer=None):
    """Save a text table to a file.

    Args:
        filename (str): Name of the file to save to.
        values (list/tuple/object): List of table values. Example: [(a1, b1, c1), (a2, b2, c2)].
        head (dict)[None]: Dictionary of Name Value pairs to save before the table.
        header (list/str)[None]: String table column headers.
        delimiter (str)['\t']: Delimiter to separate column values by.
        head_delimiter (str)[' = ']: Delimiter to separate the head Name Value pairs saved before the table.
        row_delimiter (str)['\n']: Delimiter to separate rows by.
        serializer (callable/function)[None/str]: Function to convert the cell value to a string.
            This may be important for saving and loading files (json.dumps may be desired).
        head_serializer (callable/function)[None/str]: Function to convert the head values to a string.
            This may be important for saving and loading files (json.dumps may be desired).
    """
    # Convert the data to text
    text = get_table_text(values, head=head, header=header,
                          delimiter=delimiter, head_delimiter=head_delimiter, row_delimiter=row_delimiter,
                          serializer=serializer, head_serializer=head_serializer)

    # Save the file.
    with open(filename, 'w') as f:
        f.write(text)


def text_to_table(text, delimiter=',', head_delimiter=' = ', row_delimiter='\n',
                  deserializer=None, head_deserializer=None):
    """Parse a table from the given text.

    Args:
        text (str): Text to parse.
        delimiter (str)['\t']: Delimiter to separate column values by.
        head_delimiter (str)[' = ']: Delimiter to separate the head Name Value pairs saved before the table.
        row_delimiter (str)['\n']: Delimiter to separate rows by.
        deserializer (callable/function)[None/str]: Function to convert the string to a python value.
            This may be important for loading files (json.loads may be desired).
        head_deserializer (callable/function)[None/str]: Function to convert the head string value to a python value.
            This may be important for loading files (json.loads may be desired).

    Returns:
        head (dict)[{}]: Dictionary of Name Value pairs to save before the table.
        header (list)[None]: List of string header column names. None if first sign of table looks like values.
        values (list/tuple): List of table values. Example: [(a1, b1, c1), (a2, b2, c2)].
    """
    head = {}
    header = None
    values = [tuple()] * 1000000  # Pre-initialize array for speed
    values_idx = 0

    # Convert text to lines
    if isinstance(text, str):
        lines = text.split(row_delimiter)
    else:
        lines = text

    # Parse the lines
    table_found = False
    for line in lines:
        if table_found:
            # Save the table data
            values[values_idx] = tuple(deserializer(v.strip()) for v in line.split(delimiter))
            values_idx += 1

            # Check to increment large array
            if values_idx >= len(values):
                values += [[]] * 1000000

        elif head_delimiter in line:
            # Save the head Name Value pairs
            name, value = line.split(head_delimiter, 1)
            head[name.strip()] = head_deserializer(value.strip())

        elif line.count(delimiter) > 1:
            # Check if header
            try:
                vals = tuple(deserializer(v.strip()) for v in line.split(delimiter))
                if all(isinstance(v, str) for v in vals):
                    if all(len(v) == 0 for v in vals):
                        raise RuntimeError('Ignore this row. It is empty')
                    raise ValueError('Save the column header names.')

                # Table values found. There is no header
                values[values_idx] = vals
                values_idx += 1
            except RuntimeError:
                continue  # All of the values were empty. Try to find the table again.
            except (ValueError, TypeError, Exception):
                # This is a header not table values
                header = [str(v.strip()) for v in line.split(delimiter)]
            table_found = True

    # Trim values
    values = values[:values_idx]

    # Return values
    return head, header, values


def parse_table(filename, delimiter=',', head_delimiter=' = ', row_delimiter='\n',
              deserializer=None, head_deserializer=None):
    """Parse a table from the file.

    Args:
        filename (str): Name of the file to save to.
        delimiter (str)['\t']: Delimiter to separate column values by.
        head_delimiter (str)[' = ']: Delimiter to separate the head Name Value pairs saved before the table.
        row_delimiter (str)['\n']: Delimiter to separate rows by.
        deserializer (callable/function)[None/str]: Function to convert the string to a python value.
            This may be important for loading files (json.loads may be desired).
        head_deserializer (callable/function)[None/str]: Function to convert the head string value to a python value.
            This may be important for loading files (json.loads may be desired).

    Returns:
        head (dict)[{}]: Dictionary of Name Value pairs to save before the table.
        header (list)[None]: List of string header column names. None if first sign of table looks like values.
        values (list/tuple): List of table values. Example: [(a1, b1, c1), (a2, b2, c2)].
    """
    with open(filename, 'r') as f:
        lines = f
        if row_delimiter != '\n' and row_delimiter != '\r\n':
            lines = f.read()
        return text_to_table(lines, delimiter=delimiter, head_delimiter=head_delimiter, row_delimiter=row_delimiter,
                             deserializer=deserializer, head_deserializer=head_deserializer)


class Table(Workbook):
    get_row_text = staticmethod(get_row_text)
    get_table_text = staticmethod(get_table_text)
    save_table = staticmethod(save_table)
    text_to_table = staticmethod(text_to_table)
    parse_table = staticmethod(parse_table)

