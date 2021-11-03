import string
import wrapt
from itertools import takewhile
from .dtypes import datetime, date, time


__all__ = ['CustomProperty', 'extract_single', 'is_iterable', 'decode_value', 'encode_value', 'excel_column_name',
           'Field', 'Item',
           'RangeItem', 'RowItem', 'ColumnItem', 'CellItem', 'ConstantItem',
           'Range', 'Row', 'Column', 'Cell', 'Constant', 'BuiltinDocumentPropertyItem', 'BuiltinDocumentProperty',
           'DateTime', 'Date', 'Time'
           ]


class CustomProperty(object):
    def __init__(self, fget=None, fset=None, fdel=None, doc=None):
        super().__init__()
        if fget is not None:
            self.fget = fget
        if fset is not None:
            self.fset = fset
        if fdel is not None:
            self.fdel = fdel

        if doc is None and self.fget is not None:
            doc = self.fget.__doc__
        self.__doc__ = doc

    def __get__(self, instance, owner):
        if instance is None:
            return self
        if self.fget is None:
            raise AttributeError("unreadable attribute")
        return self.fget(instance)

    def __set__(self, instance, value):
        if self.fset is None:
            raise AttributeError("can't set attribute")
        self.fset(instance, value)

    def __delete__(self, instance):
        if self.fdel is None:
            raise AttributeError("can't delete attribute")
        self.fdel(instance)

    def getter(self, fget):
        self.fget = fget
        return self

    def setter(self, fset):
        self.fset = fset
        return self

    def deleter(self, fdel):
        self.fdel = fdel
        return self

    def fget(self, instance):
        text = 'Example:\n' \
               '    class MyFile(Table):\n' \
               '        field = Field()\n' \
               '        @field.getter\n' \
               '        def field(self):\n' \
               '            sheet = self.wb.Sheets("Sheet1")\n' \
               '            return sheet.Range("A1:A2").Value\n' \
               'unreadable attribute\n'
        raise AttributeError(text)

    def fset(self, instance, value):
        text = 'Example:\n' \
               '    class MyFile(Table):\n' \
               '        field = Field()\n' \
               '        @field.setter\n' \
               '        def field(self, value):\n' \
               '            sheet = self.wb.Sheets("Sheet1")\n' \
               '            sheet.Range("A1:A2").Value = value\n' \
               'can\'t set attribute\n'
        raise AttributeError(text)

    def fdel(self, instance):
        text = 'Example:\n' \
               '    class MyFile(Table):\n' \
               '        field = Field()\n' \
               '        @field.getter\n' \
               '        def field(self):\n' \
               '            sheet = self.wb.Sheets("Sheet1")\n' \
               '            sheet.Range("A1:A2").Value = None\n' \
               'can\'t delete attribute\n'
        raise AttributeError(text)


def extract_single(value):
    """Return a single value if the given value is of len 1."""
    try:
        if len(value) == 1:
            return value[0]
    except(ValueError, TypeError, Exception):
        pass
    return value


def is_iterable(value, allow_str=False):
    """Return if iterable."""
    try:
        if isinstance(value, str) and not allow_str:
            return False

        iter(value)
        return True
    except(ValueError, TypeError, Exception):
        return False


def is_contiguous(items):
    """Return if the given items are contiguous."""
    for i in range(len(items)-1):
        if (items[i] - items[i+1]) > 1:
            return False
    return True


def excel_column_name(col):
    """Take in a number and return the excel column name."""
    if isinstance(col, int):
        num = col
        col = ''
        while num > 0:
            num, r = divmod(num-1, 26)
            col = string.ascii_uppercase[r] + col
    return col


class DecodedObject(wrapt.ObjectProxy):
    """Custom Excel Item object that has array and list methods.

    Args:
        item (Excel Item/object): Excel Item object (Row, Column, Range, Cell)

    Returns:
        value (DecodedObject): Python item/object/value that was read
    """
    def __init__(self, wrapped):
        self._self_xl_parent = None
        super().__init__(wrapped)

    def to_tuple(self):
        """Return a tuple of contiguous values."""
        parent = self._self_xl_parent

        if self.Areas.Count > 1:
            # Could have multiple areas where .Value will only return the first Area (Range) value.
            value = tuple(area.Value for area in self.Areas)
        else:
            value = self.Value

        if is_iterable(value, allow_str=False):
            value = extract_single(value)
            if is_iterable(value, allow_str=False):
                # Extract single converts single column to single value
                value = tuple(takewhile(lambda x: extract_single(x) is not None, value))
                value = tuple(extract_single(val) for val in value)
        return value

    def to_list(self):
        """Return a list of contiguous values."""
        return list(self.to_tuple())

    def length(self):
        """Return the length of contiguous values."""
        return len(self.to_tuple())

    def pop(self, idx=None):
        """Pop off an item at the end."""
        parent = self._self_xl_parent
        if idx is None:
            idx = self.length() + 1  # Excel has an index 1 offset

        idx -= 1
        if parent.cols:
            val = self.Cells(idx, 1).Value
            self.Cells(idx, 1).Value = None
        else:
            val = self.Cells(1, idx).Value
            self.Cells(1, idx).Value = None
        return val

    def append(self, value):
        """Append an item to the end of the list."""
        parent = self._self_xl_parent
        last_idx = self.length() + 1  # Excel has an index 1 offset
        if parent.cols:
            self.Cells(last_idx, 1).Value = value
        else:
            self.Cells(1, last_idx).Value = value

    def extend(self, value):
        """Extend an item with the given list."""
        parent = self._self_xl_parent
        last_idx = self.length() + 1  # Excel has an index 1 offset
        if parent.cols:
            for v in value:
                self.Cells(last_idx, 1).Value = v
                last_idx += 1
        else:
            for v in value:
                self.Cells(1, last_idx).Value = v
                last_idx += 1


def decode_value(item):
    """Convert an Excel item object into a Python value.

    Args:
        item (Excel Item/object): Excel Item object (Row, Column, Range, Cell)

    Returns:
        value (object/str/float/int): Python item/object/value that was read
    """
    if item.Areas.Count > 1:
        # Could have multiple areas where .Value will only return the first Area (Range) value.
        value = tuple(area.Value for area in item.Areas)
    else:
        value = item.Value

    if is_iterable(value, allow_str=False):
        value = extract_single(value)
        if is_iterable(value, allow_str=False):
            # Extract single converts single column to single value
            value = tuple(extract_single(val) for val in value)
    return value


def encode_value(item, value):
    """Set the given excel item object with the given value.

    Args:
        item (Excel Item/object): Excel Item object (Row, Column, Range, Cell)
        value (object/str/float/int): Value to set
    """
    if not is_iterable(value, allow_str=False):
        item.Value = value
    else:
        # Create Item Cell index iterator
        cell = iter(item.Cells)
        for r, obj in enumerate(value):
            if is_iterable(obj, allow_str=False):
                for c, val in enumerate(obj):
                    next(cell).Value = val
            else:
                next(cell).Value = obj


class Field(CustomProperty):
    def __init__(self, sheet=1, dtype=None, decoder=None, encoder=None):
        self.sheet = sheet
        self._dtype = None
        self.orig_decode = self.decode
        self.orig_encode = self.encode
        if decoder is not None:
            self.decode = decoder
        if encoder is not None:
            self.encode = encoder
        super().__init__()

        if dtype is not None:
            self.set_dtype(dtype)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        sheet = instance.get_sheet(self.sheet)
        return sheet

    def get_dtype(self):
        """Return the data type which should have an encode and a decode function."""
        return self._dtype

    def set_dtype(self, dtype):
        """Set the data type which should have an encode and a decode function."""
        self._dtype = dtype

        # Set the encoder
        if hasattr(self._dtype, 'encode'):
            self.encoder(self._dtype.encode)

        # Set the decoder
        if hasattr(self._dtype, 'decode'):
            self.decoder(self._dtype.decode)
        elif not hasattr(self._dtype, 'encode') and callable(self._dtype):
            # If callable function assume it is a decoder (take in an item.Value and return a python object)
            def decode(item):
                return self._dtype(item.Value)
            self.decoder(decode)

    dtype = property(get_dtype, set_dtype)

    decode = staticmethod(DecodedObject)
    encode = staticmethod(encode_value)

    def decoder(self, func):
        """Change the decode function to a custom function.

        This should take in an Excel item object and return a Python value.
        """
        if func is None:
            func = self.orig_decode
        self.decode = func

    def encoder(self, func):
        """Change the encode function to a custom function.

        This should take in an excel item object and Python value setting the value to excel item object
        """
        if func is None:
            func = self.orig_encode
        self.encode = func


def create_range_property(name):
    attr = '_'+str(name)

    def get(self):
        return getattr(self, attr, None)

    def set(self, value):
        try:
            self.range_str = self.get_range_str(**{name: value})
        except Exception as err:
            raise ValueError('Invalid value given {} for {}!'.format(value, name)) from err
        setattr(self, attr, value)

    return property(get, set)


class Item(Field):
    def __init__(self, cells=None, rows=None, row_length=None, cols=None, col_length=None, ranges=None,
                 sheet=1, dtype=None, decoder=None, encoder=None):
        self.range_str = None
        self._cells = cells
        self._rows = rows
        self._row_length = row_length
        self._cols = cols
        self._col_length = col_length
        self._ranges = ranges

        super().__init__(sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)
        self.range_str = self.get_range_str()

    cells = create_range_property('cells')
    rows = create_range_property('rows')
    row_length = create_range_property('row_length')
    cols = create_range_property('cols')
    col_length = create_range_property('col_length')
    ranges = create_range_property('ranges')

    @staticmethod
    def is_cell(item):
        try:
            return hasattr(item, 'Address') or \
                    (hasattr(item, 'Row') and hasattr(item, 'Column')) or \
                    (isinstance(item, str) and ':' not in item and ',' not in item) or \
                    (len(item) >= 2 and isinstance(item[0], int) and isinstance(item[1], (int, str)))
        except (TypeError, ValueError, Exception):
            pass
        return False

    @staticmethod
    def get_cell_str(item):
        if isinstance(item, str):
            return str(item)
        elif hasattr(item, 'Address'):
            return str(item.Address)
        elif hasattr(item, 'Row') and hasattr(item, 'Column'):
            return '${}${}'.format(excel_column_name(item.Column), item.Row)
        else:
            return '${}${}'.format(excel_column_name(item[1]), item[0])

    def get_range_str(self, **kwargs):
        """Return the string for the Range.

        Args:
            cells (tuple/Cell)[None]: Single Cell position (int, (int,st)) or list of Cells (or single cell position).
            rows (tuple/str/int)[None]: Rows for these items.
            row_length (int)[None]: Length of the rows.
            cols (tuple/str/int/)[None]: Cols for these items.
            col_length (int)[None]: Length of the columns.
            ranges (tuple/str/int)[None]: Range for these items.

        Returns:
            range_str (str): String to make the Range with.
        """
        # Get Args
        cells = kwargs.get('cells', self.cells)
        rows = kwargs.get('rows', self.rows)
        row_length = kwargs.get('row_length', self.row_length)
        cols = kwargs.get('cols', self.cols)
        col_length = kwargs.get('col_length', self.col_length)
        ranges = kwargs.get('ranges', self.ranges)

        # Create the initial ranges string
        range_str = []
        if ranges is not None:
            if isinstance(ranges, str):
                ranges = ranges.split(',')
            elif not is_iterable(ranges):
                ranges = [ranges]

            # Check if two cells given for range
            if is_iterable(ranges) and len(ranges) == 2 and self.is_cell(ranges[0]) and self.is_cell(ranges[1]):
                addr = '{}:{}'.format(self.get_cell_str(ranges[0]), self.get_cell_str(ranges[1]))
                range_str.append(addr)
            else:
                for r in ranges:
                    if isinstance(r, str):
                        range_str.append(r)
                    elif hasattr(r, 'Address'):
                        range_str.append(r.Address)
                    elif is_iterable(r) and len(r) == 2 and self.is_cell(ranges[0]) and self.is_cell(ranges[1]):
                        addr = '{}:{}'.format(self.get_cell_str(ranges[0]), self.get_cell_str(ranges[1]))
                        range_str.append(addr)

        # Add columns to the ranges
        if cols is not None:
            if col_length is None:
                col_length = 16840  # Default column size in excel

            if not is_iterable(cols):
                cols = [cols]

            if is_contiguous(cols):
                addr = '${0}$1:${1}${2}'.format(excel_column_name(cols[0]), excel_column_name(cols[-1]), col_length)
                range_str.append(addr)
            else:
                for c in cols:
                    if hasattr(c, 'Address'):
                        addr = str(c.Address)
                    else:
                        addr = '${0}$1:${0}${1}'.format(excel_column_name(c), col_length)
                    range_str.append(addr)

        if rows is not None:
            if row_length is None:
                row_length = 702  # Default row size in excel ($ZZ)

            if not is_iterable(rows):
                rows = [rows]

            if is_contiguous(rows):
                addr = '$A${0}:${2}${1}'.format(rows[0], rows[-1], excel_column_name(row_length))
                range_str.append(addr)
            else:
                for r in rows:
                    if hasattr(r, 'Address'):
                        addr = str(r.Address)
                    else:
                        addr = '$A${0}:${1}${0}'.format(r, excel_column_name(row_length))
                    range_str.append(addr)

        if cells is not None:
            if self.is_cell(cells) or not is_iterable(cells):
                cells = [cells]
            range_str.extend([self.get_cell_str(cell) for cell in cells])

        return ', '.join(range_str)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        sheet = instance.get_sheet(self.sheet)
        items = sheet.Range(self.range_str)
        return items

    def fget(self, instance):
        """Return the Range Item Object"""
        item = self.get_item(instance)
        item = self.decode(item)
        if hasattr(item, '_self_xl_parent'):
            setattr(item, '_self_xl_parent', self)
        return item

    def fset(self, instance, value):
        """Set the Range value."""
        item = self.get_item(instance)
        self.encode(item, value)

    def fdel(self, instance):
        """Delete the range."""
        item = self.get_item(instance)
        item.Delete()


class RangeItem(Item):
    """Everything in excel is essentially a Range."""
    def __init__(self, *ranges, sheet=1, dtype=None, decoder=None, encoder=None):
        super().__init__(ranges=ranges, sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        if self.ranges is None or (is_iterable(self.ranges) and len(self.ranges) == 0):
            item = instance.get_sheet(self.sheet).Range
        else:
            item = super().get_item(instance)  # instance.get_sheet(self.sheet).Range(self.ranges)
        return item


class Range(RangeItem):
    """Range Value"""
    decode = staticmethod(decode_value)
    encode = staticmethod(encode_value)


class RowItem(Item):
    def __init__(self, *rows, row_length=None, sheet=1, dtype=None, decoder=None, encoder=None):
        super().__init__(rows=rows, row_length=row_length, sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        # Get the Row(s)
        if self.rows is None or (is_iterable(self.rows) and len(self.rows) == 0):
            item = instance.get_sheet(self.sheet).Rows
        else:
            item = super().get_item(instance)  # instance.get_sheet(self.sheet).Rows(self.rows)
        return item


class Row(RowItem):
    """Row Value"""
    decode = staticmethod(decode_value)
    encode = staticmethod(encode_value)


class ColumnItem(Item):
    def __init__(self, *cols, col_length=None, sheet=1, dtype=None, decoder=None, encoder=None):
        super().__init__(cols=cols, col_length=col_length, sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        # Get the Column(s)
        if self.cols is None or (is_iterable(self.cols) and len(self.cols) == 0):
            item = instance.get_sheet(self.sheet).Columns
        else:
            item = super().get_item(instance)  # instance.get_sheet(self.sheet).Columns(self.cols)
        return item


class Column(ColumnItem):
    """Column Value"""
    decode = staticmethod(decode_value)
    encode = staticmethod(encode_value)


class CellItem(Item):
    def __init__(self, *cells, sheet=1, dtype=None, decoder=None, encoder=None):
        super().__init__(cells=cells, sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        if self.cells is None or (is_iterable(self.cells) and len(self.cells) == 0):
            item = instance.get_sheet(self.sheet).Cells
        else:
            item = super().get_item(instance)  # instance.get_sheet(self.sheet).Cells(*self.cells)
        return item


class Cell(CellItem):
    """Cell Value"""
    decode = staticmethod(decode_value)
    encode = staticmethod(encode_value)


class ConstantItem(Item):
    """Constant Value that is set when the table is initialized"""
    def __init__(self, value, *cells, rows=None, row_length=None, cols=None, col_length=None, ranges=None,
                 sheet=1, dtype=None, decoder=None, encoder=None):
        self.value = value
        super().__init__(cells=cells, rows=rows, row_length=row_length, cols=cols, col_length=col_length, ranges=ranges,
                         sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def init_table(self, instance):
        """Set the table value with this value on table initialization."""
        self.fset(instance, self.value)


class Constant(ConstantItem):
    """Constant Value that is set when the table is initialized"""
    decode = staticmethod(decode_value)
    encode = staticmethod(encode_value)


class BuiltinDocumentPropertyItem(Item):
    def __init__(self, name, sheet=1, dtype=None, decoder=None, encoder=None):
        self.name = name
        super().__init__(sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)

    def get_item(self, instance):
        """Return the item for the instance and settings."""
        return instance.wb.BuiltinDocumentProperties(self.name)


class BuiltinDocumentProperty(BuiltinDocumentPropertyItem):
    def decode(self, item):
        """Convert an Excel item object into a Python value.

        Args:
            item (Excel Item/object): Excel Item object (Row, Column, Range, Cell)

        Returns:
            value (object/str/float/int): Python item/object/value that was read
        """
        return item.Value

    def encode(self, item, value):
        """Set the given excel item object with the given value.

        Args:
            item (Excel Item/object): Excel Item object (Row, Column, Range, Cell)
            value (object/str/float/int): Value to set
        """
        item.Value = value


class DateTime(Item):
    def __init__(self, *cells, rows=None, row_length=None, cols=None, col_length=None, ranges=None,
                 sheet=1, dtype=None, decoder=None, encoder=None, str_format=None, formats=None):

        if dtype is None and decoder is None and encoder is None:
            dtype = datetime(str_format=str_format, formats=formats)

        super().__init__(cells=cells, rows=rows, row_length=row_length, cols=cols, col_length=col_length, ranges=ranges,
                         sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)


class Date(Item):
    def __init__(self, *cells, rows=None, row_length=None, cols=None, col_length=None, ranges=None,
                 sheet=1, dtype=None, decoder=None, encoder=None, str_format=None, formats=None):

        if dtype is None and decoder is None and encoder is None:
            dtype = date(str_format=str_format, formats=formats)

        super().__init__(cells=cells, rows=rows, row_length=row_length, cols=cols, col_length=col_length, ranges=ranges,
                         sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)


class Time(Item):
    def __init__(self, *cells, rows=None, row_length=None, cols=None, col_length=None, ranges=None,
                 sheet=1, dtype=None, decoder=None, encoder=None, str_format=None, formats=None):

        if dtype is None and decoder is None and encoder is None:
            dtype = time(str_format=str_format, formats=formats)

        super().__init__(cells=cells, rows=rows, row_length=row_length, cols=cols, col_length=col_length, ranges=ranges,
                         sheet=sheet, dtype=dtype, decoder=decoder, encoder=encoder)
