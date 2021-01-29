import datetime as dt_module
from dynamicmethod import dynamicmethod
from collections import OrderedDict


__all__ = ['datetime', 'date', 'time',
           'make_datetime', 'str_datetime',
           'DATETIME_FORMATS', 'TIME_FORMATS', 'DATETIME_FORMATS']


DATE_FORMATS = [
    '%m/%d/%Y', '%Y-%m-%d', '%y-%m-%d', '%m/%d/%y',  # '2019-04-17', '04/17/2019'
    '%b %d %Y', '%b %d, %Y',  # 'Apr 17 2019', 'Apr 17, 2019'
    '%d %b %Y', '%d %b, %Y',  # '17 Apr 2019', '17 Apr, 2019'
    '%B %d %Y', '%B %d, %Y',  # 'April 17 2019', 'April 17, 2019'
    '%d %B %Y', '%d %B, %Y',  # '17 April 2019', '17 April, 2019'
    ]

TIME_FORMATS = [
    '%H:%M:%S',        # '14:24:55'
    '%I:%M:%S %p',     # '02:24:55 PM'
    '%I:%M:%S.%f %p',  # '02:24:55.000200 PM'
    '%I:%M %p',        # '02:24 PM'
    '%H:%M:%S.%f',     # '14:24:55.000200'
    '%H:%M',           # '14:24'
    '%H:%M:%S+00:00',        # '14:24:55+00:00'
    ]

DATETIME_FORMATS = [d + ' ' + t for t in TIME_FORMATS for d in DATE_FORMATS] + DATE_FORMATS + TIME_FORMATS


DATETIME_FORMAT_TO_EXCEL = OrderedDict([
    ('%a', 'ddd'), ('%A', 'dddd'), ('%w', 'ddd'), ('%d', 'dd'),
    ('%b', 'mmm'), ('%B', 'mmmm'), ('%m', 'mm'),
    ('%y', 'yy'), ('%Y', 'yyyy'),
    ('%H', 'hh'), ('%I', 'h'), ('%p', 'AM/PM'),
    ('%M', 'mm'),
    ('%S', 'ss'),
    ('%f', '00'),
    ('%z', ''), ('%Z', ''), ('%j', ''), ('%U', ''), ('%W', ''),
    ('%c', 'ddd mmm dd hh:mm:ss yyyy'),
    ('%x', 'mm/dd/yyyy'),
    ('%X', 'hh:mm:ss'),
    ('%%', '%'),
    ])


def make_datetime(dt_string, formats=None):
    """Make the datetime from the given date time string.
    Args:
        dt_string (str): Datetime string '04:00 PM' ...
        formats (list)[None]: List of acceptable datetime string formats.
    Returns:
        dt (datetime.datetime): Datetime object or None.
    """
    if not isinstance(dt_string, str):
        return dt_string

    if formats is None:
        formats = DATETIME_FORMATS

    for fmt in formats:
        try:
            return dt_module.datetime.strptime(dt_string, fmt)
        except (TypeError, ValueError, Exception):
            pass

    try:  # Try ISO format
        return dt_module.datetime.fromisoformat(dt_string)
    except (TypeError, ValueError, AttributeError, Exception):
        pass

    raise ValueError('Invalid datetime format {}. Allowed formats are {}'.format(repr(dt_string), repr(formats)))


def str_datetime(dt, formats=None):
    """Return the datetime as a string."""
    if isinstance(dt, str):
        return dt

    if isinstance(formats, str):
        return dt.strftime(formats)

    if formats is None:
        formats = DATETIME_FORMATS
    return dt.strftime(formats[0])


def datetime_fmt_to_excel(fmt):
    for k, v in DATETIME_FORMAT_TO_EXCEL.items():
        fmt = fmt.replace(k, v)
    return fmt


DEFAULT_DT = dt_module.datetime.utcfromtimestamp(0)


class DtMixin(object):
    str_format = DATETIME_FORMATS[0]
    formats = DATETIME_FORMATS

    DEFAULT_PARAMS = OrderedDict([('year', DEFAULT_DT.year), ('month', DEFAULT_DT.month), ('day', DEFAULT_DT.day),
                                  ('hour', DEFAULT_DT.hour), ('minute', DEFAULT_DT.minute),
                                  ('second', DEFAULT_DT.second), ('microsecond', DEFAULT_DT.microsecond),
                                  ('tzinfo', DEFAULT_DT.tzinfo), ('fold', DEFAULT_DT.fold)])

    def __init__(self, dt=None, *args, str_format=None, formats=None, **kwargs):
        super().__init__()  # Object init

    @classmethod
    def get_params(cls, defaults, dt, args, kwargs, formats=None):
        """Return a list of keyword arguments for the given positional, keyword and default values.

        Args:
            defaults (OrderedDict): Ordered dict of key word names with default values. None value means optional.
            dt (object/int)[None]: Object datetime, date, time or integer year or hour is assumed.
            args (tuple): Positional arguments
            kwargs (dict): Key word arguments
            formats (list)[None]: List of acceptable datetime string formats.

        Returns:
            params (dict): Dictionary of mapped name arguments.
        """
        defaults = defaults.copy()
        params = {}

        # If datetime get the values as default values (override with kwargs for duplicate values)
        if isinstance(dt, str):
            dt = make_datetime(dt, formats or cls.formats)
        elif isinstance(dt, (int, float)) and len(args) + len(kwargs) < 2:
            dt = dt_module.datetime.utcfromtimestamp(dt)
        elif isinstance(dt, (int, float)):
            # Assume integer is positional argument year or hour.
            args = (dt,) + args
            dt = None

        # Get datetime attributes
        if isinstance(dt, (dt_module.datetime, dt_module.date, dt_module.time)):
            for name in defaults:
                value = getattr(dt, name, None)
                if value is not None:
                    defaults[name] = value

        # Find positional values and keyword values
        arg_len = len(args)
        for i, name in enumerate(defaults.keys()):
            if i < arg_len:
                params[name] = args[i]
            else:
                value = kwargs.get(name, defaults.get(name, None))
                if value is not None:
                    params[name] = value
        return params

    @classmethod
    def get_init_formats(cls, str_format=None, formats=None):
        """Get the str_format and formats from the initial args."""
        if formats is None:
            formats = cls.formats
        if str_format is None:
            str_format = cls.str_format

        if isinstance(str_format, (list, tuple)):
            formats = str_format
            str_format = str_format[0]
        elif str_format is None:
            str_format = formats[0]
        return str_format, formats

    @dynamicmethod  # Run as a classmethod or instancemethod
    def decode(self, item):
        # Get the class object
        cls = self
        if isinstance(self, (dt_module.datetime, dt_module.date, dt_module.time)):
            cls = self.__class__

        # Get the item value
        try:
            value = item.Value
            if isinstance(value, (int, float)) and value < 1:
                # Convert excel hours to seconds? 86400 seconds == 24 hr?  Only gets here from time value.
                value = value * 86400
        except (ValueError, TypeError, AttributeError, Exception):
            value = item

        # Convert the value to a datetime object
        if not isinstance(value, cls):
            value = cls(value, str_format=self.str_format, formats=self.formats)

        return value

    @dynamicmethod  # Run as a classmethod or instancemethod
    def encode(self, item, value):
        # Get the class object
        cls = self
        if isinstance(self, (dt_module.datetime, dt_module.date, dt_module.time)):
            cls = self.__class__

        # Convert to this object type
        if not isinstance(value, datetime):
            value = cls(dt=value, str_format=self.str_format, formats=self.formats)

        item.NumberFormat = datetime_fmt_to_excel(self.str_format or self.formats[0])
        item.Value = str(value)

    def __str__(self):
        return str_datetime(self, self.str_format or self.formats)


class datetime(dt_module.datetime, DtMixin):
    formats = DATETIME_FORMATS
    str_format = DATETIME_FORMATS[0]

    DEFAULT_PARAMS = OrderedDict([('year', DEFAULT_DT.year), ('month', DEFAULT_DT.month), ('day', DEFAULT_DT.day),
                                  ('hour', DEFAULT_DT.hour), ('minute', DEFAULT_DT.minute),
                                  ('second', DEFAULT_DT.second), ('microsecond', DEFAULT_DT.microsecond),
                                  ('tzinfo', DEFAULT_DT.tzinfo), ('fold', DEFAULT_DT.fold)])

    def __new__(cls, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Create the datetime object.

        Args:
            dt (int/float/str/datetime): Datetime, str datetime, timestamp, or year positional argument.
            *args (tuple): Positional datetime arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of datetime keyword arguments.
        """
        # Get the parameters and their defaults
        params = cls.get_params(cls.DEFAULT_PARAMS, dt, args, kwargs, formats)

        # Create this object type
        dt = super().__new__(cls, **params)
        dt.str_format, dt.formats = dt.get_init_formats(str_format, formats)
        return dt  # Return will run __init__

    def __init__(self, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Initialize the datetime object.

        Args:
            dt (int/float/str/datetime): Datetime, str datetime, timestamp, or year positional argument.
            *args (tuple): Positional datetime arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of datetime keyword arguments.
        """
        super().__init__()

    def __str__(self):
        return str_datetime(self, self.str_format or self.formats)


class date(datetime):  # (dt_module.date, DtMixin):  # excel cannot use date use datetime with different format.
    formats = DATETIME_FORMATS
    str_format = DATE_FORMATS[0]

    DEFAULT_PARAMS = OrderedDict([('year', DEFAULT_DT.year), ('month', DEFAULT_DT.month), ('day', DEFAULT_DT.day)])

    def __new__(cls, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Create the date object.

        Args:
            dt (int/float/str/datetime): date, str date, timestamp, or year positional argument.
            *args (tuple): Positional date arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of date keyword arguments.
        """
        # Get the parameters and their defaults
        params = cls.get_params(cls.DEFAULT_PARAMS, dt, args, kwargs)

        # Create this object type
        dt = super().__new__(cls, **params)
        dt.str_format, dt.formats = dt.get_init_formats(str_format, formats)
        return dt  # Return will run __init__

    def __init__(self, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Initialize the date object.

        Args:
            dt (int/float/str/datetime): date, str date, timestamp, or year positional argument.
            *args (tuple): Positional date arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of date keyword arguments.
        """
        super().__init__()

    def __str__(self):
        return str_datetime(self, self.str_format or self.formats)


class time(dt_module.time, DtMixin):
    formats = TIME_FORMATS
    str_format = TIME_FORMATS[0]

    DEFAULT_PARAMS = OrderedDict([('hour', DEFAULT_DT.hour), ('minute', DEFAULT_DT.minute),
                                  ('second', DEFAULT_DT.second), ('microsecond', DEFAULT_DT.microsecond)])

    def __new__(cls, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Create the time object.

        Args:
            dt (int/float/str/datetime): time, str time, timestamp, or hour positional argument.
            *args (tuple): Positional time arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of time keyword arguments.
        """
        # Get the parameters and their defaults
        params = cls.get_params(cls.DEFAULT_PARAMS, dt, args, kwargs)

        # Create this object type
        dt = super().__new__(cls, **params)
        dt.str_format, dt.formats = dt.get_init_formats(str_format, formats)
        return dt  # Return will run __init__

    def __init__(self, dt=None, *args, str_format=None, formats=None, **kwargs):
        """Initialize the time object.

        Args:
            dt (int/float/str/datetime): time, str time, timestamp, or hour positional argument.
            *args (tuple): Positional time arguments.
            str_format (str)[None]: String format to convert the object to a string with.
            formats (list)[None]: List of string formats to parse and decode information with.
            **kwargs (dict): Dictionary of time keyword arguments.
        """
        super().__init__()

    def __str__(self):
        return str_datetime(self, self.str_format or self.formats)
