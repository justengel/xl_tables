"""
Constants taken from win32com.client. Import this after gencache (excel_attributes.py has been imported).
"""
import sys
import contextlib
import win32com.client as com


__all__ = ['com']


with contextlib.suppress(AttributeError, Exception):
    MY_MODULE = sys.modules[__name__]

    class ConstantsModule(MY_MODULE.__class__):
        def __dir__(self):
            return self.__all__

    # Override the module make it callable
    try:
        MY_MODULE.__class__ = ConstantsModule  # Override __class__ (Python 3.6+)
        MY_MODULE.__doc__ = ConstantsModule.__call__.__doc__
    except (TypeError, Exception):
        # < Python 3.6 Create the module and make the attributes accessible
        sys.modules[__name__] = MY_MODULE = ConstantsModule(__name__)
        for ATTR in __all__:
            setattr(MY_MODULE, ATTR, vars()[ATTR])

    consts = {k: v for k, v in com.constants.__dict__['__dicts__'][0].items()
              if not k.startswith('_')}
    MY_MODULE.__dict__.update(consts)
    MY_MODULE.__all__ += list(consts.keys())
