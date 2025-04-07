from .__meta__ import version as __version__

from .prop_utils import ProxyProperty, ProxyMethod, HashDict

from .dtypes import (
    datetime,
    date,
    time,
    datetime_fmt_to_excel,
    DATETIME_FORMATS,
    TIME_FORMATS,
    DATETIME_FORMATS,
)

from .fields import (
    CustomProperty,
    extract_single,
    is_iterable,
    decode_value,
    encode_value,
    excel_column_name,
    Field,
    Item,
    RangeItem,
    RowItem,
    ColumnItem,
    CellItem,
    ConstantItem,
    Range,
    Row,
    Column,
    Cell,
    Constant,
    BuiltinDocumentPropertyItem,
    BuiltinDocumentProperty,
    DateTime,
    Date,
    Time,
    get_row_text,
    get_table_text,
    save_table,
    text_to_table,
    parse_table,
)

from .windows import (
    CAN_USE_WINDOWS_EXCEL,
    constants as windows_constants,
    Excel as WindowsExcel,
    Workbook as WindowsWorkbook,
    Table as WindowsTable,
    should_init_sig as windows_should_init_sig,
    set_init_sig as windows_set_init_sig,
    init_sig_shutdown as windows_init_sig_shutdown,
    shutdown as windows_shutdown,
)

from .openpyxl_support import (
    CAN_USE_OPENPYXL,
    constants as openpyxl_constants,
    Excel as OpenpyxlExcel,
    Workbook as OpenpyxlWorkbook,
    Table as OpenpyxlTable,
    should_init_sig as openpyxl_should_init_sig,
    set_init_sig as openpyxl_set_init_sig,
    init_sig_shutdown as openpyxl_init_sig_shutdown,
    shutdown as openpyxl_shutdown,
)

# Set initial values
constants = openpyxl_constants
Excel = OpenpyxlExcel
Workbook = OpenpyxlWorkbook
Table = OpenpyxlTable
should_init_sig = openpyxl_should_init_sig
set_init_sig = openpyxl_set_init_sig
init_sig_shutdown = openpyxl_init_sig_shutdown
shutdown = openpyxl_shutdown

USING_OPENPYXL = True
USING_EXCEL = not USING_OPENPYXL


def change_backend(windows):
    global USING_OPENPYXL, USING_EXCEL, constants, Excel, Workbook, Table
    global should_init_sig, set_init_sig, init_sig_shutdown, shutdown
    import atexit

    if windows:
        if not CAN_USE_WINDOWS_EXCEL:
            raise RuntimeError("Cannot use Windows backend. Windows Excel is not availabel!")
        
        constants = windows_constants
        Excel = WindowsExcel
        Workbook = WindowsWorkbook
        Table = WindowsTable
        should_init_sig = windows_should_init_sig
        set_init_sig = windows_set_init_sig
        init_sig_shutdown = windows_init_sig_shutdown
        shutdown = windows_shutdown
            
        # Automatically register the shutdown function with atexit
        atexit.register(windows_shutdown, sys_exit=False)

        USING_EXCEL = True
        USING_OPENPYXL = not USING_EXCEL

    else:
        constants = openpyxl_constants
        Excel = OpenpyxlExcel
        Workbook = OpenpyxlWorkbook
        Table = OpenpyxlTable
        should_init_sig = openpyxl_should_init_sig
        set_init_sig = openpyxl_set_init_sig
        init_sig_shutdown = openpyxl_init_sig_shutdown
        shutdown = openpyxl_shutdown

        # Automatically register the shutdown function with atexit
        atexit.unregister(windows_shutdown)

        USING_OPENPYXL = True
        USING_EXCEL = not USING_OPENPYXL


# Change the default backend (prefer windows)
change_backend(windows=CAN_USE_WINDOWS_EXCEL)
