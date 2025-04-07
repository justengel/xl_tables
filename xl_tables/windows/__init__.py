try:
    from .excel_attributes import Excel, Workbook, \
        is_excel_installed, should_init_sig, set_init_sig, init_sig_shutdown, shutdown

    from .table import Table

    from .constants import *
    
    CAN_USE_WINDOWS_EXCEL = is_excel_installed()
except (ImportError, Exception):
    CAN_USE_WINDOWS_EXCEL = False


if not CAN_USE_WINDOWS_EXCEL:
    CAN_USE_WINDOWS_EXCEL = False

    def raise_not_available(*args, **kwargs):
        raise RuntimeError("Windows Excel is not available!")
    

    class NotAvailable:
        __new__ = raise_not_available
            
        
    Excel = NotAvailable
    Workbook = NotAvailable
    Table = NotAvailable
    should_init_sig = raise_not_available
    set_init_sig = raise_not_available
    init_sig_shutdown = raise_not_available
    shutdown = raise_not_available