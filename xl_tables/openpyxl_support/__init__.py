from .fake_excel import Excel
from .workbook import Workbook, should_init_sig, set_init_sig, init_sig_shutdown, shutdown
from .table import Table

CAN_USE_OPENPYXL = True