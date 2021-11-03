"""
Typical hierarchy for Microsoft Excel is Excel Application > Workbooks > Workbook > Sheet.

https://docs.microsoft.com/en-us/office/vba/api/excel.workbook
"""
import os
import sys
import atexit
import signal
import shutil
import win32com.client
from .constants import populate_constants
from .prop_utils import ProxyProperty, ProxyMethod, HashDict, ItemStorage
from .fields import ConstantItem

__all__ = ['Excel', 'Workbook',
           'should_init_sig', 'set_init_sig', 'init_sig_shutdown', 'shutdown']


def is_gencache_available():
    """Determine if make_py has been run (i.e. early binding is setup).

    Early binding helps get class/object documentation in development.
    """
    try:
        return win32com.client.gencache.GetClassForProgID('Excel.Application') is not None
    except (AttributeError, Exception):
        return False


def del_gencache():
    """Delete the gencache to create it again."""
    try:
        shutil.rmtree(os.path.expanduser('~\\AppData\\Local\\Temp\\gen_py'))
    except (PermissionError, OSError, FileNotFoundError, Exception):
        pass


def init_gencache():
    """Initialize the gencache for early binding."""
    if not is_gencache_available():
        try:  # Create the gencache application and delete it
            gen = win32com.client.gencache.EnsureDispatch('Excel.Application')
            gen.Quit()
            del gen
        except (AttributeError, TypeError, ValueError, Exception):
            try:  # Create the gencache application and delete it
                del_gencache()

                gen = win32com.client.gencache.EnsureDispatch('Excel.Application')
                gen.Quit()
                del gen
            except (AttributeError, TypeError, ValueError, Exception):
                pass


# Call init gencache. Required for constants!
init_gencache()
populate_constants()


SETTINGS = {
    'is_parallel': True,  # Custom setting

    # 'ActiveCell': None,
    # 'ActiveChart': None,
    # 'ActiveEncryptionSession': -1,
    # 'ActivePrinter': "",
    # 'ActiveProtectedViewWindow': None,
    # 'ActiveSheet': None,
    # 'ActiveWindow': None,
    # 'ActiveWorkbook': None,
    'AlertBeforeOverwriting': True,
    'AltStartupPath': "",
    'AlwaysUseClearType': False,
    'AskToUpdateLinks': False,
    'AutoFormatAsYouTypeReplaceHyperlinks': True,
    'AutomationSecurity': 3,  # Optimized for load time
    'AutoPercentEntry': True,
    # 'Calculation': constants.xlCalculationAutomatic,
    # 'CalculateBeforeSave': False,
    'CalculationInterruptKey': 2,
    # 'CalculationState': 0,
    # 'CalculationVersion': 181029,
    # 'Caller': -2146826265,
    # 'CanPlaySounds': True,
    # 'CanRecordSounds': True,
    'Caption': "Excel",
    'CellDragAndDrop': True,
    # 'Cells': None,
    'ChartDataPointTrack': True,
    # 'Charts': None,
    'ClusterConnector': "",
    # 'Columns': None,
    'CommandUnderlines': -4105,
    # 'ConstrainNumeric': False,
    'CopyObjectsWithCells': True,
    # 'Creator': 1480803660,
    'Cursor': -4143,
    'CutCopyMode': 0,
    # 'DataEntryMode': -4146,
    'DecimalSeparator': ".",
    'DefaultSaveFormat': 51,
    'DefaultSheetDirection': -5003,
    'DeferAsyncQueries': False,
    'DisplayAlerts': False,  # Optimized for load time
    'DisplayClipboardWindow': False,
    'DisplayCommentIndicator': -1,
    # 'DisplayDocumentActionTaskPane': False,
    'DisplayDocumentInformationPanel': False,
    'DisplayExcel4Menus': False,
    'DisplayFormulaAutoComplete': True,
    'DisplayFormulaBar': True,
    'DisplayFullScreen': False,
    'DisplayFunctionToolTips': True,
    'DisplayInsertOptions': True,
    'DisplayNoteIndicator': True,
    'DisplayPasteOptions': True,
    'DisplayRecentFiles': True,
    'DisplayScrollBars': True,
    'DisplayStatusBar': True,
    'EditDirectlyInCell': True,
    'EnableAnimations': True,
    'EnableAutoComplete': True,
    'EnableCancelKey': 1,
    'EnableCheckFileExtensions': True,
    'EnableEvents': True,
    'EnableLargeOperationAlert': True,
    'EnableLivePreview': True,
    'EnableMacroAnimations': False,
    'EnableSound': False,
    # 'Excel4IntlMacroSheets': None,
    # 'Excel4MacroSheets': None,
    'ExtendList': True,
    'FeatureInstall': 0,
    # 'FileConverters': None,
    'FileValidation': 0,
    'FileValidationPivot': 0,
    'FixedDecimal': False,
    'FixedDecimalPlaces': 2,
    'FlashFill': True,
    'FlashFillMode': False,
    'FormulaBarHeight': 1,
    'GenerateGetPivotData': False,
    'GenerateTableRefs': 1,
    'HighQualityModeForGraphics': False,
    # 'Hinstance': 16384000,
    # 'HinstancePtr': 16384000,
    # 'Hwnd': 1051558,
    'IgnoreRemoteRequests': False,
    'Interactive': False,  # Optimized for load time
    # 'IsSandboxed': False,
    # 'Iteration': -2146826246,
    'LargeOperationCellThousandCount': 33554,
    # 'MailSession': None,
    'MapPaperSize': True,
    # 'MathCoprocessorAvailable': True,
    # 'MaxChange': -2146826246,
    # 'MaxIterations': -2146826246,
    'MeasurementUnit': 0,
    'MergeInstances': True,
    # 'MouseAvailable': True,
    'MoveAfterReturn': True,
    'MoveAfterReturnDirection': -4121,
    # 'Names': None,
    # 'NetworkTemplatesPath': "",
    'ODBCTimeout': 45,
    'OnWindow': None,
    # 'OrganizationName': "",
    # 'PathSeparator': "\\",
    'PivotTableSelection': False,
    # 'PreviousSelections': None,
    'PrintCommunication': False,  # Optimized for load time
    'PromptForSummaryInfo': False,
    # 'Ready': True,
    # 'RecordRelative': False,
    'ReferenceStyle': 1,
    # 'RegisteredFunctions': None,
    'RollZoom': False,
    # 'Rows': None,
    'ScreenUpdating': False,  # Optimized for load time
    # 'Selection': None,
    # 'Sheets': None,
    'ShowChartTipNames': True,
    'ShowChartTipValues': True,
    'ShowDevTools': True,
    'ShowMenuFloaties': True,
    'ShowQuickAnalysis': True,
    'ShowSelectionFloaties': False,
    'ShowStartupDialog': False,
    'ShowToolTips': True,
    'StandardFont': "Calibri",
    'StatusBar': False,  # Optimized for load time
    # 'ThisCell': None,
    # 'ThisWorkbook': None,
    'ThousandsSeparator': ",",
    # 'TransitionMenuKey': "/",
    # 'TransitionNavigKeys': False,
    'UseClusterConnector': False,
    # 'UseClusterConnector': False,
    'UserControl': False,
    'UseSystemSeparators': True,
    # 'Value': "Microsoft Excel",
    # 'VBE': None,
    'Visible': False,
    'WarnOnFunctionNameConflict': False,
    # 'WindowsForPens': False,
    'WindowState': -4143,
    # 'Worksheets': None,
    }


def compare_xl_com(xl_com, settings):
    """Return if the given xl_com object has all of the same settings."""
    if not isinstance(xl_com, dict):
        xl_com = get_xl_com_settings(xl_com)
    return xl_com == settings


def get_xl_com_settings(xl, **settings):
    """Get the current settings."""
    values = {}
    for k in SETTINGS:
        obj = getattr(xl, k, None)
        if isinstance(getattr(Excel, k, None), ProxyMethod):
            # Get the value from the method
            values[k] = obj()
        else:
            # Get the value from the property
            values[k] = obj

    # Update with the given settings
    values.update(settings)
    return values


def set_xl_com_settings(xl, **settings):
    """Set the given attribute settings."""
    # Set the given attributes
    for k, v in settings.items():
        try:
            # Set the property
            setattr(xl, k, v)
        except (AttributeError, Exception) as err:
            try:
                # Call the setter method
                if str(err).startswith(Excel.METHOD_SETTER_ERROR):
                    func = getattr(xl, k, None)
                    if callable(func):
                        func(v)
            except (AttributeError, TypeError, ValueError, Exception):
                raise err


def get_xl_com_for_settings(**xl_settings):
    """Create or return a global excel object to minimize the number of existing excel objects."""
    # Get the settings with defaults
    settings = SETTINGS.copy()
    settings.update(Excel.DEFAULT_SETTINGS)
    settings.update(xl_settings)

    try:
        # Get the existing excel settings
        xl = Excel.GLOBAL_EXCELS[settings]
    except (KeyError, Exception):
        # Create a new excel object for the settings and return the excel object
        is_parallel = settings.get('is_parallel', None)
        if (is_parallel or is_parallel is None) and is_gencache_available():
            xl = win32com.client.DispatchEx("Excel.Application")
        else:
            xl = win32com.client.Dispatch("Excel.Application")

        # Set the given attributes
        set_xl_com_settings(xl, **settings)

        # Save the object
        Excel.GLOBAL_EXCELS.append(xl)

    # Return the excel object
    return xl


class Excel(object):
    METHOD_SETTER_ERROR = 'Cannot set property '
    
    DEFAULT_SETTINGS = {
        'DisplayAlerts': False,  # If False hides prompt to save
        'AutomationSecurity': 1,
        'Interactive': False,
        'PrintCommunication': False,
        }

    GLOBAL_EXCELS = ItemStorage(compare_func=compare_xl_com)

    def __new__(cls, **kwargs):
        obj = super().__new__(cls)
        obj.is_parallel = kwargs.get('is_parallel', True)
        obj._xl = None

        # Set the settings
        settings = SETTINGS.copy()
        settings.update(cls.DEFAULT_SETTINGS)
        settings.update(kwargs)
        obj._settings = settings
        return obj

    def __init__(self, **kwargs):
        super().__init__()
        if not hasattr(self, '_xl'):
            self._xl = None
        if not hasattr(self, '_settings'):
            settings = SETTINGS.copy()
            settings.update(self.DEFAULT_SETTINGS)
            settings.update(kwargs)
            self._settings = settings  # Use settings to create excel application object.

    def get_settings(self, **settings):
        """Get the current settings."""
        if self._xl is None:
            return self._settings.copy()

        settings = get_xl_com_settings(self, **settings)
        self._settings = settings.copy()  # Might as well save current settings
        return settings

    def set_settings(self, **settings):
        """Set the given attribute settings."""
        # Set the given attributes
        set_xl_com_settings(self, **settings)

    @property
    def xl(self):
        """Get (or create) the Excel Application attribute."""
        if self._xl is None:
            # Catch SIGTERM/SIGINT to close excel safely!
            if should_init_sig() and len(Excel.GLOBAL_EXCELS) == 1:
                init_sig_shutdown()

            # Find or create the application using settings
            self._xl = get_xl_com_for_settings(**self._settings.copy())

        return self._xl

    @xl.setter
    def xl(self, value):
        """Set the Excel Application attribute."""
        if value is None:
            # Save settings in case of recreate
            try:
                self._settings = get_xl_com_settings(self._xl)
            except (AttributeError, ValueError, TypeError, Exception):
                pass
        else:
            self._settings = get_xl_com_settings(value)

        # Save the given xl com object or None.
        self._xl = value

    def __enter__(self):
        return self

    def __exit__(self, exception_type, exception_value, traceback):
        self.Quit()
        return exception_type is None

    # Methods Application object? https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)
    ActivateMicrosoftApp = ProxyMethod('xl.ActivateMicrosoftApp', setter_error=METHOD_SETTER_ERROR)
    AddCustomList = ProxyMethod('xl.AddCustomList', setter_error=METHOD_SETTER_ERROR)
    Calculate = ProxyMethod('xl.Calculate', setter_error=METHOD_SETTER_ERROR)
    CalculateFull = ProxyMethod('xl.CalculateFull', setter_error=METHOD_SETTER_ERROR)
    CalculateFullRebuild = ProxyMethod('xl.CalculateFullRebuild', setter_error=METHOD_SETTER_ERROR)
    CalculateUntilAsyncQueriesDone = ProxyMethod('xl.CalculateUntilAsyncQueriesDone', setter_error=METHOD_SETTER_ERROR)
    CentimetersToPoints = ProxyMethod('xl.CentimetersToPoints', setter_error=METHOD_SETTER_ERROR)
    CheckAbort = ProxyMethod('xl.CheckAbort', setter_error=METHOD_SETTER_ERROR)
    CheckSpelling = ProxyMethod('xl.CheckSpelling', setter_error=METHOD_SETTER_ERROR)
    ConvertFormula = ProxyMethod('xl.ConvertFormula', setter_error=METHOD_SETTER_ERROR)
    DDEExecute = ProxyMethod('xl.DDEExecute', setter_error=METHOD_SETTER_ERROR)
    DDEInitiate = ProxyMethod('xl.DDEInitiate', setter_error=METHOD_SETTER_ERROR)
    DDEPoke = ProxyMethod('xl.DDEPoke', setter_error=METHOD_SETTER_ERROR)
    DDERequest = ProxyMethod('xl.DDERequest', setter_error=METHOD_SETTER_ERROR)
    DDETerminate = ProxyMethod('xl.DDETerminate', setter_error=METHOD_SETTER_ERROR)
    DeleteCustomList = ProxyMethod('xl.DeleteCustomList', setter_error=METHOD_SETTER_ERROR)
    DisplayXMLSourcePane = ProxyMethod('xl.DisplayXMLSourcePane', setter_error=METHOD_SETTER_ERROR)
    DoubleClick = ProxyMethod('xl.DoubleClick', setter_error=METHOD_SETTER_ERROR)
    Evaluate = ProxyMethod('xl.Evaluate', setter_error=METHOD_SETTER_ERROR)
    ExecuteExcel4Macro = ProxyMethod('xl.ExecuteExcel4Macro', setter_error=METHOD_SETTER_ERROR)
    FindFile = ProxyMethod('xl.FindFile', setter_error=METHOD_SETTER_ERROR)
    GetCustomListContents = ProxyMethod('xl.GetCustomListContents', setter_error=METHOD_SETTER_ERROR)
    GetCustomListNum = ProxyMethod('xl.GetCustomListNum', setter_error=METHOD_SETTER_ERROR)
    GetOpenFilename = ProxyMethod('xl.GetOpenFilename', setter_error=METHOD_SETTER_ERROR)
    GetPhonetic = ProxyMethod('xl.GetPhonetic', setter_error=METHOD_SETTER_ERROR)
    GetSaveAsFilename = ProxyMethod('xl.GetSaveAsFilename', setter_error=METHOD_SETTER_ERROR)
    Goto = ProxyMethod('xl.Goto', setter_error=METHOD_SETTER_ERROR)
    Help = ProxyMethod('xl.Help', setter_error=METHOD_SETTER_ERROR)
    InchesToPoints = ProxyMethod('xl.InchesToPoints', setter_error=METHOD_SETTER_ERROR)
    InputBox = ProxyMethod('xl.InputBox', setter_error=METHOD_SETTER_ERROR)
    Intersect = ProxyMethod('xl.Intersect', setter_error=METHOD_SETTER_ERROR)
    MacroOptions = ProxyMethod('xl.MacroOptions', setter_error=METHOD_SETTER_ERROR)
    MailLogoff = ProxyMethod('xl.MailLogoff', setter_error=METHOD_SETTER_ERROR)
    MailLogon = ProxyMethod('xl.MailLogon', setter_error=METHOD_SETTER_ERROR)
    NextLetter = ProxyMethod('xl.NextLetter', setter_error=METHOD_SETTER_ERROR)
    OnKey = ProxyMethod('xl.OnKey', setter_error=METHOD_SETTER_ERROR)
    OnRepeat = ProxyMethod('xl.OnRepeat', setter_error=METHOD_SETTER_ERROR)
    OnTime = ProxyMethod('xl.OnTime', setter_error=METHOD_SETTER_ERROR)
    OnUndo = ProxyMethod('xl.OnUndo', setter_error=METHOD_SETTER_ERROR)
    Quit = ProxyMethod('xl.Quit', setter_error=METHOD_SETTER_ERROR)
    RecordMacro = ProxyMethod('xl.RecordMacro', setter_error=METHOD_SETTER_ERROR)
    RegisterXLL = ProxyMethod('xl.RegisterXLL', setter_error=METHOD_SETTER_ERROR)
    Repeat = ProxyMethod('xl.Repeat', setter_error=METHOD_SETTER_ERROR)
    Run = ProxyMethod('xl.Run', setter_error=METHOD_SETTER_ERROR)
    SendKeys = ProxyMethod('xl.SendKeys', setter_error=METHOD_SETTER_ERROR)
    SharePointVersion = ProxyMethod('xl.SharePointVersion', setter_error=METHOD_SETTER_ERROR)
    Undo = ProxyMethod('xl.Undo', setter_error=METHOD_SETTER_ERROR)
    Union = ProxyMethod('xl.Union', setter_error=METHOD_SETTER_ERROR)
    Volatile = ProxyMethod('xl.Volatile', setter_error=METHOD_SETTER_ERROR)
    Wait = ProxyMethod('xl.Wait', setter_error=METHOD_SETTER_ERROR)

    # Properties
    ActiveCell = ProxyProperty('xl.ActiveCell')
    ActiveChart = ProxyProperty('xl.ActiveChart')
    ActiveEncryptionSession = ProxyProperty('xl.ActiveEncryptionSession')
    ActivePrinter = ProxyProperty('xl.ActivePrinter')
    ActiveProtectedViewWindow = ProxyProperty('xl.ActiveProtectedViewWindow')
    ActiveSheet = ProxyProperty('xl.ActiveSheet')
    ActiveWindow = ProxyProperty('xl.ActiveWindow')
    ActiveWorkbook = ProxyProperty('xl.ActiveWorkbook')
    AddIns = ProxyProperty('xl.AddIns')
    AddIns2 = ProxyProperty('xl.AddIns2')
    AlertBeforeOverwriting = ProxyProperty('xl.AlertBeforeOverwriting')
    AltStartupPath = ProxyProperty('xl.AltStartupPath')
    AlwaysUseClearType = ProxyProperty('xl.AlwaysUseClearType')
    Application = ProxyProperty('xl.Application')
    ArbitraryXMLSupportAvailable = ProxyProperty('xl.ArbitraryXMLSupportAvailable')
    AskToUpdateLinks = ProxyProperty('xl.AskToUpdateLinks')
    Assistance = ProxyProperty('xl.Assistance')
    AutoCorrect = ProxyProperty('xl.AutoCorrect')
    AutoFormatAsYouTypeReplaceHyperlinks = ProxyProperty('xl.AutoFormatAsYouTypeReplaceHyperlinks')
    AutomationSecurity = ProxyProperty('xl.AutomationSecurity')
    AutoPercentEntry = ProxyProperty('xl.AutoPercentEntry')
    AutoRecover = ProxyProperty('xl.AutoRecover')
    Build = ProxyProperty('xl.Build')
    CalculateBeforeSave = ProxyProperty('xl.CalculateBeforeSave')
    Calculation = ProxyProperty('xl.Calculation')
    CalculationInterruptKey = ProxyProperty('xl.CalculationInterruptKey')
    CalculationState = ProxyProperty('xl.CalculationState')
    CalculationVersion = ProxyProperty('xl.CalculationVersion')
    Caller = ProxyProperty('xl.Caller')
    CanPlaySounds = ProxyProperty('xl.CanPlaySounds')
    CanRecordSounds = ProxyProperty('xl.CanRecordSounds')
    Caption = ProxyProperty('xl.Caption')
    CellDragAndDrop = ProxyProperty('xl.CellDragAndDrop')
    Cells = ProxyProperty('xl.Cells')
    ChartDataPointTrack = ProxyProperty('xl.ChartDataPointTrack')
    Charts = ProxyProperty('xl.Charts')
    ClipboardFormats = ProxyProperty('xl.ClipboardFormats')
    ClusterConnector = ProxyProperty('xl.ClusterConnector')
    Columns = ProxyProperty('xl.Columns')
    COMAddIns = ProxyProperty('xl.COMAddIns')
    CommandBars = ProxyProperty('xl.CommandBars')
    CommandUnderlines = ProxyProperty('xl.CommandUnderlines')
    ConstrainNumeric = ProxyProperty('xl.ConstrainNumeric')
    ControlCharacters = ProxyProperty('xl.ControlCharacters')
    CopyObjectsWithCells = ProxyProperty('xl.CopyObjectsWithCells')
    Creator = ProxyProperty('xl.Creator')
    Cursor = ProxyProperty('xl.Cursor')
    CursorMovement = ProxyProperty('xl.CursorMovement')
    CustomListCount = ProxyProperty('xl.CustomListCount')
    CutCopyMode = ProxyProperty('xl.CutCopyMode')
    DataEntryMode = ProxyProperty('xl.DataEntryMode')
    DDEAppReturnCode = ProxyProperty('xl.DDEAppReturnCode')
    DecimalSeparator = ProxyProperty('xl.DecimalSeparator')
    DefaultFilePath = ProxyProperty('xl.DefaultFilePath')
    DefaultSaveFormat = ProxyProperty('xl.DefaultSaveFormat')
    DefaultSheetDirection = ProxyProperty('xl.DefaultSheetDirection')
    DefaultWebOptions = ProxyProperty('xl.DefaultWebOptions')
    DeferAsyncQueries = ProxyProperty('xl.DeferAsyncQueries')
    Dialogs = ProxyProperty('xl.Dialogs')
    DisplayAlerts = ProxyProperty('xl.DisplayAlerts')
    DisplayClipboardWindow = ProxyProperty('xl.DisplayClipboardWindow')
    DisplayCommentIndicator = ProxyProperty('xl.DisplayCommentIndicator')
    DisplayDocumentActionTaskPane = ProxyProperty('xl.DisplayDocumentActionTaskPane')
    DisplayDocumentInformationPanel = ProxyProperty('xl.DisplayDocumentInformationPanel')
    DisplayExcel4Menus = ProxyProperty('xl.DisplayExcel4Menus')
    DisplayFormulaAutoComplete = ProxyProperty('xl.DisplayFormulaAutoComplete')
    DisplayFormulaBar = ProxyProperty('xl.DisplayFormulaBar')
    DisplayFullScreen = ProxyProperty('xl.DisplayFullScreen')
    DisplayFunctionToolTips = ProxyProperty('xl.DisplayFunctionToolTips')
    DisplayInsertOptions = ProxyProperty('xl.DisplayInsertOptions')
    DisplayNoteIndicator = ProxyProperty('xl.DisplayNoteIndicator')
    DisplayPasteOptions = ProxyProperty('xl.DisplayPasteOptions')
    DisplayRecentFiles = ProxyProperty('xl.DisplayRecentFiles')
    DisplayScrollBars = ProxyProperty('xl.DisplayScrollBars')
    DisplayStatusBar = ProxyProperty('xl.DisplayStatusBar')
    EditDirectlyInCell = ProxyProperty('xl.EditDirectlyInCell')
    EnableAnimations = ProxyProperty('xl.EnableAnimations')
    EnableAutoComplete = ProxyProperty('xl.EnableAutoComplete')
    EnableCancelKey = ProxyProperty('xl.EnableCancelKey')
    EnableCheckFileExtensions = ProxyProperty('xl.EnableCheckFileExtensions')
    EnableEvents = ProxyProperty('xl.EnableEvents')
    EnableLargeOperationAlert = ProxyProperty('xl.EnableLargeOperationAlert')
    EnableLivePreview = ProxyProperty('xl.EnableLivePreview')
    EnableMacroAnimations = ProxyProperty('xl.EnableMacroAnimations')
    EnableSound = ProxyProperty('xl.EnableSound')
    ErrorCheckingOptions = ProxyProperty('xl.ErrorCheckingOptions')
    Excel4IntlMacroSheets = ProxyProperty('xl.Excel4IntlMacroSheets')
    Excel4MacroSheets = ProxyProperty('xl.Excel4MacroSheets')
    ExtendList = ProxyProperty('xl.ExtendList')
    FeatureInstall = ProxyProperty('xl.FeatureInstall')
    FileConverters = ProxyProperty('xl.FileConverters')
    FileDialog = ProxyProperty('xl.FileDialog')
    FileExportConverters = ProxyProperty('xl.FileExportConverters')
    FileValidation = ProxyProperty('xl.FileValidation')
    FileValidationPivot = ProxyProperty('xl.FileValidationPivot')
    FindFormat = ProxyProperty('xl.FindFormat')
    FixedDecimal = ProxyProperty('xl.FixedDecimal')
    FixedDecimalPlaces = ProxyProperty('xl.FixedDecimalPlaces')
    FlashFill = ProxyProperty('xl.FlashFill')
    FlashFillMode = ProxyProperty('xl.FlashFillMode')
    FormulaBarHeight = ProxyProperty('xl.FormulaBarHeight')
    GenerateGetPivotData = ProxyProperty('xl.GenerateGetPivotData')
    GenerateTableRefs = ProxyProperty('xl.GenerateTableRefs')
    Height = ProxyProperty('xl.Height')
    HighQualityModeForGraphics = ProxyProperty('xl.HighQualityModeForGraphics')
    Hinstance = ProxyProperty('xl.Hinstance')
    HinstancePtr = ProxyProperty('xl.HinstancePtr')
    Hwnd = ProxyProperty('xl.Hwnd')
    IgnoreRemoteRequests = ProxyProperty('xl.IgnoreRemoteRequests')
    Interactive = ProxyProperty('xl.Interactive')
    International = ProxyProperty('xl.International')
    IsSandboxed = ProxyProperty('xl.IsSandboxed')
    Iteration = ProxyProperty('xl.Iteration')
    LanguageSettings = ProxyProperty('xl.LanguageSettings')
    LargeOperationCellThousandCount = ProxyProperty('xl.LargeOperationCellThousandCount')
    Left = ProxyProperty('xl.Left')
    LibraryPath = ProxyProperty('xl.LibraryPath')
    MailSession = ProxyProperty('xl.MailSession')
    MailSystem = ProxyProperty('xl.MailSystem')
    MapPaperSize = ProxyProperty('xl.MapPaperSize')
    MathCoprocessorAvailable = ProxyProperty('xl.MathCoprocessorAvailable')
    MaxChange = ProxyProperty('xl.MaxChange')
    MaxIterations = ProxyProperty('xl.MaxIterations')
    MeasurementUnit = ProxyProperty('xl.MeasurementUnit')
    MergeInstances = ProxyProperty('xl.MergeInstances')
    MouseAvailable = ProxyProperty('xl.MouseAvailable')
    MoveAfterReturn = ProxyProperty('xl.MoveAfterReturn')
    MoveAfterReturnDirection = ProxyProperty('xl.MoveAfterReturnDirection')
    MultiThreadedCalculation = ProxyProperty('xl.MultiThreadedCalculation')
    Name = ProxyProperty('xl.Name')
    Names = ProxyProperty('xl.Names')
    NetworkTemplatesPath = ProxyProperty('xl.NetworkTemplatesPath')
    NewWorkbook = ProxyProperty('xl.NewWorkbook')
    ODBCErrors = ProxyProperty('xl.ODBCErrors')
    ODBCTimeout = ProxyProperty('xl.ODBCTimeout')
    OLEDBErrors = ProxyProperty('xl.OLEDBErrors')
    OnWindow = ProxyProperty('xl.OnWindow')
    OperatingSystem = ProxyProperty('xl.OperatingSystem')
    OrganizationName = ProxyProperty('xl.OrganizationName')
    Parent = ProxyProperty('xl.Parent')
    Path = ProxyProperty('xl.Path')
    PathSeparator = ProxyProperty('xl.PathSeparator')
    PivotTableSelection = ProxyProperty('xl.PivotTableSelection')
    PreviousSelections = ProxyProperty('xl.PreviousSelections')
    PrintCommunication = ProxyProperty('xl.PrintCommunication')
    ProductCode = ProxyProperty('xl.ProductCode')
    PromptForSummaryInfo = ProxyProperty('xl.PromptForSummaryInfo')
    ProtectedViewWindows = ProxyProperty('xl.ProtectedViewWindows')
    QuickAnalysis = ProxyProperty('xl.QuickAnalysis')
    Range = ProxyProperty('xl.Range')
    Ready = ProxyProperty('xl.Ready')
    RecentFiles = ProxyProperty('xl.RecentFiles')
    RecordRelative = ProxyProperty('xl.RecordRelative')
    ReferenceStyle = ProxyProperty('xl.ReferenceStyle')
    RegisteredFunctions = ProxyProperty('xl.RegisteredFunctions')
    ReplaceFormat = ProxyProperty('xl.ReplaceFormat')
    RollZoom = ProxyProperty('xl.RollZoom')
    Rows = ProxyProperty('xl.Rows')
    RTD = ProxyProperty('xl.RTD')
    ScreenUpdating = ProxyProperty('xl.ScreenUpdating')
    Selection = ProxyProperty('xl.Selection')
    Sheets = ProxyProperty('xl.Sheets')
    SheetsInNewWorkbook = ProxyProperty('xl.SheetsInNewWorkbook')
    ShowChartTipNames = ProxyProperty('xl.ShowChartTipNames')
    ShowChartTipValues = ProxyProperty('xl.ShowChartTipValues')
    ShowDevTools = ProxyProperty('xl.ShowDevTools')
    ShowMenuFloaties = ProxyProperty('xl.ShowMenuFloaties')
    ShowQuickAnalysis = ProxyProperty('xl.ShowQuickAnalysis')
    ShowSelectionFloaties = ProxyProperty('xl.ShowSelectionFloaties')
    ShowStartupDialog = ProxyProperty('xl.ShowStartupDialog')
    ShowToolTips = ProxyProperty('xl.ShowToolTips')
    SmartArtColors = ProxyProperty('xl.SmartArtColors')
    SmartArtLayouts = ProxyProperty('xl.SmartArtLayouts')
    SmartArtQuickStyles = ProxyProperty('xl.SmartArtQuickStyles')
    Speech = ProxyProperty('xl.Speech')
    SpellingOptions = ProxyProperty('xl.SpellingOptions')
    StandardFont = ProxyProperty('xl.StandardFont')
    StandardFontSize = ProxyProperty('xl.StandardFontSize')
    StartupPath = ProxyProperty('xl.StartupPath')
    StatusBar = ProxyProperty('xl.StatusBar')
    TemplatesPath = ProxyProperty('xl.TemplatesPath')
    ThisCell = ProxyProperty('xl.ThisCell')
    ThisWorkbook = ProxyProperty('xl.ThisWorkbook')
    ThousandsSeparator = ProxyProperty('xl.ThousandsSeparator')
    Top = ProxyProperty('xl.Top')
    TransitionMenuKey = ProxyProperty('xl.TransitionMenuKey')
    TransitionMenuKeyAction = ProxyProperty('xl.TransitionMenuKeyAction')
    TransitionNavigKeys = ProxyProperty('xl.TransitionNavigKeys')
    UsableHeight = ProxyProperty('xl.UsableHeight')
    UsableWidth = ProxyProperty('xl.UsableWidth')
    UseClusterConnector = ProxyProperty('xl.UseClusterConnector')
    UsedObjects = ProxyProperty('xl.UsedObjects')
    UserControl = ProxyProperty('xl.UserControl')
    UserLibraryPath = ProxyProperty('xl.UserLibraryPath')
    UserName = ProxyProperty('xl.UserName')
    UseSystemSeparators = ProxyProperty('xl.UseSystemSeparators')
    Value = ProxyProperty('xl.Value')
    VBE = ProxyProperty('xl.VBE')
    Version = ProxyProperty('xl.Version')
    Visible = ProxyProperty('xl.Visible')
    WarnOnFunctionNameConflict = ProxyProperty('xl.WarnOnFunctionNameConflict')
    Watches = ProxyProperty('xl.Watches')
    Width = ProxyProperty('xl.Width')
    Windows = ProxyProperty('xl.Windows')
    WindowsForPens = ProxyProperty('xl.WindowsForPens')
    WindowState = ProxyProperty('xl.WindowState')
    Workbooks = ProxyProperty('xl.Workbooks')
    WorksheetFunction = ProxyProperty('xl.WorksheetFunction')
    Worksheets = ProxyProperty('xl.Worksheets')


class Workbook(object):
    METHOD_SETTER_ERROR = 'Cannot set property '
    SAVE_ON_CLOSE = False

    def __init__(self, filename=None, *args, xl=None, wb=None, **xl_settings):
        # Variables
        self._xl = xl
        self._wb = wb
        self._filename = None  # Save the filename as a variable

        # Initialize Excel
        if self._xl is None:
            self._xl = Excel(**xl_settings)

        # Set the filename
        self.set_filename(filename)

        # Check to open the filename
        if isinstance(self.filename, str) and os.path.exists(self.filename) and os.path.isfile(self.filename):
            self.open(filename)

        # Initialize constants
        self.init_constants()

    def init_constants(self):
        """Set all of the constant values."""
        for k, field in self.__class__.__dict__.items():
            if isinstance(field, ConstantItem):
                field.init_table(self)

    @property
    def xl(self):
        """Get (or create) the Excel Application object."""
        if self._xl is None:
            self._xl = Excel()
        return self._xl

    @xl.setter
    def xl(self, value):
        """Set the Excel Application object."""
        self._xl = value

    @property
    def wb(self):
        """Get (or Add) a Workbook to the Excel Application Workbooks collection."""
        if self._wb is None:
            self._wb = self.xl.Workbooks.Add()
        return self._wb

    @wb.setter
    def wb(self, value):
        """Set the Workbook object."""
        try:
            self._wb.Close(self.SAVE_ON_CLOSE)
        except (AttributeError, Exception):
            pass
        self._wb = value

    def get_filename(self):
        """Return the filename."""
        return self._filename

    def set_filename(self, filename):
        """Set the filename.

        Args:
            filename (str/object): Filename to save and open from.
        """
        if isinstance(filename, str):
            filename = os.path.abspath(filename)
        self._filename = filename

    filename = property(get_filename, set_filename)

    def open(self, filename=None):
        """Open a workbook with the given filename and use this workbook."""
        if filename is not None:
            self.set_filename(filename)

        filename = self.get_filename()
        if isinstance(filename, str) and os.path.exists(filename) and os.path.isfile(filename):
            self._wb = self.xl.Workbooks.Open(filename)
        return self

    EXT_TO_FMT = {
        # Lower case extension to Format https://docs.microsoft.com/en-us/office/vba/api/excel.xlfileformat
        '.xla': getattr(win32com.client.constants, 'xlAddIn', 18),
        '.csv': getattr(win32com.client.constants, 'xlCSV', 6),
        '.txt': getattr(win32com.client.constants, 'xlCurrentPlatformText', -4158),
        '.dbf': getattr(win32com.client.constants, 'xlDBF4', 11),
        '.dif': getattr(win32com.client.constants, 'xlDIF', 9),
        '.xlsb': getattr(win32com.client.constants, 'xlExcel12', 50),
        # '.xls': getattr(win32com.client.constants, 'xlExcel8', 56),
        '.htm': getattr(win32com.client.constants, 'xlHtml', 44),
        '.html': getattr(win32com.client.constants, 'xlHtml', 44),
        '.ods': getattr(win32com.client.constants, 'xlOpenDocumentSpreadsheet', 60),
        '.xlam': getattr(win32com.client.constants, 'xlOpenXMLAddIn', 55),
        '.xltx': getattr(win32com.client.constants, 'xlOpenXMLTemplate', 54),
        # getattr(win32com.client.constants, 'xlTemplate', 17),
        '.xlsm': getattr(win32com.client.constants, 'xlOpenXMLWorkbookMacroEnabled', 52),
        '.slk': getattr(win32com.client.constants, 'xlSYLK', 2),
        '.xlt': getattr(win32com.client.constants, 'xlTemplate', 17),
        '.prn': getattr(win32com.client.constants, 'xlTextPrinter', 36),
        '.mht': getattr(win32com.client.constants, 'xlWebArchive', 45),
        '.mhtml': getattr(win32com.client.constants, 'xlWebArchive', 45),
        '.wj2': getattr(win32com.client.constants, 'xlWJ2WD1', 14),
        '.wj3': getattr(win32com.client.constants, 'xlWJ3FJ3', 41),
        '.wk1': getattr(win32com.client.constants, 'xlWK3FM3', 32),
        '.wk3': getattr(win32com.client.constants, 'xlWK3', 15),
        '.wk4': getattr(win32com.client.constants, 'xlWK4', 38),
        '.wks': getattr(win32com.client.constants, 'xlWKS', 4),
        '.xlsx': getattr(win32com.client.constants, 'xlWorkbookDefault', 51),
        '.xls': getattr(win32com.client.constants, 'xlWorkbookNormal', -4143),
        '.wq1': getattr(win32com.client.constants, 'xlWQ1', 34),
        '.xml': getattr(win32com.client.constants, 'xlXMLSpreadsheet', 46),
        }

    def save(self, filename=None):
        """Save the given filename or set filename."""
        if filename is not None:
            self.set_filename(filename)

        filename = self.get_filename()

        ext = os.path.splitext(filename)[-1].lower()
        file_fmt = self.EXT_TO_FMT.get(ext, None)

        # Saving as CSV or non excel type renames the active sheet to the base filename.
        sheet_name = self.wb.ActiveSheet.Name
        self.wb.SaveAs(filename, file_fmt)
        self.wb.ActiveSheet.Name = sheet_name

    def get_sheet(self, sheet, create=True):
        """Return the sheet for an index or name."""
        try:
            if isinstance(sheet, int):
                obj = self.wb.Sheets[sheet]  # Get the sheet
            else:
                obj = self.wb.Sheets(sheet)  # Get the sheet
            obj.Activate()
        except (ValueError, TypeError, Exception):
            if create:
                obj = self.wb.Sheets.Add()  # Create the sheet
                obj.Name = sheet
            else:
                obj = None
        return obj

    def has_sheet(self, sheet):
        """Return if the given sheet name or index exists"""
        return self.get_sheet(sheet, create=False) is not None

    # ===== Workbook Object Methods ===== https://docs.microsoft.com/en-us/office/vba/api/excel.workbook#methods
    AcceptAllChanges = ProxyMethod('wb.AcceptAllChanges', setter_error=METHOD_SETTER_ERROR)
    Activate = ProxyMethod('wb.Activate', setter_error=METHOD_SETTER_ERROR)
    AddToFavorites = ProxyMethod('wb.AddToFavorites', setter_error=METHOD_SETTER_ERROR)
    ApplyTheme = ProxyMethod('wb.ApplyTheme', setter_error=METHOD_SETTER_ERROR)
    BreakLink = ProxyMethod('wb.BreakLink', setter_error=METHOD_SETTER_ERROR)
    CanCheckIn = ProxyMethod('wb.CanCheckIn', setter_error=METHOD_SETTER_ERROR)
    ChangeFileAccess = ProxyMethod('wb.ChangeFileAccess', setter_error=METHOD_SETTER_ERROR)
    ChangeLink = ProxyMethod('wb.ChangeLink', setter_error=METHOD_SETTER_ERROR)
    CheckIn = ProxyMethod('wb.CheckIn', setter_error=METHOD_SETTER_ERROR)
    CheckInWithVersion = ProxyMethod('wb.CheckInWithVersion', setter_error=METHOD_SETTER_ERROR)
    Close = ProxyMethod('wb.Close', setter_error=METHOD_SETTER_ERROR)
    ConvertComments = ProxyMethod('wb.ConvertComments', setter_error=METHOD_SETTER_ERROR)
    CreateForecastSheet = ProxyMethod('wb.CreateForecastSheet', setter_error=METHOD_SETTER_ERROR)
    DeleteNumberFormat = ProxyMethod('wb.DeleteNumberFormat', setter_error=METHOD_SETTER_ERROR)
    EnableConnections = ProxyMethod('wb.EnableConnections', setter_error=METHOD_SETTER_ERROR)
    EndReview = ProxyMethod('wb.EndReview', setter_error=METHOD_SETTER_ERROR)
    ExclusiveAccess = ProxyMethod('wb.ExclusiveAccess', setter_error=METHOD_SETTER_ERROR)
    ExportAsFixedFormat = ProxyMethod('wb.ExportAsFixedFormat', setter_error=METHOD_SETTER_ERROR)
    FollowHyperlink = ProxyMethod('wb.FollowHyperlink', setter_error=METHOD_SETTER_ERROR)
    ForwardMailer = ProxyMethod('wb.ForwardMailer', setter_error=METHOD_SETTER_ERROR)
    GetWorkflowTasks = ProxyMethod('wb.GetWorkflowTasks', setter_error=METHOD_SETTER_ERROR)
    GetWorkflowTemplates = ProxyMethod('wb.GetWorkflowTemplates', setter_error=METHOD_SETTER_ERROR)
    HighlightChangesOptions = ProxyMethod('wb.HighlightChangesOptions', setter_error=METHOD_SETTER_ERROR)
    LinkInfo = ProxyMethod('wb.LinkInfo', setter_error=METHOD_SETTER_ERROR)
    LinkSources = ProxyMethod('wb.LinkSources', setter_error=METHOD_SETTER_ERROR)
    LockServerFile = ProxyMethod('wb.LockServerFile', setter_error=METHOD_SETTER_ERROR)
    MergeWorkbook = ProxyMethod('wb.MergeWorkbook', setter_error=METHOD_SETTER_ERROR)
    NewWindow = ProxyMethod('wb.NewWindow', setter_error=METHOD_SETTER_ERROR)
    OpenLinks = ProxyMethod('wb.OpenLinks', setter_error=METHOD_SETTER_ERROR)
    PivotCaches = ProxyMethod('wb.PivotCaches', setter_error=METHOD_SETTER_ERROR)
    Post = ProxyMethod('wb.Post', setter_error=METHOD_SETTER_ERROR)
    PrintOut = ProxyMethod('wb.PrintOut', setter_error=METHOD_SETTER_ERROR)
    PrintPreview = ProxyMethod('wb.PrintPreview', setter_error=METHOD_SETTER_ERROR)
    Protect = ProxyMethod('wb.Protect', setter_error=METHOD_SETTER_ERROR)
    ProtectSharing = ProxyMethod('wb.ProtectSharing', setter_error=METHOD_SETTER_ERROR)
    PublishToDocs = ProxyMethod('wb.PublishToDocs', setter_error=METHOD_SETTER_ERROR)
    PurgeChangeHistoryNow = ProxyMethod('wb.PurgeChangeHistoryNow', setter_error=METHOD_SETTER_ERROR)
    RefreshAll = ProxyMethod('wb.RefreshAll', setter_error=METHOD_SETTER_ERROR)
    RejectAllChanges = ProxyMethod('wb.RejectAllChanges', setter_error=METHOD_SETTER_ERROR)
    ReloadAs = ProxyMethod('wb.ReloadAs', setter_error=METHOD_SETTER_ERROR)
    RemoveDocumentInformation = ProxyMethod('wb.RemoveDocumentInformation', setter_error=METHOD_SETTER_ERROR)
    RemoveUser = ProxyMethod('wb.RemoveUser', setter_error=METHOD_SETTER_ERROR)
    Reply = ProxyMethod('wb.Reply', setter_error=METHOD_SETTER_ERROR)
    ReplyAll = ProxyMethod('wb.ReplyAll', setter_error=METHOD_SETTER_ERROR)
    ReplyWithChanges = ProxyMethod('wb.ReplyWithChanges', setter_error=METHOD_SETTER_ERROR)
    ResetColors = ProxyMethod('wb.ResetColors', setter_error=METHOD_SETTER_ERROR)
    RunAutoMacros = ProxyMethod('wb.RunAutoMacros', setter_error=METHOD_SETTER_ERROR)
    Save = ProxyMethod('wb.Save', setter_error=METHOD_SETTER_ERROR)
    SaveAs = ProxyMethod('wb.SaveAs', setter_error=METHOD_SETTER_ERROR)
    SaveAsXMLData = ProxyMethod('wb.SaveAsXMLData', setter_error=METHOD_SETTER_ERROR)
    SaveCopyAs = ProxyMethod('wb.SaveCopyAs', setter_error=METHOD_SETTER_ERROR)
    SendFaxOverInternet = ProxyMethod('wb.SendFaxOverInternet', setter_error=METHOD_SETTER_ERROR)
    SendForReview = ProxyMethod('wb.SendForReview', setter_error=METHOD_SETTER_ERROR)
    SendMail = ProxyMethod('wb.SendMail', setter_error=METHOD_SETTER_ERROR)
    SendMailer = ProxyMethod('wb.SendMailer', setter_error=METHOD_SETTER_ERROR)
    SetLinkOnData = ProxyMethod('wb.SetLinkOnData', setter_error=METHOD_SETTER_ERROR)
    SetPasswordEncryptionOptions = ProxyMethod('wb.SetPasswordEncryptionOptions', setter_error=METHOD_SETTER_ERROR)
    ToggleFormsDesign = ProxyMethod('wb.ToggleFormsDesign', setter_error=METHOD_SETTER_ERROR)
    Unprotect = ProxyMethod('wb.Unprotect', setter_error=METHOD_SETTER_ERROR)
    UnprotectSharing = ProxyMethod('wb.UnprotectSharing', setter_error=METHOD_SETTER_ERROR)
    UpdateFromFile = ProxyMethod('wb.UpdateFromFile', setter_error=METHOD_SETTER_ERROR)
    UpdateLink = ProxyMethod('wb.UpdateLink', setter_error=METHOD_SETTER_ERROR)
    WebPagePreview = ProxyMethod('wb.WebPagePreview', setter_error=METHOD_SETTER_ERROR)
    XmlImport = ProxyMethod('wb.XmlImport', setter_error=METHOD_SETTER_ERROR)
    XmlImportXml = ProxyMethod('wb.XmlImportXml', setter_error=METHOD_SETTER_ERROR)

    # ===== Workbook Object Properties ===== https://docs.microsoft.com/en-us/office/vba/api/excel.workbook#properties
    AccuracyVersion = ProxyProperty('wb.AccuracyVersion')
    ActiveChart = ProxyProperty('wb.ActiveChart')
    ActiveSheet = ProxyProperty('wb.ActiveSheet')
    ActiveSlicer = ProxyProperty('wb.ActiveSlicer')
    Application = ProxyProperty('wb.Application')
    AutoSaveOn = ProxyProperty('wb.AutoSaveOn')
    AutoUpdateFrequency = ProxyProperty('wb.AutoUpdateFrequency')
    AutoUpdateSaveChanges = ProxyProperty('wb.AutoUpdateSaveChanges')
    BuiltinDocumentProperties = ProxyProperty('wb.BuiltinDocumentProperties')
    CalculationVersion = ProxyProperty('wb.CalculationVersion')
    CaseSensitive = ProxyProperty('wb.CaseSensitive')
    ChangeHistoryDuration = ProxyProperty('wb.ChangeHistoryDuration')
    ChartDataPointTrack = ProxyProperty('wb.ChartDataPointTrack')
    Charts = ProxyProperty('wb.Charts')
    CheckCompatibility = ProxyProperty('wb.CheckCompatibility')
    CodeName = ProxyProperty('wb.CodeName')
    Colors = ProxyProperty('wb.Colors')
    CommandBars = ProxyProperty('wb.CommandBars')
    ConflictResolution = ProxyProperty('wb.ConflictResolution')
    Connections = ProxyProperty('wb.Connections')
    ConnectionsDisabled = ProxyProperty('wb.ConnectionsDisabled')
    Container = ProxyProperty('wb.Container')
    ContentTypeProperties = ProxyProperty('wb.ContentTypeProperties')
    CreateBackup = ProxyProperty('wb.CreateBackup')
    Creator = ProxyProperty('wb.Creator')
    CustomDocumentProperties = ProxyProperty('wb.CustomDocumentProperties')
    CustomViews = ProxyProperty('wb.CustomViews')
    CustomXMLParts = ProxyProperty('wb.CustomXMLParts')
    Date1904 = ProxyProperty('wb.Date1904')
    DefaultPivotTableStyle = ProxyProperty('wb.DefaultPivotTableStyle')
    DefaultSlicerStyle = ProxyProperty('wb.DefaultSlicerStyle')
    DefaultTableStyle = ProxyProperty('wb.DefaultTableStyle')
    DefaultTimelineStyle = ProxyProperty('wb.DefaultTimelineStyle')
    DisplayDrawingObjects = ProxyProperty('wb.DisplayDrawingObjects')
    DisplayInkComments = ProxyProperty('wb.DisplayInkComments')
    DocumentInspectors = ProxyProperty('wb.DocumentInspectors')
    DocumentLibraryVersions = ProxyProperty('wb.DocumentLibraryVersions')
    DoNotPromptForConvert = ProxyProperty('wb.DoNotPromptForConvert')
    EnableAutoRecover = ProxyProperty('wb.EnableAutoRecover')
    EncryptionProvider = ProxyProperty('wb.EncryptionProvider')
    EnvelopeVisible = ProxyProperty('wb.EnvelopeVisible')
    Excel4IntlMacroSheets = ProxyProperty('wb.Excel4IntlMacroSheets')
    Excel4MacroSheets = ProxyProperty('wb.Excel4MacroSheets')
    Excel8CompatibilityMode = ProxyProperty('wb.Excel8CompatibilityMode')
    FileFormat = ProxyProperty('wb.FileFormat')
    Final = ProxyProperty('wb.Final')
    ForceFullCalculation = ProxyProperty('wb.ForceFullCalculation')
    FullName = ProxyProperty('wb.FullName')
    FullNameURLEncoded = ProxyProperty('wb.FullNameURLEncoded')
    HasPassword = ProxyProperty('wb.HasPassword')
    HasVBProject = ProxyProperty('wb.HasVBProject')
    HighlightChangesOnScreen = ProxyProperty('wb.HighlightChangesOnScreen')
    IconSets = ProxyProperty('wb.IconSets')
    InactiveListBorderVisible = ProxyProperty('wb.InactiveListBorderVisible')
    IsAddin = ProxyProperty('wb.IsAddin')
    IsInplace = ProxyProperty('wb.IsInplace')
    KeepChangeHistory = ProxyProperty('wb.KeepChangeHistory')
    ListChangesOnNewSheet = ProxyProperty('wb.ListChangesOnNewSheet')
    Mailer = ProxyProperty('wb.Mailer')
    Model = ProxyProperty('wb.Model')
    MultiUserEditing = ProxyProperty('wb.MultiUserEditing')
    Name = ProxyProperty('wb.Name')
    Names = ProxyProperty('wb.Names')
    Parent = ProxyProperty('wb.Parent')
    Password = ProxyProperty('wb.Password')
    PasswordEncryptionAlgorithm = ProxyProperty('wb.PasswordEncryptionAlgorithm')
    PasswordEncryptionFileProperties = ProxyProperty('wb.PasswordEncryptionFileProperties')
    PasswordEncryptionKeyLength = ProxyProperty('wb.PasswordEncryptionKeyLength')
    PasswordEncryptionProvider = ProxyProperty('wb.PasswordEncryptionProvider')
    Path = ProxyProperty('wb.Path')
    Permission = ProxyProperty('wb.Permission')
    PersonalViewListSettings = ProxyProperty('wb.PersonalViewListSettings')
    PersonalViewPrintSettings = ProxyProperty('wb.PersonalViewPrintSettings')
    PivotTables = ProxyProperty('wb.PivotTables')
    PrecisionAsDisplayed = ProxyProperty('wb.PrecisionAsDisplayed')
    ProtectStructure = ProxyProperty('wb.ProtectStructure')
    ProtectWindows = ProxyProperty('wb.ProtectWindows')
    PublishObjects = ProxyProperty('wb.PublishObjects')
    Queries = ProxyProperty('wb.Queries')
    ReadOnly = ProxyProperty('wb.ReadOnly')
    ReadOnlyRecommended = ProxyProperty('wb.ReadOnlyRecommended')
    RemovePersonalInformation = ProxyProperty('wb.RemovePersonalInformation')
    Research = ProxyProperty('wb.Research')
    RevisionNumber = ProxyProperty('wb.RevisionNumber')
    Saved = ProxyProperty('wb.Saved')
    SaveLinkValues = ProxyProperty('wb.SaveLinkValues')
    ServerPolicy = ProxyProperty('wb.ServerPolicy')
    ServerViewableItems = ProxyProperty('wb.ServerViewableItems')
    SharedWorkspace = ProxyProperty('wb.SharedWorkspace')
    # Sheets = ProxyProperty('wb.Sheets')
    ShowConflictHistory = ProxyProperty('wb.ShowConflictHistory')
    ShowPivotChartActiveFields = ProxyProperty('wb.ShowPivotChartActiveFields')
    ShowPivotTableFieldList = ProxyProperty('wb.ShowPivotTableFieldList')
    Signatures = ProxyProperty('wb.Signatures')
    SlicerCaches = ProxyProperty('wb.SlicerCaches')
    SmartDocument = ProxyProperty('wb.SmartDocument')
    Styles = ProxyProperty('wb.Styles')
    Sync = ProxyProperty('wb.Sync')
    TableStyles = ProxyProperty('wb.TableStyles')
    TemplateRemoveExtData = ProxyProperty('wb.TemplateRemoveExtData')
    Theme = ProxyProperty('wb.Theme')
    UpdateLinks = ProxyProperty('wb.UpdateLinks')
    UpdateRemoteReferences = ProxyProperty('wb.UpdateRemoteReferences')
    UserStatus = ProxyProperty('wb.UserStatus')
    UseWholeCellCriteria = ProxyProperty('wb.UseWholeCellCriteria')
    UseWildcards = ProxyProperty('wb.UseWildcards')
    VBASigned = ProxyProperty('wb.VBASigned')
    VBProject = ProxyProperty('wb.VBProject')
    WebOptions = ProxyProperty('wb.WebOptions')
    Windows = ProxyProperty('wb.Windows')
    Worksheets = ProxyProperty('wb.Worksheets')
    WritePassword = ProxyProperty('wb.WritePassword')
    WriteReserved = ProxyProperty('wb.WriteReserved')
    WriteReservedBy = ProxyProperty('wb.WriteReservedBy')
    XmlMaps = ProxyProperty('wb.XmlMaps')
    XmlNamespaces = ProxyProperty('wb.XmlNamespaces')

    # ===== Application Methods ===== https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)
    ActivateMicrosoftApp = ProxyMethod('xl.ActivateMicrosoftApp', setter_error=METHOD_SETTER_ERROR)
    AddCustomList = ProxyMethod('xl.AddCustomList', setter_error=METHOD_SETTER_ERROR)
    Calculate = ProxyMethod('xl.Calculate', setter_error=METHOD_SETTER_ERROR)
    CalculateFull = ProxyMethod('xl.CalculateFull', setter_error=METHOD_SETTER_ERROR)
    CalculateFullRebuild = ProxyMethod('xl.CalculateFullRebuild', setter_error=METHOD_SETTER_ERROR)
    CalculateUntilAsyncQueriesDone = ProxyMethod('xl.CalculateUntilAsyncQueriesDone', setter_error=METHOD_SETTER_ERROR)
    CentimetersToPoints = ProxyMethod('xl.CentimetersToPoints', setter_error=METHOD_SETTER_ERROR)
    CheckAbort = ProxyMethod('xl.CheckAbort', setter_error=METHOD_SETTER_ERROR)
    CheckSpelling = ProxyMethod('xl.CheckSpelling', setter_error=METHOD_SETTER_ERROR)
    ConvertFormula = ProxyMethod('xl.ConvertFormula', setter_error=METHOD_SETTER_ERROR)
    DDEExecute = ProxyMethod('xl.DDEExecute', setter_error=METHOD_SETTER_ERROR)
    DDEInitiate = ProxyMethod('xl.DDEInitiate', setter_error=METHOD_SETTER_ERROR)
    DDEPoke = ProxyMethod('xl.DDEPoke', setter_error=METHOD_SETTER_ERROR)
    DDERequest = ProxyMethod('xl.DDERequest', setter_error=METHOD_SETTER_ERROR)
    DDETerminate = ProxyMethod('xl.DDETerminate', setter_error=METHOD_SETTER_ERROR)
    DeleteCustomList = ProxyMethod('xl.DeleteCustomList', setter_error=METHOD_SETTER_ERROR)
    DisplayXMLSourcePane = ProxyMethod('xl.DisplayXMLSourcePane', setter_error=METHOD_SETTER_ERROR)
    DoubleClick = ProxyMethod('xl.DoubleClick', setter_error=METHOD_SETTER_ERROR)
    Evaluate = ProxyMethod('xl.Evaluate', setter_error=METHOD_SETTER_ERROR)
    ExecuteExcel4Macro = ProxyMethod('xl.ExecuteExcel4Macro', setter_error=METHOD_SETTER_ERROR)
    FindFile = ProxyMethod('xl.FindFile', setter_error=METHOD_SETTER_ERROR)
    GetCustomListContents = ProxyMethod('xl.GetCustomListContents', setter_error=METHOD_SETTER_ERROR)
    GetCustomListNum = ProxyMethod('xl.GetCustomListNum', setter_error=METHOD_SETTER_ERROR)
    GetOpenFilename = ProxyMethod('xl.GetOpenFilename', setter_error=METHOD_SETTER_ERROR)
    GetPhonetic = ProxyMethod('xl.GetPhonetic', setter_error=METHOD_SETTER_ERROR)
    GetSaveAsFilename = ProxyMethod('xl.GetSaveAsFilename', setter_error=METHOD_SETTER_ERROR)
    Goto = ProxyMethod('xl.Goto', setter_error=METHOD_SETTER_ERROR)
    Help = ProxyMethod('xl.Help', setter_error=METHOD_SETTER_ERROR)
    InchesToPoints = ProxyMethod('xl.InchesToPoints', setter_error=METHOD_SETTER_ERROR)
    InputBox = ProxyMethod('xl.InputBox', setter_error=METHOD_SETTER_ERROR)
    Intersect = ProxyMethod('xl.Intersect', setter_error=METHOD_SETTER_ERROR)
    MacroOptions = ProxyMethod('xl.MacroOptions', setter_error=METHOD_SETTER_ERROR)
    MailLogoff = ProxyMethod('xl.MailLogoff', setter_error=METHOD_SETTER_ERROR)
    MailLogon = ProxyMethod('xl.MailLogon', setter_error=METHOD_SETTER_ERROR)
    NextLetter = ProxyMethod('xl.NextLetter', setter_error=METHOD_SETTER_ERROR)
    OnKey = ProxyMethod('xl.OnKey', setter_error=METHOD_SETTER_ERROR)
    OnRepeat = ProxyMethod('xl.OnRepeat', setter_error=METHOD_SETTER_ERROR)
    OnTime = ProxyMethod('xl.OnTime', setter_error=METHOD_SETTER_ERROR)
    OnUndo = ProxyMethod('xl.OnUndo', setter_error=METHOD_SETTER_ERROR)
    Quit = ProxyMethod('xl.Quit', setter_error=METHOD_SETTER_ERROR)
    RecordMacro = ProxyMethod('xl.RecordMacro', setter_error=METHOD_SETTER_ERROR)
    RegisterXLL = ProxyMethod('xl.RegisterXLL', setter_error=METHOD_SETTER_ERROR)
    Repeat = ProxyMethod('xl.Repeat', setter_error=METHOD_SETTER_ERROR)
    Run = ProxyMethod('xl.Run', setter_error=METHOD_SETTER_ERROR)
    SendKeys = ProxyMethod('xl.SendKeys', setter_error=METHOD_SETTER_ERROR)
    SharePointVersion = ProxyMethod('xl.SharePointVersion', setter_error=METHOD_SETTER_ERROR)
    Undo = ProxyMethod('xl.Undo', setter_error=METHOD_SETTER_ERROR)
    Union = ProxyMethod('xl.Union', setter_error=METHOD_SETTER_ERROR)
    Volatile = ProxyMethod('xl.Volatile', setter_error=METHOD_SETTER_ERROR)
    Wait = ProxyMethod('xl.Wait', setter_error=METHOD_SETTER_ERROR)

    # ===== Application Properties ===== https://docs.microsoft.com/en-us/office/vba/api/excel.application(object)
    ActiveCell = ProxyProperty('xl.ActiveCell')
    # ActiveChart = ProxyProperty('xl.ActiveChart')
    ActiveEncryptionSession = ProxyProperty('xl.ActiveEncryptionSession')
    ActivePrinter = ProxyProperty('xl.ActivePrinter')
    ActiveProtectedViewWindow = ProxyProperty('xl.ActiveProtectedViewWindow')
    # ActiveSheet = ProxyProperty('xl.ActiveSheet')
    ActiveWindow = ProxyProperty('xl.ActiveWindow')
    ActiveWorkbook = ProxyProperty('xl.ActiveWorkbook')
    AddIns = ProxyProperty('xl.AddIns')
    AddIns2 = ProxyProperty('xl.AddIns2')
    AlertBeforeOverwriting = ProxyProperty('xl.AlertBeforeOverwriting')
    AltStartupPath = ProxyProperty('xl.AltStartupPath')
    AlwaysUseClearType = ProxyProperty('xl.AlwaysUseClearType')
    # Application = ProxyProperty('xl.Application')
    ArbitraryXMLSupportAvailable = ProxyProperty('xl.ArbitraryXMLSupportAvailable')
    AskToUpdateLinks = ProxyProperty('xl.AskToUpdateLinks')
    Assistance = ProxyProperty('xl.Assistance')
    AutoCorrect = ProxyProperty('xl.AutoCorrect')
    AutoFormatAsYouTypeReplaceHyperlinks = ProxyProperty('xl.AutoFormatAsYouTypeReplaceHyperlinks')
    AutomationSecurity = ProxyProperty('xl.AutomationSecurity')
    AutoPercentEntry = ProxyProperty('xl.AutoPercentEntry')
    AutoRecover = ProxyProperty('xl.AutoRecover')
    Build = ProxyProperty('xl.Build')
    CalculateBeforeSave = ProxyProperty('xl.CalculateBeforeSave')
    Calculation = ProxyProperty('xl.Calculation')
    CalculationInterruptKey = ProxyProperty('xl.CalculationInterruptKey')
    CalculationState = ProxyProperty('xl.CalculationState')
    # CalculationVersion = ProxyProperty('xl.CalculationVersion')
    Caller = ProxyProperty('xl.Caller')
    CanPlaySounds = ProxyProperty('xl.CanPlaySounds')
    CanRecordSounds = ProxyProperty('xl.CanRecordSounds')
    Caption = ProxyProperty('xl.Caption')
    CellDragAndDrop = ProxyProperty('xl.CellDragAndDrop')
    Cells = ProxyProperty('xl.Cells')
    # ChartDataPointTrack = ProxyProperty('xl.ChartDataPointTrack')
    # Charts = ProxyProperty('xl.Charts')
    ClipboardFormats = ProxyProperty('xl.ClipboardFormats')
    ClusterConnector = ProxyProperty('xl.ClusterConnector')
    Columns = ProxyProperty('xl.Columns')
    COMAddIns = ProxyProperty('xl.COMAddIns')
    # CommandBars = ProxyProperty('xl.CommandBars')
    CommandUnderlines = ProxyProperty('xl.CommandUnderlines')
    ConstrainNumeric = ProxyProperty('xl.ConstrainNumeric')
    ControlCharacters = ProxyProperty('xl.ControlCharacters')
    CopyObjectsWithCells = ProxyProperty('xl.CopyObjectsWithCells')
    # Creator = ProxyProperty('xl.Creator')
    Cursor = ProxyProperty('xl.Cursor')
    CursorMovement = ProxyProperty('xl.CursorMovement')
    CustomListCount = ProxyProperty('xl.CustomListCount')
    CutCopyMode = ProxyProperty('xl.CutCopyMode')
    DataEntryMode = ProxyProperty('xl.DataEntryMode')
    DDEAppReturnCode = ProxyProperty('xl.DDEAppReturnCode')
    DecimalSeparator = ProxyProperty('xl.DecimalSeparator')
    DefaultFilePath = ProxyProperty('xl.DefaultFilePath')
    DefaultSaveFormat = ProxyProperty('xl.DefaultSaveFormat')
    DefaultSheetDirection = ProxyProperty('xl.DefaultSheetDirection')
    DefaultWebOptions = ProxyProperty('xl.DefaultWebOptions')
    DeferAsyncQueries = ProxyProperty('xl.DeferAsyncQueries')
    Dialogs = ProxyProperty('xl.Dialogs')
    DisplayAlerts = ProxyProperty('xl.DisplayAlerts')
    DisplayClipboardWindow = ProxyProperty('xl.DisplayClipboardWindow')
    DisplayCommentIndicator = ProxyProperty('xl.DisplayCommentIndicator')
    DisplayDocumentActionTaskPane = ProxyProperty('xl.DisplayDocumentActionTaskPane')
    DisplayDocumentInformationPanel = ProxyProperty('xl.DisplayDocumentInformationPanel')
    DisplayExcel4Menus = ProxyProperty('xl.DisplayExcel4Menus')
    DisplayFormulaAutoComplete = ProxyProperty('xl.DisplayFormulaAutoComplete')
    DisplayFormulaBar = ProxyProperty('xl.DisplayFormulaBar')
    DisplayFullScreen = ProxyProperty('xl.DisplayFullScreen')
    DisplayFunctionToolTips = ProxyProperty('xl.DisplayFunctionToolTips')
    DisplayInsertOptions = ProxyProperty('xl.DisplayInsertOptions')
    DisplayNoteIndicator = ProxyProperty('xl.DisplayNoteIndicator')
    DisplayPasteOptions = ProxyProperty('xl.DisplayPasteOptions')
    DisplayRecentFiles = ProxyProperty('xl.DisplayRecentFiles')
    DisplayScrollBars = ProxyProperty('xl.DisplayScrollBars')
    DisplayStatusBar = ProxyProperty('xl.DisplayStatusBar')
    EditDirectlyInCell = ProxyProperty('xl.EditDirectlyInCell')
    EnableAnimations = ProxyProperty('xl.EnableAnimations')
    EnableAutoComplete = ProxyProperty('xl.EnableAutoComplete')
    EnableCancelKey = ProxyProperty('xl.EnableCancelKey')
    EnableCheckFileExtensions = ProxyProperty('xl.EnableCheckFileExtensions')
    EnableEvents = ProxyProperty('xl.EnableEvents')
    EnableLargeOperationAlert = ProxyProperty('xl.EnableLargeOperationAlert')
    EnableLivePreview = ProxyProperty('xl.EnableLivePreview')
    EnableMacroAnimations = ProxyProperty('xl.EnableMacroAnimations')
    EnableSound = ProxyProperty('xl.EnableSound')
    ErrorCheckingOptions = ProxyProperty('xl.ErrorCheckingOptions')
    # Excel4IntlMacroSheets = ProxyProperty('xl.Excel4IntlMacroSheets')
    # Excel4MacroSheets = ProxyProperty('xl.Excel4MacroSheets')
    ExtendList = ProxyProperty('xl.ExtendList')
    FeatureInstall = ProxyProperty('xl.FeatureInstall')
    FileConverters = ProxyProperty('xl.FileConverters')
    FileDialog = ProxyProperty('xl.FileDialog')
    FileExportConverters = ProxyProperty('xl.FileExportConverters')
    FileValidation = ProxyProperty('xl.FileValidation')
    FileValidationPivot = ProxyProperty('xl.FileValidationPivot')
    FindFormat = ProxyProperty('xl.FindFormat')
    FixedDecimal = ProxyProperty('xl.FixedDecimal')
    FixedDecimalPlaces = ProxyProperty('xl.FixedDecimalPlaces')
    FlashFill = ProxyProperty('xl.FlashFill')
    FlashFillMode = ProxyProperty('xl.FlashFillMode')
    FormulaBarHeight = ProxyProperty('xl.FormulaBarHeight')
    GenerateGetPivotData = ProxyProperty('xl.GenerateGetPivotData')
    GenerateTableRefs = ProxyProperty('xl.GenerateTableRefs')
    Height = ProxyProperty('xl.Height')
    HighQualityModeForGraphics = ProxyProperty('xl.HighQualityModeForGraphics')
    Hinstance = ProxyProperty('xl.Hinstance')
    HinstancePtr = ProxyProperty('xl.HinstancePtr')
    Hwnd = ProxyProperty('xl.Hwnd')
    IgnoreRemoteRequests = ProxyProperty('xl.IgnoreRemoteRequests')
    Interactive = ProxyProperty('xl.Interactive')
    International = ProxyProperty('xl.International')
    IsSandboxed = ProxyProperty('xl.IsSandboxed')
    Iteration = ProxyProperty('xl.Iteration')
    LanguageSettings = ProxyProperty('xl.LanguageSettings')
    LargeOperationCellThousandCount = ProxyProperty('xl.LargeOperationCellThousandCount')
    Left = ProxyProperty('xl.Left')
    LibraryPath = ProxyProperty('xl.LibraryPath')
    MailSession = ProxyProperty('xl.MailSession')
    MailSystem = ProxyProperty('xl.MailSystem')
    MapPaperSize = ProxyProperty('xl.MapPaperSize')
    MathCoprocessorAvailable = ProxyProperty('xl.MathCoprocessorAvailable')
    MaxChange = ProxyProperty('xl.MaxChange')
    MaxIterations = ProxyProperty('xl.MaxIterations')
    MeasurementUnit = ProxyProperty('xl.MeasurementUnit')
    MergeInstances = ProxyProperty('xl.MergeInstances')
    MouseAvailable = ProxyProperty('xl.MouseAvailable')
    MoveAfterReturn = ProxyProperty('xl.MoveAfterReturn')
    MoveAfterReturnDirection = ProxyProperty('xl.MoveAfterReturnDirection')
    MultiThreadedCalculation = ProxyProperty('xl.MultiThreadedCalculation')
    # Name = ProxyProperty('xl.Name')
    # Names = ProxyProperty('xl.Names')
    NetworkTemplatesPath = ProxyProperty('xl.NetworkTemplatesPath')
    NewWorkbook = ProxyProperty('xl.NewWorkbook')
    ODBCErrors = ProxyProperty('xl.ODBCErrors')
    ODBCTimeout = ProxyProperty('xl.ODBCTimeout')
    OLEDBErrors = ProxyProperty('xl.OLEDBErrors')
    OnWindow = ProxyProperty('xl.OnWindow')
    OperatingSystem = ProxyProperty('xl.OperatingSystem')
    OrganizationName = ProxyProperty('xl.OrganizationName')
    # Parent = ProxyProperty('xl.Parent')
    # Path = ProxyProperty('xl.Path')
    PathSeparator = ProxyProperty('xl.PathSeparator')
    PivotTableSelection = ProxyProperty('xl.PivotTableSelection')
    PreviousSelections = ProxyProperty('xl.PreviousSelections')
    PrintCommunication = ProxyProperty('xl.PrintCommunication')
    ProductCode = ProxyProperty('xl.ProductCode')
    PromptForSummaryInfo = ProxyProperty('xl.PromptForSummaryInfo')
    ProtectedViewWindows = ProxyProperty('xl.ProtectedViewWindows')
    QuickAnalysis = ProxyProperty('xl.QuickAnalysis')
    Range = ProxyProperty('xl.Range')
    Ready = ProxyProperty('xl.Ready')
    RecentFiles = ProxyProperty('xl.RecentFiles')
    RecordRelative = ProxyProperty('xl.RecordRelative')
    ReferenceStyle = ProxyProperty('xl.ReferenceStyle')
    RegisteredFunctions = ProxyProperty('xl.RegisteredFunctions')
    ReplaceFormat = ProxyProperty('xl.ReplaceFormat')
    RollZoom = ProxyProperty('xl.RollZoom')
    Rows = ProxyProperty('xl.Rows')
    RTD = ProxyProperty('xl.RTD')
    ScreenUpdating = ProxyProperty('xl.ScreenUpdating')
    Selection = ProxyProperty('xl.Selection')
    # Sheets = ProxyProperty('xl.Sheets')
    SheetsInNewWorkbook = ProxyProperty('xl.SheetsInNewWorkbook')
    ShowChartTipNames = ProxyProperty('xl.ShowChartTipNames')
    ShowChartTipValues = ProxyProperty('xl.ShowChartTipValues')
    ShowDevTools = ProxyProperty('xl.ShowDevTools')
    ShowMenuFloaties = ProxyProperty('xl.ShowMenuFloaties')
    ShowQuickAnalysis = ProxyProperty('xl.ShowQuickAnalysis')
    ShowSelectionFloaties = ProxyProperty('xl.ShowSelectionFloaties')
    ShowStartupDialog = ProxyProperty('xl.ShowStartupDialog')
    ShowToolTips = ProxyProperty('xl.ShowToolTips')
    SmartArtColors = ProxyProperty('xl.SmartArtColors')
    SmartArtLayouts = ProxyProperty('xl.SmartArtLayouts')
    SmartArtQuickStyles = ProxyProperty('xl.SmartArtQuickStyles')
    Speech = ProxyProperty('xl.Speech')
    SpellingOptions = ProxyProperty('xl.SpellingOptions')
    StandardFont = ProxyProperty('xl.StandardFont')
    StandardFontSize = ProxyProperty('xl.StandardFontSize')
    StartupPath = ProxyProperty('xl.StartupPath')
    StatusBar = ProxyProperty('xl.StatusBar')
    TemplatesPath = ProxyProperty('xl.TemplatesPath')
    ThisCell = ProxyProperty('xl.ThisCell')
    ThisWorkbook = ProxyProperty('xl.ThisWorkbook')
    ThousandsSeparator = ProxyProperty('xl.ThousandsSeparator')
    Top = ProxyProperty('xl.Top')
    TransitionMenuKey = ProxyProperty('xl.TransitionMenuKey')
    TransitionMenuKeyAction = ProxyProperty('xl.TransitionMenuKeyAction')
    TransitionNavigKeys = ProxyProperty('xl.TransitionNavigKeys')
    UsableHeight = ProxyProperty('xl.UsableHeight')
    UsableWidth = ProxyProperty('xl.UsableWidth')
    UseClusterConnector = ProxyProperty('xl.UseClusterConnector')
    UsedObjects = ProxyProperty('xl.UsedObjects')
    UserControl = ProxyProperty('xl.UserControl')
    UserLibraryPath = ProxyProperty('xl.UserLibraryPath')
    UserName = ProxyProperty('xl.UserName')
    UseSystemSeparators = ProxyProperty('xl.UseSystemSeparators')
    Value = ProxyProperty('xl.Value')
    VBE = ProxyProperty('xl.VBE')
    Version = ProxyProperty('xl.Version')
    Visible = ProxyProperty('xl.Visible')
    WarnOnFunctionNameConflict = ProxyProperty('xl.WarnOnFunctionNameConflict')
    Watches = ProxyProperty('xl.Watches')
    Width = ProxyProperty('xl.Width')
    # Windows = ProxyProperty('xl.Windows')
    WindowsForPens = ProxyProperty('xl.WindowsForPens')
    WindowState = ProxyProperty('xl.WindowState')
    Workbooks = ProxyProperty('xl.Workbooks')
    WorksheetFunction = ProxyProperty('xl.WorksheetFunction')
    # Worksheets = ProxyProperty('xl.Worksheets')


SHOULD_INIT_SIGNAL = True


def should_init_sig():
    """Return if init_sig_shutdown should be called on first Excel creation.

    This will set the SIGTERM and SIGINT handlers to call "shutdown" and close all Excel Applications run by this
    process.
    """
    global SHOULD_INIT_SIGNAL
    return SHOULD_INIT_SIGNAL


def set_init_sig(value):
    """Set if init_sig_shutdown should be called on first Excel creation.

    This will set the SIGTERM and SIGINT handlers to call "shutdown" and close all Excel Applications run by this
    process.
    """
    global SHOULD_INIT_SIGNAL
    SHOULD_INIT_SIGNAL = value


def init_sig_shutdown(func=None):
    """Set the SIGTERM and SIGINT handlers to call "shutdown" and close all Excel Applications run by this process."""
    if func is None:
        func = shutdown
    signal.signal(signal.SIGTERM, func)
    signal.signal(signal.SIGINT, func)


def shutdown(*args, sys_exit=True, **kwargs):
    """Close all Excel Applications run by this process."""
    for excel in Excel.GLOBAL_EXCELS:
        try:
            excel.Quit()
        except (AttributeError, ValueError, TypeError, Excel):
            pass
    if sys_exit:
        sys.exit(-1)


# Automatically register the shutdown function with atexit
atexit.register(shutdown, sys_exit=False)
