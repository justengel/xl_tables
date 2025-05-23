from ..prop_utils import HashDict, ItemStorage


def fake_proxy_method(*args, **kwargs):
    def caller(*args, **kwargs):
        pass

    return caller


def fake_proxy_property(*args, **kwargs):
    def get(self):
        return None
    
    def set(self, value):
        return
    
    return property(get, set)


class Excel:
    METHOD_SETTER_ERROR = 'Cannot set property '
    
    DEFAULT_SETTINGS = {
        'DisplayAlerts': False,  # If False hides prompt to save
        'AutomationSecurity': 1,
        'Interactive': False,
        'PrintCommunication': False,
        }

    GLOBAL_EXCELS = ItemStorage()

    def __init__(self, *args, **kwargs):
        super().__init__()

    def get_settings(self, **settings):
        """Get the current settings."""
        return {}

    def set_settings(self, **settings):
        """Set the given attribute settings."""
        pass

    @property
    def xl(self):
        """Get (or create) the Excel Application attribute."""
        return None

    @xl.setter
    def xl(self, value):
        """Set the Excel Application attribute."""
        return

    def __enter__(self):
        return self

    def __exit__(self, exception_type, exception_value, traceback):
        self.Quit()
        return exception_type is None

    ActivateMicrosoftApp = fake_proxy_method()
    AddCustomList = fake_proxy_method()
    Calculate = fake_proxy_method()
    CalculateFull = fake_proxy_method()
    CalculateFullRebuild = fake_proxy_method()
    CalculateUntilAsyncQueriesDone = fake_proxy_method()
    CentimetersToPoints = fake_proxy_method()
    CheckAbort = fake_proxy_method()
    CheckSpelling = fake_proxy_method()
    ConvertFormula = fake_proxy_method()
    DDEExecute = fake_proxy_method()
    DDEInitiate = fake_proxy_method()
    DDEPoke = fake_proxy_method()
    DDERequest = fake_proxy_method()
    DDETerminate = fake_proxy_method()
    DeleteCustomList = fake_proxy_method()
    DisplayXMLSourcePane = fake_proxy_method()
    DoubleClick = fake_proxy_method()
    Evaluate = fake_proxy_method()
    ExecuteExcel4Macro = fake_proxy_method()
    FindFile = fake_proxy_method()
    GetCustomListContents = fake_proxy_method()
    GetCustomListNum = fake_proxy_method()
    GetOpenFilename = fake_proxy_method()
    GetPhonetic = fake_proxy_method()
    GetSaveAsFilename = fake_proxy_method()
    Goto = fake_proxy_method()
    Help = fake_proxy_method()
    InchesToPoints = fake_proxy_method()
    InputBox = fake_proxy_method()
    Intersect = fake_proxy_method()
    MacroOptions = fake_proxy_method()
    MailLogoff = fake_proxy_method()
    MailLogon = fake_proxy_method()
    NextLetter = fake_proxy_method()
    OnKey = fake_proxy_method()
    OnRepeat = fake_proxy_method()
    OnTime = fake_proxy_method()
    OnUndo = fake_proxy_method()
    Quit = fake_proxy_method()
    RecordMacro = fake_proxy_method()
    RegisterXLL = fake_proxy_method()
    Repeat = fake_proxy_method()
    Run = fake_proxy_method()
    SendKeys = fake_proxy_method()
    SharePointVersion = fake_proxy_method()
    Undo = fake_proxy_method()
    Union = fake_proxy_method()
    Volatile = fake_proxy_method()
    Wait = fake_proxy_method()

    # Properties
    ActiveCell = fake_proxy_property()
    ActiveChart = fake_proxy_property()
    ActiveEncryptionSession = fake_proxy_property()
    ActivePrinter = fake_proxy_property()
    ActiveProtectedViewWindow = fake_proxy_property()
    ActiveSheet = fake_proxy_property()
    ActiveWindow = fake_proxy_property()
    ActiveWorkbook = fake_proxy_property()
    AddIns = fake_proxy_property()
    AddIns2 = fake_proxy_property()
    AlertBeforeOverwriting = fake_proxy_property()
    AltStartupPath = fake_proxy_property()
    AlwaysUseClearType = fake_proxy_property()
    Application = fake_proxy_property()
    ArbitraryXMLSupportAvailable = fake_proxy_property()
    AskToUpdateLinks = fake_proxy_property()
    Assistance = fake_proxy_property()
    AutoCorrect = fake_proxy_property()
    AutoFormatAsYouTypeReplaceHyperlinks = fake_proxy_property()
    AutomationSecurity = fake_proxy_property()
    AutoPercentEntry = fake_proxy_property()
    AutoRecover = fake_proxy_property()
    Build = fake_proxy_property()
    CalculateBeforeSave = fake_proxy_property()
    Calculation = fake_proxy_property()
    CalculationInterruptKey = fake_proxy_property()
    CalculationState = fake_proxy_property()
    CalculationVersion = fake_proxy_property()
    Caller = fake_proxy_property()
    CanPlaySounds = fake_proxy_property()
    CanRecordSounds = fake_proxy_property()
    Caption = fake_proxy_property()
    CellDragAndDrop = fake_proxy_property()
    Cells = fake_proxy_property()
    ChartDataPointTrack = fake_proxy_property()
    Charts = fake_proxy_property()
    ClipboardFormats = fake_proxy_property()
    ClusterConnector = fake_proxy_property()
    Columns = fake_proxy_property()
    COMAddIns = fake_proxy_property()
    CommandBars = fake_proxy_property()
    CommandUnderlines = fake_proxy_property()
    ConstrainNumeric = fake_proxy_property()
    ControlCharacters = fake_proxy_property()
    CopyObjectsWithCells = fake_proxy_property()
    Creator = fake_proxy_property()
    Cursor = fake_proxy_property()
    CursorMovement = fake_proxy_property()
    CustomListCount = fake_proxy_property()
    CutCopyMode = fake_proxy_property()
    DataEntryMode = fake_proxy_property()
    DDEAppReturnCode = fake_proxy_property()
    DecimalSeparator = fake_proxy_property()
    DefaultFilePath = fake_proxy_property()
    DefaultSaveFormat = fake_proxy_property()
    DefaultSheetDirection = fake_proxy_property()
    DefaultWebOptions = fake_proxy_property()
    DeferAsyncQueries = fake_proxy_property()
    Dialogs = fake_proxy_property()
    DisplayAlerts = fake_proxy_property()
    DisplayClipboardWindow = fake_proxy_property()
    DisplayCommentIndicator = fake_proxy_property()
    DisplayDocumentActionTaskPane = fake_proxy_property()
    DisplayDocumentInformationPanel = fake_proxy_property()
    DisplayExcel4Menus = fake_proxy_property()
    DisplayFormulaAutoComplete = fake_proxy_property()
    DisplayFormulaBar = fake_proxy_property()
    DisplayFullScreen = fake_proxy_property()
    DisplayFunctionToolTips = fake_proxy_property()
    DisplayInsertOptions = fake_proxy_property()
    DisplayNoteIndicator = fake_proxy_property()
    DisplayPasteOptions = fake_proxy_property()
    DisplayRecentFiles = fake_proxy_property()
    DisplayScrollBars = fake_proxy_property()
    DisplayStatusBar = fake_proxy_property()
    EditDirectlyInCell = fake_proxy_property()
    EnableAnimations = fake_proxy_property()
    EnableAutoComplete = fake_proxy_property()
    EnableCancelKey = fake_proxy_property()
    EnableCheckFileExtensions = fake_proxy_property()
    EnableEvents = fake_proxy_property()
    EnableLargeOperationAlert = fake_proxy_property()
    EnableLivePreview = fake_proxy_property()
    EnableMacroAnimations = fake_proxy_property()
    EnableSound = fake_proxy_property()
    ErrorCheckingOptions = fake_proxy_property()
    Excel4IntlMacroSheets = fake_proxy_property()
    Excel4MacroSheets = fake_proxy_property()
    ExtendList = fake_proxy_property()
    FeatureInstall = fake_proxy_property()
    FileConverters = fake_proxy_property()
    FileDialog = fake_proxy_property()
    FileExportConverters = fake_proxy_property()
    FileValidation = fake_proxy_property()
    FileValidationPivot = fake_proxy_property()
    FindFormat = fake_proxy_property()
    FixedDecimal = fake_proxy_property()
    FixedDecimalPlaces = fake_proxy_property()
    FlashFill = fake_proxy_property()
    FlashFillMode = fake_proxy_property()
    FormulaBarHeight = fake_proxy_property()
    GenerateGetPivotData = fake_proxy_property()
    GenerateTableRefs = fake_proxy_property()
    Height = fake_proxy_property()
    HighQualityModeForGraphics = fake_proxy_property()
    Hinstance = fake_proxy_property()
    HinstancePtr = fake_proxy_property()
    Hwnd = fake_proxy_property()
    IgnoreRemoteRequests = fake_proxy_property()
    Interactive = fake_proxy_property()
    International = fake_proxy_property()
    IsSandboxed = fake_proxy_property()
    Iteration = fake_proxy_property()
    LanguageSettings = fake_proxy_property()
    LargeOperationCellThousandCount = fake_proxy_property()
    Left = fake_proxy_property()
    LibraryPath = fake_proxy_property()
    MailSession = fake_proxy_property()
    MailSystem = fake_proxy_property()
    MapPaperSize = fake_proxy_property()
    MathCoprocessorAvailable = fake_proxy_property()
    MaxChange = fake_proxy_property()
    MaxIterations = fake_proxy_property()
    MeasurementUnit = fake_proxy_property()
    MergeInstances = fake_proxy_property()
    MouseAvailable = fake_proxy_property()
    MoveAfterReturn = fake_proxy_property()
    MoveAfterReturnDirection = fake_proxy_property()
    MultiThreadedCalculation = fake_proxy_property()
    Name = fake_proxy_property()
    Names = fake_proxy_property()
    NetworkTemplatesPath = fake_proxy_property()
    NewWorkbook = fake_proxy_property()
    ODBCErrors = fake_proxy_property()
    ODBCTimeout = fake_proxy_property()
    OLEDBErrors = fake_proxy_property()
    OnWindow = fake_proxy_property()
    OperatingSystem = fake_proxy_property()
    OrganizationName = fake_proxy_property()
    Parent = fake_proxy_property()
    Path = fake_proxy_property()
    PathSeparator = fake_proxy_property()
    PivotTableSelection = fake_proxy_property()
    PreviousSelections = fake_proxy_property()
    PrintCommunication = fake_proxy_property()
    ProductCode = fake_proxy_property()
    PromptForSummaryInfo = fake_proxy_property()
    ProtectedViewWindows = fake_proxy_property()
    QuickAnalysis = fake_proxy_property()
    Range = fake_proxy_property()
    Ready = fake_proxy_property()
    RecentFiles = fake_proxy_property()
    RecordRelative = fake_proxy_property()
    ReferenceStyle = fake_proxy_property()
    RegisteredFunctions = fake_proxy_property()
    ReplaceFormat = fake_proxy_property()
    RollZoom = fake_proxy_property()
    Rows = fake_proxy_property()
    RTD = fake_proxy_property()
    ScreenUpdating = fake_proxy_property()
    Selection = fake_proxy_property()
    Sheets = fake_proxy_property()
    SheetsInNewWorkbook = fake_proxy_property()
    ShowChartTipNames = fake_proxy_property()
    ShowChartTipValues = fake_proxy_property()
    ShowDevTools = fake_proxy_property()
    ShowMenuFloaties = fake_proxy_property()
    ShowQuickAnalysis = fake_proxy_property()
    ShowSelectionFloaties = fake_proxy_property()
    ShowStartupDialog = fake_proxy_property()
    ShowToolTips = fake_proxy_property()
    SmartArtColors = fake_proxy_property()
    SmartArtLayouts = fake_proxy_property()
    SmartArtQuickStyles = fake_proxy_property()
    Speech = fake_proxy_property()
    SpellingOptions = fake_proxy_property()
    StandardFont = fake_proxy_property()
    StandardFontSize = fake_proxy_property()
    StartupPath = fake_proxy_property()
    StatusBar = fake_proxy_property()
    TemplatesPath = fake_proxy_property()
    ThisCell = fake_proxy_property()
    ThisWorkbook = fake_proxy_property()
    ThousandsSeparator = fake_proxy_property()
    Top = fake_proxy_property()
    TransitionMenuKey = fake_proxy_property()
    TransitionMenuKeyAction = fake_proxy_property()
    TransitionNavigKeys = fake_proxy_property()
    UsableHeight = fake_proxy_property()
    UsableWidth = fake_proxy_property()
    UseClusterConnector = fake_proxy_property()
    UsedObjects = fake_proxy_property()
    UserControl = fake_proxy_property()
    UserLibraryPath = fake_proxy_property()
    UserName = fake_proxy_property()
    UseSystemSeparators = fake_proxy_property()
    Value = fake_proxy_property()
    VBE = fake_proxy_property()
    Version = fake_proxy_property()
    Visible = fake_proxy_property()
    WarnOnFunctionNameConflict = fake_proxy_property()
    Watches = fake_proxy_property()
    Width = fake_proxy_property()
    Windows = fake_proxy_property()
    WindowsForPens = fake_proxy_property()
    WindowState = fake_proxy_property()
    Workbooks = fake_proxy_property()
    WorksheetFunction = fake_proxy_property()
    Worksheets = fake_proxy_property()