Attribute VB_Name = "gemini1"
Option Explicit

' ============================================================
' CMS XT EXPORT MACRO - FIXED / SPEED + ROBUSTNESS UPDATE
' ------------------------------------------------------------
' UPDATED IN THIS VERSION:
'  1. Corrected TCP/top-side orientation axis mapping:
'       Y stack -> *Top / *Bottom
'       Z stack -> *Front / *Back
'       X stack -> *Right / *Left
'  2. Orientation now tries top/bottom HOLDER, POT, INS, then TCP/BCP.
'  3. CMS_TOP is saved from matched top-side orientation.
'  4. StabilizeActiveView restored.
'  5. DXF 1:1 setting remains forced for all DXF views/files.
'  6. Later parts include J BLOCK native-assembly DXF fix.
' ============================================================

' NOTE: No Windows API Declare is used. Waiting is handled by WaitMilliseconds.

' ============================================================
' USER SETTINGS
' ============================================================

Private Const ROOT_JOB_PATH As String = "C:\Users\lenovo\Desktop\000000005.May 2026"
Private Const EXTRACT_FOLDER_NAME As String = "_EXTRACTED_ZIP"
Private Const OUTPUT_FOLDER_SUFFIX As String = " PRINTS"

Private Const LOCAL_WORKSPACE_ROOT As String = "C:\CMS_Local_Workspace"

Private Const PUSH_OUTPUTS_TO_NETWORK As Boolean = False
Private Const DELETE_EXTRACTED_ZIP_AFTER_FLATTEN As Boolean = True

Private Const USE_ACTIVE_SOLIDWORKS_DOC_FIRST As Boolean = False
Private Const SHOW_ERROR_MESSAGES As Boolean = True

Private Const READ_PDF_BOM_WITH_PDFTOTEXT As Boolean = True
Private Const PDFTOTEXT_EXE As String = "C:\Users\lenovo\Downloads\New folder (9)\poppler-26.02.0\Library\bin\pdftotext.exe"

Private Const ONLY_INCLUDE_4140_BOM_ITEMS As Boolean = True
Private Const DEFAULT_STEEL_TYPE As String = "4140"

Private Const HIDE_QUARTER_INCH_THICKNESS As Boolean = True
Private Const PHYSICALLY_HIDE_250_BEFORE_SCAN As Boolean = False
Private Const QUARTER_INCH_THICKNESS As Double = 0.25
Private Const QUARTER_INCH_TOLERANCE As Double = 0.01

Private Const DIM_DECIMALS As Long = 3
Private Const DIM_OK_TOL As Double = 0.03
Private Const DIM_REVIEW_TOL As Double = 0.125
Private Const DIM_MAX_MATCH_TOTAL_DIFF As Double = 5#
Private Const MIN_STEEL_VOLUME_CUIN As Double = 1#

Private Const INCHES_PER_METER As Double = 39.3700787401575
Private Const CUIN_PER_CUBIC_METER As Double = 61023.7440947323

Private Const SAME_SIZE_PAIR_TOL As Double = 0.125

Private Const TURBO_READ_ONLY_BOM_SHEET As Boolean = True
Private Const TURBO_BOM_SHEET_NAME As String = "BOM"
Private Const BOM_HEADER_SEARCH_MAX_ROWS As Long = 150
Private Const STOP_BOM_READ_AFTER_BLANK_ROWS As Long = 20
Private Const VERBOSE_PROGRESS_LOG As Boolean = False

Private Const PROMPT_FOR_TOP_ORIENTATION As Boolean = False
Private Const CMS_TOP_VIEW_NAME As String = "CMS_TOP"

Private Const CMS_BASE_TOP_VIEW_NAME As String = "*Bottom"
Private Const CMS_BASE_TOP_VIEW_ID As Long = 6

' If the corrected top-side face is right but rotated on screen, adjust this:
' 0 = no rotation, 1 = rotate +Z once, 2 = 180, -1 = rotate -Z once.
Private Const CMS_TOP_ROTATE_Z_STEPS As Long = 0

Private Const AUTO_SELECT_TCP_TOP_ORIENTATION As Boolean = True

Private Const TCP_TOP_ORIENTATION_KEYS As String = "TCP|TOP CLAMPING|TOP CLAMPING PLATE|TOP SMED|TOP SMED PLATE|ID SMED"
Private Const BCP_BOTTOM_ORIENTATION_KEYS As String = "BCP|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|BOT CLAMPING|BOT CLAMPING PLATE|BOTTOM SMED|BOT SMED|OD SMED"

Private Const PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE As Boolean = True

' ============================================================
' SPEED / GRAPHICS SETTINGS
' ============================================================

Private Const DISABLE_STABILIZE_DELAYS As Boolean = True
Private Const DISABLE_MAIN_VIEWPORT_GRAPHICS As Boolean = True
Private Const RUN_SOLIDWORKS_INVISIBLE As Boolean = True

' ============================================================
' TOP/BOT INSERT GEOMETRY FALLBACK
' ============================================================

Private Const ADD_TOP_BOT_INS_FROM_CAD_GEOMETRY As Boolean = True
Private Const INSERT_THICKNESS_TOL As Double = 0.035
Private Const INSERT_WIDTH_MATCH_TOL As Double = 0.25
Private Const INSERT_LENGTH_MATCH_TOL As Double = 0.75
Private Const INSERT_LENGTH_MATCH_TOL_LOOSE As Double = 1.5

' ============================================================
' DXF SETTINGS
' ============================================================

Private Const CREATE_INDIVIDUAL_DXFS As Boolean = False
Private Const CREATE_DXFS_DURING_XT_SAVE As Boolean = True

' UPDATED: this is now honored for BASE/HOLDERS too in later DXF code.
Private Const FORCE_ALL_DXF_VIEWS_1_TO_1 As Boolean = True

Private Const FREEZE_DXF_DRAWING_GRAPHICS As Boolean = True

Private Const FLIP_ID_HOLDER_CENTER_VIEW_180 As Boolean = True
Private Const FLIP_ID_HOLDER_CENTER_VIEW_180_FROM_ASSEMBLY As Boolean = False

Private Const OD_HOLDER_CENTER_ROTATION_DEG As Double = 180#

Private Const DIMENSION_J_BLOCK As Boolean = True
Private Const MULTIVIEW_FIT_SAFETY As Double = 0.9

Private Const SW_DRAWING_TEMPLATE_PATH As String = "C:\ProgramData\SolidWorks\SOLIDWORKS 2023\templates\Drawing.drwdot"

Private Const E_SHEET_WIDTH_IN As Double = 44#
Private Const E_SHEET_HEIGHT_IN As Double = 34#
Private Const DXF_MARGIN_IN As Double = 1#
Private Const DXF_MAX_SCALE As Double = 1#
Private Const DXF_PROJECTED_VIEW_GAP_IN As Double = 2.25
Private Const HOLDERS_SIDE_VIEW_EXTRA_GAP_IN As Double = 6#
Private Const PI_VALUE As Double = 3.14159265358979

' ============================================================
' PULL CORE DIMENSIONS EXCEL REPORT
' ============================================================

Private Const CREATE_PULLCORE_DIMENSIONS_EXCEL As Boolean = True
Private Const PULLCORE_DIMENSIONS_REPORT_FILE As String = "Pull Core Dimensions.xlsx"

' ============================================================
' MAIN ASSEMBLY / HOLDERS PACKAGE SETTINGS
' ============================================================

Private Const CREATE_MAIN_ASSEMBLY_PACKAGE As Boolean = True
Private Const MAIN_ASSEMBLY_FILE_TOKEN As String = "BASE"
Private Const HOLDERS_ONLY_FILE_TOKEN As String = "HOLDERS"

Private Const ID_HOLDER_KEYS As String = "ID HOLDER|TOP HOLDER|IDTE HOLDER|TOP CARRIER|ID CARRIER"
Private Const OD_HOLDER_KEYS As String = "OD HOLDER|BOTTOM HOLDER|BOT HOLDER|ODTE HOLDER|BOTTOM CARRIER|OD CARRIER"

' ============================================================
' J BLOCK PACKAGE SETTINGS
' ============================================================

Private Const CREATE_J_BLOCK_PACKAGE As Boolean = True
Private Const J_BLOCK_FOLDER_NAME As String = "J BLOCK"

Private Const J_BLOCK_NAME_KEYS As String = "EJECTOR J-BLOCK|EJECTOR J BLOCK|EJECTOR JBLOCK|EJ J-BLOCK|EJ J BLOCK|EJ JBLOCK|EJ. J-BLOCK|EJ. J BLOCK|EJ.J-BLOCK|EJ.J BLOCK|J-BLOCK|J BLOCK|JBLOCK"

Private Const J_BLOCK_DIMENSION_DECIMALS As Long = 4

Private Const J_BLOCK_PARENT_VIEW_PRIMARY As String = "*Right"
Private Const J_BLOCK_PARENT_VIEW_FALLBACK As String = "*Right"

Private Const ROTATE_J_BLOCK_DXF_180 As Boolean = False

Private Const J_BLOCK_TARGET_THICKNESS As Double = 1#
Private Const J_BLOCK_TARGET_WIDTH As Double = 2.75
Private Const J_BLOCK_TARGET_LENGTH As Double = 3.838
Private Const J_BLOCK_DIM_MATCH_TOL As Double = 0.75

Private Const EJECTOR_CAM_NAME_KEYS As String = "EJECTOR CAM|EJ CAM|EJ. CAM|EJ.CAM|CAM"
Private Const EJECTOR_CAM_TARGET_THICKNESS As Double = 1#
Private Const EJECTOR_CAM_TARGET_WIDTH As Double = 1.28
Private Const EJECTOR_CAM_TARGET_LENGTH As Double = 1.63
Private Const EJECTOR_CAM_THICKNESS_TOL As Double = 0.35
Private Const EJECTOR_CAM_WIDTH_TOL As Double = 0.5
Private Const EJECTOR_CAM_LENGTH_TOL As Double = 0.6

Private Const J_BLOCK_BASE_FILE_TOKEN As String = "J BLOCK"
Private Const J_BLOCK_CAM_FILE_TOKEN As String = "CAM"

Private JBlockTgtL As Double
Private JBlockTgtW As Double
Private JBlockTgtT As Double
Private EjCamTgtL As Double
Private EjCamTgtW As Double
Private EjCamTgtT As Double

' ============================================================
' PULLCORE CAM AND KEY PACKAGE SETTINGS
' ============================================================

Private Const CREATE_PULLCORE_CAM_KEY_PACKAGE As Boolean = True
Private Const PULLCORE_CAM_KEY_FOLDER_NAME As String = "PULLCORE CAM AND KEY"

Private Const AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT As Boolean = True
Private Const PULLCORE_ID_OD_HEIGHT_AXIS As String = "AUTO"
Private Const PULLCORE_ID_IS_HIGHER As Boolean = True

Private Const PULLCORE_T_TOL As Double = 0.175
Private Const PULLCORE_W_TOL As Double = 0.25
Private Const PULLCORE_L_TOL As Double = 0.35

Private Const USE_PULLCORE_BEST_FIT_BBOX As Boolean = True
Private Const PULLCORE_BEST_FIT_ALL_PARTS As Boolean = False

Private Const STRAIGHTEN_PULLCORE_DXF As Boolean = True
Private Const PULLCORE_STRAIGHTEN_SIGN As Double = -1#
Private Const PULLCORE_STRAIGHTEN_MIN_DEG As Double = 2#
Private Const PULLCORE_FALLBACK_ANGLE_DEG As Double = 32.5

' ============================================================
' PULLCORE STOP / FLIPPER CAM COVER PACKAGE SETTINGS
' ============================================================

Private Const CREATE_PULLCORE_STOP_PACKAGE As Boolean = True
Private Const PULLCORE_STOP_FOLDER_NAME As String = "PULLCORE STOP"

Private Const PULLCORE_STOP_NAME_KEYS As String = "PULLCORE STOP|PULL CORE STOP"
Private Const FLIPPER_CAM_COVER_NAME_KEYS As String = "FLIPPER CAM COVER PLATE|FLIPPER CAM COVER|CAM COVER PLATE"

Private Const FLIPPER_CAM_COVER_TARGET_THICKNESS As Double = 2.95
Private Const FLIPPER_CAM_COVER_TARGET_WIDTH As Double = 4.225
Private Const FLIPPER_CAM_COVER_TARGET_LENGTH As Double = 5.975
Private Const FLIPPER_CAM_COVER_DIM_TOL As Double = 0.35

' ============================================================
' PYROPEL SETTINGS
' ============================================================

Private Const ROUTE_PYROPEL_TO_SEPARATE_FOLDERS As Boolean = True
Private Const PYROPEL_POTS_FOLDER_SUFFIX As String = " PYROPEL POTS"
Private Const PYROPEL_HOLDERS_FOLDER_SUFFIX As String = " PYROPEL HOLDERS"

' ============================================================
' SOLIDWORKS / EXCEL CONSTANTS
' ============================================================

Private Const swDocNONE As Long = 0
Private Const swDocPART As Long = 1
Private Const swDocASSEMBLY As Long = 2
Private Const swDocDRAWING As Long = 3

Private Const swOpenDocOptions_Silent As Long = 1
Private Const swOpenDocOptions_ReadOnly As Long = 2

Private Const swSaveAsCurrentVersion As Long = 0
Private Const swSaveAsOptions_Silent As Long = 1

Private Const swSolidBody As Long = 0

Private Const swComponentHidden As Long = 0
Private Const swComponentVisible As Long = 1

Private Const xlCalculationManual As Long = -4135
Private Const xlOpenXMLWorkbook As Long = 51

' ============================================================
' TYPES
' ============================================================

Private Type PartInfo
    componentName As String
    cleanName As String
    filePath As String
    configName As String
    bodyName As String
    Quantity As Long

    Length As Double
    Width As Double
    Thickness As Double
    BBoxVolume As Double
    massValue As Double

    hasOriginalAsmBBox As Boolean
    OriginalAsmLength As Double
    OriginalAsmWidth As Double
    OriginalAsmThickness As Double

    hasAsmCenter As Boolean
    AsmCenterX As Double
    AsmCenterY As Double
    AsmCenterZ As Double

    hasTopRotation As Boolean
    topRotationRad As Double

    UsedForBomMatch As Boolean
    isBodyOnly As Boolean
End Type

Private Type BomInfo
    Description As String
    quoteName As String
    TypeField As String
    Quantity As Long
    material As String

    BomLength As Double
    BomWidth As Double
    BomThickness As Double
    hasDims As Boolean
End Type

Private Type ExportInfo
    quoteName As String
    Quantity As Long
    material As String

    CadPartIndex As Long
    HasCad As Boolean

    Thickness As Double
    Width As Double
    Length As Double

    BomThickness As Double
    BomWidth As Double
    BomLength As Double
    HasBomDims As Boolean

    Status As String
End Type

Private Type PullcoreMatchInfo
    quoteName As String
    Description As String
    material As String
    Quantity As Long

    CadPartIndex As Long
    isCam As Boolean

    BomThickness As Double
    BomWidth As Double
    BomLength As Double

    CadThickness As Double
    CadWidth As Double
    CadLength As Double

    OriginalThickness As Double
    OriginalWidth As Double
    OriginalLength As Double

    FittedThickness As Double
    FittedWidth As Double
    FittedLength As Double

    DetectedAngleDeg As Double
    DxfRotationDeg As Double

    Status As String
End Type

' ============================================================
' GLOBALS
' ============================================================

Private swApp As Object
Private swModel As Object
Private swAssy As Object

Private parts() As PartInfo
Private PartCount As Long

Private BomRows() As BomInfo
Private BomCount As Long

Private ExportRows() As ExportInfo
Private ExportCount As Long

Private PullcoreMatches() As PullcoreMatchInfo
Private PullcoreMatchCount As Long

Private RunLogPath As String
Private StartupLogPath As String
Private CurrentJobFolder As String
Private CurrentJobNumber As String

Private NetworkJobFolder As String
Private LocalJobFolder As String
Private BaseNativeAssemblyPath As String

Private CustomerNumber As String
Private DateCode As String
Private NamingSourceText As String

Private ExportFilePaths As Object
Private SpecialBomCadMatches As Object
Private SpecialBomCadQuoteNames As Object
Private PullcoreBestFitDimCache As Object

Private MacroStartTime As Date
Private StepStartTime As Date
Private CurrentStepName As String

Private MainCadOpenedByMacro As Boolean
Private MainCadTitleForClose As String

Private DxfFreezeDoc As Object
Private MainViewportGraphicsDisabled As Boolean

Private LastJobFailReason As String

Private CurrentDxfForce1to1 As Boolean
Private CurrentIdHolderDxfFromAssembly As Boolean
Private CurrentDxfKeepImportedOrientation As Boolean
Private CurrentDxfStraightenAngleRad As Double

Private CurrentPullcoreStraightenCadIndex As Long

Private PullCoreDimensionsReportPath As String
Private PullcoreIdOdHeightAxisUsed As String

' ============================================================
' MAIN - BATCH CONTROLLER
' ============================================================

Sub main()
On Error GoTo ErrHandler

    Set swApp = Application.SldWorks

    If RUN_SOLIDWORKS_INVISIBLE Then
        On Error Resume Next
        swApp.Visible = False
        On Error GoTo ErrHandler
    End If

    MacroStartTime = Now
    StartupLogPath = Environ$("USERPROFILE") & "\Desktop\CMS_XT_Export_STARTUP_Log.txt"
    RunLogPath = StartupLogPath

    LogLine "========================================"
    LogLine "MACRO STARTED"
    LogLine "Root path: " & ROOT_JOB_PATH
    LogLine "Local workspace root: " & LOCAL_WORKSPACE_ROOT
    LogLine "Push outputs to network: " & CStr(PUSH_OUTPUTS_TO_NETWORK)
    LogLine "Run SolidWorks invisible: " & CStr(RUN_SOLIDWORKS_INVISIBLE)
    LogLine "Create Pull Core Dimensions Excel: " & CStr(CREATE_PULLCORE_DIMENSIONS_EXCEL)
    LogLine "Auto-select TCP top orientation: " & CStr(AUTO_SELECT_TCP_TOP_ORIENTATION)
    LogLine "Auto-label pullcore ID/OD by height: " & CStr(AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT)
    LogLine "OD holder center rotation deg: " & CStr(OD_HOLDER_CENTER_ROTATION_DEG)
    LogLine "FORCE_ALL_DXF_VIEWS_1_TO_1: " & CStr(FORCE_ALL_DXF_VIEWS_1_TO_1)
    LogLine "========================================"

    Dim jobInput As String
    jobInput = Trim(InputBox("Enter one or multiple job numbers/search texts." & vbCrLf & _
                             "Examples:" & vbCrLf & _
                             "J8410" & vbCrLf & _
                             "J8410, J8411, J8412" & vbCrLf & _
                             "833200084 J8413", _
                             "CMS XT Export Batch"))

    If jobInput = "" Then GoTo NormalEnd

    Dim jobs As Collection
    Set jobs = ParseJobInputList(jobInput)

    If jobs Is Nothing Or jobs.count = 0 Then
        MsgBox "No valid job numbers entered.", vbExclamation
        GoTo NormalEnd
    End If

    Dim completed As Collection
    Dim failed As Collection

    Set completed = New Collection
    Set failed = New Collection

    Dim i As Long
    Dim jobText As String
    Dim ok As Boolean

    For i = 1 To jobs.count

        jobText = UCase(Trim(CStr(jobs(i))))

        If jobText <> "" Then

            RunLogPath = StartupLogPath
            LogLine "========================================"
            LogLine "BATCH ITEM " & i & "/" & jobs.count & ": " & jobText
            LogLine "========================================"

            ok = ProcessOneJob(jobText)

            If ok Then
                completed.Add jobText
            Else
                failed.Add jobText & IIf(LastJobFailReason <> "", "  ->  " & LastJobFailReason, "")
            End If

            CloseAllDocumentsSafely
            Set swModel = Nothing
            Set swAssy = Nothing
            Set SpecialBomCadMatches = Nothing
            Set SpecialBomCadQuoteNames = Nothing
            Set PullcoreBestFitDimCache = Nothing

            Erase parts
            Erase BomRows
            Erase ExportRows
            Erase PullcoreMatches

            PartCount = 0
            BomCount = 0
            ExportCount = 0
            PullcoreMatchCount = 0
            CurrentPullcoreStraightenCadIndex = 0
            PullCoreDimensionsReportPath = ""
            PullcoreIdOdHeightAxisUsed = ""

            DoEvents
            WaitMilliseconds 500
            DoEvents

        End If

    Next i

    Dim summary As String
    summary = BuildBatchSummary(completed, failed)

    LogLine "========================================"
    LogLine "BATCH DONE"
    LogLine Replace(summary, vbCrLf, " | ")
    LogLine "========================================"

    CloseAllDocumentsSafely

    On Error Resume Next
    If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo ErrHandler

    MsgBox summary, IIf(failed.count > 0, vbExclamation, vbInformation)

NormalEnd:
    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0
    Exit Sub

ErrHandler:
    LogLine "FATAL BATCH ERROR"
    LogLine "Step: " & CurrentStepName
    LogLine "Err " & Err.Number & ": " & Err.Description

    On Error Resume Next
    RestoreMainViewportGraphics
    CloseAllDocumentsSafely
    If Not swApp Is Nothing And Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    On Error GoTo 0

    MsgBox "Macro batch error at step: " & CurrentStepName & vbCrLf & _
           Err.Description & vbCrLf & RunLogPath, vbCritical
End Sub

Private Function ProcessOneJob(ByVal jobSearchText As String) As Boolean
On Error GoTo ErrHandler

    ProcessOneJob = False
    LastJobFailReason = ""

    Erase parts
    Erase BomRows
    Erase ExportRows
    Erase PullcoreMatches

    PartCount = 0
    BomCount = 0
    ExportCount = 0
    PullcoreMatchCount = 0
    CurrentPullcoreStraightenCadIndex = 0
    PullCoreDimensionsReportPath = ""
    PullcoreIdOdHeightAxisUsed = ""

    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
    MainViewportGraphicsDisabled = False

    Set swModel = Nothing
    Set swAssy = Nothing
    Set SpecialBomCadMatches = CreateObject("Scripting.Dictionary")
    Set SpecialBomCadQuoteNames = CreateObject("Scripting.Dictionary")
    Set PullcoreBestFitDimCache = CreateObject("Scripting.Dictionary")

    CurrentJobNumber = UCase(Trim(jobSearchText))
    CurrentJobFolder = ""
    NetworkJobFolder = ""
    LocalJobFolder = ""
    BaseNativeAssemblyPath = ""

    CustomerNumber = ""
    DateCode = ""
    NamingSourceText = ""

    ReDim parts(1 To 1)
    ReDim BomRows(1 To 1)
    ReDim ExportRows(1 To 1)
    ReDim PullcoreMatches(1 To 1)

    LogStart "Find network job folder"

    NetworkJobFolder = FindJobFolderByText(ROOT_JOB_PATH, CurrentJobNumber)
    LogLine "Network job folder result: " & NetworkJobFolder

    If NetworkJobFolder = "" Then
        LogErrorText "Could not find job folder for: " & CurrentJobNumber
        GoTo CleanExit
    End If

    LogDone "Find network job folder"

    LogStart "Prepare local job workspace"

    PrepareLocalJobWorkspace NetworkJobFolder, CurrentJobNumber, LocalJobFolder

    If LocalJobFolder = "" Then
        LogErrorText "Could not create local workspace for: " & CurrentJobNumber
        GoTo CleanExit
    End If

    CurrentJobFolder = LocalJobFolder
    RunLogPath = CurrentJobFolder & "\CMS_XT_Export_Log.txt"

    PullCoreDimensionsReportPath = CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE

    LogLine "========================================"
    LogLine "CMS XT export started"
    LogLine "Job/search text: " & CurrentJobNumber
    LogLine "Local job folder: " & CurrentJobFolder
    LogLine "Pull Core Dimensions report path: " & PullCoreDimensionsReportPath
    LogLine "========================================"

    LogDone "Prepare local job workspace"

    Dim extractFolder As String
    extractFolder = CurrentJobFolder & "\" & EXTRACT_FOLDER_NAME

    LogStart "Extract ZIP files"
    EnsureFolderDeep extractFolder
    ExtractAllZipFilesInJobFolder CurrentJobFolder, extractFolder
    FlattenExtractedZipContentsIntoJobFolder CurrentJobFolder, extractFolder

    If DELETE_EXTRACTED_ZIP_AFTER_FLATTEN Then
        LogLine "Deleting " & EXTRACT_FOLDER_NAME & " folder after flattening."
        DeleteFolderSafe extractFolder
    End If

    LogDone "Extract ZIP files"

    LogStart "Determine output naming info"
    DetermineOutputNamingInfo CurrentJobFolder
    LogLine "Customer/program number: " & CustomerNumber
    LogLine "Date code: " & DateCode
    LogLine "Naming source: " & NamingSourceText
    LogDone "Determine output naming info"

    LogStart "Find CAD file"

    If USE_ACTIVE_SOLIDWORKS_DOC_FIRST Then
        Set swModel = swApp.activeDoc
    End If

    Dim cadPath As String

    If swModel Is Nothing Then

        Dim cadCandidates As Collection
        Set cadCandidates = FindAllCadModelsRanked(CurrentJobFolder)
        AppendCadCandidates cadCandidates, FindAllCadModelsRanked(extractFolder)

        If cadCandidates.count = 0 Then
            LogErrorText "No CAD file found."
            GoTo CleanExit
        End If

        LogLine "CAD candidates found: " & cadCandidates.count

        Dim ci As Long
        For ci = 1 To cadCandidates.count
            LogLine "  candidate " & ci & ": " & CStr(cadCandidates(ci))
        Next ci

        LogDone "Find CAD file"
        LogStart "Open CAD"

        For ci = 1 To cadCandidates.count
            cadPath = CStr(cadCandidates(ci))
            LogLine "Trying to open CAD candidate " & ci & "/" & cadCandidates.count & ": " & cadPath
            Set swModel = OpenCadFile(cadPath)

            If Not swModel Is Nothing Then
                LogLine "CAD opened successfully: " & cadPath
                Exit For
            End If

            LogLine "Candidate failed to open, trying next."
        Next ci

        If swModel Is Nothing Then
            LogErrorText "Open CAD failed; no CAD candidate opened."
            GoTo CleanExit
        End If

        MainCadOpenedByMacro = True
        MainCadTitleForClose = swModel.GetTitle

        LogLine "Opened CAD title: " & swModel.GetTitle
        LogDone "Open CAD"

    Else

        MainCadOpenedByMacro = False
        MainCadTitleForClose = swModel.GetTitle

        LogLine "Using active document: " & swModel.GetTitle
        LogDone "Find CAD file"

    End If

    Dim errs As Long
    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden

    MainCadTitleForClose = swModel.GetTitle

    swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    EnsureSwHidden

    DisableMainViewportGraphics

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        LogStart "Physically hide .250 items"
        HideQuarterInchThicknessItems swModel
        LogDone "Physically hide .250 items"
    ElseIf HIDE_QUARTER_INCH_THICKNESS Then
        LogLine ".250 physical hide skipped for speed. Filtering by BOM/CAD logic."
    End If

    StabilizeActiveView swModel, 300

    LogStart "Scan CAD"

    ScanActiveSolidWorksDocument
    SortPartsByVolumeDescending

    LogLine "CAD PartCount=" & PartCount
    LogDone "Scan CAD"

    If PartCount = 0 Then
        LogErrorText "No measurable CAD parts found."
        GoTo CleanExit
    End If

    WritePartDimensionCsv CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv"

    LogStart "Find BOM"

    Dim bomPath As String
    bomPath = FindCustomerBomFile(CurrentJobFolder)

    If bomPath = "" Then
        LogErrorText "No BOM found."
        GoTo CleanExit
    End If

    LogLine "BOM selected: " & bomPath
    LogDone "Find BOM"

    If LCase(GetFileExtension(bomPath)) = "pdf" Then

        If READ_PDF_BOM_WITH_PDFTOTEXT Then
            LogStart "Read PDF BOM with Poppler"
            ReadCustomerBomPdfUsingPdfToText bomPath
            LogDone "Read PDF BOM with Poppler"
        Else
            LogErrorText "PDF BOM found but PDF reading disabled."
            GoTo CleanExit
        End If

    Else

        LogStart "Read Excel BOM TURBO"
        ReadCustomerBom bomPath
        LogDone "Read Excel BOM TURBO"

    End If

    LogLine "BomCount=" & BomCount

    If BomCount = 0 Then
        LogErrorText "No usable BOM rows found."
        GoTo CleanExit
    End If

    LogStart "Match BOM to CAD"

    BuildExportRowsFromBom

    If ADD_TOP_BOT_INS_FROM_CAD_GEOMETRY Then
        LogStart "Add missing TOP/BOT INS from CAD geometry"
        AddMissingTopBotInsFromCadGeometry
        LogDone "Add missing TOP/BOT INS from CAD geometry"
    End If

    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    SaveImportedBaseAssemblyToLocal swModel
    MainCadTitleForClose = swModel.GetTitle
    LogDone "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    WriteExportCheckCsv CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv"

    If CREATE_PULLCORE_DIMENSIONS_EXCEL Then
        LogStart "Write Pull Core Dimensions Excel"
        WritePullCoreDimensionsExcel PullCoreDimensionsReportPath
        LogDone "Write Pull Core Dimensions Excel"
    End If

    LogLine "ExportCount=" & ExportCount
    LogLine "PullcoreMatchCount=" & PullcoreMatchCount
    LogDone "Match BOM to CAD"

    If ExportCount = 0 And PullcoreMatchCount = 0 Then
        LogErrorText "No matched export rows."
        GoTo CleanExit
    End If

    Dim outputFolder As String
    outputFolder = CurrentJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX

    EnsureFolderDeep outputFolder

    Set ExportFilePaths = CreateObject("Scripting.Dictionary")

    LogStart "Export XT/DXF files"
    ExportMatchedPartsAsXt outputFolder
    LogDone "Export XT/DXF files"

    If CREATE_INDIVIDUAL_DXFS Then
        LogStart "Export individual four-view DXFs"
        ExportIndividualHolderAndClampingDxfs outputFolder
        LogDone "Export individual four-view DXFs"
    End If

    LogStart "Restore visibility/state"

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ShowAllPartBodies swModel

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        HideQuarterInchThicknessItems swModel
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    LogDone "Restore visibility/state"

    CollectJobPdfsAndReports

    If PUSH_OUTPUTS_TO_NETWORK Then
        LogStart "Push finished outputs back to network"
        PushFinishedOutputsBackToNetwork
        LogDone "Push finished outputs back to network"
    Else
        LogLine "Network push DISABLED. Outputs kept locally only at:"
        LogLine "  " & CurrentJobFolder
    End If

    LogLine "========================================"
    LogLine "DONE JOB. Total seconds=" & DateDiff("s", MacroStartTime, Now)
    LogLine "Local output folder: " & outputFolder
    LogLine "Pull Core Dimensions report: " & PullCoreDimensionsReportPath
    LogLine "========================================"

    ProcessOneJob = True

CleanExit:
    On Error Resume Next

    RestoreMainViewportGraphics

    If Not swApp Is Nothing Then
        If Not RUN_SOLIDWORKS_INVISIBLE Then swApp.Visible = True
    End If

    CloseCurrentJobCadIfNeeded
    CloseAllDocumentsSafely

    Set swAssy = Nothing
    Set swModel = Nothing
    Set ExportFilePaths = Nothing
    Set SpecialBomCadMatches = Nothing
    Set SpecialBomCadQuoteNames = Nothing
    Set PullcoreBestFitDimCache = Nothing

    CurrentPullcoreStraightenCadIndex = 0
    PullcoreIdOdHeightAxisUsed = ""

    Exit Function

ErrHandler:
    LogLine "FATAL JOB ERROR"
    LogLine "Job: " & CurrentJobNumber
    LogLine "Step: " & CurrentStepName
    LogLine "Err " & Err.Number & ": " & Err.Description
    LastJobFailReason = "Step '" & CurrentStepName & "' - Err " & Err.Number & ": " & Err.Description
    ProcessOneJob = False
    Resume CleanExit
End Function

Private Function ParseJobInputList(ByVal inputText As String) As Collection
On Error GoTo ErrHandler

    Dim result As New Collection

    inputText = Replace(inputText, vbCr, " ")
    inputText = Replace(inputText, vbLf, " ")
    inputText = Replace(inputText, vbTab, " ")
    inputText = Replace(inputText, ",", " ")
    inputText = Replace(inputText, ";", " ")

    Do While InStr(inputText, "  ") > 0
        inputText = Replace(inputText, "  ", " ")
    Loop

    inputText = Trim(inputText)

    If inputText = "" Then
        Set ParseJobInputList = result
        Exit Function
    End If

    Dim arr() As String
    arr = Split(inputText, " ")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim token As String

    For i = LBound(arr) To UBound(arr)
        token = UCase(Trim(CStr(arr(i))))

        If token <> "" Then
            If Not dict.Exists(token) Then
                dict(token) = True
                result.Add token
            End If
        End If
    Next i

    Set ParseJobInputList = result
    Exit Function

ErrHandler:
    Set ParseJobInputList = New Collection
End Function

Private Function BuildBatchSummary(ByVal completed As Collection, ByVal failed As Collection) As String
On Error GoTo ErrHandler

    Dim s As String
    Dim i As Long

    s = "Batch complete." & vbCrLf & vbCrLf
    s = s & "Completed: " & completed.count & vbCrLf

    If completed.count > 0 Then
        For i = 1 To completed.count
            s = s & "  - " & CStr(completed(i)) & vbCrLf
        Next i
    End If

    s = s & vbCrLf & "Failed: " & failed.count & vbCrLf

    If failed.count > 0 Then
        For i = 1 To failed.count
            s = s & "  - " & CStr(failed(i)) & vbCrLf
        Next i
    End If

    BuildBatchSummary = s
    Exit Function

ErrHandler:
    BuildBatchSummary = "Batch complete."
End Function

Private Sub EnsureSwHidden()
On Error Resume Next

    If RUN_SOLIDWORKS_INVISIBLE Then
        If Not swApp Is Nothing Then
            If swApp.Visible Then swApp.Visible = False
        End If
    End If
End Sub

Private Sub CloseCurrentJobCadIfNeeded()
On Error Resume Next

    If MainCadOpenedByMacro Then
        If MainCadTitleForClose <> "" Then
            LogLine "Closing main CAD document for finished job: " & MainCadTitleForClose
            swApp.CloseDoc MainCadTitleForClose
        End If
    Else
        If MainCadTitleForClose <> "" Then
            LogLine "Main CAD was not opened by macro. Leaving open: " & MainCadTitleForClose
        End If
    End If

    MainCadOpenedByMacro = False
    MainCadTitleForClose = ""
End Sub

Private Sub CloseAllDocumentsSafely()
On Error Resume Next

    If swApp Is Nothing Then Exit Sub

    swApp.CloseAllDocuments True

    Dim swDoc As Object
    Dim nextDoc As Object

    Set swDoc = swApp.GetFirstDocument

    Do While Not swDoc Is Nothing
        Set nextDoc = swDoc.GetNext
        swApp.CloseDoc swDoc.GetTitle
        Set swDoc = nextDoc
    Loop
End Sub

Private Sub DeleteFolderSafe(ByVal folderPath As String)
On Error Resume Next

    If folderPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        fso.DeleteFolder folderPath, True
        LogLine "Deleted folder: " & folderPath
    End If
End Sub

' ============================================================
' LOCAL STAGING / NETWORK PUSH-BACK
' ============================================================

Private Sub PrepareLocalJobWorkspace(ByVal sourceNetworkFolder As String, _
                                     ByVal jobNumber As String, _
                                     ByRef localFolderOut As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    EnsureFolderDeep LOCAL_WORKSPACE_ROOT

    localFolderOut = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(jobNumber)

    If fso.FolderExists(localFolderOut) Then
        LogLine "Clearing old local workspace: " & localFolderOut
        On Error Resume Next
        fso.DeleteFolder localFolderOut, True
        On Error GoTo ErrHandler
    End If

    EnsureFolderDeep localFolderOut

    If fso.FolderExists(localFolderOut) = False Then
        LogLine "Could not create local workspace folder: " & localFolderOut
        localFolderOut = ""
        Exit Sub
    End If

    LogLine "Copying network job folder to local workspace."
    LogLine "  FROM: " & sourceNetworkFolder
    LogLine "  TO  : " & localFolderOut

    If CopyFolderWithRobocopy(sourceNetworkFolder, localFolderOut) Then
        LogLine "Local copy complete (robocopy)."
    Else
        LogLine "Robocopy unavailable/failed. Using VBA copy fallback."
        CopyFolderContentsFiltered sourceNetworkFolder, localFolderOut
        LogLine "Local copy complete (VBA fallback)."
    End If

    Exit Sub

ErrHandler:
    LogLine "PrepareLocalJobWorkspace error: " & Err.Description

    On Error Resume Next
    If fso Is Nothing Then
        localFolderOut = ""
    ElseIf fso.FolderExists(localFolderOut) = False Then
        localFolderOut = ""
    End If
End Sub

Private Function CopyFolderWithRobocopy(ByVal src As String, ByVal dst As String) As Boolean
On Error GoTo ErrHandler

    CopyFolderWithRobocopy = False

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cleanSrc As String
    cleanSrc = src

    Do While Right(cleanSrc, 1) = "\"
        cleanSrc = Left(cleanSrc, Len(cleanSrc) - 1)
    Loop

    Dim excludeExtract As String
    excludeExtract = cleanSrc & "\" & EXTRACT_FOLDER_NAME

    Dim cmd As String
    cmd = "cmd /c robocopy " & Chr(34) & cleanSrc & Chr(34) & " " & Chr(34) & dst & Chr(34) & _
          " /MIR /XD " & Chr(34) & excludeExtract & Chr(34) & _
          " /R:1 /W:1 /MT:16 /NFL /NDL /NJH /NJS /NP"

    Dim rc As Long
    rc = sh.Run(cmd, 0, True)

    LogLine "robocopy exit code: " & rc

    If rc < 8 Then CopyFolderWithRobocopy = True
    Exit Function

ErrHandler:
    LogLine "CopyFolderWithRobocopy error: " & Err.Description
    CopyFolderWithRobocopy = False
End Function

Private Sub CopyFolderContentsFiltered(ByVal sourceFolder As String, ByVal destFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(sourceFolder) Then Exit Sub

    EnsureFolderDeep destFolder

    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)

    Dim file As Object
    For Each file In src.Files
        fso.CopyFile file.path, destFolder & "\" & file.Name, True
    Next file

    Dim subFolder As Object
    Dim folderNameUpper As String

    For Each subFolder In src.SubFolders

        folderNameUpper = UCase(subFolder.Name)

        If folderNameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextFolder
        If InStr(folderNameUpper, " PRINTS") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "J BLOCK") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PULLCORE") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PULL CORE") > 0 Then GoTo NextFolder
        If InStr(folderNameUpper, "PYROPEL") > 0 Then GoTo NextFolder
        If folderNameUpper = "BASE" Then GoTo NextFolder
        If InStr(folderNameUpper, " BASE") > 0 Then GoTo NextFolder

        EnsureFolderDeep destFolder & "\" & subFolder.Name
        CopyFolderContentsFiltered subFolder.path, destFolder & "\" & subFolder.Name

NextFolder:
    Next subFolder

    Exit Sub

ErrHandler:
    LogLine "CopyFolderContentsFiltered error: " & Err.Description
End Sub

Private Sub PushFinishedOutputsBackToNetwork()
On Error GoTo ErrHandler

    If PUSH_OUTPUTS_TO_NETWORK = False Then
        LogLine "PushFinishedOutputsBackToNetwork called but push is disabled. Skipping."
        Exit Sub
    End If

    If NetworkJobFolder = "" Then
        LogLine "Push back skipped: NetworkJobFolder is blank."
        Exit Sub
    End If

    If CurrentJobFolder = "" Then
        LogLine "Push back skipped: CurrentJobFolder is blank."
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(CurrentJobFolder) = False Then
        LogLine "Push back skipped: local CurrentJobFolder does not exist."
        Exit Sub
    End If

    EnsureFolderDeep NetworkJobFolder

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX, _
                       NetworkJobFolder & "\" & CurrentJobNumber & OUTPUT_FOLDER_SUFFIX

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME

    CopyFolderIfExists CurrentJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME, _
                       NetworkJobFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME

    CopyFolderIfExists LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber) & "\" & CurrentJobNumber & " Base", _
                       NetworkJobFolder & "\" & CurrentJobNumber & " Base"

    CopyRootOutputFilesToNetwork CurrentJobFolder, NetworkJobFolder

    CopyFileIfExists CurrentJobFolder & "\XT_Export_CAD_Dimensions.csv", _
                     NetworkJobFolder & "\XT_Export_CAD_Dimensions.csv"

    CopyFileIfExists CurrentJobFolder & "\XT_Export_BOM_Match_Report.csv", _
                     NetworkJobFolder & "\XT_Export_BOM_Match_Report.csv"

    CopyFileIfExists CurrentJobFolder & "\XT_Export_BOM_PDF_Text.txt", _
                     NetworkJobFolder & "\XT_Export_BOM_PDF_Text.txt"

    CopyFileIfExists CurrentJobFolder & "\CMS_XT_Export_Log.txt", _
                     NetworkJobFolder & "\CMS_XT_Export_Log.txt"

    CopyFileIfExists CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE, _
                     NetworkJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE

    LogLine "Push back to network complete."
    Exit Sub

ErrHandler:
    LogLine "PushFinishedOutputsBackToNetwork error: " & Err.Description
End Sub

Private Sub CopyFolderIfExists(ByVal sourceFolder As String, ByVal destFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sourceFolder) = False Then
        LogLine "CopyFolderIfExists skipped missing folder: " & sourceFolder
        Exit Sub
    End If

    If fso.FolderExists(destFolder) Then
        LogLine "Deleting old network output folder: " & destFolder
        fso.DeleteFolder destFolder, True
    End If

    EnsureFolderDeep destFolder
    CopyFolderContents sourceFolder, destFolder

    LogLine "Copied folder to network:"
    LogLine "  FROM: " & sourceFolder
    LogLine "  TO  : " & destFolder
    Exit Sub

ErrHandler:
    LogLine "CopyFolderIfExists error: " & Err.Description
End Sub

Private Sub CopyFileIfExists(ByVal sourceFile As String, ByVal destFile As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFile) = False Then Exit Sub

    EnsureFolderDeep fso.GetParentFolderName(destFile)
    fso.CopyFile sourceFile, destFile, True
    LogLine "Copied file to network: " & destFile
    Exit Sub

ErrHandler:
    LogLine "CopyFileIfExists error: " & Err.Description
End Sub

Private Sub CopyRootOutputFilesToNetwork(ByVal localRoot As String, ByVal networkRoot As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(localRoot) = False Then Exit Sub

    Dim folder As Object
    Set folder = fso.GetFolder(localRoot)

    Dim file As Object
    Dim ext As String
    Dim nameUpper As String

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        nameUpper = UCase(file.Name)

        If Left(nameUpper, Len(UCase(CurrentJobNumber))) = UCase(CurrentJobNumber) Then
            Select Case ext
                Case "x_t", "igs", "easm", "dxf", "sldasm"
                    fso.CopyFile file.path, networkRoot & "\" & file.Name, True
                    LogLine "Copied root output file to network: " & file.Name
            End Select
        End If

    Next file

    Exit Sub

ErrHandler:
    LogLine "CopyRootOutputFilesToNetwork error: " & Err.Description
End Sub

' ============================================================
' OPEN CAD / BASE SAVE / GRAPHICS / ORIENTATION
' ============================================================

Private Function OpenCadFile(ByVal cadPath As String) As Object
On Error GoTo ErrHandler

    LogLine "OpenCadFile: " & cadPath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(cadPath) Then
        LogLine "OpenCadFile: file does not exist."
        Set OpenCadFile = Nothing
        Exit Function
    End If

    LogLine "CAD size bytes: " & fso.GetFile(cadPath).Size

    Dim ext As String
    ext = LCase(fso.GetExtensionName(cadPath))

    Dim errs As Long
    Dim warns As Long
    Dim importErrors As Long
    Dim m As Object

    Set m = Nothing

    If ext = "sldasm" Then

        errs = 0
        warns = 0
        Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
        LogLine "OpenDoc6(ASSEMBLY) errs=" & errs & " warns=" & warns & " ok=" & CStr(Not m Is Nothing)

    ElseIf ext = "sldprt" Then

        errs = 0
        warns = 0
        Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
        LogLine "OpenDoc6(PART) errs=" & errs & " warns=" & warns & " ok=" & CStr(Not m Is Nothing)

    ElseIf ext = "slddrw" Then

        LogLine "OpenCadFile: drawing file skipped."
        Set m = Nothing

    Else

        importErrors = 0
        Set m = swApp.LoadFile4(cadPath, "r", Nothing, importErrors)
        LogLine "LoadFile4 importErrors=" & importErrors & " ok=" & CStr(Not m Is Nothing)

        If m Is Nothing Then
            importErrors = 0
            Set m = swApp.LoadFile4(cadPath, "", Nothing, importErrors)
            LogLine "LoadFile4(no opts) importErrors=" & importErrors & " ok=" & CStr(Not m Is Nothing)
        End If

        If m Is Nothing Then
            errs = 0
            warns = 0
            Set m = swApp.OpenDoc6(cadPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
            LogLine "Fallback OpenDoc6(ASSEMBLY) errs=" & errs & " ok=" & CStr(Not m Is Nothing)
        End If

        If m Is Nothing Then
            errs = 0
            warns = 0
            Set m = swApp.OpenDoc6(cadPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
            LogLine "Fallback OpenDoc6(PART) errs=" & errs & " ok=" & CStr(Not m Is Nothing)
        End If

    End If

    If m Is Nothing Then
        LogLine "OpenCadFile: ALL open methods failed for ext '" & ext & "'."
    End If

    Set OpenCadFile = m
    Exit Function

ErrHandler:
    LogLine "OpenCadFile error: " & Err.Description
    Set OpenCadFile = Nothing
End Function

Private Sub SaveImportedBaseAssemblyToLocal(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub

    If model.GetType <> swDocASSEMBLY Then
        LogLine "SaveImportedBaseAssemblyToLocal skipped: model is not an assembly."
        Exit Sub
    End If

    Dim baseFolder As String
    baseFolder = LOCAL_WORKSPACE_ROOT & "\" & CleanFileName(CurrentJobNumber) & "\" & CurrentJobNumber & " Base"

    EnsureFolderDeep baseFolder

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    BaseNativeAssemblyPath = baseFolder & "\" & _
                             CurrentJobNumber & "_BASE_" & _
                             custToken & "_" & dateToken & ".sldasm"

    LogLine "Saving imported BASE assembly locally as native SLDASM:"
    LogLine "  " & BaseNativeAssemblyPath

    On Error Resume Next
    model.ResolveAllLightWeightComponents True
    On Error GoTo ErrHandler

    UnsuppressAllAssemblyComponents model
    ShowAllAssemblyComponents model

    ApplyCmsTopView model
    StabilizeActiveView model, 50

    Dim errs As Long
    Dim warns As Long

    model.Extension.SaveAs3 BaseNativeAssemblyPath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent, _
                            Nothing, Nothing, errs, warns

    LogLine "BASE native SLDASM save done. Errors=" & errs & " Warnings=" & warns

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(BaseNativeAssemblyPath) = False Then
        LogLine "WARNING: BASE native SLDASM was not created."
    Else
        LogLine "BASE native SLDASM created successfully."
    End If

    Exit Sub

ErrHandler:
    LogLine "SaveImportedBaseAssemblyToLocal error: " & Err.Description
End Sub

Private Sub DisableMainViewportGraphics()
On Error Resume Next

    If DISABLE_MAIN_VIEWPORT_GRAPHICS = False Then Exit Sub
    If swModel Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = False
        MainViewportGraphicsDisabled = True
        LogLine "Main 3D viewport graphics updates disabled."
    End If
End Sub

Private Sub RestoreMainViewportGraphics()
On Error Resume Next

    If swModel Is Nothing Then Exit Sub

    Dim swView As Object
    Set swView = swModel.ActiveView

    If Not swView Is Nothing Then
        swView.EnableGraphicsUpdate = True
    End If

    If MainViewportGraphicsDisabled Then
        LogLine "Main 3D viewport graphics updates restored."
    End If

    MainViewportGraphicsDisabled = False
End Sub

Private Sub SetCmsTopOrientation(ByVal model As Object, Optional ByVal persistAsStandardTop As Boolean = False)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    If PROMPT_FOR_TOP_ORIENTATION Then

        MsgBox "Rotate model so you are looking top-down at the TCP / TOP SMED plate, then click OK.", vbInformation

    Else

        Dim autoOriented As Boolean
        autoOriented = False

        If AUTO_SELECT_TCP_TOP_ORIENTATION Then
            autoOriented = TryShowTcpTopViewFromComponentCenters(model)
        End If

        If autoOriented = False Then
            model.ShowNamedView2 CMS_BASE_TOP_VIEW_NAME, CMS_BASE_TOP_VIEW_ID

            LogLine "TCP-top auto orientation unavailable. Fallback view used: " & _
                    CMS_BASE_TOP_VIEW_NAME

            RotateViewZSteps model, CMS_TOP_ROTATE_Z_STEPS
        End If

    End If

    model.DeleteNamedView CMS_TOP_VIEW_NAME
    model.NameView CMS_TOP_VIEW_NAME

    Dim persisted As Boolean
    persisted = False

    If persistAsStandardTop Then

        persisted = PersistCurrentViewAsStandardTop(model)

        If persisted Then
            model.ShowNamedView2 "*Top", 5
            model.DeleteNamedView CMS_TOP_VIEW_NAME
            model.NameView CMS_TOP_VIEW_NAME
            LogLine "CMS_TOP rebuilt from persisted SolidWorks *Top view."
        Else
            model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
            LogLine "WARNING: Could not persist standard views; CMS_TOP named view was still saved."
        End If

    End If

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    StabilizeActiveView model, 300
End Sub

Private Function TryShowTcpTopViewFromComponentCenters(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryShowTcpTopViewFromComponentCenters = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    If TryOrientTcpUpByViewProjection(model) Then
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
    End If

    LogLine "TCP-top auto orientation: component-name method unavailable."
    Exit Function

ErrHandler:
    LogLine "TryShowTcpTopViewFromComponentCenters error: " & Err.Description
    TryShowTcpTopViewFromComponentCenters = False
End Function

Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryOrientTcpUpByViewProjection = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim tcpComp As Object
    Dim bcpComp As Object

    Set tcpComp = FindComponentByKeys(model, TCP_TOP_ORIENTATION_KEYS)
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpComp Is Nothing Or bcpComp Is Nothing Then
        LogLine "World-axis orientation skipped: TCP or BCP component not found."
        Exit Function
    End If

    Dim tcpX As Double, tcpY As Double, tcpZ As Double
    Dim bcpX As Double, bcpY As Double, bcpZ As Double

    If TryGetComponentCenterInches(tcpComp, tcpX, tcpY, tcpZ) = False Then
        LogLine "World-axis orientation skipped: could not read TCP center."
        Exit Function
    End If

    If TryGetComponentCenterInches(bcpComp, bcpX, bcpY, bcpZ) = False Then
        LogLine "World-axis orientation skipped: could not read BCP center."
        Exit Function
    End If

    TryOrientTcpUpByViewProjection = OrientTcpTopFromCenters(model, _
        tcpX, tcpY, tcpZ, bcpX, bcpY, bcpZ, _
        "World-axis raw components")

    Exit Function

ErrHandler:
    LogLine "TryOrientTcpUpByViewProjection error: " & Err.Description
    TryOrientTcpUpByViewProjection = False
End Function

Private Function OrientTcpTopFromCenters(ByVal model As Object, _
                                         ByVal tcpX As Double, _
                                         ByVal tcpY As Double, _
                                         ByVal tcpZ As Double, _
                                         ByVal bcpX As Double, _
                                         ByVal bcpY As Double, _
                                         ByVal bcpZ As Double, _
                                         ByVal sourceLabel As String) As Boolean
On Error GoTo ErrHandler

    OrientTcpTopFromCenters = False

    If model Is Nothing Then Exit Function

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = tcpX - bcpX
    dy = tcpY - bcpY
    dz = tcpZ - bcpZ

    Dim aX As Double
    Dim aY As Double
    Dim aZ As Double

    aX = Abs(dx)
    aY = Abs(dy)
    aZ = Abs(dz)

    LogLine sourceLabel & " TCP-BCP separation: " & _
            "X=" & FormatNumberForCsv(dx) & _
            " Y=" & FormatNumberForCsv(dy) & _
            " Z=" & FormatNumberForCsv(dz)

    Dim axisName As String
    Dim tcpHigh As Boolean

    If aY >= aX And aY >= aZ Then
        axisName = "Y"
        tcpHigh = (dy >= 0#)
    ElseIf aZ >= aX And aZ >= aY Then
        axisName = "Z"
        tcpHigh = (dz >= 0#)
    Else
        axisName = "X"
        tcpHigh = (dx >= 0#)
    End If

    Dim viewName As String
    Dim viewId As Long

    ' Correct SolidWorks default standard-view mapping:
    '   Y stack -> *Top / *Bottom
    '   Z stack -> *Front / *Back
    '   X stack -> *Right / *Left

    Select Case axisName

        Case "Y"
            If tcpHigh Then
                viewName = "*Top"
                viewId = 5
            Else
                viewName = "*Bottom"
                viewId = 6
            End If

        Case "Z"
            If tcpHigh Then
                viewName = "*Front"
                viewId = 1
            Else
                viewName = "*Back"
                viewId = 2
            End If

        Case "X"
            If tcpHigh Then
                viewName = "*Right"
                viewId = 4
            Else
                viewName = "*Left"
                viewId = 3
            End If

    End Select

    If viewName = "" Then
        LogLine sourceLabel & " orientation failed: no standard view selected."
        Exit Function
    End If

    model.ShowNamedView2 viewName, viewId

    LogLine sourceLabel & " selected " & viewName & _
            ". Stack axis=" & axisName & _
            ", TCP at high end=" & CStr(tcpHigh)

    RotateViewZSteps model, CMS_TOP_ROTATE_Z_STEPS
    StabilizeActiveView model, 100

    OrientTcpTopFromCenters = True
    Exit Function

ErrHandler:
    LogLine "OrientTcpTopFromCenters error: " & Err.Description
    OrientTcpTopFromCenters = False
End Function

Private Function TryOrientFromMatchedQuotePair(ByVal model As Object, _
                                               ByVal topQuoteName As String, _
                                               ByVal bottomQuoteName As String, _
                                               ByVal sourceLabel As String) As Boolean
On Error GoTo ErrHandler

    TryOrientFromMatchedQuotePair = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    Dim topIdx As Long
    Dim botIdx As Long

    topIdx = FindCadIndexFromExportQuote(topQuoteName)
    botIdx = FindCadIndexFromExportQuote(bottomQuoteName)

    If topIdx <= 0 Or botIdx <= 0 Then
        LogLine sourceLabel & " orientation skipped: missing matched quote pair. " & _
                "TopQuote='" & topQuoteName & "' idx=" & CStr(topIdx) & _
                ", BottomQuote='" & bottomQuoteName & "' idx=" & CStr(botIdx)
        Exit Function
    End If

    If topIdx > PartCount Or botIdx > PartCount Then Exit Function

    If parts(topIdx).hasAsmCenter = False Or parts(botIdx).hasAsmCenter = False Then
        LogLine sourceLabel & " orientation skipped: assembly centers unavailable. " & _
                "TopQuote='" & topQuoteName & "', BottomQuote='" & bottomQuoteName & "'"
        Exit Function
    End If

    LogLine sourceLabel & " orientation source:"
    LogLine "  TOP SIDE quote '" & topQuoteName & "' -> CAD '" & _
            parts(topIdx).componentName & "' center X/Y/Z=" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterX) & "/" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterY) & "/" & _
            FormatNumberForCsv(parts(topIdx).AsmCenterZ)

    LogLine "  BOTTOM SIDE quote '" & bottomQuoteName & "' -> CAD '" & _
            parts(botIdx).componentName & "' center X/Y/Z=" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterX) & "/" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterY) & "/" & _
            FormatNumberForCsv(parts(botIdx).AsmCenterZ)

    TryOrientFromMatchedQuotePair = OrientTcpTopFromCenters(model, _
        parts(topIdx).AsmCenterX, parts(topIdx).AsmCenterY, parts(topIdx).AsmCenterZ, _
        parts(botIdx).AsmCenterX, parts(botIdx).AsmCenterY, parts(botIdx).AsmCenterZ, _
        sourceLabel)

    Exit Function

ErrHandler:
    LogLine "TryOrientFromMatchedQuotePair error (" & sourceLabel & "): " & Err.Description
    TryOrientFromMatchedQuotePair = False
End Function

Private Sub EnsureCmsTopOrientationFromMatchedTcpBcp(ByVal model As Object, _
                                                     Optional ByVal persistAsStandardTop As Boolean = False)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    Dim oriented As Boolean
    oriented = False

    ' More reliable than TCP/BCP when imported parts have generic names:
    ' try top-side/bottom-side holder/pot/insert pairs first.

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "ID HOLDER", _
                    "OD HOLDER", _
                    "Matched TOP/BOTTOM HOLDER")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "ID POT BLOCK", _
                    "OD POT BLOCK", _
                    "Matched TOP/BOTTOM POT")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "TOP INS", _
                    "BOT INS", _
                    "Matched TOP/BOTTOM INS")
    End If

    If oriented = False Then
        oriented = TryOrientFromMatchedQuotePair(model, _
                    "TCP", _
                    "BCP", _
                    "Matched TCP/BCP")
    End If

    If oriented = False Then
        LogLine "Matched top-side orientation failed from holder/pot/ins/TCP pairs."
        LogLine "Falling back to SetCmsTopOrientation."
        SetCmsTopOrientation model, persistAsStandardTop
        Exit Sub
    End If

    On Error Resume Next
    model.DeleteNamedView CMS_TOP_VIEW_NAME
    Err.Clear
    model.NameView CMS_TOP_VIEW_NAME
    On Error GoTo ErrHandler

    LogLine "CMS_TOP named view saved from matched top-side orientation."

    If persistAsStandardTop Then

        If PersistCurrentViewAsStandardTop(model) Then

            model.ShowNamedView2 "*Top", 5

            On Error Resume Next
            model.DeleteNamedView CMS_TOP_VIEW_NAME
            Err.Clear
            model.NameView CMS_TOP_VIEW_NAME
            On Error GoTo ErrHandler

            LogLine "Matched top-side orientation persisted as SolidWorks *Top."

        Else

            LogLine "WARNING: Matched top-side orientation could not be persisted as standard top."

        End If

    End If

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    StabilizeActiveView model, 100

    Exit Sub

ErrHandler:
    LogLine "EnsureCmsTopOrientationFromMatchedTcpBcp error: " & Err.Description
End Sub

Private Function PersistCurrentViewAsStandardTop(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    PersistCurrentViewAsStandardTop = False

    If model Is Nothing Then Exit Function

    Dim errs As Long
    swApp.ActivateDoc3 model.GetTitle, False, 0, errs
    EnsureSwHidden

    Dim ok As Boolean
    ok = False

    On Error Resume Next

    Err.Clear
    ok = CBool(model.Extension.UpdateStandardViews("*Top", 5))
    If Err.Number <> 0 Then
        Err.Clear
        ok = False
    End If

    If ok = False Then
        Err.Clear
        ok = CBool(model.UpdateStandardViews("*Top", 5))
        If Err.Number <> 0 Then
            Err.Clear
            ok = False
        End If
    End If

    On Error GoTo ErrHandler

    If ok Then
        PersistCurrentViewAsStandardTop = True
        LogLine "Standard views REDEFINED: current TCP/top-side orientation assigned to *Top (ViewId 5)."
    Else
        LogLine "WARNING: UpdateStandardViews(*Top,5) did not succeed; orientation still carried by CMS_TOP named view."
    End If

    On Error Resume Next
    model.ForceRebuild3 False
    model.ViewZoomtofit2
    On Error GoTo ErrHandler

    Exit Function

ErrHandler:
    LogLine "PersistCurrentViewAsStandardTop error: " & Err.Description
    PersistCurrentViewAsStandardTop = False
End Function

Private Sub ApplyCmsTopView(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
End Sub

Private Sub StabilizeActiveView(ByVal model As Object, Optional ByVal waitMs As Long = 200)
On Error Resume Next

    If model Is Nothing Then Exit Sub

    If DISABLE_STABILIZE_DELAYS Then
        waitMs = 0
    End If

    model.ViewZoomtofit2
    model.GraphicsRedraw2
    DoEvents

    If waitMs > 0 Then WaitMilliseconds waitMs

    model.GraphicsRedraw2
    DoEvents
End Sub

Private Sub RotateViewZSteps(ByVal model As Object, ByVal steps As Long)
On Error Resume Next

    Dim i As Long

    If model Is Nothing Then Exit Sub

    If steps < 0 Then
        For i = 1 To Abs(steps)
            model.ViewRotateminusz
        Next i
    Else
        For i = 1 To steps
            model.ViewRotateplusz
        Next i
    End If
End Sub

Private Function TryGetComponentViewY(ByVal model As Object, _
                                      ByVal swComp As Object, _
                                      ByRef screenY As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentViewY = False
    screenY = 0#

    If model Is Nothing Or swComp Is Nothing Then Exit Function

    Dim vBox As Variant
    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    cx = (CDbl(vBox(0)) + CDbl(vBox(3))) / 2#
    cy = (CDbl(vBox(1)) + CDbl(vBox(4))) / 2#
    cz = (CDbl(vBox(2)) + CDbl(vBox(5))) / 2#

    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then Exit Function

    Dim mView As Variant
    mView = swView.Orientation3.ArrayData

    If IsEmpty(mView) Then Exit Function
    If IsArray(mView) = False Then Exit Function
    If UBound(mView) < 8 Then Exit Function

    screenY = (cx * CDbl(mView(1))) + (cy * CDbl(mView(4))) + (cz * CDbl(mView(7)))

    TryGetComponentViewY = True
    Exit Function

ErrHandler:
    TryGetComponentViewY = False
End Function

Private Function FindComponentByKeys(ByVal assyModel As Object, ByVal pipeKeys As String) As Object
On Error GoTo ErrHandler

    Set FindComponentByKeys = Nothing

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function
    If IsArray(vComps) = False Then Exit Function

    Dim bestComp As Object
    Dim bestScore As Double
    bestScore = -1#

    Dim i As Long
    Dim swComp As Object
    Dim hay As String
    Dim score As Double

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)
        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                hay = swComp.Name2 & " " & swComp.GetPathName
                If ContainsAnyPipeKey(hay, pipeKeys) Then
                    score = Len(hay)
                    If score > bestScore Then
                        bestScore = score
                        Set bestComp = swComp
                    End If
                End If
            End If
        End If
    Next i

    Set FindComponentByKeys = bestComp
    Exit Function

ErrHandler:
    Set FindComponentByKeys = Nothing
End Function

Private Function TryFindBestComponentCenterByKeys(ByVal assyModel As Object, _
                                                  ByVal pipeKeys As String, _
                                                  ByRef cx As Double, _
                                                  ByRef cy As Double, _
                                                  ByRef cz As Double, _
                                                  ByRef foundName As String) As Boolean
On Error GoTo ErrHandler

    TryFindBestComponentCenterByKeys = False

    cx = 0#
    cy = 0#
    cz = 0#
    foundName = ""

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function
    If IsArray(vComps) = False Then Exit Function

    Dim bestScore As Double
    bestScore = -1#

    Dim i As Long
    Dim swComp As Object
    Dim hay As String
    Dim score As Double

    Dim tx As Double
    Dim ty As Double
    Dim tz As Double

    For i = 0 To UBound(vComps)

        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                hay = swComp.Name2 & " " & swComp.GetPathName

                If ContainsAnyPipeKey(hay, pipeKeys) Then

                    score = Len(hay)

                    If TryGetComponentCenterInches(swComp, tx, ty, tz) Then

                        If score > bestScore Then
                            bestScore = score
                            cx = tx
                            cy = ty
                            cz = tz
                            foundName = swComp.Name2
                            TryFindBestComponentCenterByKeys = True
                        End If

                    End If

                End If

            End If
        End If

    Next i

    Exit Function

ErrHandler:
    TryFindBestComponentCenterByKeys = False
End Function

' ============================================================
' END OF PART 1
' Paste Part 2 immediately after this.
' ============================================================
' ============================================================
' ZIP
' ============================================================

Private Sub ExtractAllZipFilesInJobFolder(ByVal jobFolder As String, ByVal extractRoot As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(jobFolder) Then Exit Sub

    EnsureFolderDeep extractRoot

    Dim zips As Collection
    Set zips = New Collection

    SearchZipFilesRecursive fso.GetFolder(jobFolder), zips

    LogLine "ZIP count=" & zips.count

    Dim i As Long
    For i = 1 To zips.count
        LogLine "Extracting ZIP " & i & "/" & zips.count & ": " & CStr(zips(i))
        If UnzipOneZipRobust(CStr(zips(i)), extractRoot) Then
            LogLine "ZIP extracted OK."
        Else
            LogLine "ZIP extract failed."
        End If
    Next i

    Exit Sub

ErrHandler:
    LogLine "ExtractAllZipFilesInJobFolder error: " & Err.Description
End Sub

Private Sub SearchZipFilesRecursive(ByVal folder As Object, ByRef zips As Collection)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderName As String
    folderName = UCase(folder.Name)

    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub
    If InStr(folderName, "PRINT") > 0 Then Exit Sub

    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.path)) = "zip" Then
            zips.Add file.path
        End If
    Next file

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        SearchZipFilesRecursive subFolder, zips
    Next subFolder
End Sub

Private Function UnzipOneZipRobust(ByVal zipPath As String, ByVal extractRoot As String) As Boolean
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(zipPath) Then
        UnzipOneZipRobust = False
        Exit Function
    End If

    Dim zipBaseName As String
    zipBaseName = CleanFileName(GetFileBaseName(zipPath))

    Dim finalDest As String
    finalDest = extractRoot & "\" & zipBaseName
    EnsureFolderDeep finalDest

    Dim tempRoot As String
    tempRoot = Environ$("TEMP") & "\CMS_ZIP_" & Format(Now, "yyyymmdd_hhnnss") & "_" & zipBaseName

    Dim tempOut As String
    tempOut = tempRoot & "\OUT"

    EnsureFolderDeep tempRoot
    EnsureFolderDeep tempOut

    Dim localZip As String
    localZip = tempRoot & "\" & zipBaseName & ".zip"

    FileCopy zipPath, localZip

    Dim ok As Boolean
    ok = ExtractZipUsingPowerShell(localZip, tempOut)

    If Not ok Then
        ok = ExtractZipUsingShell(localZip, tempOut)
    End If

    If Not ok Then
        UnzipOneZipRobust = False
        Exit Function
    End If

    CopyFolderContents tempOut, finalDest

    On Error Resume Next
    fso.DeleteFolder tempRoot, True
    On Error GoTo 0

    UnzipOneZipRobust = True
    Exit Function

ErrHandler:
    LogLine "UnzipOneZipRobust error: " & Err.Description
    UnzipOneZipRobust = False
End Function

Private Function ExtractZipUsingPowerShell(ByVal zipPath As String, ByVal destFolder As String) As Boolean
On Error GoTo ErrHandler

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & _
          "Expand-Archive -LiteralPath " & PowerShellQuote(zipPath) & _
          " -DestinationPath " & PowerShellQuote(destFolder) & " -Force" & Chr(34)

    ExtractZipUsingPowerShell = (sh.Run(cmd, 0, True) = 0)
    Exit Function

ErrHandler:
    ExtractZipUsingPowerShell = False
End Function

Private Function ExtractZipUsingShell(ByVal zipPath As String, ByVal destFolder As String) As Boolean
On Error GoTo ErrHandler

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")

    Dim z As Object
    Dim d As Object

    Set z = shellApp.NameSpace(zipPath)
    Set d = shellApp.NameSpace(destFolder)

    If z Is Nothing Or d Is Nothing Then
        ExtractZipUsingShell = False
        Exit Function
    End If

    d.CopyHere z.Items, 16 + 4
    WaitMilliseconds 3000

    ExtractZipUsingShell = True
    Exit Function

ErrHandler:
    ExtractZipUsingShell = False
End Function

Private Sub FlattenExtractedZipContentsIntoJobFolder(ByVal jobFolder As String, ByVal extractRoot As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If jobFolder = "" Then Exit Sub
    If extractRoot = "" Then Exit Sub
    If fso.FolderExists(jobFolder) = False Then Exit Sub
    If fso.FolderExists(extractRoot) = False Then Exit Sub

    LogLine "Flattening extracted ZIP contents into main job folder."
    LogLine "  FROM: " & extractRoot
    LogLine "  TO  : " & jobFolder

    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(extractRoot)

    Dim subFolder As Object
    Dim file As Object

    For Each file In rootFolder.Files
        fso.CopyFile file.path, jobFolder & "\" & file.Name, True
        LogLine "Flatten copied file: " & file.Name
    Next file

    For Each subFolder In rootFolder.SubFolders
        CopyExtractedFolderContentsToMain subFolder.path, jobFolder
    Next subFolder

    LogLine "Flatten extracted ZIP contents complete."
    Exit Sub

ErrHandler:
    LogLine "FlattenExtractedZipContentsIntoJobFolder error: " & Err.Description
End Sub

Private Sub CopyExtractedFolderContentsToMain(ByVal sourceFolder As String, ByVal mainJobFolder As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sourceFolder) = False Then Exit Sub
    If fso.FolderExists(mainJobFolder) = False Then Exit Sub

    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)

    Dim file As Object
    Dim subFolder As Object
    Dim destSub As String

    For Each file In src.Files
        fso.CopyFile file.path, mainJobFolder & "\" & file.Name, True
        LogLine "Flatten copied file: " & file.Name
    Next file

    For Each subFolder In src.SubFolders
        destSub = mainJobFolder & "\" & subFolder.Name
        EnsureFolderDeep destSub
        CopyFolderContents subFolder.path, destSub
        LogLine "Flatten copied folder: " & subFolder.Name
    Next subFolder

    Exit Sub

ErrHandler:
    LogLine "CopyExtractedFolderContentsToMain error: " & Err.Description
End Sub

Private Sub CopyFolderContents(ByVal sourceFolder As String, ByVal destFolder As String)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(sourceFolder) Then Exit Sub

    EnsureFolderDeep destFolder

    Dim src As Object
    Set src = fso.GetFolder(sourceFolder)

    Dim file As Object
    For Each file In src.Files
        fso.CopyFile file.path, destFolder & "\" & file.Name, True
    Next file

    Dim subFolder As Object
    For Each subFolder In src.SubFolders
        EnsureFolderDeep destFolder & "\" & subFolder.Name
        CopyFolderContents subFolder.path, destFolder & "\" & subFolder.Name
    Next subFolder
End Sub

' ============================================================
' CAD SCAN
' ============================================================

Private Sub ScanActiveSolidWorksDocument()
On Error GoTo ErrHandler

    Set swModel = swApp.activeDoc
    If swModel Is Nothing Then Exit Sub

    If swModel.GetType = swDocASSEMBLY Then

        Set swAssy = swModel

        On Error Resume Next
        swAssy.ResolveAllLightWeightComponents True
        On Error GoTo ErrHandler

        Dim vComps As Variant
        vComps = swAssy.GetComponents(False)
        If IsEmpty(vComps) Then Exit Sub

        Dim i As Long
        For i = 0 To UBound(vComps)
            If VERBOSE_PROGRESS_LOG And i Mod 10 = 0 Then
                LogProgress "Scanning comp " & (i + 1) & "/" & (UBound(vComps) + 1)
            End If

            ProcessAssemblyComponent vComps(i)
            DoEvents
        Next i

    ElseIf swModel.GetType = swDocPART Then

        ScanPartBodies swModel

    End If

    Exit Sub

ErrHandler:
    LogLine "ScanActiveSolidWorksDocument error: " & Err.Description
End Sub

Private Sub ProcessAssemblyComponent(ByVal swComp As Object)
On Error GoTo SafeExit

    If swComp Is Nothing Then Exit Sub
    If swComp.IsSuppressed Then Exit Sub

    If HIDE_QUARTER_INCH_THICKNESS And PHYSICALLY_HIDE_250_BEFORE_SCAN Then
        If IsComponentHidden(swComp) Then Exit Sub
    End If

    Dim compPath As String
    Dim compName As String
    Dim configName As String

    compPath = swComp.GetPathName
    compName = swComp.Name2
    configName = swComp.ReferencedConfiguration

    Dim swCompModel As Object
    Set swCompModel = swComp.GetModelDoc2

    If swCompModel Is Nothing And compPath <> "" Then
        Dim errs As Long
        Dim warns As Long

        Set swCompModel = swApp.OpenDoc6(compPath, _
                                         swDocPART, _
                                         swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, _
                                         configName, errs, warns)
    End If

    If swCompModel Is Nothing Then Exit Sub
    If swCompModel.GetType <> swDocPART Then Exit Sub

    Dim existingIndex As Long
    existingIndex = FindExistingPart(compPath, configName, compName)

    If existingIndex > 0 Then
        parts(existingIndex).Quantity = parts(existingIndex).Quantity + 1
        Exit Sub
    End If

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    If Not GetPartBoundingBoxInches(swCompModel, dx, dy, dz) Then Exit Sub

    Dim massValue As Double
    massValue = GetModelMassOrVolumeValue(swCompModel)

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double
    Dim hasCenter As Boolean

    hasCenter = TryGetComponentCenterInches(swComp, cx, cy, cz)

    Dim asmBoxL As Double
    Dim asmBoxW As Double
    Dim asmBoxT As Double
    Dim hasAsmBox As Boolean

    hasAsmBox = TryGetComponentAssemblyBoundingBoxInches(swComp, asmBoxL, asmBoxW, asmBoxT)

    If hasAsmBox Then
        LogLine "Assembly/world bbox for " & compName & " = " & _
                FormatNumberForCsv(asmBoxL) & "/" & _
                FormatNumberForCsv(asmBoxW) & "/" & _
                FormatNumberForCsv(asmBoxT)
    Else
        LogLine "Assembly/world bbox unavailable for " & compName & "; using part-local bbox."
    End If

    AddCadPart compName, _
               CleanDisplayName(compName, compPath), _
               compPath, configName, "", 1, _
               dx, dy, dz, massValue, False, _
               hasCenter, cx, cy, cz, _
               False, 0#, _
               hasAsmBox, asmBoxL, asmBoxW, asmBoxT

SafeExit:
End Sub

Private Sub ScanPartBodies(ByVal swPartModel As Object)
On Error GoTo ErrHandler

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If Not IsEmpty(vBox) Then

                Dim dx As Double
                Dim dy As Double
                Dim dz As Double

                dx = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * INCHES_PER_METER
                dy = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * INCHES_PER_METER
                dz = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * INCHES_PER_METER

                Dim massValue As Double
                massValue = GetBodyMassOrVolumeValue(swBody)
                If massValue <= 0 Then massValue = dx * dy * dz

                AddCadPart swBody.Name, swBody.Name, _
                           swPartModel.GetPathName, "", swBody.Name, 1, _
                           dx, dy, dz, massValue, True
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "ScanPartBodies error: " & Err.Description
End Sub

Private Sub AddCadPart(ByVal componentName As String, _
                       ByVal cleanName As String, _
                       ByVal filePath As String, _
                       ByVal configName As String, _
                       ByVal bodyName As String, _
                       ByVal qty As Long, _
                       ByVal dx As Double, _
                       ByVal dy As Double, _
                       ByVal dz As Double, _
                       ByVal massValue As Double, _
                       ByVal isBodyOnly As Boolean, _
                       Optional ByVal hasAsmCenter As Boolean = False, _
                       Optional ByVal asmCx As Double = 0#, _
                       Optional ByVal asmCy As Double = 0#, _
                       Optional ByVal asmCz As Double = 0#, _
                       Optional ByVal hasTopRotation As Boolean = False, _
                       Optional ByVal topRotationRad As Double = 0#, _
                       Optional ByVal hasOriginalAsmBBox As Boolean = False, _
                       Optional ByVal originalAsmL As Double = 0#, _
                       Optional ByVal originalAsmW As Double = 0#, _
                       Optional ByVal originalAsmT As Double = 0#)

    Dim L As Double
    Dim W As Double
    Dim T As Double

    SortThreeDimensions dx, dy, dz, L, W, T

    L = Round(L, DIM_DECIMALS)
    W = Round(W, DIM_DECIMALS)
    T = Round(T, DIM_DECIMALS)

    If L * W * T < MIN_STEEL_VOLUME_CUIN Then Exit Sub

    PartCount = PartCount + 1
    ReDim Preserve parts(1 To PartCount)

    parts(PartCount).componentName = componentName
    parts(PartCount).cleanName = cleanName
    parts(PartCount).filePath = filePath
    parts(PartCount).configName = configName
    parts(PartCount).bodyName = bodyName
    parts(PartCount).Quantity = qty

    parts(PartCount).Length = L
    parts(PartCount).Width = W
    parts(PartCount).Thickness = T
    parts(PartCount).BBoxVolume = L * W * T

    parts(PartCount).hasOriginalAsmBBox = hasOriginalAsmBBox

    If hasOriginalAsmBBox Then
        parts(PartCount).OriginalAsmLength = originalAsmL
        parts(PartCount).OriginalAsmWidth = originalAsmW
        parts(PartCount).OriginalAsmThickness = originalAsmT
    Else
        parts(PartCount).OriginalAsmLength = L
        parts(PartCount).OriginalAsmWidth = W
        parts(PartCount).OriginalAsmThickness = T
    End If

    If massValue > 0 Then
        parts(PartCount).massValue = massValue
    Else
        parts(PartCount).massValue = parts(PartCount).BBoxVolume
    End If

    parts(PartCount).hasAsmCenter = hasAsmCenter
    parts(PartCount).AsmCenterX = asmCx
    parts(PartCount).AsmCenterY = asmCy
    parts(PartCount).AsmCenterZ = asmCz

    parts(PartCount).hasTopRotation = hasTopRotation
    parts(PartCount).topRotationRad = topRotationRad

    parts(PartCount).UsedForBomMatch = False
    parts(PartCount).isBodyOnly = isBodyOnly
End Sub

Private Function FindExistingPart(ByVal filePath As String, ByVal configName As String, ByVal compName As String) As Long

    Dim i As Long

    For i = 1 To PartCount

        If filePath <> "" Then

            If LCase(parts(i).filePath) = LCase(filePath) And _
               LCase(parts(i).configName) = LCase(configName) Then

                FindExistingPart = i
                Exit Function

            End If

        Else

            If LCase(parts(i).componentName) = LCase(compName) Then
                FindExistingPart = i
                Exit Function
            End If

        End If

    Next i

    FindExistingPart = 0
End Function

Private Function GetPartBoundingBoxInches(ByVal swPartModel As Object, _
                                          ByRef dxIn As Double, _
                                          ByRef dyIn As Double, _
                                          ByRef dzIn As Double) As Boolean
On Error GoTo ErrHandler

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If

    Dim firstBody As Boolean
    firstBody = True

    Dim xmin As Double
    Dim yMin As Double
    Dim zMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim zMax As Double

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If Not IsEmpty(vBox) Then

                If firstBody Then
                    xmin = CDbl(vBox(0))
                    yMin = CDbl(vBox(1))
                    zMin = CDbl(vBox(2))
                    xmax = CDbl(vBox(3))
                    yMax = CDbl(vBox(4))
                    zMax = CDbl(vBox(5))
                    firstBody = False
                Else
                    If CDbl(vBox(0)) < xmin Then xmin = CDbl(vBox(0))
                    If CDbl(vBox(1)) < yMin Then yMin = CDbl(vBox(1))
                    If CDbl(vBox(2)) < zMin Then zMin = CDbl(vBox(2))
                    If CDbl(vBox(3)) > xmax Then xmax = CDbl(vBox(3))
                    If CDbl(vBox(4)) > yMax Then yMax = CDbl(vBox(4))
                    If CDbl(vBox(5)) > zMax Then zMax = CDbl(vBox(5))
                End If

            End If

        End If

    Next i

    If firstBody Then
        GetPartBoundingBoxInches = False
        Exit Function
    End If

    dxIn = Abs(xmax - xmin) * INCHES_PER_METER
    dyIn = Abs(yMax - yMin) * INCHES_PER_METER
    dzIn = Abs(zMax - zMin) * INCHES_PER_METER

    GetPartBoundingBoxInches = True
    Exit Function

ErrHandler:
    GetPartBoundingBoxInches = False
End Function

Private Function TryGetComponentAssemblyBoundingBoxInches(ByVal swComp As Object, _
                                                          ByRef outL As Double, _
                                                          ByRef outW As Double, _
                                                          ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentAssemblyBoundingBoxInches = False

    outL = 0#
    outW = 0#
    outT = 0#

    If swComp Is Nothing Then Exit Function

    Dim swPartModel As Object
    Set swPartModel = swComp.GetModelDoc2

    If swPartModel Is Nothing Then Exit Function
    If swPartModel.GetType <> swDocPART Then Exit Function

    Dim xmin As Double
    Dim yMin As Double
    Dim zMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim zMax As Double

    If TryGetPartRawBoxMeters(swPartModel, xmin, yMin, zMin, xmax, yMax, zMax) = False Then Exit Function

    Dim xform As Object
    Set xform = swComp.Transform2

    If xform Is Nothing Then Exit Function

    Dim m As Variant
    m = xform.ArrayData

    If IsEmpty(m) Then Exit Function
    If IsArray(m) = False Then Exit Function
    If UBound(m) < 12 Then Exit Function

    Dim scaleVal As Double
    scaleVal = CDbl(m(12))
    If Abs(scaleVal) < 0.0000001 Then scaleVal = 1#

    Dim xs(0 To 7) As Double
    Dim ys(0 To 7) As Double
    Dim zs(0 To 7) As Double

    xs(0) = xmin: ys(0) = yMin: zs(0) = zMin
    xs(1) = xmax: ys(1) = yMin: zs(1) = zMin
    xs(2) = xmin: ys(2) = yMax: zs(2) = zMin
    xs(3) = xmax: ys(3) = yMax: zs(3) = zMin
    xs(4) = xmin: ys(4) = yMin: zs(4) = zMax
    xs(5) = xmax: ys(5) = yMin: zs(5) = zMax
    xs(6) = xmin: ys(6) = yMax: zs(6) = zMax
    xs(7) = xmax: ys(7) = yMax: zs(7) = zMax

    Dim firstPoint As Boolean
    firstPoint = True

    Dim axmin As Double
    Dim aymin As Double
    Dim azmin As Double
    Dim axmax As Double
    Dim aymax As Double
    Dim azmax As Double

    Dim i As Long

    For i = 0 To 7

        Dim tx As Double
        Dim ty As Double
        Dim tz As Double

        tx = scaleVal * ((xs(i) * CDbl(m(0))) + (ys(i) * CDbl(m(3))) + (zs(i) * CDbl(m(6)))) + CDbl(m(9))
        ty = scaleVal * ((xs(i) * CDbl(m(1))) + (ys(i) * CDbl(m(4))) + (zs(i) * CDbl(m(7)))) + CDbl(m(10))
        tz = scaleVal * ((xs(i) * CDbl(m(2))) + (ys(i) * CDbl(m(5))) + (zs(i) * CDbl(m(8)))) + CDbl(m(11))

        If firstPoint Then
            axmin = tx
            aymin = ty
            azmin = tz
            axmax = tx
            aymax = ty
            azmax = tz
            firstPoint = False
        Else
            If tx < axmin Then axmin = tx
            If ty < aymin Then aymin = ty
            If tz < azmin Then azmin = tz

            If tx > axmax Then axmax = tx
            If ty > aymax Then aymax = ty
            If tz > azmax Then azmax = tz
        End If

    Next i

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    dx = Abs(axmax - axmin) * INCHES_PER_METER
    dy = Abs(aymax - aymin) * INCHES_PER_METER
    dz = Abs(azmax - azmin) * INCHES_PER_METER

    SortThreeDimensions dx, dy, dz, outL, outW, outT

    outL = Round(outL, DIM_DECIMALS)
    outW = Round(outW, DIM_DECIMALS)
    outT = Round(outT, DIM_DECIMALS)

    TryGetComponentAssemblyBoundingBoxInches = (outL > 0# And outW > 0# And outT > 0#)

    Exit Function

ErrHandler:
    TryGetComponentAssemblyBoundingBoxInches = False
End Function

Private Function TryGetPartRawBoxMeters(ByVal swPartModel As Object, _
                                        ByRef xmin As Double, _
                                        ByRef yMin As Double, _
                                        ByRef zMin As Double, _
                                        ByRef xmax As Double, _
                                        ByRef yMax As Double, _
                                        ByRef zMax As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPartRawBoxMeters = False

    If swPartModel Is Nothing Then Exit Function
    If swPartModel.GetType <> swDocPART Then Exit Function

    Dim vBodies As Variant
    vBodies = swPartModel.GetBodies2(swSolidBody, False)

    If IsEmpty(vBodies) Then Exit Function
    If IsArray(vBodies) = False Then Exit Function

    Dim firstBody As Boolean
    firstBody = True

    Dim i As Long
    Dim swBody As Object
    Dim vBox As Variant

    For i = 0 To UBound(vBodies)

        Set swBody = vBodies(i)

        If Not swBody Is Nothing Then

            vBox = swBody.GetBodyBox

            If IsEmpty(vBox) = False Then
                If IsArray(vBox) Then
                    If UBound(vBox) >= 5 Then

                        If firstBody Then
                            xmin = CDbl(vBox(0))
                            yMin = CDbl(vBox(1))
                            zMin = CDbl(vBox(2))
                            xmax = CDbl(vBox(3))
                            yMax = CDbl(vBox(4))
                            zMax = CDbl(vBox(5))
                            firstBody = False
                        Else
                            If CDbl(vBox(0)) < xmin Then xmin = CDbl(vBox(0))
                            If CDbl(vBox(1)) < yMin Then yMin = CDbl(vBox(1))
                            If CDbl(vBox(2)) < zMin Then zMin = CDbl(vBox(2))

                            If CDbl(vBox(3)) > xmax Then xmax = CDbl(vBox(3))
                            If CDbl(vBox(4)) > yMax Then yMax = CDbl(vBox(4))
                            If CDbl(vBox(5)) > zMax Then zMax = CDbl(vBox(5))
                        End If

                    End If
                End If
            End If

        End If

    Next i

    TryGetPartRawBoxMeters = Not firstBody

    Exit Function

ErrHandler:
    TryGetPartRawBoxMeters = False
End Function

Private Function TryGetComponentCenterInches(ByVal swComp As Object, _
                                             ByRef cx As Double, _
                                             ByRef cy As Double, _
                                             ByRef cz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentCenterInches = False

    cx = 0#
    cy = 0#
    cz = 0#

    If swComp Is Nothing Then Exit Function

    Dim vBox As Variant

    On Error Resume Next
    vBox = swComp.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    cx = ((CDbl(vBox(0)) + CDbl(vBox(3))) / 2#) * INCHES_PER_METER
    cy = ((CDbl(vBox(1)) + CDbl(vBox(4))) / 2#) * INCHES_PER_METER
    cz = ((CDbl(vBox(2)) + CDbl(vBox(5))) / 2#) * INCHES_PER_METER

    TryGetComponentCenterInches = True
    Exit Function

ErrHandler:
    TryGetComponentCenterInches = False
End Function

Private Function GetModelMassOrVolumeValue(ByVal model As Object) As Double
On Error GoTo ErrHandler

    If model Is Nothing Then
        GetModelMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim mp As Object
    Set mp = model.Extension.CreateMassProperty

    If mp Is Nothing Then
        GetModelMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim m As Double
    Dim v As Double

    m = 0#
    v = 0#

    On Error Resume Next
    m = CDbl(mp.Mass)
    v = CDbl(mp.Volume) * CUIN_PER_CUBIC_METER
    On Error GoTo ErrHandler

    If m > 0# Then
        GetModelMassOrVolumeValue = m
        Exit Function
    End If

    If v > 0# Then
        GetModelMassOrVolumeValue = v
        Exit Function
    End If

    GetModelMassOrVolumeValue = 0#
    Exit Function

ErrHandler:
    GetModelMassOrVolumeValue = 0#
End Function

Private Function GetBodyMassOrVolumeValue(ByVal swBody As Object) As Double
On Error GoTo ErrHandler

    If swBody Is Nothing Then
        GetBodyMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim vProps As Variant

    On Error Resume Next
    vProps = swBody.GetMassProperties(1#)
    On Error GoTo ErrHandler

    If IsEmpty(vProps) Then
        GetBodyMassOrVolumeValue = 0#
        Exit Function
    End If

    Dim v As Double
    v = 0#

    On Error Resume Next
    v = CDbl(vProps(3)) * CUIN_PER_CUBIC_METER
    On Error GoTo ErrHandler

    GetBodyMassOrVolumeValue = v
    Exit Function

ErrHandler:
    GetBodyMassOrVolumeValue = 0#
End Function

Private Sub SortThreeDimensions(ByVal a As Double, _
                                ByVal b As Double, _
                                ByVal c As Double, _
                                ByRef L As Double, _
                                ByRef W As Double, _
                                ByRef T As Double)

    Dim arr(1 To 3) As Double
    Dim i As Long
    Dim j As Long
    Dim tmp As Double

    arr(1) = a
    arr(2) = b
    arr(3) = c

    For i = 1 To 2
        For j = i + 1 To 3
            If arr(j) > arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    L = arr(1)
    W = arr(2)
    T = arr(3)
End Sub

Private Sub SortPartsByVolumeDescending()

    Dim i As Long
    Dim j As Long
    Dim tmp As PartInfo

    For i = 1 To PartCount - 1
        For j = i + 1 To PartCount
            If parts(j).BBoxVolume > parts(i).BBoxVolume Then
                tmp = parts(i)
                parts(i) = parts(j)
                parts(j) = tmp
            End If
        Next j
    Next i
End Sub

' ============================================================
' EXPORT X_T / DXF
' ============================================================

Private Sub ExportMatchedPartsAsXt(ByVal outputFolder As String)
On Error GoTo ErrHandler

    EnsureFolderDeep outputFolder
    EnsureSwHidden

    Dim i As Long

    For i = 1 To ExportCount
        LogLine "Exporting item " & i & "/" & ExportCount & ": " & ExportRows(i).quoteName
        ExportOneMatchedPartAsXt ExportRows(i), outputFolder, i
        DoEvents
    Next i

    If CREATE_MAIN_ASSEMBLY_PACKAGE Then
        LogStart "Create MAIN ASSEMBLY / HOLDERS package"
        ExportMainAssemblyAndHoldersPackage CurrentJobFolder
        LogDone "Create MAIN ASSEMBLY / HOLDERS package"
    End If

    If CREATE_J_BLOCK_PACKAGE Then
        LogStart "Create J BLOCK package"
        ExportJBlockPackage CurrentJobFolder
        LogDone "Create J BLOCK package"
    End If

    If CREATE_PULLCORE_CAM_KEY_PACKAGE Then
        LogStart "Create PULLCORE CAM AND KEY package"
        ExportPullcoreCamKeyPackage CurrentJobFolder
        LogDone "Create PULLCORE CAM AND KEY package"
    End If

    If CREATE_PULLCORE_STOP_PACKAGE Then
        LogStart "Create PULLCORE STOP package"
        ExportPullcoreStopPackage CurrentJobFolder
        LogDone "Create PULLCORE STOP package"
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportMatchedPartsAsXt error: " & Err.Description
End Sub

Private Sub ExportOneMatchedPartAsXt(ByRef q As ExportInfo, ByVal outputFolder As String, ByVal itemNumber As Long)
On Error GoTo ErrHandler

    If q.CadPartIndex <= 0 Or q.CadPartIndex > PartCount Then
        LogLine "  Skipped export (no CAD geometry for this BOM row): " & q.quoteName
        Exit Sub
    End If

    Dim p As PartInfo
    p = parts(q.CadPartIndex)

    Dim quoteToken As String
    quoteToken = CleanQuoteTokenForFile(q.quoteName)
    If quoteToken = "" Then quoteToken = "ITEM"

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim targetFolder As String
    Dim isPyropel As Boolean

    isPyropel = IsPyropelExportRow(q, p)

    If isPyropel And ROUTE_PYROPEL_TO_SEPARATE_FOLDERS Then

        If InStr(UCase(p.componentName), "POT") > 0 _
           Or InStr(UCase(p.cleanName), "POT") > 0 _
           Or InStr(UCase(q.quoteName), "POT") > 0 Then

            targetFolder = outputFolder & "\" & CurrentJobNumber & PYROPEL_POTS_FOLDER_SUFFIX
        Else
            targetFolder = outputFolder & "\" & CurrentJobNumber & PYROPEL_HOLDERS_FOLDER_SUFFIX
        End If

        LogLine "Pyropel routing -> " & targetFolder

    Else

        targetFolder = GetOutputFolderForRegularExport(q.quoteName, outputFolder)

    End If

    EnsureFolderDeep targetFolder

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = targetFolder & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".x_t"
    xtPath = GetUniqueFilePath(xtPath)

    dxfPath = targetFolder & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".dxf"
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Export file name: " & xtPath
    LogLine "  Quote name: " & q.quoteName
    LogLine "  Material: " & q.material
    LogLine "  CAD component: " & p.componentName
    LogLine "  Target folder: " & targetFolder

    Dim makeDxf As Boolean
    makeDxf = ShouldCreateStandardPrintDxf(q.quoteName)

    If isPyropel Then
        makeDxf = True
        LogLine "Pyropel export: creating X_T + DXF for routing."
    End If

    If swModel.GetType = swDocASSEMBLY Then

        ExportAssemblyOneComponentHideRestWithOptionalDxf swModel, p, xtPath, dxfPath, q.quoteName, makeDxf

    ElseIf swModel.GetType = swDocPART Then

        If p.isBodyOnly Then

            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath

            If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, q.quoteName
            End If

            ShowAllPartBodies swModel

        Else

            SaveModelAs swModel, xtPath

            If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, q.quoteName
            End If

        End If

    End If

    If Not ExportFilePaths Is Nothing Then
        ExportFilePaths(NormalizeKey(q.quoteName)) = xtPath
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportOneMatchedPartAsXt error: " & Err.Description
End Sub

Private Function IsPyropelExportRow(ByRef q As ExportInfo, ByRef p As PartInfo) As Boolean
On Error Resume Next

    Dim hay As String
    hay = UCase(q.material & " " & q.quoteName & " " & p.componentName & " " & p.cleanName & " " & p.filePath)

    IsPyropelExportRow = (InStr(hay, "PYROPEL") > 0)
End Function

Private Function ShouldCreateStandardPrintDxf(ByVal quoteName As String) As Boolean

    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "IDHOLDER" Then ShouldCreateStandardPrintDxf = True: Exit Function
    If k = "ODHOLDER" Then ShouldCreateStandardPrintDxf = True: Exit Function
    If k = "TCP" Then ShouldCreateStandardPrintDxf = True: Exit Function
    If k = "BCP" Then ShouldCreateStandardPrintDxf = True: Exit Function

End Function

' ============================================================
' SUPPRESS-ISOLATE EXPORT ROUTINES
' ============================================================

Private Sub ExportAssemblyOneComponentHideRestWithOptionalDxf(ByVal assyModel As Object, _
                                                              ByRef p As PartInfo, _
                                                              ByVal xtPath As String, _
                                                              ByVal dxfPath As String, _
                                                              ByVal quoteName As String, _
                                                              ByVal makeDxf As Boolean)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub

    LogLine "Suppress-isolate export. Target component: " & p.componentName

    If ExportAssemblyComponentBySuppressRestWithOptionalDxf(assyModel, p, xtPath, dxfPath, quoteName, makeDxf) Then
        LogLine "Suppress-isolate export succeeded."
    Else
        LogLine "WARNING: Suppress-isolate export failed for: " & p.componentName
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportAssemblyOneComponentHideRestWithOptionalDxf error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents assyModel
    ShowAllAssemblyComponents assyModel
End Sub

Private Function TryExportAssemblyComponentDirectWithOptionalDxf(ByVal assyModel As Object, _
                                                                 ByRef p As PartInfo, _
                                                                 ByVal xtPath As String, _
                                                                 ByVal dxfPath As String, _
                                                                 ByVal quoteName As String, _
                                                                 ByVal makeDxf As Boolean) As Boolean
On Error GoTo ErrHandler

    TryExportAssemblyComponentDirectWithOptionalDxf = _
        ExportAssemblyComponentBySuppressRestWithOptionalDxf(assyModel, p, xtPath, dxfPath, quoteName, makeDxf)
    Exit Function

ErrHandler:
    LogLine "TryExportAssemblyComponentDirectWithOptionalDxf error: " & Err.Description
    TryExportAssemblyComponentDirectWithOptionalDxf = False
End Function

Private Function ExportAssemblyComponentBySuppressRestWithOptionalDxf(ByVal assyModel As Object, _
                                                                      ByRef p As PartInfo, _
                                                                      ByVal xtPath As String, _
                                                                      ByVal dxfPath As String, _
                                                                      ByVal quoteName As String, _
                                                                      ByVal makeDxf As Boolean) As Boolean
On Error GoTo ErrHandler

    ExportAssemblyComponentBySuppressRestWithOptionalDxf = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    Dim i As Long
    Dim swComp As Object
    Dim foundTarget As Boolean
    Dim suppressCount As Long

    Dim suppressedObjects As Collection
    Set suppressedObjects = New Collection

    assyModel.ClearSelection2 True

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) = LCase(p.componentName) Then
                    If swComp.Select4(False, Nothing, False) Then foundTarget = True
                    Exit For
                End If
            End If
        End If
    Next i

    If foundTarget = False Then
        LogLine "Suppress-isolate failed: target component not found: " & p.componentName
        assyModel.ClearSelection2 True
        Exit Function
    End If

    assyModel.ClearSelection2 True

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) <> LCase(p.componentName) Then
                    If swComp.Select4(True, Nothing, False) Then
                        suppressCount = suppressCount + 1
                        suppressedObjects.Add swComp
                    End If
                End If
            End If
        End If
    Next i

    If suppressCount > 0 Then
        On Error Resume Next
        assyModel.EditSuppress2
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrHandler
    End If

    assyModel.ClearSelection2 True

    SaveModelAs assyModel, xtPath

    If fso.FileExists(xtPath) = False Then
        LogLine "Suppress-isolate failed: XT was not created."
        GoTo CleanExit
    End If

    If makeDxf And CREATE_DXFS_DURING_XT_SAVE Then

        Dim qk As String
        qk = NormalizeKey(quoteName)

        Select Case qk

            Case "IDHOLDER", "ODHOLDER"
                If CreateHolderDxfFromAssemblyBottomView(assyModel, dxfPath, quoteName) = False Then
                    LogLine "WARNING: " & quoteName & " assembly-bottom DXF failed; using part-based fallback."
                    CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName
                End If

            Case "TCP", "BCP"
                If CreateClampingPlateDxfFromAssemblyTopView(assyModel, dxfPath, quoteName) = False Then
                    LogLine "WARNING: " & quoteName & " assembly-top DXF failed; using part-based fallback."
                    CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName
                End If

            Case Else
                CreateStandardPrintDxfFromXtPath xtPath, dxfPath, quoteName

        End Select
    End If

    ExportAssemblyComponentBySuppressRestWithOptionalDxf = True

CleanExit:
    On Error Resume Next

    assyModel.ClearSelection2 True

    Dim obj As Object

    If Not suppressedObjects Is Nothing Then
        For i = 1 To suppressedObjects.count
            Set obj = suppressedObjects(i)
            If Not obj Is Nothing Then obj.Select4 True, Nothing, False
        Next i

        If suppressedObjects.count > 0 Then assyModel.EditUnsuppress2
    End If

    assyModel.ClearSelection2 True
    Exit Function

ErrHandler:
    LogLine "ExportAssemblyComponentBySuppressRestWithOptionalDxf error: " & Err.Description
    Resume CleanExit
End Function

Private Sub UnsuppressAllAssemblyComponents(ByVal assyModel As Object)
On Error Resume Next

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    assyModel.ClearSelection2 True

    Dim i As Long
    Dim swComp As Object
    Dim selectedCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed Then
                If swComp.Select4(True, Nothing, False) Then selectedCount = selectedCount + 1
            End If
        End If
    Next i

    If selectedCount > 0 Then assyModel.EditUnsuppress2

    assyModel.ClearSelection2 True
End Sub

Private Function FindAssemblyComponentByName(ByVal assyModel As Object, ByVal componentName As String) As Object
On Error GoTo ErrHandler

    Set FindAssemblyComponentByName = Nothing

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If componentName = "" Then Exit Function

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    Dim i As Long
    Dim swComp As Object

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If LCase(swComp.Name2) = LCase(componentName) Then
                    Set FindAssemblyComponentByName = swComp
                    Exit Function
                End If
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    Set FindAssemblyComponentByName = Nothing
End Function

' ============================================================
' ASSEMBLY/PART VISIBILITY HELPERS
' ============================================================

Private Function HideAllExceptTwoTargetComponentsOnce(ByVal assyModel As Object, _
                                                      ByVal targetComponentName1 As String, _
                                                      ByVal targetComponentName2 As String, _
                                                      ByRef hiddenNames As Collection) As Boolean
On Error GoTo ErrHandler

    HideAllExceptTwoTargetComponentsOnce = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    assyModel.ClearSelection2 True

    Dim i As Long
    Dim swComp As Object
    Dim foundTarget1 As Boolean
    Dim foundTarget2 As Boolean
    Dim selectedCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                If LCase(swComp.Name2) = LCase(targetComponentName1) Then
                    foundTarget1 = True
                ElseIf LCase(swComp.Name2) = LCase(targetComponentName2) Then
                    foundTarget2 = True
                Else
                    If swComp.Select4(True, Nothing, False) Then
                        hiddenNames.Add swComp.Name2
                        selectedCount = selectedCount + 1
                    End If
                End If

            End If
        End If
    Next i

    If foundTarget1 = False Or foundTarget2 = False Then
        LogLine "WARNING: Did not find both holder target components."
        assyModel.ClearSelection2 True
        Exit Function
    End If

    If selectedCount > 0 Then
        LogLine "Selected non-holder components to hide: " & selectedCount
        swAssembly.HideComponent2
    End If

    assyModel.ClearSelection2 True
    HideAllExceptTwoTargetComponentsOnce = True
    Exit Function

ErrHandler:
    LogLine "HideAllExceptTwoTargetComponentsOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    HideAllExceptTwoTargetComponentsOnce = False
End Function

Private Function HideAllExceptComponentNamesOnce(ByVal assyModel As Object, _
                                                 ByVal keepNames As Collection, _
                                                 ByRef hiddenNames As Collection) As Boolean
On Error GoTo ErrHandler

    HideAllExceptComponentNamesOnce = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function
    If keepNames Is Nothing Then Exit Function
    If keepNames.count = 0 Then Exit Function

    If hiddenNames Is Nothing Then Set hiddenNames = New Collection

    Dim keepDict As Object
    Set keepDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To keepNames.count
        keepDict(LCase(CStr(keepNames(i)))) = True
    Next i

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Function

    assyModel.ClearSelection2 True

    Dim swComp As Object
    Dim selectedCount As Long
    Dim keepFoundCount As Long

    selectedCount = 0
    keepFoundCount = 0

    For i = 0 To UBound(vComps)

        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                If keepDict.Exists(LCase(swComp.Name2)) Then
                    keepFoundCount = keepFoundCount + 1
                    swComp.Visible = swComponentVisible
                Else
                    If swComp.Select4(True, Nothing, False) Then
                        hiddenNames.Add swComp.Name2
                        selectedCount = selectedCount + 1
                    End If
                End If

            End If
        End If

    Next i

    If keepFoundCount = 0 Then
        LogLine "HideAllExceptComponentNamesOnce: none of the requested keep components were found."
        assyModel.ClearSelection2 True
        Exit Function
    End If

    If selectedCount > 0 Then
        LogLine "BASE DXF isolation: hiding non-selected components = " & selectedCount
        swAssembly.HideComponent2
    End If

    assyModel.ClearSelection2 True

    LogLine "BASE DXF isolation: selected components found = " & keepFoundCount
    HideAllExceptComponentNamesOnce = True
    Exit Function

ErrHandler:
    LogLine "HideAllExceptComponentNamesOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    HideAllExceptComponentNamesOnce = False
End Function

Private Sub ShowNamedComponentsOnce(ByVal assyModel As Object, ByVal componentNames As Collection)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub
    If componentNames Is Nothing Then Exit Sub
    If componentNames.count = 0 Then Exit Sub

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    assyModel.ClearSelection2 True

    Dim nameDict As Object
    Set nameDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To componentNames.count
        nameDict(LCase(CStr(componentNames(i)))) = True
    Next i

    Dim swComp As Object
    Dim selectedCount As Long

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then
                If nameDict.Exists(LCase(swComp.Name2)) Then
                    If swComp.Select4(True, Nothing, False) Then selectedCount = selectedCount + 1
                End If
            End If
        End If
    Next i

    If selectedCount > 0 Then
        swAssembly.ShowComponent2
    Else
        ShowAllAssemblyComponents assyModel
    End If

    assyModel.ClearSelection2 True
    Exit Sub

ErrHandler:
    LogLine "ShowNamedComponentsOnce error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    ShowAllAssemblyComponents assyModel
End Sub

Private Sub ShowOnlyTwoAssemblyComponents(ByVal assyModel As Object, _
                                          ByVal targetComponentName1 As String, _
                                          ByVal targetComponentName2 As String)
On Error Resume Next

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    Dim swAssembly As Object
    Set swAssembly = assyModel

    Dim vComps As Variant
    vComps = swAssembly.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long
    Dim swComp As Object

    For i = 0 To UBound(vComps)
        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If Not swComp.IsSuppressed Then
                If LCase(swComp.Name2) = LCase(targetComponentName1) Or _
                   LCase(swComp.Name2) = LCase(targetComponentName2) Then
                    swComp.Visible = swComponentVisible
                Else
                    swComp.Visible = swComponentHidden
                End If
            End If
        End If
    Next i
End Sub

Private Sub ShowAllAssemblyComponents(ByVal model As Object)
On Error Resume Next

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    Dim vComps As Variant
    vComps = model.GetComponents(False)
    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vComps)
        If Not vComps(i) Is Nothing Then
            If Not vComps(i).IsSuppressed Then vComps(i).Visible = swComponentVisible
        End If
    Next i
End Sub

Private Sub ShowOnlyPartBody(ByVal partModel As Object, ByVal targetBodyName As String)
On Error Resume Next

    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vBodies)
        If LCase(vBodies(i).Name) = LCase(targetBodyName) Then
            vBodies(i).Hide2 False
        Else
            vBodies(i).Hide2 True
        End If
    Next i
End Sub

Private Sub ShowAllPartBodies(ByVal partModel As Object)
On Error Resume Next

    If partModel Is Nothing Then Exit Sub
    If partModel.GetType <> swDocPART Then Exit Sub

    Dim vBodies As Variant
    vBodies = partModel.GetBodies2(swSolidBody, False)
    If IsEmpty(vBodies) Then Exit Sub

    Dim i As Long

    For i = 0 To UBound(vBodies)
        vBodies(i).Hide2 False
    Next i
End Sub

Private Sub SaveModelAs(ByVal model As Object, ByVal fullPath As String)
On Error GoTo ErrHandler

    Dim errs As Long
    Dim warns As Long

    LogLine "Saving: " & fullPath
    EnsureSwHidden

    model.Extension.SaveAs3 fullPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns

    LogLine "Save done. Errors=" & errs & " Warnings=" & warns
    Exit Sub

ErrHandler:
    LogLine "SaveModelAs error: " & Err.Description
End Sub

Private Sub ExportIndividualHolderAndClampingDxfs(ByVal outputFolder As String)
On Error Resume Next
    LogLine "ExportIndividualHolderAndClampingDxfs skipped. DXFs created during XT export."
End Sub

' ============================================================
' END OF PART 2
' Paste Part 3 immediately after this.
' ============================================================
' ============================================================
' MAIN ASSEMBLY / HOLDERS PACKAGE
' ============================================================

Private Sub ExportMainAssemblyAndHoldersPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "MAIN ASSEMBLY package skipped: active model is not an assembly."
        Exit Sub
    End If

    EnsureFolderDeep outputFolder

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim mainToken As String
    mainToken = CleanQuoteTokenForFile(MAIN_ASSEMBLY_FILE_TOKEN)
    If mainToken = "" Then mainToken = "BASE"

    Dim mainEasmPath As String
    Dim mainXtPath As String
    Dim mainIgsPath As String
    Dim mainDxfPath As String

    mainEasmPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".easm"
    mainXtPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".x_t"
    mainIgsPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".igs"
    mainDxfPath = outputFolder & "\" & CurrentJobNumber & "_" & mainToken & "_" & custToken & "_" & dateToken & ".dxf"

    mainEasmPath = GetUniqueFilePath(mainEasmPath)
    mainXtPath = GetUniqueFilePath(mainXtPath)
    mainIgsPath = GetUniqueFilePath(mainIgsPath)
    mainDxfPath = GetUniqueFilePath(mainDxfPath)

    LogLine "Exporting MAIN ASSEMBLY BASE package."

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    SaveModelAs swModel, mainEasmPath
    SaveModelAs swModel, mainXtPath
    SaveModelAs swModel, mainIgsPath

    CreateBaseDxfFromSelectedComponents mainXtPath, mainDxfPath

    Dim idHolderIdx As Long
    Dim odHolderIdx As Long

    idHolderIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odHolderIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    If idHolderIdx <= 0 Or odHolderIdx <= 0 Then
        LogLine "HOLDERS package skipped: could not find both holders."
    Else
        ExportHoldersOnlyMainFolderPackage CurrentJobFolder, idHolderIdx, odHolderIdx
    End If

    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    Exit Sub

ErrHandler:
    LogLine "ExportMainAssemblyAndHoldersPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Sub CreateBaseDxfFromSelectedComponents(ByVal fullBaseXtPath As String, _
                                                ByVal mainDxfPath As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If swModel.GetType <> swDocASSEMBLY Then
        LogLine "BASE selected-components DXF skipped: active model is not an assembly."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        Exit Sub
    End If

    Dim keepNames As Collection
    Set keepNames = BuildBaseDxfKeepComponentNames()

    If keepNames Is Nothing Or keepNames.count = 0 Then
        LogLine "WARNING: BASE DXF selected component list is empty. Falling back to full BASE DXF."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        Exit Sub
    End If

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_BASE_DXF_SELECTED_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Use native SLDASM for DXF creation.
    ' X_T does not preserve CMS_TOP / redefined standard views.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_BASE_DXF_SELECTED_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    LogLine "Creating BASE DXF from selected components only. Count=" & keepNames.count

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "WARNING: Could not isolate selected BASE DXF components. Falling back to full BASE DXF."

        If BaseNativeAssemblyPath <> "" Then
            CreateProjectedDxfFromNativePath BaseNativeAssemblyPath, mainDxfPath, "MAIN ASSEMBLY", _
                                             CMS_TOP_VIEW_NAME, "*Top", _
                                             False, False, False, True
        Else
            CreateProjectedDxfFromXtPath fullBaseXtPath, mainDxfPath, "MAIN ASSEMBLY", _
                                         CMS_TOP_VIEW_NAME, "*Top", False, False, False, True
        End If

        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    LogLine "BASE DXF: saving selected-component native temp SLDASM:"
    LogLine "  " & tempNativePath

    SaveModelAs swModel, tempNativePath

    CreateProjectedDxfFromNativePath tempNativePath, mainDxfPath, "MAIN ASSEMBLY", _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     False, False, False, True

CleanExit:
    On Error Resume Next

    If Not hiddenNames Is Nothing Then
        If hiddenNames.count > 0 Then
            ShowNamedComponentsOnce swModel, hiddenNames
        Else
            ShowAllAssemblyComponents swModel
        End If
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso Is Nothing Then
        If tempFolder <> "" Then
            If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
        End If
    End If

    Exit Sub

ErrHandler:
    LogLine "CreateBaseDxfFromSelectedComponents error: " & Err.Description
    Resume CleanExit
End Sub

Private Function BuildBaseDxfKeepComponentNames() As Collection
On Error GoTo ErrHandler

    Dim keepNames As New Collection

    AddBaseDxfKeepComponentFromQuote keepNames, "TCP", "TCP", _
        "TCP|TOP SMED|TOP SMED PLATE|TOP CLAMPING PLATE|TOP CLAMPING|ID SMED"

    AddBaseDxfKeepComponentFromQuote keepNames, "BCP", "BCP", _
        "BCP|BOTTOM SMED|BOT SMED|BOTTOM SMED PLATE|BOT SMED PLATE|BOTTOM CLAMPING PLATE|BOT CLAMPING PLATE|BOTTOM CLAMPING|BOT CLAMPING|OD SMED"

    AddBaseDxfKeepComponentFromQuote keepNames, "ID HOLDER", "ID HOLDER", _
        ID_HOLDER_KEYS

    AddBaseDxfKeepComponentFromQuote keepNames, "OD HOLDER", "OD HOLDER", _
        OD_HOLDER_KEYS

    AddBaseDxfKeepComponentFromQuote keepNames, "ID POT", "ID POT BLOCK", _
        "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT"

    AddBaseDxfKeepComponentFromQuote keepNames, "OD POT", "OD POT BLOCK", _
        "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT"

    Set BuildBaseDxfKeepComponentNames = keepNames
    Exit Function

ErrHandler:
    LogLine "BuildBaseDxfKeepComponentNames error: " & Err.Description
    Set BuildBaseDxfKeepComponentNames = New Collection
End Function

Private Sub AddBaseDxfKeepComponentFromQuote(ByVal keepNames As Collection, _
                                             ByVal label As String, _
                                             ByVal quoteName As String, _
                                             ByVal fallbackKeys As String)
On Error GoTo ErrHandler

    If keepNames Is Nothing Then Exit Sub

    Dim cadIdx As Long

    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then

        AddUniqueComponentName keepNames, parts(cadIdx).componentName

        LogLine "BASE DXF include " & label & ": CAD '" & _
                parts(cadIdx).componentName & "'  L/W/T=" & _
                FormatNumberForCsv(parts(cadIdx).Length) & "/" & _
                FormatNumberForCsv(parts(cadIdx).Width) & "/" & _
                FormatNumberForCsv(parts(cadIdx).Thickness)
    Else
        LogLine "WARNING: BASE DXF could not find component for " & label & _
                " using quote '" & quoteName & "'."
    End If

    Exit Sub

ErrHandler:
    LogLine "AddBaseDxfKeepComponentFromQuote error (" & label & "): " & Err.Description
End Sub

Private Function FindCadIndexFromExportQuote(ByVal quoteName As String) As Long
On Error GoTo ErrHandler

    FindCadIndexFromExportQuote = 0

    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                FindCadIndexFromExportQuote = ExportRows(i).CadPartIndex
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    FindCadIndexFromExportQuote = 0
End Function

Private Sub AddUniqueComponentName(ByVal names As Collection, ByVal componentName As String)
On Error Resume Next

    If names Is Nothing Then Exit Sub

    componentName = Trim(componentName)
    If componentName = "" Then Exit Sub

    Dim i As Long

    For i = 1 To names.count
        If LCase(CStr(names(i))) = LCase(componentName) Then Exit Sub
    Next i

    names.Add componentName
End Sub

Private Sub ExportHoldersOnlyMainFolderPackage(ByVal outputFolder As String, _
                                               ByVal idHolderIdx As Long, _
                                               ByVal odHolderIdx As Long)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub
    If swModel.GetType <> swDocASSEMBLY Then Exit Sub
    If idHolderIdx <= 0 Or idHolderIdx > PartCount Then Exit Sub
    If odHolderIdx <= 0 Or odHolderIdx > PartCount Then Exit Sub

    EnsureFolderDeep outputFolder

    Dim idPart As PartInfo
    Dim odPart As PartInfo

    idPart = parts(idHolderIdx)
    odPart = parts(odHolderIdx)

    Dim custToken As String
    Dim dateToken As String
    Dim holderToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)
    holderToken = CleanQuoteTokenForFile(HOLDERS_ONLY_FILE_TOKEN)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")
    If holderToken = "" Then holderToken = "HOLDERS"

    Dim holdersIgsPath As String
    Dim holdersDxfPath As String

    holdersIgsPath = outputFolder & "\" & CurrentJobNumber & "_" & holderToken & "_" & custToken & "_" & dateToken & ".igs"
    holdersDxfPath = outputFolder & "\" & CurrentJobNumber & "_" & holderToken & "_" & custToken & "_" & dateToken & ".dxf"

    holdersIgsPath = GetUniqueFilePath(holdersIgsPath)
    holdersDxfPath = GetUniqueFilePath(holdersDxfPath)

    Dim tempFolder As String
    Dim holdersTempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_HOLDERS_ONLY_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Native SLDASM for DXF so CMS_TOP / standard orientation survives.
    holdersTempNativePath = tempFolder & "\" & CurrentJobNumber & "_HOLDERS_TEMP.sldasm"

    LogLine "Exporting HOLDERS package."

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptTwoTargetComponentsOnce(swModel, idPart.componentName, odPart.componentName, hiddenNames) = False Then
        ShowOnlyTwoAssemblyComponents swModel, idPart.componentName, odPart.componentName
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 200

    SaveModelAs swModel, holdersIgsPath

    LogLine "HOLDERS DXF: saving native temp SLDASM:"
    LogLine "  " & holdersTempNativePath

    SaveModelAs swModel, holdersTempNativePath

    CreateProjectedDxfFromNativePath holdersTempNativePath, holdersDxfPath, "HOLDERS", _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     False, False, False, True

    If hiddenNames.count > 0 Then
        ShowNamedComponentsOnce swModel, hiddenNames
    Else
        ShowAllAssemblyComponents swModel
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

CleanExit:
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso Is Nothing Then
        If fso.FolderExists(tempFolder) Then fso.DeleteFolder tempFolder, True
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportHoldersOnlyMainFolderPackage error: " & Err.Description
    On Error Resume Next
    ShowAllAssemblyComponents swModel
    Resume CleanExit
End Sub

Private Function FindHolderCadIndexForMainPackage(ByVal holderQuoteName As String) As Long
On Error GoTo ErrHandler

    Dim k As String
    k = NormalizeKey(holderQuoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                FindHolderCadIndexForMainPackage = ExportRows(i).CadPartIndex
                Exit Function
            End If
        End If
    Next i

    If k = "IDHOLDER" Then
        FindHolderCadIndexForMainPackage = FindCadPartIndexByQuoteOrKeys("ID HOLDER", ID_HOLDER_KEYS)
        Exit Function
    ElseIf k = "ODHOLDER" Then
        FindHolderCadIndexForMainPackage = FindCadPartIndexByQuoteOrKeys("OD HOLDER", OD_HOLDER_KEYS)
        Exit Function
    End If

    FindHolderCadIndexForMainPackage = 0
    Exit Function

ErrHandler:
    LogLine "FindHolderCadIndexForMainPackage error for " & holderQuoteName & ": " & Err.Description
    FindHolderCadIndexForMainPackage = 0
End Function

' ============================================================
' STANDARD DXF FROM SAVED XT
' ============================================================

Private Sub CreateStandardPrintDxfFromXtPath(ByVal xtPath As String, _
                                             ByVal dxfPath As String, _
                                             ByVal quoteName As String)
On Error GoTo ErrHandler

    If xtPath = "" Then Exit Sub
    If dxfPath = "" Then Exit Sub

    Dim parentPrimary As String
    Dim parentFallback As String

    If NormalizeKey(quoteName) = "IDHOLDER" Then
        parentPrimary = "*Bottom"
        parentFallback = CMS_TOP_VIEW_NAME
    Else
        parentPrimary = CMS_TOP_VIEW_NAME
        parentFallback = "*Top"
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, quoteName, _
                                 parentPrimary, parentFallback, True, False, False

    Exit Sub

ErrHandler:
    LogLine "CreateStandardPrintDxfFromXtPath error for " & quoteName & ": " & Err.Description
End Sub

Private Function CreateHolderDxfFromAssemblyBottomView(ByVal assyModel As Object, _
                                                       ByVal dxfPath As String, _
                                                       ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    CreateHolderDxfFromAssemblyBottomView = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_HOLDER_DXF_" & Format(Now, "yyyymmdd_hhnnss") & "_" & NormalizeKey(quoteName)
    EnsureFolderDeep tempFolder

    ' Native assembly, not X_T.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & NormalizeKey(quoteName) & "_BOTTOMVIEW_TEMP.sldasm"

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 100

    assyModel.ShowNamedView2 "*Bottom", 6
    StabilizeActiveView assyModel, 100

    LogLine quoteName & " holder DXF: saving native temp SLDASM:"
    LogLine "  " & tempNativePath

    SaveModelAs assyModel, tempNativePath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Holder assembly-bottom DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CurrentIdHolderDxfFromAssembly = True

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, quoteName, _
                                     "*Bottom", "*Top", _
                                     True, False, False, False, False

    CurrentIdHolderDxfFromAssembly = False

    CreateHolderDxfFromAssemblyBottomView = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    CurrentIdHolderDxfFromAssembly = False

    ApplyCmsTopView assyModel

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateHolderDxfFromAssemblyBottomView error: " & Err.Description
    CreateHolderDxfFromAssemblyBottomView = False
    Resume CleanExit
End Function

Private Function CreateClampingPlateDxfFromAssemblyTopView(ByVal assyModel As Object, _
                                                          ByVal dxfPath As String, _
                                                          ByVal quoteName As String) As Boolean
On Error GoTo ErrHandler

    CreateClampingPlateDxfFromAssemblyTopView = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_CLAMP_DXF_" & Format(Now, "yyyymmdd_hhnnss") & "_" & NormalizeKey(quoteName)
    EnsureFolderDeep tempFolder

    ' Native assembly, not X_T.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & NormalizeKey(quoteName) & "_TOPVIEW_TEMP.sldasm"

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 100

    LogLine quoteName & " clamping plate DXF: saving native temp SLDASM:"
    LogLine "  " & tempNativePath

    SaveModelAs assyModel, tempNativePath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Clamping-plate assembly-top DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, quoteName, _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     True, False, False, False, False

    CreateClampingPlateDxfFromAssemblyTopView = fso.FileExists(dxfPath)

CleanExit:
    On Error Resume Next

    ApplyCmsTopView assyModel

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateClampingPlateDxfFromAssemblyTopView error: " & Err.Description
    CreateClampingPlateDxfFromAssemblyTopView = False
    Resume CleanExit
End Function

Private Sub CreateProjectedDxfFromXtPath(ByVal xtPath As String, _
                                         ByVal dxfPath As String, _
                                         ByVal quoteName As String, _
                                         ByVal parentPrimaryView As String, _
                                         ByVal parentFallbackView As String, _
                                         ByVal allWireframe As Boolean, _
                                         ByVal jBlockMixedDisplay As Boolean, _
                                         ByVal dimensionedJBlock As Boolean, _
                                         Optional ByVal allSolid As Boolean = False, _
                                         Optional ByVal pullcoreAux As Boolean = False)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(xtPath) = False Then
        LogLine "DXF skipped. XT source missing: " & xtPath
        Exit Sub
    End If

    Dim tempFolder As String
    tempFolder = Environ$("TEMP") & "\CMS_DXF_FROM_XT_" & Format(Now, "yyyymmdd_hhnnss") & "_" & CleanFileName(quoteName)

    EnsureFolderDeep tempFolder

    Dim nativePath As String
    nativePath = OpenXtAndSaveNativeForDrawing_Fixed(xtPath, tempFolder, NormalizeKey(quoteName))

    If nativePath = "" Then
        LogLine "DXF skipped. Could not open XT/native source for: " & xtPath
        GoTo CleanExit
    End If

    CreateProjectedDxfFromNativePath nativePath, dxfPath, quoteName, _
                                     parentPrimaryView, parentFallbackView, _
                                     allWireframe, jBlockMixedDisplay, dimensionedJBlock, allSolid, pullcoreAux

CleanExit:
    On Error Resume Next

    If tempFolder <> "" Then fso.DeleteFolder tempFolder, True
    Exit Sub

ErrHandler:
    LogLine "CreateProjectedDxfFromXtPath error for " & quoteName & ": " & Err.Description
    Resume CleanExit
End Sub

Private Function OpenXtAndSaveNativeForDrawing_Fixed(ByVal xtPath As String, _
                                                     ByVal tempFolder As String, _
                                                     ByVal baseName As String) As String
On Error GoTo ErrHandler

    OpenXtAndSaveNativeForDrawing_Fixed = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(xtPath) = False Then Exit Function

    EnsureFolderDeep tempFolder

    Dim importErrors As Long
    Dim errs As Long
    Dim warns As Long

    Dim mdl As Object
    Set mdl = swApp.LoadFile4(xtPath, "", Nothing, importErrors)

    If mdl Is Nothing Then
        Set mdl = swApp.OpenDoc6(xtPath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then
        Set mdl = swApp.OpenDoc6(xtPath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then
        LogLine "OpenXtAndSaveNativeForDrawing_Fixed failed to open: " & xtPath
        Exit Function
    End If

    Dim nativePath As String

    If mdl.GetType = swDocASSEMBLY Then
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldasm"
    Else
        nativePath = tempFolder & "\" & CleanFileName(baseName) & ".sldprt"
    End If

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    EnsureSwHidden

    If CurrentDxfKeepImportedOrientation Then
        LogLine "Native drawing source kept in imported orientation (no re-orient)."
    Else
        SetCmsTopOrientation mdl
        ApplyCmsTopView mdl
        StabilizeActiveView mdl, 150
    End If

    mdl.Extension.SaveAs3 nativePath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, errs, warns

    If fso.FileExists(nativePath) Then
        OpenXtAndSaveNativeForDrawing_Fixed = nativePath
    Else
        LogLine "Native temp save failed: " & nativePath
    End If

    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    Exit Function

ErrHandler:
    LogLine "OpenXtAndSaveNativeForDrawing_Fixed error: " & Err.Description
    OpenXtAndSaveNativeForDrawing_Fixed = ""
End Function

Private Sub CreateProjectedDxfFromNativePath(ByVal nativePath As String, _
                                             ByVal dxfPath As String, _
                                             ByVal quoteName As String, _
                                             ByVal parentPrimaryView As String, _
                                             ByVal parentFallbackView As String, _
                                             ByVal allWireframe As Boolean, _
                                             ByVal jBlockMixedDisplay As Boolean, _
                                             ByVal dimensionedJBlock As Boolean, _
                                             Optional ByVal allSolid As Boolean = False, _
                                             Optional ByVal pullcoreAux As Boolean = False)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim drawTitle As String
    drawTitle = ""

    Dim freezeApplied As Boolean
    freezeApplied = False

    If fso.FileExists(SW_DRAWING_TEMPLATE_PATH) = False Then
        LogLine "DXF skipped. Drawing template not found: " & SW_DRAWING_TEMPLATE_PATH
        Exit Sub
    End If

    If fso.FileExists(nativePath) = False Then
        LogLine "DXF skipped. Native source missing: " & nativePath
        Exit Sub
    End If

    Dim partL As Double
    Dim partW As Double
    Dim partT As Double

    partL = 0#
    partW = 0#
    partT = 0#

    GetBestCurrentOutputDimensions quoteName, partL, partW, partT

    Dim liveL As Double
    Dim liveW As Double
    Dim liveT As Double

    If TryGetNativeModelDimsInches(nativePath, liveL, liveW, liveT) Then
        If liveL > 0# Then partL = liveL
        If liveW > 0# Then partW = liveW
        If liveT > 0# Then partT = liveT
    End If

    If IsHoldersDxfQuote(quoteName) Then
        If partL <= 0# Then partL = GetCombinedHoldersLengthForDxf()
        If partW <= 0# Then partW = GetCombinedHoldersWidthForDxf()
        If partT <= 0# Then partT = GetCombinedHoldersThicknessForDxf()

        If partL <= 0# Then partL = 42#
        If partW <= 0# Then partW = 28#
        If partT <= 0# Then partT = 10#
    End If

    If IsMainBaseDxfQuote(quoteName) Then
        If partL <= 0# Then partL = 42#
        If partW <= 0# Then partW = 30#
        If partT <= 0# Then partT = 10#
    End If

    If IsJBlockDxfQuote(quoteName) Then
        If partL <= 0# Then partL = JBlockTgtL
        If partW <= 0# Then partW = JBlockTgtW
        If partT <= 0# Then partT = JBlockTgtT
    End If

    If partL <= 0# Then partL = 10#
    If partW <= 0# Then partW = 10#
    If partT <= 0# Then partT = 2#

    ' UPDATED:
    ' Force 1:1 for ALL DXF views/files when enabled.
    ' No BASE/HOLDERS override.
    CurrentDxfForce1to1 = FORCE_ALL_DXF_VIEWS_1_TO_1

    Dim scaleVal As Double

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Then

        CurrentDxfForce1to1 = True
        scaleVal = 1#

        LogLine "DXF scale forced 1:1 for " & quoteName

    Else

        CurrentDxfForce1to1 = False

        scaleVal = CalculateProjectedFourViewDxfScale(partL, partW, partT)
        scaleVal = scaleVal * MULTIVIEW_FIT_SAFETY

        If scaleVal <= 0# Then scaleVal = 0.1

        LogLine "DXF auto-fit scale for " & quoteName & " = " & Format(scaleVal, "0.0000")

    End If

    If scaleVal = 1# Then
        Dim layoutWCheck As Double
        Dim layoutHCheck As Double

        layoutWCheck = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
        layoutHCheck = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)

        If layoutWCheck > (E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) Or _
           layoutHCheck > (E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)) Then

            LogLine "WARNING: 1:1 DXF layout may not fit E-size sheet for " & quoteName & _
                    ". Required W/H=" & FormatNumberForCsv(layoutWCheck) & "/" & _
                    FormatNumberForCsv(layoutHCheck) & _
                    ", usable W/H=" & FormatNumberForCsv(E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) & "/" & _
                    FormatNumberForCsv(E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN))
        End If
    End If

    If IsDxfInPrintsFolder(dxfPath) Then

        allWireframe = True
        allSolid = False
        jBlockMixedDisplay = False
        dimensionedJBlock = False

        If IsClampingPlateDxfQuote(quoteName) Then
            LogLine "TCP/BCP DXF detected. Starting hidden-lines-visible; left/bottom will be forced solid."
        Else
            LogLine "PRINTS folder DXF detected. Forcing HIDDEN LINES VISIBLE: " & dxfPath
        End If

    End If

    LogLine "Creating DXF: quote=" & quoteName & " native=" & nativePath
    LogLine "  L/W/T=" & FormatNumberForCsv(partL) & "/" & FormatNumberForCsv(partW) & "/" & FormatNumberForCsv(partT) & _
            "  scale=" & Format(scaleVal, "0.0000")

    CurrentDxfStraightenAngleRad = 0#

    If STRAIGHTEN_PULLCORE_DXF And IsPullcoreDxfQuote(quoteName) Then

        Dim pcDetectedDeg As Double
        Dim pcGotAngle As Boolean

        pcDetectedDeg = 0#
        pcGotAngle = False

        If CurrentPullcoreStraightenCadIndex > 0 Then
            pcGotAngle = TryComputePullcoreStraightenAngleFromMatchedCad(CurrentPullcoreStraightenCadIndex, pcDetectedDeg)
        End If

        If pcGotAngle = False Then
            pcGotAngle = TryComputePullcoreBestFitStraightenAngleFromNativePath(nativePath, pcDetectedDeg)
        End If

        If pcGotAngle = False And PULLCORE_FALLBACK_ANGLE_DEG <> 0# Then
            pcDetectedDeg = PULLCORE_FALLBACK_ANGLE_DEG
            pcGotAngle = True
            LogLine "PULLCORE: using fixed fallback angle " & _
                    Format(pcDetectedDeg, "0.00") & " deg because geometry detection failed."
        End If

        If pcGotAngle Then
            If Abs(pcDetectedDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                CurrentDxfStraightenAngleRad = PULLCORE_STRAIGHTEN_SIGN * pcDetectedDeg * PI_VALUE / 180#

                LogLine "PULLCORE geometry straighten angle detected: " & _
                        Format(pcDetectedDeg, "0.00") & " deg"

                LogLine "PULLCORE DXF correction angle: " & _
                        Format(RadToDeg(CurrentDxfStraightenAngleRad), "0.00") & " deg"
            Else
                CurrentDxfStraightenAngleRad = 0#
                LogLine "PULLCORE geometry angle " & Format(pcDetectedDeg, "0.00") & _
                        " deg is below threshold; DXF left unrotated."
            End If
        Else
            CurrentDxfStraightenAngleRad = 0#
            LogLine "PULLCORE: no straightening angle available; DXF left unrotated."
        End If

    End If

    Dim swDraw As Object
    Set swDraw = swApp.NewDocument(SW_DRAWING_TEMPLATE_PATH, 0, _
                                   E_SHEET_WIDTH_IN / INCHES_PER_METER, _
                                   E_SHEET_HEIGHT_IN / INCHES_PER_METER)

    If swDraw Is Nothing Then
        LogLine "DXF skipped. Could not create drawing."
        GoTo CleanExit
    End If

    drawTitle = swDraw.GetTitle

    Dim errs As Long
    swApp.ActivateDoc3 drawTitle, False, 0, errs
    EnsureSwHidden

    SetupDrawingAsESize swDraw

    If FREEZE_DXF_DRAWING_GRAPHICS Then
        FreezeDxfDrawingGraphics swDraw
        freezeApplied = True
    End If

    Dim centerX As Double
    Dim centerY As Double

    centerX = E_SHEET_WIDTH_IN / 2#
    centerY = E_SHEET_HEIGHT_IN / 2#

    Dim projectedXOffset As Double
    Dim projectedYOffset As Double

    projectedXOffset = ((partL / 2#) + DXF_PROJECTED_VIEW_GAP_IN + (partT / 2#)) * scaleVal
    projectedYOffset = ((partW / 2#) + DXF_PROJECTED_VIEW_GAP_IN + (partT / 2#)) * scaleVal

    If IsHoldersDxfQuote(quoteName) Then
        projectedXOffset = projectedXOffset + (HOLDERS_SIDE_VIEW_EXTRA_GAP_IN * scaleVal)
        LogLine "HOLDERS DXF: spreading left/right views by extra " & _
                FormatNumberForCsv(HOLDERS_SIDE_VIEW_EXTRA_GAP_IN) & " in."
    End If

    Dim xLeft As Double
    Dim yLeft As Double
    Dim xRight As Double
    Dim yRight As Double
    Dim xTop As Double
    Dim yTop As Double
    Dim xBottom As Double
    Dim yBottom As Double

    xLeft = centerX - projectedXOffset
    yLeft = centerY

    xRight = centerX + projectedXOffset
    yRight = centerY

    xTop = centerX
    yTop = centerY + projectedYOffset

    xBottom = centerX
    yBottom = centerY - projectedYOffset

    If xLeft < DXF_MARGIN_IN Then xLeft = DXF_MARGIN_IN
    If xRight > E_SHEET_WIDTH_IN - DXF_MARGIN_IN Then xRight = E_SHEET_WIDTH_IN - DXF_MARGIN_IN
    If yTop > E_SHEET_HEIGHT_IN - DXF_MARGIN_IN Then yTop = E_SHEET_HEIGHT_IN - DXF_MARGIN_IN
    If yBottom < DXF_MARGIN_IN Then yBottom = DXF_MARGIN_IN

    Dim parentView As Object
    Dim viewLeft As Object
    Dim viewRight As Object
    Dim viewTop As Object
    Dim viewBottom As Object

    Set parentView = CreateParentDrawingView(swDraw, nativePath, parentPrimaryView, parentFallbackView, centerX, centerY, scaleVal)

    If parentView Is Nothing Then
        LogLine "DXF skipped. Could not create parent drawing view."
        GoTo CleanExit
    End If

    Dim rotateJBlockParentBeforeProjection As Boolean
    rotateJBlockParentBeforeProjection = (ROTATE_J_BLOCK_DXF_180 And IsJBlockDxfQuote(quoteName))

    Dim parentPreProjectionAngleRad As Double
    Dim qKey As String

    parentPreProjectionAngleRad = 0#
    qKey = NormalizeKey(quoteName)

    If rotateJBlockParentBeforeProjection Then
        parentPreProjectionAngleRad = parentPreProjectionAngleRad + PI_VALUE
        LogLine "J BLOCK: rotating center/base view 180 BEFORE projected views are created."
    End If

    If qKey = "IDHOLDER" Then
        Dim idHolderFlip As Boolean
        If CurrentIdHolderDxfFromAssembly Then
            idHolderFlip = FLIP_ID_HOLDER_CENTER_VIEW_180_FROM_ASSEMBLY
        Else
            idHolderFlip = FLIP_ID_HOLDER_CENTER_VIEW_180
        End If

        If idHolderFlip Then
            parentPreProjectionAngleRad = parentPreProjectionAngleRad + PI_VALUE
            LogLine "ID HOLDER: rotating center/base view 180 BEFORE projected views are created. (fromAssembly=" & CStr(CurrentIdHolderDxfFromAssembly) & ")"
        Else
            LogLine "ID HOLDER: center/base view used as-is BEFORE projected views are created. (fromAssembly=" & CStr(CurrentIdHolderDxfFromAssembly) & ")"
        End If
    End If

    If qKey = "ODHOLDER" Then
        If Abs(OD_HOLDER_CENTER_ROTATION_DEG) > 0.000001 Then
            parentPreProjectionAngleRad = parentPreProjectionAngleRad + DegToRad(OD_HOLDER_CENTER_ROTATION_DEG)

            LogLine "OD HOLDER: rotating center/base view " & _
                    Format(OD_HOLDER_CENTER_ROTATION_DEG, "0.00") & _
                    " deg BEFORE projected views are created."
        Else
            LogLine "OD HOLDER: center/base view used as-is BEFORE projected views are created."
        End If
    End If

    If Abs(CurrentDxfStraightenAngleRad) > 0.000001 Then
        parentPreProjectionAngleRad = parentPreProjectionAngleRad + CurrentDxfStraightenAngleRad

        LogLine "PULLCORE: applying center/base view straightening BEFORE projected views are created: " & _
                Format(RadToDeg(CurrentDxfStraightenAngleRad), "0.00") & " deg"
    End If

    If Abs(parentPreProjectionAngleRad) > 0.000001 Then
        RotateDrawingViewByAngle parentView, parentPreProjectionAngleRad

        swDraw.ClearSelection2 True
        swDraw.ForceRebuild3 False
        swDraw.GraphicsRedraw2
    End If

    Set viewLeft = Nothing
    Set viewRight = Nothing
    Set viewTop = Nothing
    Set viewBottom = Nothing

    If pullcoreAux Then
        LogLine "PULLCORE: building 4 edge auxiliary views around the already-oriented front view."
        AddFourEdgeAuxViews swDraw, parentView, scaleVal
    Else
        Set viewLeft = CreateProjectedDrawingView(swDraw, parentView, xLeft, yLeft, scaleVal, "PROJECTED LEFT")
        Set viewRight = CreateProjectedDrawingView(swDraw, parentView, xRight, yRight, scaleVal, "PROJECTED RIGHT")
        Set viewTop = CreateProjectedDrawingView(swDraw, parentView, xTop, yTop, scaleVal, "PROJECTED TOP")
        Set viewBottom = CreateProjectedDrawingView(swDraw, parentView, xBottom, yBottom, scaleVal, "PROJECTED BOTTOM")
    End If

    If allSolid Then
        ForceAllDrawingViewsSolid swDraw
    ElseIf allWireframe Then
        ForceAllDrawingViewsWireframe swDraw
    ElseIf jBlockMixedDisplay Then
        SetDrawingViewSolid parentView
        SetDrawingViewSolid viewLeft
        SetDrawingViewSolid viewRight
        SetDrawingViewWireframe viewTop
        SetDrawingViewWireframe viewBottom
    Else
        ForceAllDrawingViewsWireframe swDraw
    End If

    If dimensionedJBlock Then
        AddBestEffortJBlockDimensions swDraw, parentView, viewTop, viewBottom, viewLeft, viewRight, partL, partW, partT, scaleVal
    End If

    If allSolid Then ForceAllDrawingViewsSolid swDraw
    If allWireframe Then ForceAllDrawingViewsWireframe swDraw

    If IsDxfInPrintsFolder(dxfPath) Then

        If IsClampingPlateDxfQuote(quoteName) Then
            ForceTcpBcpClampingPlateViewModes parentView, viewLeft, viewRight, viewTop, viewBottom
            LogLine "TCP/BCP DXF: center/right/top forced wireframe; left/bottom forced solid."
        Else
            ForceAllDrawingViewsWireframe swDraw
        End If

    End If

    If IsMainBaseDxfQuote(quoteName) Then
        ForceAllDrawingViewsSolid swDraw
    End If

    If IsHoldersDxfQuote(quoteName) Then
        ForceAllDrawingViewsSolid swDraw
        LogLine "HOLDERS DXF: all views forced SOLID."
    End If

    If IsJBlockDxfQuote(quoteName) Then
        ForceJBlockLeftRightViewsSolid viewLeft, viewRight
        LogLine "J BLOCK DXF: left/right projected views forced SOLID."
    End If

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    Dim saveErrs As Long
    Dim saveWarns As Long

    LogLine "Saving DXF: " & dxfPath
    EnsureSwHidden

    swDraw.Extension.SaveAs3 dxfPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, Nothing, saveErrs, saveWarns

    LogLine "DXF save done. Errors=" & saveErrs & " Warnings=" & saveWarns

CleanExit:
    On Error Resume Next

    If freezeApplied Then
        UnfreezeDxfDrawingGraphics
        freezeApplied = False
    End If

    If drawTitle <> "" Then swApp.CloseDoc drawTitle
    Exit Sub

ErrHandler:
    LogLine "CreateProjectedDxfFromNativePath error for " & quoteName & ": " & Err.Description
    Resume CleanExit
End Sub

' ============================================================
' END OF PART 3
' Paste Part 4 immediately after this.
' ============================================================
Private Sub FreezeDxfDrawingGraphics(Optional ByVal docToFreeze As Object = Nothing)
On Error Resume Next

    Set DxfFreezeDoc = docToFreeze

    swApp.UserControl = False

    If Not DxfFreezeDoc Is Nothing Then DxfFreezeDoc.FeatureManager.EnableFeatureTree = False
    If Not swModel Is Nothing Then swModel.FeatureManager.EnableFeatureTree = False

    DoEvents
End Sub

Private Sub UnfreezeDxfDrawingGraphics()
On Error Resume Next

    If Not DxfFreezeDoc Is Nothing Then DxfFreezeDoc.FeatureManager.EnableFeatureTree = True
    If Not swModel Is Nothing Then swModel.FeatureManager.EnableFeatureTree = True

    swApp.UserControl = True

    Set DxfFreezeDoc = Nothing

    DoEvents
End Sub

Private Function CalculateProjectedFourViewDxfScale(ByVal partL As Double, _
                                                    ByVal partW As Double, _
                                                    ByVal partT As Double) As Double

    Dim scaleVal As Double
    scaleVal = DXF_MAX_SCALE

    Dim usableW As Double
    Dim usableH As Double

    usableW = E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)
    usableH = E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)

    Dim layoutW As Double
    Dim layoutH As Double

    layoutW = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
    layoutH = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)

    If layoutW > 0 Then
        If usableW / layoutW < scaleVal Then scaleVal = usableW / layoutW
    End If

    If layoutH > 0 Then
        If usableH / layoutH < scaleVal Then scaleVal = usableH / layoutH
    End If

    If scaleVal <= 0 Then scaleVal = 1#
    If scaleVal > DXF_MAX_SCALE Then scaleVal = DXF_MAX_SCALE

    CalculateProjectedFourViewDxfScale = scaleVal
End Function

Private Function CreateParentDrawingView(ByVal swDraw As Object, _
                                         ByVal nativePath As String, _
                                         ByVal primaryViewName As String, _
                                         ByVal fallbackViewName As String, _
                                         ByVal xIn As Double, _
                                         ByVal yIn As Double, _
                                         ByVal scaleVal As Double) As Object
On Error GoTo ErrHandler

    Set CreateParentDrawingView = Nothing

    If swDraw Is Nothing Then Exit Function

    Dim swView As Object
    Set swView = Nothing

    On Error Resume Next

    Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, primaryViewName, _
                                                      xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)

    If swView Is Nothing Then
        Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, fallbackViewName, _
                                                          xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    End If

    If swView Is Nothing Then
        Set swView = swDraw.CreateDrawViewFromModelView3(nativePath, "*Top", _
                                                          xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#)
    End If

    On Error GoTo ErrHandler

    If swView Is Nothing Then
        LogLine "Could not create parent drawing view."
        Exit Function
    End If

    SetDrawingViewScale swView, scaleVal

    Set CreateParentDrawingView = swView
    Exit Function

ErrHandler:
    LogLine "CreateParentDrawingView error: " & Err.Description
    Set CreateParentDrawingView = Nothing
End Function

Private Function CreateProjectedDrawingView(ByVal swDraw As Object, _
                                            ByVal parentView As Object, _
                                            ByVal xIn As Double, _
                                            ByVal yIn As Double, _
                                            ByVal scaleVal As Double, _
                                            ByVal labelText As String) As Object
On Error GoTo ErrHandler

    Set CreateProjectedDrawingView = Nothing

    If swDraw Is Nothing Then Exit Function
    If parentView Is Nothing Then Exit Function

    swDraw.ClearSelection2 True

    If SelectDrawingView(swDraw, parentView) = False Then
        LogLine "Could not select parent view for " & labelText
        Exit Function
    End If

    Dim projView As Object
    Set projView = Nothing

    On Error Resume Next
    Set projView = swDraw.CreateUnfoldedViewAt3(xIn / INCHES_PER_METER, yIn / INCHES_PER_METER, 0#, False)
    On Error GoTo ErrHandler

    swDraw.ClearSelection2 True

    If projView Is Nothing Then
        LogLine "Could not create projected view: " & labelText
        Exit Function
    End If

    SetDrawingViewScale projView, scaleVal

    Set CreateProjectedDrawingView = projView
    Exit Function

ErrHandler:
    LogLine "CreateProjectedDrawingView error for " & labelText & ": " & Err.Description
    On Error Resume Next
    swDraw.ClearSelection2 True
    Set CreateProjectedDrawingView = Nothing
End Function

Private Sub AddFourEdgeAuxViews(ByVal swDraw As Object, ByVal parentView As Object, ByVal scaleVal As Double)
On Error GoTo ErrHandler

    If swDraw Is Nothing Or parentView Is Nothing Then Exit Sub

    Dim vOut As Variant
    vOut = parentView.GetOutline

    If Not IsArray(vOut) Then Exit Sub
    If UBound(vOut) < 3 Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double

    xmin = CDbl(vOut(0))
    yMin = CDbl(vOut(1))
    xmax = CDbl(vOut(2))
    yMax = CDbl(vOut(3))

    Dim mx As Double
    Dim my As Double
    Dim W As Double
    Dim h As Double

    mx = (xmin + xmax) / 2#
    my = (yMin + yMax) / 2#
    W = xmax - xmin
    h = yMax - yMin

    Dim dx As Double
    Dim dy As Double

    dx = W * 1.4 + (1.5 / INCHES_PER_METER)
    dy = h * 1.4 + (1.5 / INCHES_PER_METER)

    TryMakeEdgeAuxView swDraw, parentView, scaleVal, xmin, my, True, mx - dx, my + dy, "VIEW B"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, xmax, my, True, mx + dx, my + dy, "VIEW C"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, mx, yMin, False, mx - dx, my - dy, "VIEW A"
    TryMakeEdgeAuxView swDraw, parentView, scaleVal, mx, yMax, False, mx + dx, my - dy, "VIEW D"

    Exit Sub

ErrHandler:
    LogLine "AddFourEdgeAuxViews error: " & Err.Description
End Sub

Private Sub TryMakeEdgeAuxView(ByVal swDraw As Object, ByVal parentView As Object, ByVal scaleVal As Double, _
                               ByVal selX As Double, ByVal selY As Double, ByVal edgeVertical As Boolean, _
                               ByVal placeX As Double, ByVal placeY As Double, ByVal label As String)
On Error GoTo ErrHandler

    swDraw.ClearSelection2 True

    If TrySelectViewSilhouetteEdge(swDraw, selX, selY, edgeVertical, False) = False Then
        LogLine "Aux " & label & ": could not grab an edge to project from."
        swDraw.ClearSelection2 True
        Exit Sub
    End If

    Dim auxV As Object
    Set auxV = Nothing

    On Error Resume Next
    Set auxV = swDraw.CreateAuxiliaryViewAt2(placeX, placeY, 0#, label)
    On Error GoTo ErrHandler

    swDraw.ClearSelection2 True

    If auxV Is Nothing Then
        LogLine "Aux " & label & ": CreateAuxiliaryViewAt2 returned nothing."
    Else
        SetDrawingViewScale auxV, scaleVal
        LogLine "Aux " & label & " created."
    End If

    Exit Sub

ErrHandler:
    LogLine "TryMakeEdgeAuxView error (" & label & "): " & Err.Description
    On Error Resume Next
    swDraw.ClearSelection2 True
End Sub

Private Function SelectDrawingView(ByVal swDraw As Object, ByVal swView As Object) As Boolean
On Error GoTo ErrHandler

    SelectDrawingView = False

    If swDraw Is Nothing Then Exit Function
    If swView Is Nothing Then Exit Function

    swDraw.ClearSelection2 True

    On Error Resume Next
    SelectDrawingView = CBool(swView.SelectEntity(False))
    On Error GoTo ErrHandler

    If SelectDrawingView Then Exit Function

    Dim viewName As String
    viewName = ""

    On Error Resume Next
    viewName = CStr(swView.GetName2)
    On Error GoTo ErrHandler

    If viewName <> "" Then
        Dim ok As Boolean

        On Error Resume Next
        ok = CBool(swDraw.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0))
        On Error GoTo ErrHandler

        SelectDrawingView = ok
        Exit Function
    End If

    SelectDrawingView = False
    Exit Function

ErrHandler:
    SelectDrawingView = False
End Function

Private Sub RotateDrawingViewByAngle(ByVal swView As Object, ByVal angleRad As Double)
On Error Resume Next

    If swView Is Nothing Then Exit Sub
    If Abs(angleRad) < 0.000001 Then Exit Sub

    swView.Angle = CDbl(swView.Angle) + angleRad
End Sub

Private Sub RotateDrawingView180(ByVal swView As Object)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    Dim a As Double
    a = CDbl(swView.Angle)

    swView.Angle = a + PI_VALUE
End Sub

Private Sub RotateJBlockProjectedViewSet180(ByVal swDraw As Object, _
                                            ByVal parentView As Object, _
                                            ByVal viewLeft As Object, _
                                            ByVal viewRight As Object, _
                                            ByVal viewTop As Object, _
                                            ByVal viewBottom As Object)
On Error Resume Next

    RotateDrawingView180 parentView
    RotateDrawingView180 viewLeft
    RotateDrawingView180 viewRight
    RotateDrawingView180 viewTop
    RotateDrawingView180 viewBottom

    SwapDrawingViewPositions viewLeft, viewRight
    SwapDrawingViewPositions viewTop, viewBottom

    If Not swDraw Is Nothing Then
        swDraw.ForceRebuild3 False
        swDraw.GraphicsRedraw2
    End If
End Sub

Private Sub SwapDrawingViewPositions(ByVal viewA As Object, ByVal viewB As Object)
On Error Resume Next

    If viewA Is Nothing Then Exit Sub
    If viewB Is Nothing Then Exit Sub

    Dim posA As Variant
    Dim posB As Variant

    posA = viewA.Position
    posB = viewB.Position

    If IsArray(posA) = False Then Exit Sub
    If IsArray(posB) = False Then Exit Sub
    If UBound(posA) < 1 Then Exit Sub
    If UBound(posB) < 1 Then Exit Sub

    Dim newA As Variant
    Dim newB As Variant

    newA = Array(CDbl(posB(0)), CDbl(posB(1)))
    newB = Array(CDbl(posA(0)), CDbl(posA(1)))

    viewA.Position = newA
    viewB.Position = newB
End Sub

' ============================================================
' ANGLE / DRAWING DISPLAY HELPERS
' ============================================================

Private Function Atan2Safe(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atan2Safe = Atn(y / x)
    ElseIf x < 0# And y >= 0# Then
        Atan2Safe = Atn(y / x) + PI_VALUE
    ElseIf x < 0# And y < 0# Then
        Atan2Safe = Atn(y / x) - PI_VALUE
    ElseIf x = 0# And y > 0# Then
        Atan2Safe = PI_VALUE / 2#
    ElseIf x = 0# And y < 0# Then
        Atan2Safe = -PI_VALUE / 2#
    Else
        Atan2Safe = 0#
    End If
End Function

Private Function RadToDeg(ByVal rad As Double) As Double
    RadToDeg = rad * 180# / PI_VALUE
End Function

Private Function DegToRad(ByVal deg As Double) As Double
    DegToRad = deg * PI_VALUE / 180#
End Function

Private Function NormalizeAngleToPlusMinus90(ByVal angleRad As Double) As Double
    Do While angleRad > PI_VALUE / 2#
        angleRad = angleRad - PI_VALUE
    Loop

    Do While angleRad < -PI_VALUE / 2#
        angleRad = angleRad + PI_VALUE
    Loop

    NormalizeAngleToPlusMinus90 = angleRad
End Function

Private Function TryGetComponentTopRotationRad(ByVal swComp As Object, ByRef angleRad As Double) As Boolean
On Error GoTo ErrHandler

    TryGetComponentTopRotationRad = False
    angleRad = 0#

    If swComp Is Nothing Then Exit Function

    Dim xform As Object
    Set xform = swComp.Transform2

    If xform Is Nothing Then Exit Function

    Dim v As Variant
    v = xform.ArrayData

    If IsEmpty(v) Then Exit Function
    If IsArray(v) = False Then Exit Function
    If UBound(v) < 8 Then Exit Function

    Dim xAxisX As Double
    Dim xAxisY As Double

    xAxisX = CDbl(v(0))
    xAxisY = CDbl(v(1))

    angleRad = Atan2Safe(xAxisY, xAxisX)
    angleRad = NormalizeAngleToPlusMinus90(angleRad)

    TryGetComponentTopRotationRad = True
    Exit Function

ErrHandler:
    TryGetComponentTopRotationRad = False
End Function

Private Sub RotateAllModelDrawingViews(ByVal swDraw As Object, ByVal angleRad As Double)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub
    If Abs(angleRad) < 0.000001 Then Exit Sub

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        v.Angle = CDbl(v.Angle) + angleRad
        Set v = v.GetNextView
    Loop

    swDraw.ForceRebuild3 False
    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceAllDrawingViewsWireframe(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        SetDrawingViewWireframe v
        Set v = v.GetNextView
    Loop

    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceAllDrawingViewsSolid(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        SetDrawingViewSolid v
        Set v = v.GetNextView
    Loop

    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceTcpBcpClampingPlateViewModes(ByVal parentView As Object, _
                                              ByVal viewLeft As Object, _
                                              ByVal viewRight As Object, _
                                              ByVal viewTop As Object, _
                                              ByVal viewBottom As Object)
On Error Resume Next

    SetDrawingViewWireframe parentView
    SetDrawingViewWireframe viewRight
    SetDrawingViewWireframe viewTop

    SetDrawingViewSolid viewLeft
    SetDrawingViewSolid viewBottom
End Sub

Private Sub ForceJBlockLeftRightViewsSolid(ByVal viewLeft As Object, _
                                           ByVal viewRight As Object)
On Error Resume Next

    SetDrawingViewSolid viewLeft
    SetDrawingViewSolid viewRight
End Sub

Private Sub SetupDrawingAsESize(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet

    If Not swSheet Is Nothing Then

        swSheet.SetSize 12, E_SHEET_WIDTH_IN / INCHES_PER_METER, E_SHEET_HEIGHT_IN / INCHES_PER_METER

        ' Force sheet scale 1:1 where supported.
        Err.Clear
        swSheet.SetScale 1#, 1#, True, True
        Err.Clear

    End If

    swDraw.GraphicsRedraw2
End Sub

Private Sub SetDrawingViewWireframe(ByVal swView As Object)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    swView.UseParentStyle = False
    swView.SetDisplayMode3 False, 1, False, True
    swView.DisplayMode = 1
End Sub

Private Sub SetDrawingViewSolid(ByVal swView As Object)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    swView.UseParentStyle = False
    swView.SetDisplayMode3 False, 2, False, True
    swView.DisplayMode = 2
End Sub

Private Sub SetDrawingViewScale(ByVal swView As Object, ByVal scaleVal As Double)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Or CurrentDxfForce1to1 Then
        scaleVal = 1#
    End If

    If scaleVal <= 0# Then scaleVal = 1#

    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal

    If Abs(scaleVal - 1#) < 0.000001 Then
        swView.ScaleRatio = "1:1"
    End If
End Sub

' ============================================================
' J BLOCK OVERALL DIMENSIONS
' ============================================================

Private Sub AddBestEffortJBlockDimensions(ByVal swDraw As Object, _
                                          ByVal parentView As Object, _
                                          ByVal viewTop As Object, _
                                          ByVal viewBottom As Object, _
                                          ByVal viewLeft As Object, _
                                          ByVal viewRight As Object, _
                                          ByVal partL As Double, _
                                          ByVal partW As Double, _
                                          ByVal partT As Double, _
                                          ByVal scaleVal As Double)
On Error GoTo ErrHandler

    If DIMENSION_J_BLOCK = False Then
        LogLine "J BLOCK dimensioning disabled by setting."
        Exit Sub
    End If

    If swDraw Is Nothing Then Exit Sub

    LogLine "Adding overall J BLOCK dimensions."

    AddViewOverallDimension swDraw, parentView, False
    AddViewOverallDimension swDraw, parentView, True

    If Not viewRight Is Nothing Then
        AddViewOverallDimension swDraw, viewRight, True
    ElseIf Not viewLeft Is Nothing Then
        AddViewOverallDimension swDraw, viewLeft, True
    ElseIf Not viewTop Is Nothing Then
        AddViewOverallDimension swDraw, viewTop, False
    End If

    AddJBlockDimensionNotes swDraw, parentView, partL, partW, partT

    swDraw.GraphicsRedraw2
    Exit Sub

ErrHandler:
    LogLine "AddBestEffortJBlockDimensions error: " & Err.Description
End Sub

Private Sub AddViewOverallDimension(ByVal swDraw As Object, _
                                    ByVal swView As Object, _
                                    ByVal horizontal As Boolean)
On Error GoTo ErrHandler

    If swDraw Is Nothing Then Exit Sub
    If swView Is Nothing Then Exit Sub

    Dim vOut As Variant
    vOut = swView.GetOutline

    If IsEmpty(vOut) Then Exit Sub
    If IsArray(vOut) = False Then Exit Sub
    If UBound(vOut) < 3 Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double
    Dim midx As Double
    Dim midy As Double
    Dim gap As Double

    xmin = CDbl(vOut(0))
    yMin = CDbl(vOut(1))
    xmax = CDbl(vOut(2))
    yMax = CDbl(vOut(3))

    midx = (xmin + xmax) / 2#
    midy = (yMin + yMax) / 2#

    gap = 0.625 / INCHES_PER_METER

    swDraw.ClearSelection2 True

    Dim ok1 As Boolean
    Dim ok2 As Boolean
    Dim dispDim As Object

    If horizontal Then

        ok1 = TrySelectViewSilhouetteEdge(swDraw, xmin, midy, True, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, xmax, midy, True, True)

        If ok1 And ok2 Then
            Set dispDim = swDraw.AddHorizontalDimension2(midx, yMin - gap, 0#)
        End If

    Else

        ok1 = TrySelectViewSilhouetteEdge(swDraw, midx, yMin, False, False)
        ok2 = TrySelectViewSilhouetteEdge(swDraw, midx, yMax, False, True)

        If ok1 And ok2 Then
            Set dispDim = swDraw.AddVerticalDimension2(xmax + gap, midy, 0#)
        End If

    End If

    swDraw.ClearSelection2 True
    Exit Sub

ErrHandler:
    On Error Resume Next
    swDraw.ClearSelection2 True
    LogLine "AddViewOverallDimension error: " & Err.Description
End Sub

Private Function TrySelectViewSilhouetteEdge(ByVal swDraw As Object, _
                                             ByVal x As Double, _
                                             ByVal y As Double, _
                                             ByVal edgeIsVertical As Boolean, _
                                             ByVal appendToSelection As Boolean) As Boolean
On Error GoTo ErrHandler

    TrySelectViewSilhouetteEdge = False

    Dim nudges(0 To 4) As Double

    nudges(0) = 0#
    nudges(1) = 0.015 / INCHES_PER_METER
    nudges(2) = -0.015 / INCHES_PER_METER
    nudges(3) = 0.04 / INCHES_PER_METER
    nudges(4) = -0.04 / INCHES_PER_METER

    Dim i As Long
    Dim tx As Double
    Dim ty As Double
    Dim okSel As Boolean

    For i = 0 To 4

        tx = x
        ty = y

        If edgeIsVertical Then
            tx = x + nudges(i)
        Else
            ty = y + nudges(i)
        End If

        okSel = CBool(swDraw.Extension.SelectByID2("", "EDGE", tx, ty, 0#, appendToSelection, 0, Nothing, 0))

        If okSel Then
            TrySelectViewSilhouetteEdge = True
            Exit Function
        End If

    Next i

    Exit Function

ErrHandler:
    TrySelectViewSilhouetteEdge = False
End Function

Private Sub TryInsertModelAnnotations(ByVal swView As Object)
    ' Reserved. Imported solids usually carry no model dimensions.
End Sub

Private Sub AddJBlockDimensionNotes(ByVal swDraw As Object, ByVal swView As Object, _
                                    ByVal partL As Double, ByVal partW As Double, ByVal partT As Double)
On Error GoTo ErrHandler

    If swDraw Is Nothing Then Exit Sub

    Dim xmin As Double
    Dim yMin As Double
    Dim xmax As Double
    Dim yMax As Double

    xmin = E_SHEET_WIDTH_IN / 2# / INCHES_PER_METER
    yMin = E_SHEET_HEIGHT_IN / 2# / INCHES_PER_METER
    xmax = xmin
    yMax = yMin

    If Not swView Is Nothing Then

        Dim vOut As Variant
        vOut = swView.GetOutline

        If IsArray(vOut) Then
            If UBound(vOut) >= 3 Then
                xmin = CDbl(vOut(0))
                yMin = CDbl(vOut(1))
                xmax = CDbl(vOut(2))
                yMax = CDbl(vOut(3))
            End If
        End If

    End If

    Dim midx As Double
    Dim midy As Double
    Dim gap As Double

    midx = (xmin + xmax) / 2#
    midy = (yMin + yMax) / 2#
    gap = 0.6 / INCHES_PER_METER

    PlaceDrawingNote swDraw, midx - gap, yMin - gap, "OVERALL LENGTH (L) = " & FormatJBlockDim(partL)
    PlaceDrawingNote swDraw, xmin - (2.4 / INCHES_PER_METER), midy, "OVERALL HEIGHT (W) = " & FormatJBlockDim(partW)
    PlaceDrawingNote swDraw, xmax + gap, yMax + gap, "THICKNESS (T) = " & FormatJBlockDim(partT)

    swDraw.ClearSelection2 True
    Exit Sub

ErrHandler:
    LogLine "AddJBlockDimensionNotes error: " & Err.Description
End Sub

Private Sub PlaceDrawingNote(ByVal swDraw As Object, ByVal xMeters As Double, _
                             ByVal yMeters As Double, ByVal text As String)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    swDraw.ClearSelection2 True

    Dim swNote As Object
    Set swNote = swDraw.InsertNote(text)

    If Not swNote Is Nothing Then

        Dim swAnn As Object
        Set swAnn = swNote.GetAnnotation

        If Not swAnn Is Nothing Then swAnn.SetPosition xMeters, yMeters, 0#

    End If

    swDraw.ClearSelection2 True
End Sub

Private Function FormatJBlockDim(ByVal v As Double) As String
    FormatJBlockDim = Format(v, "0.000")
End Function

' ============================================================
' DXF DIMENSION / QUOTE HELPERS
' ============================================================

Private Sub GetBestCurrentOutputDimensions(ByVal quoteName As String, _
                                           ByRef L As Double, _
                                           ByRef W As Double, _
                                           ByRef T As Double)
On Error Resume Next

    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            If ExportRows(i).HasCad Then
                L = parts(ExportRows(i).CadPartIndex).Length
                W = parts(ExportRows(i).CadPartIndex).Width
                T = parts(ExportRows(i).CadPartIndex).Thickness
                Exit Sub
            End If
        End If
    Next i

    For i = 1 To PullcoreMatchCount
        If NormalizeKey(PullcoreMatches(i).quoteName) = k Then
            If PullcoreMatches(i).CadPartIndex > 0 Then
                L = PullcoreMatches(i).FittedLength
                W = PullcoreMatches(i).FittedWidth
                T = PullcoreMatches(i).FittedThickness

                If L <= 0 Then L = PullcoreMatches(i).CadLength
                If W <= 0 Then W = PullcoreMatches(i).CadWidth
                If T <= 0 Then T = PullcoreMatches(i).CadThickness

                If L <= 0 Then L = parts(PullcoreMatches(i).CadPartIndex).Length
                If W <= 0 Then W = parts(PullcoreMatches(i).CadPartIndex).Width
                If T <= 0 Then T = parts(PullcoreMatches(i).CadPartIndex).Thickness

                Exit Sub
            End If
        End If
    Next i
End Sub

Private Function GetCombinedHoldersLengthForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim best As Double
    best = 0#

    If idIdx > 0 Then If parts(idIdx).Length > best Then best = parts(idIdx).Length
    If odIdx > 0 Then If parts(odIdx).Length > best Then best = parts(odIdx).Length

    GetCombinedHoldersLengthForDxf = best
End Function

Private Function GetCombinedHoldersWidthForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim total As Double
    total = 0#

    If idIdx > 0 Then total = total + parts(idIdx).Width
    If odIdx > 0 Then total = total + parts(odIdx).Width

    GetCombinedHoldersWidthForDxf = total
End Function

Private Function GetCombinedHoldersThicknessForDxf() As Double
    Dim idIdx As Long
    Dim odIdx As Long

    idIdx = FindHolderCadIndexForMainPackage("ID HOLDER")
    odIdx = FindHolderCadIndexForMainPackage("OD HOLDER")

    Dim best As Double
    best = 0#

    If idIdx > 0 Then If parts(idIdx).Thickness > best Then best = parts(idIdx).Thickness
    If odIdx > 0 Then If parts(odIdx).Thickness > best Then best = parts(odIdx).Thickness

    GetCombinedHoldersThicknessForDxf = best
End Function

Private Function IsDxfInPrintsFolder(ByVal dxfPath As String) As Boolean
    Dim p As String
    p = UCase(dxfPath)

    If InStr(p, "\" & UCase(CurrentJobNumber & OUTPUT_FOLDER_SUFFIX) & "\") > 0 Then
        If InStr(p, "PYROPEL") > 0 Then
            IsDxfInPrintsFolder = False
            Exit Function
        End If

        IsDxfInPrintsFolder = True
        Exit Function
    End If

    If InStr(p, "\PRINTS\") > 0 Then IsDxfInPrintsFolder = True
End Function

Private Function IsHoldersDxfQuote(ByVal quoteName As String) As Boolean
    IsHoldersDxfQuote = (UCase(Trim(quoteName)) = "HOLDERS")
End Function

Private Function IsMainBaseDxfQuote(ByVal quoteName As String) As Boolean
    IsMainBaseDxfQuote = (UCase(Trim(quoteName)) = "MAIN ASSEMBLY")
End Function

Private Function IsJBlockDxfQuote(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    IsJBlockDxfQuote = (InStr(s, "J BLOCK") > 0 Or InStr(s, "J-BLOCK") > 0 Or InStr(s, "JBLOCK") > 0)
End Function

Private Function IsClampingPlateDxfQuote(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "TCP", "BCP", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", _
             "TOPSMED", "BOTTOMSMED"
            IsClampingPlateDxfQuote = True
    End Select
End Function

Private Function IsPullcoreDxfQuote(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    If s = "PULLCORE" Then
        IsPullcoreDxfQuote = True
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then
        IsPullcoreDxfQuote = True
    End If
End Function

Private Function TryGetNativeModelDimsInches(ByVal nativePath As String, _
                                             ByRef L As Double, _
                                             ByRef W As Double, _
                                             ByRef T As Double) As Boolean
On Error GoTo ErrHandler

    TryGetNativeModelDimsInches = False

    L = 0#
    W = 0#
    T = 0#

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(nativePath) = False Then Exit Function

    Dim ext As String
    ext = LCase(fso.GetExtensionName(nativePath))

    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object

    If ext = "sldasm" Then
        Set mdl = swApp.OpenDoc6(nativePath, swDocASSEMBLY, swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, "", errs, warns)
    Else
        Set mdl = swApp.OpenDoc6(nativePath, swDocPART, swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, "", errs, warns)
    End If

    If mdl Is Nothing Then Exit Function

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim gotBox As Boolean

    gotBox = TryGetModelDocBoxDimsInches(mdl, dx, dy, dz)

    On Error Resume Next
    swApp.CloseDoc mdl.GetTitle
    On Error GoTo ErrHandler

    If gotBox = False Then Exit Function

    SortThreeDimensions dx, dy, dz, L, W, T

    L = Round(L, DIM_DECIMALS)
    W = Round(W, DIM_DECIMALS)
    T = Round(T, DIM_DECIMALS)

    TryGetNativeModelDimsInches = (L > 0 And W > 0 And T > 0)
    Exit Function

ErrHandler:
    TryGetNativeModelDimsInches = False
End Function

Private Function TryGetModelDocBoxDimsInches(ByVal mdl As Object, _
                                             ByRef dx As Double, _
                                             ByRef dy As Double, _
                                             ByRef dz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetModelDocBoxDimsInches = False

    If mdl Is Nothing Then Exit Function

    If mdl.GetType = swDocPART Then
        TryGetModelDocBoxDimsInches = GetPartBoundingBoxInches(mdl, dx, dy, dz)
        Exit Function
    End If

    Dim vBox As Variant

    On Error Resume Next
    vBox = mdl.GetBox(False, False)
    On Error GoTo ErrHandler

    If IsValidBoxArray(vBox) = False Then Exit Function

    dx = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * INCHES_PER_METER
    dy = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * INCHES_PER_METER
    dz = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * INCHES_PER_METER

    TryGetModelDocBoxDimsInches = True
    Exit Function

ErrHandler:
    TryGetModelDocBoxDimsInches = False
End Function

Private Function IsValidBoxArray(ByVal vBox As Variant) As Boolean
    If IsEmpty(vBox) Then Exit Function
    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    IsValidBoxArray = True
End Function

' ============================================================
' END OF PART 4
' Paste Part 5 immediately after this.
' ============================================================
' ============================================================
' BOM FIND / READ (EXCEL)
' ============================================================

Private Function FindCustomerBomFile(ByVal jobFolder As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(jobFolder) Then Exit Function

    Dim bestPath As String
    Dim bestScore As Long

    bestPath = ""
    bestScore = -1

    SearchBomFilesRecursive fso.GetFolder(jobFolder), bestPath, bestScore

    If bestPath <> "" Then
        LogLine "BOM file chosen (score " & bestScore & "): " & bestPath
    Else
        LogLine "No BOM/Excel/PDF file found anywhere under: " & jobFolder
        LogLine "  Files present at the job folder root:"

        Dim ff As Object
        For Each ff In fso.GetFolder(jobFolder).Files
            LogLine "    - " & ff.Name
        Next ff
    End If

    FindCustomerBomFile = bestPath
    Exit Function

ErrHandler:
    LogLine "FindCustomerBomFile error: " & Err.Description
    FindCustomerBomFile = ""
End Function

Private Sub SearchBomFilesRecursive(ByVal folder As Object, ByRef bestPath As String, ByRef bestScore As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderName As String
    folderName = UCase(folder.Name)

    If folderName = UCase(EXTRACT_FOLDER_NAME) Then Exit Sub

    Dim file As Object
    Dim ext As String
    Dim nameUpper As String
    Dim score As Long

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        nameUpper = UCase(file.Name)

        If ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "pdf" Then

            score = 0

            If InStr(nameUpper, "BOM") > 0 Then score = score + 100
            If InStr(nameUpper, "BILL OF MATERIAL") > 0 Then score = score + 100
            If InStr(nameUpper, "QUOTE") > 0 Then score = score + 40
            If InStr(nameUpper, "MATERIAL") > 0 Then score = score + 30
            If InStr(nameUpper, CurrentJobNumber) > 0 Then score = score + 20

            If ext = "xlsx" Then score = score + 8
            If ext = "xlsm" Then score = score + 7
            If ext = "xls" Then score = score + 6
            If ext = "pdf" Then score = score + 5

            If score > bestScore Then
                bestScore = score
                bestPath = file.path
            End If

        End If

    Next file

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        SearchBomFilesRecursive subFolder, bestPath, bestScore
    Next subFolder
End Sub

Private Sub ReadCustomerBom(ByVal bomPath As String)
On Error GoTo ErrHandler

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False

    Set xlWb = xlApp.Workbooks.Open(bomPath, False, True)

    On Error Resume Next
    xlApp.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    Dim usedThisSheet As Boolean
    usedThisSheet = False

    If TURBO_READ_ONLY_BOM_SHEET Then
        On Error Resume Next
        Set xlWs = xlWb.Worksheets(TURBO_BOM_SHEET_NAME)
        On Error GoTo ErrHandler

        If Not xlWs Is Nothing Then
            LogLine "TURBO: reading only sheet '" & TURBO_BOM_SHEET_NAME & "'"
            ReadBomWorksheetFastArray xlWs
            usedThisSheet = True
        End If
    End If

    If usedThisSheet = False Then

        Dim ws As Object

        For Each ws In xlWb.Worksheets
            If IsLikelyBomWorksheet(ws) And ShouldSkipBomWorksheet(ws) = False Then
                LogLine "Reading worksheet: " & ws.Name
                ReadBomWorksheetFastArray ws
                If BomCount > 0 Then Exit For
            End If
        Next ws

    End If

CleanExit:
    On Error Resume Next

    If Not xlWb Is Nothing Then xlWb.Close False

    If Not xlApp Is Nothing Then
        xlApp.Calculation = -4105
        xlApp.EnableEvents = True
        xlApp.ScreenUpdating = True
        xlApp.Quit
    End If

    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrHandler:
    LogLine "ReadCustomerBom error: " & Err.Description
    Resume CleanExit
End Sub

Private Sub ReadBomWorksheetFastArray(ByVal ws As Object)
On Error GoTo ErrHandler

    Dim usedRange As Object
    Set usedRange = ws.usedRange

    Dim rowCount As Long
    Dim colCount As Long

    rowCount = usedRange.Rows.count
    colCount = usedRange.Columns.count

    If rowCount = 0 Or colCount = 0 Then Exit Sub

    Dim data As Variant
    data = usedRange.Value

    If IsEmpty(data) Then Exit Sub
    If IsArray(data) = False Then Exit Sub

    Dim headerRow As Long
    Dim descCol As Long
    Dim qtyCol As Long
    Dim matCol As Long
    Dim lenCol As Long
    Dim widCol As Long
    Dim thkCol As Long
    Dim typeCol As Long

    headerRow = 0
    descCol = 0
    qtyCol = 0
    matCol = 0
    lenCol = 0
    widCol = 0
    thkCol = 0
    typeCol = 0

    Dim r As Long
    Dim maxHeaderRow As Long

    maxHeaderRow = rowCount
    If maxHeaderRow > BOM_HEADER_SEARCH_MAX_ROWS Then maxHeaderRow = BOM_HEADER_SEARCH_MAX_ROWS

    For r = 1 To maxHeaderRow
        descCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("DESCRIPTION", "DESC", "PART NAME", "ITEM DESCRIPTION", "COMPONENT"))
        If descCol > 0 Then
            headerRow = r
            Exit For
        End If
    Next r

    If headerRow = 0 Then
        LogLine "No header row found by description."
        Exit Sub
    End If

    qtyCol = FindBomQtyColumnInArrayRow(data, headerRow, colCount)
    matCol = FindBomMaterialColumnInArrayRow(data, headerRow, colCount)
    typeCol = FindBomHeaderLikeInArrayRow(data, headerRow, colCount, Array("TYPE", "SOURCE", "MAKE/BUY", "MAKE BUY", "PROCESS"))

    FindBomDimensionColumnsInArrayRow data, headerRow, colCount, lenCol, widCol, thkCol

    LogLine "BOM header row=" & headerRow & " descCol=" & descCol & " qtyCol=" & qtyCol & _
            " matCol=" & matCol & " typeCol=" & typeCol & _
            " L=" & lenCol & " W=" & widCol & " T=" & thkCol

    Dim blanks As Long
    blanks = 0

    For r = headerRow + 1 To rowCount

        Dim desc As String
        Dim typeText As String
        Dim matText As String
        Dim qtyText As String

        desc = GetArrayValue(data, r, descCol)
        typeText = GetArrayValue(data, r, typeCol)
        matText = GetArrayValue(data, r, matCol)
        qtyText = GetArrayValue(data, r, qtyCol)

        If Trim(desc) = "" Then

            blanks = blanks + 1
            If blanks >= STOP_BOM_READ_AFTER_BLANK_ROWS Then Exit For

        Else

            blanks = 0

            Dim lenVal As Double
            Dim widVal As Double
            Dim thkVal As Double
            Dim hasDims As Boolean

            lenVal = 0#
            widVal = 0#
            thkVal = 0#

            If lenCol > 0 Then lenVal = val(GetArrayValue(data, r, lenCol))
            If widCol > 0 Then widVal = val(GetArrayValue(data, r, widCol))
            If thkCol > 0 Then thkVal = val(GetArrayValue(data, r, thkCol))

            hasDims = (lenVal > 0 And widVal > 0 And thkVal > 0)

            If ShouldUseBomItem(desc, typeText, matText, hasDims) Then

                Dim qtyVal As Long
                qtyVal = CLng(val(qtyText))

                If qtyVal <= 0 Then qtyVal = 1

                AddBomRow desc, typeText, matText, qtyVal, lenVal, widVal, thkVal, hasDims

            End If

        End If

    Next r

    Exit Sub

ErrHandler:
    LogLine "ReadBomWorksheetFastArray error: " & Err.Description
End Sub

Private Function ShouldUseBomItem(ByVal desc As String, _
                                  ByVal typeText As String, _
                                  ByVal materialText As String, _
                                  ByVal hasDims As Boolean) As Boolean

    Dim d As String
    d = NormalizeText(desc)

    If d = "" Then Exit Function
    If IsHardwareName(d) Then Exit Function

    If InStr(d, "PYROPEL") > 0 Or InStr(NormalizeText(materialText), "PYROPEL") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "EJ J-BLOCK") > 0 Or InStr(d, "EJECTOR J-BLOCK") > 0 Or _
       InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then
        ShouldUseBomItem = True
        Exit Function
    End If

    If InStr(d, "HOLDER") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "SMED") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "TCP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "BCP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "INSERT") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "INS") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "CAM") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "KEY") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "PLATE") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "STOP") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "COVER") > 0 Then ShouldUseBomItem = True: Exit Function
    If InStr(d, "POT") > 0 Then ShouldUseBomItem = True: Exit Function

    If hasDims Then ShouldUseBomItem = True: Exit Function

    ShouldUseBomItem = False
End Function

Private Sub AddBomRow(ByVal desc As String, _
                      ByVal typeText As String, _
                      ByVal matText As String, _
                      ByVal qty As Long, _
                      ByVal lenVal As Double, _
                      ByVal widVal As Double, _
                      ByVal thkVal As Double, _
                      ByVal hasDims As Boolean)

    Dim L As Double
    Dim W As Double
    Dim T As Double

    If hasDims Then
        SortThreeDimensions lenVal, widVal, thkVal, L, W, T
    Else
        L = 0#
        W = 0#
        T = 0#
    End If

    Dim isPyropel As Boolean
    isPyropel = (InStr(NormalizeText(desc), "PYROPEL") > 0 Or InStr(NormalizeText(matText), "PYROPEL") > 0)

    If HIDE_QUARTER_INCH_THICKNESS And hasDims And isPyropel = False Then
        If IsQuarterInchThickness(T) Then
            If IsInsertQuoteName(StandardPlateName(desc)) = False Then
                LogLine "BOM .250 item hidden: " & desc
                Exit Sub
            End If
        End If
    End If

    BomCount = BomCount + 1
    ReDim Preserve BomRows(1 To BomCount)

    BomRows(BomCount).Description = Trim(desc)
    BomRows(BomCount).quoteName = StandardPlateName(desc)
    BomRows(BomCount).TypeField = Trim(typeText)
    BomRows(BomCount).Quantity = qty
    BomRows(BomCount).material = NormalizeSteelType(matText)

    BomRows(BomCount).BomLength = Round(L, DIM_DECIMALS)
    BomRows(BomCount).BomWidth = Round(W, DIM_DECIMALS)
    BomRows(BomCount).BomThickness = Round(T, DIM_DECIMALS)
    BomRows(BomCount).hasDims = hasDims
End Sub

' ============================================================
' BOM READ (PDF via Poppler pdftotext)
' ============================================================

Private Sub ReadCustomerBomPdfUsingPdfToText(ByVal pdfPath As String)
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(PDFTOTEXT_EXE) = False Then
        LogErrorText "pdftotext.exe not found at: " & PDFTOTEXT_EXE
        Exit Sub
    End If

    If fso.FileExists(pdfPath) = False Then
        LogErrorText "PDF not found: " & pdfPath
        Exit Sub
    End If

    Dim txtPath As String
    txtPath = CurrentJobFolder & "\XT_Export_BOM_PDF_Text.txt"

    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = Chr(34) & PDFTOTEXT_EXE & Chr(34) & " -layout " & _
          Chr(34) & pdfPath & Chr(34) & " " & Chr(34) & txtPath & Chr(34)

    LogLine "Running pdftotext: " & cmd
    sh.Run cmd, 0, True

    If fso.FileExists(txtPath) = False Then
        LogErrorText "pdftotext produced no output file."
        Exit Sub
    End If

    Dim allText As String
    allText = ReadAllTextFile(txtPath)

    If Trim(allText) = "" Then
        LogErrorText "pdftotext output is empty."
        Exit Sub
    End If

    ParseBomTextFromPdf allText
    Exit Sub

ErrHandler:
    LogLine "ReadCustomerBomPdfUsingPdfToText error: " & Err.Description
End Sub

Private Sub ParseBomTextFromPdf(ByVal allText As String)
On Error GoTo ErrHandler

    Dim lines() As String

    allText = Replace(allText, vbCr, vbLf)

    Do While InStr(allText, vbLf & vbLf) > 0
        allText = Replace(allText, vbLf & vbLf, vbLf)
    Loop

    lines = Split(allText, vbLf)

    Dim i As Long

    For i = LBound(lines) To UBound(lines)
        ParseBomPdfTextLine lines(i)
    Next i

    LogLine "=== Parsed " & BomCount & " BOM rows from customer BOM ==="

    For i = 1 To BomCount
        LogLine "  BOMROW: '" & BomRows(i).Description & "' mat=" & BomRows(i).material & _
                " L/W/T=" & FormatNumberForCsv(BomRows(i).BomLength) & "/" & _
                FormatNumberForCsv(BomRows(i).BomWidth) & "/" & _
                FormatNumberForCsv(BomRows(i).BomThickness) & _
                " pullcore=" & CStr(IsPullcoreBomRow(BomRows(i)))
    Next i

    Exit Sub

ErrHandler:
    LogLine "ParseBomTextFromPdf error: " & Err.Description
End Sub

Private Sub ParseBomPdfTextLine(ByVal rawLine As String)
On Error GoTo ErrHandler

    Dim line As String
    line = Trim(rawLine)

    If line = "" Then Exit Sub

    Dim upperLine As String
    upperLine = UCase(line)

    If InStr(upperLine, "LTH") > 0 And InStr(upperLine, "WTH") > 0 Then Exit Sub
    If InStr(upperLine, "DESCRIPTION") > 0 And InStr(upperLine, "QTY") > 0 Then Exit Sub

    If TryParseTempcraftBasePdfMaterialLine(line) = False Then Exit Sub

    Exit Sub

ErrHandler:
    LogLine "ParseBomPdfTextLine error: " & Err.Description
End Sub

Private Function TryParseTempcraftBasePdfMaterialLine(ByVal line As String) As Boolean
On Error GoTo ErrHandler

    TryParseTempcraftBasePdfMaterialLine = False

    Dim upperLine As String
    upperLine = UCase(line)

    Dim desc As String
    Dim qty As Long
    Dim mat As String

    desc = ExtractTempcraftPdfDescription(line)
    If desc = "" Then Exit Function

    qty = ExtractTempcraftPdfQty(upperLine)
    mat = ExtractTempcraftPdfMaterial(upperLine)

    Dim nums() As Double
    Dim numCount As Long

    ExtractDecimalNumbers line, nums, numCount

    Dim lenVal As Double
    Dim widVal As Double
    Dim thkVal As Double
    Dim hasDims As Boolean

    lenVal = 0#
    widVal = 0#
    thkVal = 0#
    hasDims = False

    If numCount >= 3 Then
        lenVal = nums(0)
        widVal = nums(1)
        thkVal = nums(2)
        hasDims = (lenVal > 0 And widVal > 0 And thkVal > 0)
    End If

    If qty <= 0 Then qty = 1
    If mat = "" Then mat = DEFAULT_STEEL_TYPE

    If ShouldUseBomItem(desc, "Outsource", mat, hasDims) Then
        AddBomRow desc, "Outsource", mat, qty, lenVal, widVal, thkVal, hasDims
        TryParseTempcraftBasePdfMaterialLine = True
    End If

    Exit Function

ErrHandler:
    LogLine "TryParseTempcraftBasePdfMaterialLine error: " & Err.Description
    TryParseTempcraftBasePdfMaterialLine = False
End Function

Private Function ExtractTempcraftPdfDescription(ByVal line As String) As String
On Error GoTo ErrHandler

    Dim work As String
    work = RemoveLeadingItemNumber(line)

    Dim cutPos As Long
    cutPos = 0

    Dim upperWork As String
    upperWork = UCase(work)

    Dim kw As Variant
    Dim p As Long

    For Each kw In Array("OUTSOURCE", "PURCHASE", "MAKE", "BUY", "STOCK")
        p = InStr(upperWork, CStr(kw))
        If p > 0 Then
            If cutPos = 0 Or p < cutPos Then cutPos = p
        End If
    Next kw

    If cutPos > 1 Then
        ExtractTempcraftPdfDescription = Trim(Left(work, cutPos - 1))
    Else
        Dim i As Long

        For i = 1 To Len(work)
            If IsNumeric(Mid(work, i, 1)) Then
                cutPos = i
                Exit For
            End If
        Next i

        If cutPos > 1 Then
            ExtractTempcraftPdfDescription = Trim(Left(work, cutPos - 1))
        Else
            ExtractTempcraftPdfDescription = Trim(work)
        End If
    End If

    Exit Function

ErrHandler:
    ExtractTempcraftPdfDescription = ""
End Function

Private Function ExtractTempcraftPdfQty(ByVal upperLine As String) As Long
On Error GoTo ErrHandler
    ExtractTempcraftPdfQty = ExtractQtyAfterOutsource(upperLine)
    Exit Function
ErrHandler:
    ExtractTempcraftPdfQty = 1
End Function

Private Function ExtractTempcraftPdfMaterial(ByVal upperLine As String) As String

    If InStr(upperLine, "4140") > 0 Then ExtractTempcraftPdfMaterial = "4140": Exit Function
    If InStr(upperLine, "A-2") > 0 Or InStr(upperLine, "A2") > 0 Then ExtractTempcraftPdfMaterial = "A-2": Exit Function
    If InStr(upperLine, "H-13") > 0 Or InStr(upperLine, "H13") > 0 Then ExtractTempcraftPdfMaterial = "H13": Exit Function
    If InStr(upperLine, "D-2") > 0 Or InStr(upperLine, "D2") > 0 Then ExtractTempcraftPdfMaterial = "D-2": Exit Function
    If InStr(upperLine, "DRILL ROD") > 0 Then ExtractTempcraftPdfMaterial = "Drill Rod": Exit Function
    If InStr(upperLine, "PYROPEL") > 0 Then ExtractTempcraftPdfMaterial = "Pyropel": Exit Function

    ExtractTempcraftPdfMaterial = ""
End Function

' ============================================================
' MATCH BOM TO CAD
' ============================================================

Private Sub BuildExportRowsFromBom()
On Error GoTo ErrHandler

    Dim i As Long

    Dim pullcoreRowCount As Long
    pullcoreRowCount = 0

    For i = 1 To BomCount
        If IsPullcoreBomRow(BomRows(i)) Then pullcoreRowCount = pullcoreRowCount + 1
    Next i

    LogLine "BOM rows total=" & BomCount & "  pullcore rows=" & pullcoreRowCount

    SetPackageTargetsFromBom

    MatchPullcoreRowsByClass

    For i = 1 To BomCount

        If IsPullcoreBomRow(BomRows(i)) = False Then

            Dim b As BomInfo
            b = BomRows(i)

            Dim cadIdx As Long
            cadIdx = FindBestCadMatchForBom(b)

            AddExportRow b, cadIdx

        End If

    Next i

    LogLine "After matching: ExportRows=" & ExportCount & "  PullcoreMatches=" & PullcoreMatchCount

    Exit Sub

ErrHandler:
    LogLine "BuildExportRowsFromBom error: " & Err.Description
End Sub

Private Sub SetPackageTargetsFromBom()
On Error GoTo ErrHandler

    JBlockTgtL = J_BLOCK_TARGET_LENGTH
    JBlockTgtW = J_BLOCK_TARGET_WIDTH
    JBlockTgtT = J_BLOCK_TARGET_THICKNESS

    EjCamTgtL = EJECTOR_CAM_TARGET_LENGTH
    EjCamTgtW = EJECTOR_CAM_TARGET_WIDTH
    EjCamTgtT = EJECTOR_CAM_TARGET_THICKNESS

    Dim i As Long
    Dim d As String

    For i = 1 To BomCount

        d = NormalizeText(BomRows(i).Description)

        If (InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Or InStr(d, "JBLOCK") > 0) Then
            If BomRows(i).BomLength > 0 Then
                JBlockTgtL = BomRows(i).BomLength
                JBlockTgtW = BomRows(i).BomWidth
                JBlockTgtT = BomRows(i).BomThickness

                LogLine "J BLOCK target from BOM: L=" & FormatNumberForCsv(JBlockTgtL) & _
                        " W=" & FormatNumberForCsv(JBlockTgtW) & _
                        " T=" & FormatNumberForCsv(JBlockTgtT)
            End If
        End If

        If (InStr(d, "CAM") > 0) And (InStr(d, "EJ") > 0 Or InStr(d, "EJECTOR") > 0) _
           And InStr(d, "COVER") = 0 And InStr(d, "J-BLOCK") = 0 And InStr(d, "J BLOCK") = 0 Then

            If BomRows(i).BomLength > 0 Then
                EjCamTgtL = BomRows(i).BomLength
                EjCamTgtW = BomRows(i).BomWidth
                EjCamTgtT = BomRows(i).BomThickness

                LogLine "EJECTOR CAM target from BOM: L=" & FormatNumberForCsv(EjCamTgtL) & _
                        " W=" & FormatNumberForCsv(EjCamTgtW) & _
                        " T=" & FormatNumberForCsv(EjCamTgtT)
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "SetPackageTargetsFromBom error: " & Err.Description
End Sub

Private Function IsPullcoreBomRow(ByRef b As BomInfo) As Boolean
    Dim d As String
    d = NormalizeText(b.Description)

    If InStr(d, "CAM") = 0 And InStr(d, "KEY") = 0 Then Exit Function

    If InStr(d, "EJECTOR") > 0 Then Exit Function
    If InStr(d, "FLIPPER") > 0 Then Exit Function
    If InStr(d, "COVER") > 0 Then Exit Function
    If InStr(d, "DIRT") > 0 Then Exit Function
    If InStr(d, "STOP") > 0 Then Exit Function
    If InStr(d, "J-BLOCK") > 0 Or InStr(d, "J BLOCK") > 0 Or InStr(d, "JBLOCK") > 0 Then Exit Function
    If InStr(d, "HOLDER") > 0 Then Exit Function
    If InStr(d, "PLATE") > 0 Then Exit Function
    If InStr(d, "SMED") > 0 Then Exit Function
    If InStr(d, "POT") > 0 Then Exit Function

    If InStr(d, "PULLCORE") > 0 Or InStr(d, "PULL CORE") > 0 Then
        IsPullcoreBomRow = True
        Exit Function
    End If

    If HasPullcoreLocationToken(d) Then
        IsPullcoreBomRow = True
    End If
End Function

Private Function HasPullcoreLocationToken(ByVal d As String) As Boolean
    If GetPullcoreLocationCode(d) <> "" Then
        HasPullcoreLocationToken = True
        Exit Function
    End If

    Dim toks() As String
    toks = Split(d, " ")

    Dim i As Long
    For i = LBound(toks) To UBound(toks)
        Select Case toks(i)
            Case "TE", "LE", "ID", "OD", "IDTE", "IDLE", "ODTE", "ODLE"
                HasPullcoreLocationToken = True
                Exit Function
        End Select
    Next i
End Function

Private Sub AddPullcoreMatchRow(ByRef b As BomInfo, ByVal cadIdx As Long, _
                                ByVal cadL As Double, ByVal cadW As Double, ByVal cadT As Double)

    PullcoreMatchCount = PullcoreMatchCount + 1
    ReDim Preserve PullcoreMatches(1 To PullcoreMatchCount)

    Dim displayName As String
    displayName = CleanPullcoreDisplayName(b.Description)

    If displayName = "" Then displayName = b.quoteName

    PullcoreMatches(PullcoreMatchCount).quoteName = displayName
    PullcoreMatches(PullcoreMatchCount).Description = b.Description
    PullcoreMatches(PullcoreMatchCount).material = b.material
    PullcoreMatches(PullcoreMatchCount).Quantity = b.Quantity

    PullcoreMatches(PullcoreMatchCount).CadPartIndex = cadIdx
    PullcoreMatches(PullcoreMatchCount).isCam = (InStr(NormalizeText(b.Description), "CAM") > 0)

    PullcoreMatches(PullcoreMatchCount).BomThickness = b.BomThickness
    PullcoreMatches(PullcoreMatchCount).BomWidth = b.BomWidth
    PullcoreMatches(PullcoreMatchCount).BomLength = b.BomLength

    If cadIdx > 0 Then

        PullcoreMatches(PullcoreMatchCount).CadThickness = cadT
        PullcoreMatches(PullcoreMatchCount).CadWidth = cadW
        PullcoreMatches(PullcoreMatchCount).CadLength = cadL

        PullcoreMatches(PullcoreMatchCount).FittedThickness = cadT
        PullcoreMatches(PullcoreMatchCount).FittedWidth = cadW
        PullcoreMatches(PullcoreMatchCount).FittedLength = cadL

        If parts(cadIdx).hasOriginalAsmBBox Then
            PullcoreMatches(PullcoreMatchCount).OriginalLength = parts(cadIdx).OriginalAsmLength
            PullcoreMatches(PullcoreMatchCount).OriginalWidth = parts(cadIdx).OriginalAsmWidth
            PullcoreMatches(PullcoreMatchCount).OriginalThickness = parts(cadIdx).OriginalAsmThickness
        Else
            PullcoreMatches(PullcoreMatchCount).OriginalLength = parts(cadIdx).Length
            PullcoreMatches(PullcoreMatchCount).OriginalWidth = parts(cadIdx).Width
            PullcoreMatches(PullcoreMatchCount).OriginalThickness = parts(cadIdx).Thickness
        End If

        PullcoreMatches(PullcoreMatchCount).Status = ComparePullcoreDimsToBomStatus(b, cadL, cadW, cadT)

        PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg = _
            EstimatePullcoreAngleFromFittedAndOriginalDeg( _
                cadL, cadW, cadT, _
                PullcoreMatches(PullcoreMatchCount).OriginalLength, _
                PullcoreMatches(PullcoreMatchCount).OriginalWidth, _
                PullcoreMatches(PullcoreMatchCount).OriginalThickness)

        If Abs(PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            PullcoreMatches(PullcoreMatchCount).DxfRotationDeg = _
                PULLCORE_STRAIGHTEN_SIGN * PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg
        Else
            PullcoreMatches(PullcoreMatchCount).DxfRotationDeg = 0#
        End If

    Else
        PullcoreMatches(PullcoreMatchCount).Status = "NO CAD MATCH"
    End If

    Dim cadNameForLog As String

    If cadIdx > 0 Then
        cadNameForLog = parts(cadIdx).componentName
    Else
        cadNameForLog = "(none)"
    End If

    LogLine "PULLCORE matched: '" & displayName & "' -> CAD '" & cadNameForLog & "'"

    If cadIdx > 0 Then
        LogLine "  Original assembly/world bbox L/W/T=" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalLength) & "/" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalWidth) & "/" & _
                FormatNumberForCsv(PullcoreMatches(PullcoreMatchCount).OriginalThickness)

        LogLine "  Fitted bbox L/W/T=" & _
                FormatNumberForCsv(cadL) & "/" & _
                FormatNumberForCsv(cadW) & "/" & _
                FormatNumberForCsv(cadT)

        LogLine "  Pullcore angle=" & _
                Format(PullcoreMatches(PullcoreMatchCount).DetectedAngleDeg, "0.00") & _
                " deg, DXF rotation=" & _
                Format(PullcoreMatches(PullcoreMatchCount).DxfRotationDeg, "0.00") & " deg"

        If parts(cadIdx).hasAsmCenter Then
            LogLine "  Pullcore assembly center X/Y/Z=" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterX) & "/" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterY) & "/" & _
                    FormatNumberForCsv(parts(cadIdx).AsmCenterZ)
        End If
    End If
End Sub

Private Function ComparePullcoreDimsToBomStatus(ByRef b As BomInfo, _
                                                ByVal cadL As Double, _
                                                ByVal cadW As Double, _
                                                ByVal cadT As Double) As String
On Error GoTo ErrHandler

    If b.hasDims = False Then
        ComparePullcoreDimsToBomStatus = "MATCH (no BOM dims)"
        Exit Function
    End If

    Dim worst As Double

    worst = Abs(cadL - b.BomLength)

    If Abs(cadW - b.BomWidth) > worst Then worst = Abs(cadW - b.BomWidth)
    If Abs(cadT - b.BomThickness) > worst Then worst = Abs(cadT - b.BomThickness)

    If worst <= DIM_OK_TOL Then
        ComparePullcoreDimsToBomStatus = "OK"
    ElseIf worst <= DIM_REVIEW_TOL Then
        ComparePullcoreDimsToBomStatus = "REVIEW (" & FormatNumberForCsv(worst) & ")"
    Else
        ComparePullcoreDimsToBomStatus = "CHECK (" & FormatNumberForCsv(worst) & ")"
    End If

    Exit Function

ErrHandler:
    ComparePullcoreDimsToBomStatus = "ERROR"
End Function

Private Function EstimatePullcoreAngleFromFittedAndOriginalDeg( _
    ByVal fitL As Double, ByVal fitW As Double, ByVal fitT As Double, _
    ByVal origL As Double, ByVal origW As Double, ByVal origT As Double) As Double
On Error GoTo ErrHandler

    EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#

    If fitL <= 0 Or fitW <= 0 Or fitT <= 0 Then Exit Function
    If origL <= 0 Or origW <= 0 Or origT <= 0 Then Exit Function

    Dim b(1 To 3) As Double
    Dim n(1 To 3) As Double

    b(1) = fitL
    b(2) = fitW
    b(3) = fitT

    n(1) = origL
    n(2) = origW
    n(3) = origT

    SortArray3Descending b
    SortArray3Descending n

    If Abs(b(1) - n(1)) <= 0.001 And _
       Abs(b(2) - n(2)) <= 0.001 And _
       Abs(b(3) - n(3)) <= 0.001 Then
        EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#
        Exit Function
    End If

    Dim i As Long
    Dim unchangedIdx As Long
    Dim bestDiff As Double
    Dim d As Double

    unchangedIdx = 1
    bestDiff = 1E+99

    For i = 1 To 3
        d = Abs(b(i) - n(i))
        If d < bestDiff Then
            bestDiff = d
            unchangedIdx = i
        End If
    Next i

    Dim bA As Double
    Dim bB As Double
    Dim nA As Double
    Dim nB As Double
    Dim c As Long

    c = 0

    For i = 1 To 3
        If i <> unchangedIdx Then
            c = c + 1

            If c = 1 Then
                bA = b(i)
                nA = n(i)
            ElseIf c = 2 Then
                bB = b(i)
                nB = n(i)
            End If
        End If
    Next i

    If bB > bA Then
        SwapDoubleValues bA, bB
        SwapDoubleValues nA, nB
    End If

    EstimatePullcoreAngleFromFittedAndOriginalDeg = SearchBoxRotationDeg(bA, bB, nA, nB)

    Exit Function

ErrHandler:
    EstimatePullcoreAngleFromFittedAndOriginalDeg = 0#
End Function

Private Sub SortArray3Descending(ByRef a() As Double)
    If a(2) > a(1) Then SwapDoubleValues a(1), a(2)
    If a(3) > a(2) Then SwapDoubleValues a(2), a(3)
    If a(2) > a(1) Then SwapDoubleValues a(1), a(2)
End Sub

Private Sub SwapDoubleValues(ByRef a As Double, ByRef b As Double)
    Dim T As Double
    T = a
    a = b
    b = T
End Sub

' ============================================================
' END OF PART 5
' Paste Part 6 immediately after this.
' ============================================================

