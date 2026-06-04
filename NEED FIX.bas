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
Private Const ENABLE_EXPORT_LOG As Boolean = False

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

Private Const OD_HOLDER_CENTER_ROTATION_DEG As Double = 0#

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
Private Const JOB_SIGNATURE_REPORT_FILE As String = "XT_Export_Job_Signature.csv"

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
Private Const swSaveAsOptions_Copy As Long = 2

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
    WriteJobSignatureCsv CurrentJobFolder & "\" & JOB_SIGNATURE_REPORT_FILE

    If CREATE_PULLCORE_DIMENSIONS_EXCEL And PullcoreMatchCount > 0 Then
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

    If DISABLE_STABILIZE_DELAYS Then Exit Sub

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
                    LogLine "WARNING: " & quoteName & " assembly holder-face DXF failed; using part-based fallback."
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

Private Function SaveModelCopyAs(ByVal model As Object, ByVal fullPath As String) As Boolean
On Error GoTo ErrHandler

    SaveModelCopyAs = False

    If model Is Nothing Then Exit Function
    If fullPath = "" Then Exit Function

    Dim errs As Long
    Dim warns As Long

    LogLine "Saving copy: " & fullPath
    EnsureSwHidden

    ' Use Copy so temporary native drawing sources do not rename/repath the
    ' live assembly document. Without this, the next export can keep using a
    ' stale COM object after the temp DXF source is closed/deleted.
    model.Extension.SaveAs3 fullPath, _
                            swSaveAsCurrentVersion, _
                            swSaveAsOptions_Silent + swSaveAsOptions_Copy, _
                            Nothing, Nothing, errs, warns

    LogLine "Save copy done. Errors=" & errs & " Warnings=" & warns

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    SaveModelCopyAs = fso.FileExists(fullPath)
    Exit Function

ErrHandler:
    LogLine "SaveModelCopyAs error: " & Err.Description
    SaveModelCopyAs = False
End Function

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

    If SaveModelCopyAs(swModel, tempNativePath) = False Then GoTo CleanExit

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

    If SaveModelCopyAs(swModel, holdersTempNativePath) = False Then GoTo CleanExit

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

    Select Case NormalizeKey(quoteName)
        Case "IDHOLDER", "ODHOLDER"
            parentPrimary = "*Top"
            parentFallback = CMS_TOP_VIEW_NAME
        Case Else
            parentPrimary = CMS_TOP_VIEW_NAME
            parentFallback = "*Top"
    End Select

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

    Dim holderViewName As String
    Dim holderViewId As Long
    Dim holderViewToken As String

    ' Shop wants both holder detail DXFs from the top face.
    holderViewName = "*Top"
    holderViewId = 5
    holderViewToken = "TOPVIEW"

    ' Native assembly, not X_T.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_" & NormalizeKey(quoteName) & "_" & holderViewToken & "_TEMP.sldasm"

    ApplyCmsTopView assyModel
    StabilizeActiveView assyModel, 100

    assyModel.ShowNamedView2 holderViewName, holderViewId
    StabilizeActiveView assyModel, 100

    LogLine quoteName & " holder DXF: saving native temp SLDASM from " & holderViewName & ":"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(assyModel, tempNativePath) = False Then GoTo CleanExit

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "Holder assembly " & holderViewName & " DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    CurrentIdHolderDxfFromAssembly = (NormalizeKey(quoteName) = "IDHOLDER")

    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, quoteName, _
                                     holderViewName, CMS_TOP_VIEW_NAME, _
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

    If SaveModelCopyAs(assyModel, tempNativePath) = False Then GoTo CleanExit

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

    ForceAllDxfScales1To1 swDraw

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
    End If

    ForceDrawingSheetScale1To1 swDraw

    swDraw.GraphicsRedraw2
End Sub

Private Sub ForceDrawingSheetScale1To1(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then Exit Sub

    ' SolidWorks templates and view insertion can reset the sheet to 1:2.
    ' Force the sheet itself to 1:1 using multiple late-bound API paths.
    Err.Clear
    swSheet.SetScale 1#, 1#, False, False
    Err.Clear
    swSheet.SetScale 1#, 1#, True, True
    Err.Clear

    Dim sheetName As String
    sheetName = ""
    sheetName = CStr(swSheet.GetName)

    If sheetName <> "" Then
        Err.Clear
        swDraw.SetupSheet5 sheetName, 12, 12, 1#, 1#, False, "", _
                           E_SHEET_WIDTH_IN / INCHES_PER_METER, _
                           E_SHEET_HEIGHT_IN / INCHES_PER_METER, "", False
        Err.Clear

        Err.Clear
        swDraw.SetupSheet4 sheetName, 12, 12, 1#, 1#, False, "", _
                           E_SHEET_WIDTH_IN / INCHES_PER_METER, _
                           E_SHEET_HEIGHT_IN / INCHES_PER_METER, ""
        Err.Clear
    End If
End Sub

Private Sub ForceAllDxfScales1To1(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    ForceDrawingSheetScale1To1 swDraw

    Dim v As Object
    Set v = swDraw.GetFirstView

    If Not v Is Nothing Then Set v = v.GetNextView

    Do While Not v Is Nothing
        SetDrawingViewScale v, 1#
        Set v = v.GetNextView
    Loop

    ForceDrawingSheetScale1To1 swDraw

    swDraw.ForceRebuild3 False
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
        Dim scaleRatio(0 To 1) As Double
        scaleRatio(0) = 1#
        scaleRatio(1) = 1#

        Err.Clear
        swView.ScaleRatio = scaleRatio
        Err.Clear
        swView.ScaleRatio = "1:1"
        Err.Clear
    End If

    ' Set this again last. Some SolidWorks view operations flip back to sheet scale.
    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal
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
' ============================================================
' PULLCORE MATCHING
' ============================================================

Private Sub MatchPullcoreRowsByClass()
On Error GoTo ErrHandler

    Dim camRows() As Long
    Dim keyRows() As Long
    Dim camN As Long
    Dim keyN As Long

    camN = 0
    keyN = 0

    ReDim camRows(1 To 1)
    ReDim keyRows(1 To 1)

    Dim i As Long

    For i = 1 To BomCount

        If IsPullcoreBomRow(BomRows(i)) Then

            Dim reps As Long
            reps = BomRows(i).Quantity

            If reps < 1 Then reps = 1
            If reps > 8 Then reps = 8

            Dim isCamRow As Boolean
            isCamRow = (InStr(NormalizeText(BomRows(i).Description), "CAM") > 0)

            If (Not isCamRow) And reps < 2 Then reps = 2

            Dim r As Long
            For r = 1 To reps
                If isCamRow Then
                    camN = camN + 1
                    ReDim Preserve camRows(1 To camN)
                    camRows(camN) = i
                Else
                    keyN = keyN + 1
                    ReDim Preserve keyRows(1 To keyN)
                    keyRows(keyN) = i
                End If
            Next r

        End If

    Next i

    LogLine "Pullcore BOM rows (qty-expanded): cams=" & camN & " keys=" & keyN

    If camN > 0 Then AssignPullcoreClass camRows, camN, True
    If keyN > 0 Then AssignPullcoreClass keyRows, keyN, False

    If AUTO_LABEL_PULLCORE_ID_OD_BY_HEIGHT Then
        If LabelPullcoreCamsAndKeysByHeight() = False Then
            LabelPullcoreKeysByProximity
        End If
    Else
        LabelPullcoreKeysByProximity
    End If

    On Error Resume Next
    If Not swModel Is Nothing Then
        Dim errs As Long
        swApp.ActivateDoc3 swModel.GetTitle, False, 0, errs
    End If
    EnsureSwHidden
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    LogLine "MatchPullcoreRowsByClass error: " & Err.Description
End Sub

Private Function LabelPullcoreCamsAndKeysByHeight() As Boolean
On Error GoTo ErrHandler

    LabelPullcoreCamsAndKeysByHeight = False
    PullcoreIdOdHeightAxisUsed = ""

    If PullcoreMatchCount <= 0 Then Exit Function

    Dim camMatches As Collection
    Dim keyMatches As Collection

    Set camMatches = New Collection
    Set keyMatches = New Collection

    Dim i As Long
    Dim cadIdx As Long

    For i = 1 To PullcoreMatchCount

        cadIdx = PullcoreMatches(i).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmCenter Then

                If PullcoreMatches(i).isCam Then
                    camMatches.Add i
                Else
                    keyMatches.Add i
                End If

            End If
        End If

    Next i

    If camMatches.count < 2 Then
        LogLine "Pullcore ID/OD height labeling skipped: need at least 2 CAM matches with assembly centers."
        Exit Function
    End If

    Dim axisName As String
    axisName = DeterminePullcoreIdOdHeightAxis(camMatches, keyMatches)

    If axisName = "" Then axisName = "Y"

    PullcoreIdOdHeightAxisUsed = axisName

    LogLine "Pullcore ID/OD height labeling axis selected: " & axisName

    Dim highCamMatch As Long
    Dim lowCamMatch As Long

    FindHighLowPullcoreMatchesByAxis camMatches, axisName, highCamMatch, lowCamMatch

    If highCamMatch <= 0 Or lowCamMatch <= 0 Then
        LogLine "Pullcore ID/OD height labeling skipped: could not determine high/low CAM."
        Exit Function
    End If

    Dim idCamMatch As Long
    Dim odCamMatch As Long

    If PULLCORE_ID_IS_HIGHER Then
        idCamMatch = highCamMatch
        odCamMatch = lowCamMatch
    Else
        idCamMatch = lowCamMatch
        odCamMatch = highCamMatch
    End If

    SetPullcoreMatchLabel idCamMatch, "ID PULLCORE CAM", axisName
    SetPullcoreMatchLabel odCamMatch, "OD PULLCORE CAM", axisName

    If keyMatches.count > 0 Then
        LabelPullcoreKeysToNearestLabeledCams keyMatches, idCamMatch, odCamMatch, axisName
    End If

    LabelPullcoreCamsAndKeysByHeight = True

    Exit Function

ErrHandler:
    LogLine "LabelPullcoreCamsAndKeysByHeight error: " & Err.Description
    LabelPullcoreCamsAndKeysByHeight = False
End Function

Private Function DeterminePullcoreIdOdHeightAxis(ByVal camMatches As Collection, _
                                                 ByVal keyMatches As Collection) As String
On Error GoTo ErrHandler

    Dim setting As String
    setting = UCase(Trim(PULLCORE_ID_OD_HEIGHT_AXIS))

    If setting = "Y" Or setting = "Z" Then
        DeterminePullcoreIdOdHeightAxis = setting
        Exit Function
    End If

    Dim yMin As Double
    Dim yMax As Double
    Dim zMin As Double
    Dim zMax As Double
    Dim firstVal As Boolean

    firstVal = True

    Dim i As Long
    Dim mi As Long
    Dim cadIdx As Long

    For i = 1 To camMatches.count

        mi = CLng(camMatches(i))
        cadIdx = PullcoreMatches(mi).CadPartIndex

        If cadIdx > 0 And cadIdx <= PartCount Then
            If parts(cadIdx).hasAsmCenter Then

                If firstVal Then
                    yMin = parts(cadIdx).AsmCenterY
                    yMax = parts(cadIdx).AsmCenterY
                    zMin = parts(cadIdx).AsmCenterZ
                    zMax = parts(cadIdx).AsmCenterZ
                    firstVal = False
                Else
                    If parts(cadIdx).AsmCenterY < yMin Then yMin = parts(cadIdx).AsmCenterY
                    If parts(cadIdx).AsmCenterY > yMax Then yMax = parts(cadIdx).AsmCenterY
                    If parts(cadIdx).AsmCenterZ < zMin Then zMin = parts(cadIdx).AsmCenterZ
                    If parts(cadIdx).AsmCenterZ > zMax Then zMax = parts(cadIdx).AsmCenterZ
                End If

            End If
        End If

    Next i

    If firstVal Then
        DeterminePullcoreIdOdHeightAxis = "Y"
        Exit Function
    End If

    If Abs(zMax - zMin) > Abs(yMax - yMin) + 0.001 Then
        DeterminePullcoreIdOdHeightAxis = "Z"
    Else
        DeterminePullcoreIdOdHeightAxis = "Y"
    End If

    LogLine "Pullcore CAM center spread: Y=" & FormatNumberForCsv(Abs(yMax - yMin)) & _
            " Z=" & FormatNumberForCsv(Abs(zMax - zMin))

    Exit Function

ErrHandler:
    DeterminePullcoreIdOdHeightAxis = "Y"
End Function

Private Sub FindHighLowPullcoreMatchesByAxis(ByVal matches As Collection, _
                                             ByVal axisName As String, _
                                             ByRef highMatch As Long, _
                                             ByRef lowMatch As Long)
On Error GoTo ErrHandler

    highMatch = 0
    lowMatch = 0

    If matches Is Nothing Then Exit Sub
    If matches.count = 0 Then Exit Sub

    Dim bestHigh As Double
    Dim bestLow As Double

    bestHigh = -1E+99
    bestLow = 1E+99

    Dim i As Long
    Dim mi As Long
    Dim h As Double

    For i = 1 To matches.count

        mi = CLng(matches(i))
        h = GetPullcoreMatchHeightValue(mi, axisName)

        If h > bestHigh Then
            bestHigh = h
            highMatch = mi
        End If

        If h < bestLow Then
            bestLow = h
            lowMatch = mi
        End If

    Next i

    Exit Sub

ErrHandler:
    highMatch = 0
    lowMatch = 0
End Sub

Private Function GetPullcoreMatchHeightValue(ByVal matchIdx As Long, ByVal axisName As String) As Double
On Error GoTo ErrHandler

    GetPullcoreMatchHeightValue = 0#

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Function

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If parts(cadIdx).hasAsmCenter = False Then Exit Function

    If UCase(axisName) = "Z" Then
        GetPullcoreMatchHeightValue = parts(cadIdx).AsmCenterZ
    Else
        GetPullcoreMatchHeightValue = parts(cadIdx).AsmCenterY
    End If

    Exit Function

ErrHandler:
    GetPullcoreMatchHeightValue = 0#
End Function

Private Sub SetPullcoreMatchLabel(ByVal matchIdx As Long, _
                                  ByVal newName As String, _
                                  ByVal axisName As String)
On Error Resume Next

    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    PullcoreMatches(matchIdx).quoteName = newName

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx > 0 And cadIdx <= PartCount Then

        LogLine "Pullcore labeled '" & newName & "' -> CAD '" & _
                parts(cadIdx).componentName & "'  center X/Y/Z=" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterX) & "/" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterY) & "/" & _
                FormatNumberForCsv(parts(cadIdx).AsmCenterZ) & _
                "  ID/OD axis=" & axisName & _
                "  height=" & FormatNumberForCsv(GetPullcoreMatchHeightValue(matchIdx, axisName))

    Else

        LogLine "Pullcore labeled '" & newName & "' with no CAD component."

    End If
End Sub

Private Sub LabelPullcoreKeysToNearestLabeledCams(ByVal keyMatches As Collection, _
                                                  ByVal idCamMatch As Long, _
                                                  ByVal odCamMatch As Long, _
                                                  ByVal axisName As String)
On Error GoTo ErrHandler

    If keyMatches Is Nothing Then Exit Sub
    If keyMatches.count = 0 Then Exit Sub
    If idCamMatch <= 0 Or odCamMatch <= 0 Then Exit Sub

    If keyMatches.count = 1 Then

        Dim onlyKey As Long
        onlyKey = CLng(keyMatches(1))

        If PullcoreMatchDistanceSq(onlyKey, idCamMatch) <= PullcoreMatchDistanceSq(onlyKey, odCamMatch) Then
            SetPullcoreMatchLabel onlyKey, "ID PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel onlyKey, "OD PULLCORE KEY", axisName
        End If

        Exit Sub

    End If

    If keyMatches.count = 2 Then

        Dim k1 As Long
        Dim k2 As Long

        k1 = CLng(keyMatches(1))
        k2 = CLng(keyMatches(2))

        Dim optionA As Double
        Dim optionB As Double

        optionA = PullcoreMatchDistanceSq(k1, idCamMatch) + PullcoreMatchDistanceSq(k2, odCamMatch)
        optionB = PullcoreMatchDistanceSq(k2, idCamMatch) + PullcoreMatchDistanceSq(k1, odCamMatch)

        If optionA <= optionB Then
            SetPullcoreMatchLabel k1, "ID PULLCORE KEY", axisName
            SetPullcoreMatchLabel k2, "OD PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel k2, "ID PULLCORE KEY", axisName
            SetPullcoreMatchLabel k1, "OD PULLCORE KEY", axisName
        End If

        Exit Sub

    End If

    Dim i As Long
    Dim mi As Long

    For i = 1 To keyMatches.count

        mi = CLng(keyMatches(i))

        If PullcoreMatchDistanceSq(mi, idCamMatch) <= PullcoreMatchDistanceSq(mi, odCamMatch) Then
            SetPullcoreMatchLabel mi, "ID PULLCORE KEY", axisName
        Else
            SetPullcoreMatchLabel mi, "OD PULLCORE KEY", axisName
        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "LabelPullcoreKeysToNearestLabeledCams error: " & Err.Description
End Sub

Private Function PullcoreMatchDistanceSq(ByVal matchA As Long, ByVal matchB As Long) As Double
On Error GoTo ErrHandler

    PullcoreMatchDistanceSq = 1E+99

    If matchA <= 0 Or matchA > PullcoreMatchCount Then Exit Function
    If matchB <= 0 Or matchB > PullcoreMatchCount Then Exit Function

    Dim cadA As Long
    Dim cadB As Long

    cadA = PullcoreMatches(matchA).CadPartIndex
    cadB = PullcoreMatches(matchB).CadPartIndex

    If cadA <= 0 Or cadA > PartCount Then Exit Function
    If cadB <= 0 Or cadB > PartCount Then Exit Function

    If parts(cadA).hasAsmCenter = False Then Exit Function
    If parts(cadB).hasAsmCenter = False Then Exit Function

    PullcoreMatchDistanceSq = _
        (parts(cadA).AsmCenterX - parts(cadB).AsmCenterX) ^ 2 + _
        (parts(cadA).AsmCenterY - parts(cadB).AsmCenterY) ^ 2 + _
        (parts(cadA).AsmCenterZ - parts(cadB).AsmCenterZ) ^ 2

    Exit Function

ErrHandler:
    PullcoreMatchDistanceSq = 1E+99
End Function

Private Sub LabelPullcoreKeysByProximity()
On Error GoTo ErrHandler

    Dim haveId As Boolean
    Dim haveOd As Boolean

    Dim idx As Double
    Dim idy As Double
    Dim idz As Double
    Dim odx As Double
    Dim ody As Double
    Dim odz As Double

    Dim i As Long
    Dim ci As Long
    Dim nm As String

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).isCam And PullcoreMatches(i).CadPartIndex > 0 Then

            ci = PullcoreMatches(i).CadPartIndex

            If parts(ci).hasAsmCenter Then

                nm = NormalizeText(PullcoreMatches(i).Description)

                If InStr(nm, "OD") > 0 Then
                    odx = parts(ci).AsmCenterX
                    ody = parts(ci).AsmCenterY
                    odz = parts(ci).AsmCenterZ
                    haveOd = True
                ElseIf InStr(nm, "ID") > 0 Then
                    idx = parts(ci).AsmCenterX
                    idy = parts(ci).AsmCenterY
                    idz = parts(ci).AsmCenterZ
                    haveId = True
                End If

            End If

        End If

    Next i

    If Not (haveId And haveOd) Then
        LogLine "Pullcore key ID/OD labeling skipped (need both ID and OD cam centers)."
        Exit Sub
    End If

    Dim keyIdx() As Long
    Dim keyN As Long

    keyN = 0
    ReDim keyIdx(1 To PullcoreMatchCount)

    For i = 1 To PullcoreMatchCount
        If (Not PullcoreMatches(i).isCam) And PullcoreMatches(i).CadPartIndex > 0 Then
            If parts(PullcoreMatches(i).CadPartIndex).hasAsmCenter Then
                keyN = keyN + 1
                keyIdx(keyN) = i
            End If
        End If
    Next i

    If keyN = 0 Then Exit Sub

    Dim bias() As Double
    ReDim bias(1 To keyN)

    Dim k As Long

    For k = 1 To keyN

        ci = PullcoreMatches(keyIdx(k)).CadPartIndex

        Dim dId As Double
        Dim dOd As Double

        dId = (parts(ci).AsmCenterX - idx) ^ 2 + (parts(ci).AsmCenterY - idy) ^ 2 + (parts(ci).AsmCenterZ - idz) ^ 2
        dOd = (parts(ci).AsmCenterX - odx) ^ 2 + (parts(ci).AsmCenterY - ody) ^ 2 + (parts(ci).AsmCenterZ - odz) ^ 2

        bias(k) = dId - dOd

    Next k

    If keyN = 2 Then

        Dim idK As Long
        Dim odK As Long

        If bias(1) <= bias(2) Then
            idK = 1
            odK = 2
        Else
            idK = 2
            odK = 1
        End If

        SetPullcoreKeyLabel keyIdx(idK), "ID PULLCORE KEY"
        SetPullcoreKeyLabel keyIdx(odK), "OD PULLCORE KEY"

    Else

        For k = 1 To keyN
            If bias(k) <= 0 Then
                SetPullcoreKeyLabel keyIdx(k), "ID PULLCORE KEY"
            Else
                SetPullcoreKeyLabel keyIdx(k), "OD PULLCORE KEY"
            End If
        Next k

    End If

    Exit Sub

ErrHandler:
    LogLine "LabelPullcoreKeysByProximity error: " & Err.Description
End Sub

Private Sub SetPullcoreKeyLabel(ByVal matchIdx As Long, ByVal newName As String)
On Error Resume Next

    PullcoreMatches(matchIdx).quoteName = newName

    LogLine "Pullcore key labeled '" & newName & "' -> CAD '" & _
            parts(PullcoreMatches(matchIdx).CadPartIndex).componentName & "'"
End Sub

Private Sub AssignPullcoreClass(ByRef bomRowIdx() As Long, ByVal rowN As Long, ByVal wantCam As Boolean)
On Error GoTo ErrHandler

    Dim candIdx() As Long
    Dim candL() As Double
    Dim candW() As Double
    Dim candT() As Double
    Dim candN As Long

    candN = 0

    ReDim candIdx(1 To 1)
    ReDim candL(1 To 1)
    ReDim candW(1 To 1)
    ReDim candT(1 To 1)

    Dim i As Long
    Dim j As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim isCand As Boolean
            isCand = False

            For j = 1 To rowN
                If IsRoughPullcoreCandidateForBom(i, BomRows(bomRowIdx(j))) Then
                    isCand = True
                    Exit For
                End If
            Next j

            If isCand Then

                Dim bfL As Double
                Dim bfW As Double
                Dim bfT As Double

                bfL = parts(i).Length
                bfW = parts(i).Width
                bfT = parts(i).Thickness

                If USE_PULLCORE_BEST_FIT_BBOX Then
                    If TryGetPullcoreBestFitDims(i, bfL, bfW, bfT) = False Then
                        bfL = parts(i).Length
                        bfW = parts(i).Width
                        bfT = parts(i).Thickness
                    End If
                End If

                candN = candN + 1

                ReDim Preserve candIdx(1 To candN)
                ReDim Preserve candL(1 To candN)
                ReDim Preserve candW(1 To candN)
                ReDim Preserve candT(1 To candN)

                candIdx(candN) = i
                candL(candN) = bfL
                candW(candN) = bfW
                candT(candN) = bfT

            End If

        End If

    Next i

    LogLine "Pullcore " & IIf(wantCam, "CAM", "KEY") & " candidates found: " & candN & " for " & rowN & " BOM rows"

    If candN = 0 Then
        For j = 1 To rowN
            LogLine "PULLCORE UNMATCHED (no candidate): " & BomRows(bomRowIdx(j)).Description
            AddPullcoreMatchRow BomRows(bomRowIdx(j)), 0, 0, 0, 0
        Next j
        Exit Sub
    End If

    Dim rIdx() As Long
    ReDim rIdx(1 To rowN)

    For i = 1 To rowN
        rIdx(i) = bomRowIdx(i)
    Next i

    Dim rowUsed() As Boolean
    Dim candUsed() As Boolean

    ReDim rowUsed(1 To rowN)
    ReDim candUsed(1 To candN)

    Dim pairCount As Long

    pairCount = rowN
    If candN < pairCount Then pairCount = candN

    Dim pairStep As Long

    For pairStep = 1 To pairCount

        Dim bestR As Long
        Dim bestC As Long
        Dim bestDist As Double

        bestR = 0
        bestC = 0
        bestDist = 1E+99

        For i = 1 To rowN

            If rowUsed(i) = False Then

                For j = 1 To candN

                    If candUsed(j) = False Then

                        Dim d As Double

                        d = Abs(candL(j) - BomRows(rIdx(i)).BomLength) * 3# _
                          + Abs(candW(j) - BomRows(rIdx(i)).BomWidth) _
                          + Abs(candT(j) - BomRows(rIdx(i)).BomThickness)

                        If d < bestDist Then
                            bestDist = d
                            bestR = i
                            bestC = j
                        End If

                    End If

                Next j

            End If

        Next i

        If bestR = 0 Or bestC = 0 Then Exit For

        AddPullcoreMatchRow BomRows(rIdx(bestR)), candIdx(bestC), _
                            candL(bestC), candW(bestC), candT(bestC)

        parts(candIdx(bestC)).UsedForBomMatch = True

        rowUsed(bestR) = True
        candUsed(bestC) = True

    Next pairStep

    For i = 1 To rowN
        If rowUsed(i) = False Then
            LogLine "PULLCORE UNMATCHED (more BOM rows than candidates): " & BomRows(rIdx(i)).Description
            AddPullcoreMatchRow BomRows(rIdx(i)), 0, 0, 0, 0
        End If
    Next i

    Exit Sub

ErrHandler:
    LogLine "AssignPullcoreClass error: " & Err.Description
End Sub

Private Function IsRoughPullcoreCandidateForBom(ByVal cadIdx As Long, ByRef b As BomInfo) As Boolean
On Error GoTo ErrHandler

    IsRoughPullcoreCandidateForBom = False

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function
    If parts(cadIdx).UsedForBomMatch Then Exit Function
    If b.hasDims = False Then Exit Function

    Dim cadL As Double
    Dim cadW As Double
    Dim cadT As Double

    cadL = parts(cadIdx).Length
    cadW = parts(cadIdx).Width
    cadT = parts(cadIdx).Thickness

    If parts(cadIdx).BBoxVolume > 200# Then Exit Function
    If cadT < 0.4 Then Exit Function

    If Abs(cadL - b.BomLength) > 1# Then Exit Function
    If cadW < b.BomWidth - 0.4 Then Exit Function
    If cadW > b.BomWidth + 1# Then Exit Function

    IsRoughPullcoreCandidateForBom = True

    Exit Function

ErrHandler:
    IsRoughPullcoreCandidateForBom = False
End Function

' ============================================================
' MATCHING HELPERS / EXPORT ROWS
' ============================================================

Private Sub AddExportRow(ByRef b As BomInfo, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If ShouldSkipExportQuoteName(b.quoteName) Then
        LogLine "Export quote skipped (package-handled): " & b.quoteName
        If cadIdx > 0 Then RecordSpecialBomCadMatch b, cadIdx
        Exit Sub
    End If

    Dim isPyropel As Boolean
    isPyropel = (InStr(NormalizeText(b.material), "PYROPEL") > 0 Or InStr(NormalizeText(b.Description), "PYROPEL") > 0)

    If ONLY_INCLUDE_4140_BOM_ITEMS Then
        If NormalizeSteelType(b.material) <> "4140" Then
            If IsInsertQuoteName(b.quoteName) = False And isPyropel = False Then
                LogLine "Non-4140 item skipped: " & b.Description & " (" & b.material & ")"
                Exit Sub
            End If
        End If
    End If

    ExportCount = ExportCount + 1
    ReDim Preserve ExportRows(1 To ExportCount)

    ExportRows(ExportCount).quoteName = b.quoteName
    ExportRows(ExportCount).Quantity = b.Quantity
    ExportRows(ExportCount).material = b.material

    ExportRows(ExportCount).CadPartIndex = cadIdx
    ExportRows(ExportCount).HasCad = (cadIdx > 0)

    ExportRows(ExportCount).BomThickness = b.BomThickness
    ExportRows(ExportCount).BomWidth = b.BomWidth
    ExportRows(ExportCount).BomLength = b.BomLength
    ExportRows(ExportCount).HasBomDims = b.hasDims

    If cadIdx > 0 Then

        ExportRows(ExportCount).Thickness = parts(cadIdx).Thickness
        ExportRows(ExportCount).Width = parts(cadIdx).Width
        ExportRows(ExportCount).Length = parts(cadIdx).Length
        ExportRows(ExportCount).Status = CompareBomToCadStatus(b, cadIdx)

        parts(cadIdx).UsedForBomMatch = True

    Else

        ExportRows(ExportCount).Status = "NO CAD MATCH"

    End If

    Exit Sub

ErrHandler:
    LogLine "AddExportRow error: " & Err.Description
End Sub

Private Sub RecordSpecialBomCadMatch(ByRef b As BomInfo, ByVal cadIdx As Long)
On Error Resume Next

    If cadIdx <= 0 Then Exit Sub
    If SpecialBomCadMatches Is Nothing Then Exit Sub

    Dim k As String
    k = NormalizeKey(b.quoteName)

    SpecialBomCadMatches(k) = cadIdx
    SpecialBomCadQuoteNames(k) = b.Description
End Sub

Private Function FindBestCadMatchForBom(ByRef b As BomInfo) As Long
On Error GoTo ErrHandler

    FindBestCadMatchForBom = 0

    If ShouldUseTcpBcpMassPairRule(b) Then

        Dim pref As String
        Dim massIdx As Long

        pref = GetMassPreferenceForQuoteName(b.quoteName)

        If pref <> "" Then
            massIdx = FindBestCadMatchForBomByDimsAndMassPreference(b, pref)

            If massIdx > 0 Then
                FindBestCadMatchForBom = massIdx
                Exit Function
            End If
        End If

    End If

    Dim bestIdx As Long
    Dim bestDiff As Double

    bestIdx = 0
    bestDiff = DIM_MAX_MATCH_TOTAL_DIFF + 1#

    Dim i As Long

    If b.hasDims Then

        For i = 1 To PartCount

            If parts(i).UsedForBomMatch = False Then

                Dim dL As Double
                Dim dW As Double
                Dim dT As Double

                dL = Abs(parts(i).Length - b.BomLength)
                dW = Abs(parts(i).Width - b.BomWidth)
                dT = Abs(parts(i).Thickness - b.BomThickness)

                Dim totalDiff As Double
                totalDiff = dL + dW + dT

                Dim nameBonus As Double
                nameBonus = 0#

                If IsNameMatch(parts(i).cleanName, b.quoteName) Then nameBonus = nameBonus - 0.75
                If IsNameMatch(parts(i).componentName, b.quoteName) Then nameBonus = nameBonus - 0.5

                totalDiff = totalDiff + nameBonus

                If dL <= DIM_REVIEW_TOL * 4 And dW <= DIM_REVIEW_TOL * 4 And dT <= DIM_REVIEW_TOL * 4 Then
                    If totalDiff < bestDiff Then
                        bestDiff = totalDiff
                        bestIdx = i
                    End If
                End If

            End If

        Next i

    End If

    If bestIdx = 0 Then
        bestIdx = FindMassBasedCadMatchForBom(b)
    End If

    FindBestCadMatchForBom = bestIdx
    Exit Function

ErrHandler:
    LogLine "FindBestCadMatchForBom error: " & Err.Description
    FindBestCadMatchForBom = 0
End Function

Private Function FindBestCadMatchForBomByDimsAndMassPreference(ByRef b As BomInfo, _
                                                               ByVal massPreference As String) As Long
On Error GoTo ErrHandler

    FindBestCadMatchForBomByDimsAndMassPreference = 0

    If b.hasDims = False Then Exit Function
    If massPreference = "" Then Exit Function

    Dim i As Long
    Dim bestDimDiff As Double

    bestDimDiff = 1E+99

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim dL As Double
            Dim dW As Double
            Dim dT As Double
            Dim totalDiff As Double

            dL = Abs(parts(i).Length - b.BomLength)
            dW = Abs(parts(i).Width - b.BomWidth)
            dT = Abs(parts(i).Thickness - b.BomThickness)

            totalDiff = dL + dW + dT

            If dL <= DIM_REVIEW_TOL * 4 And _
               dW <= DIM_REVIEW_TOL * 4 And _
               dT <= DIM_REVIEW_TOL * 4 Then

                If totalDiff < bestDimDiff Then bestDimDiff = totalDiff

            End If

        End If

    Next i

    If bestDimDiff = 1E+99 Then Exit Function

    Dim dimBand As Double
    dimBand = DIM_OK_TOL
    If dimBand < 0.05 Then dimBand = 0.05

    Dim bestIdx As Long
    Dim bestMass As Double

    bestIdx = 0

    If UCase(massPreference) = "LIGHT" Then
        bestMass = 1E+99
    Else
        bestMass = -1E+99
    End If

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            dL = Abs(parts(i).Length - b.BomLength)
            dW = Abs(parts(i).Width - b.BomWidth)
            dT = Abs(parts(i).Thickness - b.BomThickness)

            totalDiff = dL + dW + dT

            If dL <= DIM_REVIEW_TOL * 4 And _
               dW <= DIM_REVIEW_TOL * 4 And _
               dT <= DIM_REVIEW_TOL * 4 Then

                If totalDiff <= bestDimDiff + dimBand Then

                    If UCase(massPreference) = "LIGHT" Then
                        If parts(i).massValue < bestMass Then
                            bestMass = parts(i).massValue
                            bestIdx = i
                        End If
                    ElseIf UCase(massPreference) = "HEAVY" Then
                        If parts(i).massValue > bestMass Then
                            bestMass = parts(i).massValue
                            bestIdx = i
                        End If
                    End If

                End If

            End If

        End If

    Next i

    If bestIdx > 0 Then
        LogLine "TCP/BCP mass match: " & b.quoteName & _
                " preference=" & massPreference & _
                " -> CAD '" & parts(bestIdx).componentName & "'" & _
                " mass=" & FormatNumberForCsv(parts(bestIdx).massValue)
    End If

    FindBestCadMatchForBomByDimsAndMassPreference = bestIdx
    Exit Function

ErrHandler:
    LogLine "FindBestCadMatchForBomByDimsAndMassPreference error: " & Err.Description
    FindBestCadMatchForBomByDimsAndMassPreference = 0
End Function

Private Function FindMassBasedCadMatchForBom(ByRef b As BomInfo) As Long
On Error GoTo ErrHandler

    FindMassBasedCadMatchForBom = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch = False Then

            Dim score As Double
            score = 0

            If IsNameMatch(parts(i).cleanName, b.quoteName) Then score = score + 10
            If IsNameMatch(parts(i).componentName, b.quoteName) Then score = score + 6

            If score > bestScore Then
                bestScore = score
                bestIdx = i
            End If

        End If

    Next i

    If bestScore <= 0 Then
        FindMassBasedCadMatchForBom = 0
    Else
        FindMassBasedCadMatchForBom = bestIdx
    End If

    Exit Function

ErrHandler:
    FindMassBasedCadMatchForBom = 0
End Function

Private Function CompareBomToCadStatus(ByRef b As BomInfo, ByVal cadIdx As Long) As String
On Error GoTo ErrHandler

    If cadIdx <= 0 Then
        CompareBomToCadStatus = "NO CAD MATCH"
        Exit Function
    End If

    If b.hasDims = False Then
        CompareBomToCadStatus = "MATCH (no BOM dims)"
        Exit Function
    End If

    Dim dL As Double
    Dim dW As Double
    Dim dT As Double

    dL = Abs(parts(cadIdx).Length - b.BomLength)
    dW = Abs(parts(cadIdx).Width - b.BomWidth)
    dT = Abs(parts(cadIdx).Thickness - b.BomThickness)

    Dim worst As Double
    worst = dL

    If dW > worst Then worst = dW
    If dT > worst Then worst = dT

    If worst <= DIM_OK_TOL Then
        CompareBomToCadStatus = "OK"
    ElseIf worst <= DIM_REVIEW_TOL Then
        CompareBomToCadStatus = "REVIEW (" & FormatNumberForCsv(worst) & ")"
    Else
        CompareBomToCadStatus = "CHECK (" & FormatNumberForCsv(worst) & ")"
    End If

    Exit Function

ErrHandler:
    CompareBomToCadStatus = "ERROR"
End Function

Private Function IsPossibleJBlockByDims(ByVal L As Double, ByVal W As Double, ByVal T As Double) As Boolean
    If Abs(T - JBlockTgtT) <= J_BLOCK_DIM_MATCH_TOL Then
        If Abs(W - JBlockTgtW) <= J_BLOCK_DIM_MATCH_TOL Then
            If Abs(L - JBlockTgtL) <= J_BLOCK_DIM_MATCH_TOL Then
                IsPossibleJBlockByDims = True
            End If
        End If
    End If
End Function

Private Function IsPossibleCamByDims(ByVal L As Double, ByVal W As Double, ByVal T As Double) As Boolean
    If Abs(T - EjCamTgtT) <= EJECTOR_CAM_THICKNESS_TOL Then
        If Abs(W - EjCamTgtW) <= EJECTOR_CAM_WIDTH_TOL Then
            If Abs(L - EjCamTgtL) <= EJECTOR_CAM_LENGTH_TOL Then
                IsPossibleCamByDims = True
            End If
        End If
    End If
End Function

Private Function CleanPullcoreDisplayName(ByVal s As String) As String
    s = Trim(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanPullcoreDisplayName = s
End Function

Private Function GetPullcoreLocationCode(ByVal text As String) As String
    Dim s As String
    s = NormalizeText(text)

    If InStr(s, "IDTE") > 0 Or InStr(s, "ID TE") > 0 Then GetPullcoreLocationCode = "IDTE": Exit Function
    If InStr(s, "IDLE") > 0 Or InStr(s, "ID LE") > 0 Then GetPullcoreLocationCode = "IDLE": Exit Function
    If InStr(s, "ODTE") > 0 Or InStr(s, "OD TE") > 0 Then GetPullcoreLocationCode = "ODTE": Exit Function
    If InStr(s, "ODLE") > 0 Or InStr(s, "OD LE") > 0 Then GetPullcoreLocationCode = "ODLE": Exit Function

    If InStr(s, "ID") > 0 And InStr(s, "OD") = 0 Then GetPullcoreLocationCode = "ID": Exit Function
    If InStr(s, "OD") > 0 And InStr(s, "ID") = 0 Then GetPullcoreLocationCode = "OD": Exit Function

    Dim toks() As String
    toks = Split(s, " ")

    Dim i As Long
    For i = LBound(toks) To UBound(toks)
        If toks(i) = "TE" Then GetPullcoreLocationCode = "TE": Exit Function
        If toks(i) = "LE" Then GetPullcoreLocationCode = "LE": Exit Function
    Next i

    GetPullcoreLocationCode = ""
End Function

Private Function PullcoreLocationScore(ByVal bomLoc As String, ByVal cadLoc As String) As Double
    If bomLoc = "" Or cadLoc = "" Then
        PullcoreLocationScore = 0
        Exit Function
    End If

    If bomLoc = cadLoc Then
        PullcoreLocationScore = -3#
        Exit Function
    End If

    Dim bSide As String
    Dim cSide As String

    bSide = Left(bomLoc, 2)
    cSide = Left(cadLoc, 2)

    If bSide = cSide Then
        PullcoreLocationScore = -1#
    Else
        PullcoreLocationScore = 5#
    End If
End Function

' ============================================================
' TOP/BOT INS FALLBACK FROM CAD GEOMETRY
' ============================================================

Private Sub AddMissingTopBotInsFromCadGeometry()
On Error GoTo ErrHandler

    Dim haveTop As Boolean
    Dim haveBot As Boolean

    haveTop = ExportQuoteExists("TOP INS")
    haveBot = ExportQuoteExists("BOT INS")

    If haveTop And haveBot Then
        LogLine "TOP/BOT INS already present. Geometry fallback not needed."
        Exit Sub
    End If

    Dim candidates As Collection
    Set candidates = FindInsertCadCandidates()

    If candidates Is Nothing Or candidates.count = 0 Then
        LogLine "No insert geometry candidates found."
        Exit Sub
    End If

    LogQuarterInchCadCandidates candidates

    Dim idxTop As Long
    Dim idxBot As Long

    idxTop = 0
    idxBot = 0

    Dim tgtLarge As Double
    Dim tgtSmall As Double

    If ComputeInsTargetFromHolderAndPot(tgtLarge, tgtSmall) Then
        LogLine "INS target from formula: " & _
                FormatNumberForCsv(tgtLarge) & " x " & FormatNumberForCsv(tgtSmall) & " x 0.25"

        FindInsPairByTargetSize candidates, tgtLarge, tgtSmall, idxTop, idxBot
    End If

    If idxTop = 0 Or idxBot = 0 Then
        FindLargestSameSizeQuarterInchPairCandidates candidates, idxTop, idxBot
    End If

    If idxTop = 0 Or idxBot = 0 Then
        FindLightestAndHeaviestCandidate candidates, idxTop, idxBot
    End If

    If idxTop > 0 And haveTop = False Then
        AddExportRowFromCadPart "TOP INS", idxTop
        parts(idxTop).UsedForBomMatch = True
        LogLine "Added TOP INS from CAD geometry: " & parts(idxTop).componentName
    End If

    If idxBot > 0 And haveBot = False Then
        AddExportRowFromCadPart "BOT INS", idxBot
        parts(idxBot).UsedForBomMatch = True
        LogLine "Added BOT INS from CAD geometry: " & parts(idxBot).componentName
    End If

    Exit Sub

ErrHandler:
    LogLine "AddMissingTopBotInsFromCadGeometry error: " & Err.Description
End Sub

Private Function ComputeInsTargetFromHolderAndPot(ByRef tgtLarge As Double, ByRef tgtSmall As Double) As Boolean
On Error GoTo ErrHandler

    Dim hL As Double, hW As Double, hT As Double
    Dim pL As Double, pW As Double, pT As Double

    If Not GetMatchedDimsByQuote("ID HOLDER", hL, hW, hT) Then
        If Not GetMatchedDimsByQuote("OD HOLDER", hL, hW, hT) Then Exit Function
    End If

    If Not GetMatchedDimsByQuote("ID POT BLOCK", pL, pW, pT) Then
        If Not GetMatchedDimsByQuote("OD POT BLOCK", pL, pW, pT) Then
            If Not GetMatchedDimsByQuote("POT BLOCK", pL, pW, pT) Then Exit Function
        End If
    End If

    tgtLarge = hL
    tgtSmall = hW + pT

    ComputeInsTargetFromHolderAndPot = (tgtLarge > 0 And tgtSmall > 0)
    Exit Function

ErrHandler:
    ComputeInsTargetFromHolderAndPot = False
End Function

Private Function GetMatchedDimsByQuote(ByVal quoteName As String, _
                                       ByRef L As Double, ByRef W As Double, ByRef T As Double) As Boolean
On Error GoTo ErrHandler

    Dim key As String
    key = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = key Then
            If ExportRows(i).Length > 0 Then
                L = ExportRows(i).Length
                W = ExportRows(i).Width
                T = ExportRows(i).Thickness
            Else
                L = ExportRows(i).BomLength
                W = ExportRows(i).BomWidth
                T = ExportRows(i).BomThickness
            End If

            GetMatchedDimsByQuote = (L > 0)
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    GetMatchedDimsByQuote = False
End Function

Private Sub FindInsPairByTargetSize(ByVal candidates As Collection, _
                                    ByVal tgtLarge As Double, ByVal tgtSmall As Double, _
                                    ByRef idxTop As Long, ByRef idxBot As Long)
On Error GoTo ErrHandler

    Dim best1 As Long
    Dim best2 As Long
    Dim d1 As Double
    Dim d2 As Double

    best1 = 0
    best2 = 0
    d1 = 1E+99
    d2 = 1E+99

    Dim n As Long
    Dim idx As Long
    Dim cL As Double
    Dim cS As Double
    Dim dist As Double

    For n = 1 To candidates.count

        idx = CLng(candidates(n))

        cL = parts(idx).Length
        cS = parts(idx).Width

        dist = Abs(cL - tgtLarge) + Abs(cS - tgtSmall)

        If dist <= 2# Then
            If dist < d1 Then
                d2 = d1
                best2 = best1
                d1 = dist
                best1 = idx
            ElseIf dist < d2 Then
                d2 = dist
                best2 = idx
            End If
        End If

    Next n

    If best1 > 0 And best2 > 0 Then

        If parts(best1).massValue <= parts(best2).massValue Then
            idxTop = best1
            idxBot = best2
        Else
            idxTop = best2
            idxBot = best1
        End If

    Else

        idxTop = best1
        idxBot = best2

    End If

    Exit Sub

ErrHandler:
    LogLine "FindInsPairByTargetSize error: " & Err.Description
End Sub

Private Function FindInsertCadCandidates() As Collection
On Error GoTo ErrHandler

    Dim result As New Collection

    Dim i As Long

    For i = 1 To PartCount
        If parts(i).UsedForBomMatch = False Then
            If IsInsertCadGeometryCandidate(i) Then
                result.Add i
            End If
        End If
    Next i

    Set FindInsertCadCandidates = result
    Exit Function

ErrHandler:
    Set FindInsertCadCandidates = New Collection
End Function

Private Function IsInsertCadGeometryCandidate(ByVal cadIdx As Long) As Boolean
    If cadIdx <= 0 Then Exit Function

    If Abs(parts(cadIdx).Thickness - QUARTER_INCH_THICKNESS) <= INSERT_THICKNESS_TOL Then
        IsInsertCadGeometryCandidate = True
        Exit Function
    End If

    Dim hay As String
    hay = UCase(parts(cadIdx).cleanName & " " & parts(cadIdx).componentName)

    If InStr(hay, "INS") > 0 Or InStr(hay, "INSERT") > 0 Then
        IsInsertCadGeometryCandidate = True
    End If
End Function

Private Sub FindLargestSameSizeQuarterInchPairCandidates(ByVal candidates As Collection, _
                                                         ByRef idxA As Long, _
                                                         ByRef idxB As Long)
On Error GoTo ErrHandler

    idxA = 0
    idxB = 0

    If candidates Is Nothing Then Exit Sub
    If candidates.count < 2 Then Exit Sub

    Dim bestVol As Double
    bestVol = -1

    Dim i As Long
    Dim j As Long
    Dim a As Long
    Dim b As Long

    For i = 1 To candidates.count - 1

        a = CLng(candidates(i))

        For j = i + 1 To candidates.count

            b = CLng(candidates(j))

            If SameWidthLengthPair(a, b) Then

                Dim vol As Double
                vol = parts(a).BBoxVolume + parts(b).BBoxVolume

                If vol > bestVol Then
                    bestVol = vol
                    idxA = a
                    idxB = b
                End If

            End If

        Next j

    Next i

    Exit Sub

ErrHandler:
    LogLine "FindLargestSameSizeQuarterInchPairCandidates error: " & Err.Description
End Sub

Private Function SameWidthLengthPair(ByVal a As Long, ByVal b As Long) As Boolean
    If a <= 0 Or b <= 0 Then Exit Function

    If Abs(parts(a).Length - parts(b).Length) <= INSERT_LENGTH_MATCH_TOL Then
        If Abs(parts(a).Width - parts(b).Width) <= INSERT_WIDTH_MATCH_TOL Then
            SameWidthLengthPair = True
        End If
    End If
End Function

Private Sub FindLightestAndHeaviestCandidate(ByVal candidates As Collection, _
                                             ByRef lightIdx As Long, _
                                             ByRef heavyIdx As Long)
On Error GoTo ErrHandler

    lightIdx = 0
    heavyIdx = 0

    If candidates Is Nothing Then Exit Sub
    If candidates.count = 0 Then Exit Sub

    Dim i As Long
    Dim c As Long
    Dim lightMass As Double
    Dim heavyMass As Double

    lightMass = 1E+15
    heavyMass = -1

    For i = 1 To candidates.count

        c = CLng(candidates(i))

        If parts(c).massValue < lightMass Then
            lightMass = parts(c).massValue
            lightIdx = c
        End If

        If parts(c).massValue > heavyMass Then
            heavyMass = parts(c).massValue
            heavyIdx = c
        End If

    Next i

    If lightIdx = heavyIdx Then heavyIdx = 0

    Exit Sub

ErrHandler:
    LogLine "FindLightestAndHeaviestCandidate error: " & Err.Description
End Sub

Private Sub AddExportRowFromCadPart(ByVal quoteName As String, ByVal cadIdx As Long)
On Error GoTo ErrHandler

    If cadIdx <= 0 Then Exit Sub

    ExportCount = ExportCount + 1
    ReDim Preserve ExportRows(1 To ExportCount)

    ExportRows(ExportCount).quoteName = quoteName
    ExportRows(ExportCount).Quantity = parts(cadIdx).Quantity
    ExportRows(ExportCount).material = DEFAULT_STEEL_TYPE

    ExportRows(ExportCount).CadPartIndex = cadIdx
    ExportRows(ExportCount).HasCad = True

    ExportRows(ExportCount).Thickness = parts(cadIdx).Thickness
    ExportRows(ExportCount).Width = parts(cadIdx).Width
    ExportRows(ExportCount).Length = parts(cadIdx).Length

    ExportRows(ExportCount).BomThickness = 0
    ExportRows(ExportCount).BomWidth = 0
    ExportRows(ExportCount).BomLength = 0
    ExportRows(ExportCount).HasBomDims = False

    ExportRows(ExportCount).Status = "FROM CAD GEOMETRY"
    Exit Sub

ErrHandler:
    LogLine "AddExportRowFromCadPart error: " & Err.Description
End Sub

Private Function ExportQuoteExists(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Dim i As Long

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k Then
            ExportQuoteExists = True
            Exit Function
        End If
    Next i
End Function

Private Sub LogQuarterInchCadCandidates(ByVal candidates As Collection)
On Error Resume Next

    If candidates Is Nothing Then Exit Sub

    Dim i As Long
    Dim c As Long

    LogLine "Insert geometry candidates: " & candidates.count

    For i = 1 To candidates.count
        c = CLng(candidates(i))

        LogLine "  CAND: " & parts(c).componentName & _
                " L=" & FormatNumberForCsv(parts(c).Length) & _
                " W=" & FormatNumberForCsv(parts(c).Width) & _
                " T=" & FormatNumberForCsv(parts(c).Thickness) & _
                " mass=" & FormatNumberForCsv(parts(c).massValue)
    Next i
End Sub

' ============================================================
' J BLOCK PACKAGE
' ============================================================

Private Sub ExportJBlockPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    Dim jbIdx As Long
    jbIdx = FindBestJBlockIndex()

    If jbIdx <= 0 Then
        LogLine "J BLOCK package: no J BLOCK CAD part found."
        Exit Sub
    End If

    Dim jbFolder As String
    jbFolder = outputFolder & "\" & CurrentJobNumber & " " & J_BLOCK_FOLDER_NAME

    EnsureFolderDeep jbFolder

    LogLine "J BLOCK CAD selected: " & parts(jbIdx).componentName

    ExportJBlockBaseExactPackage jbFolder, jbIdx

    Dim camIdx As Long
    camIdx = FindBestEjectorCamIndex()

    If camIdx > 0 Then
        LogLine "Ejector CAM CAD selected: " & parts(camIdx).componentName
        ExportJBlockCamDxfOnly jbFolder, camIdx
    Else
        LogLine "J BLOCK package: ejector CAM not found (skipping CAM DXF)."
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockPackage error: " & Err.Description
End Sub

Private Function FindBestJBlockIndex() As Long
On Error GoTo ErrHandler

    FindBestJBlockIndex = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim fbIdx As Long
    Dim fbDist As Double

    fbIdx = 0
    fbDist = 1E+99

    Dim i As Long

    For i = 1 To PartCount

        If parts(i).UsedForBomMatch Then GoTo NextPart

        Dim dist As Double

        dist = Abs(parts(i).Length - JBlockTgtL) * 2# _
             + Abs(parts(i).Width - JBlockTgtW) _
             + Abs(parts(i).Thickness - JBlockTgtT)

        If dist < fbDist Then
            fbDist = dist
            fbIdx = i
        End If

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, J_BLOCK_NAME_KEYS) Then score = score + 5000

        If IsPossibleJBlockByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then
            score = score + (1000# - dist * 100#)
        End If

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

NextPart:
    Next i

    LogJBlockCandidates

    If bestScore <= 0 Then

        If fbIdx > 0 And fbDist <= 1.6 Then
            FindBestJBlockIndex = fbIdx
            LogLine "J BLOCK fallback closest bbox: " & parts(fbIdx).componentName
        Else
            FindBestJBlockIndex = 0
        End If

    Else

        FindBestJBlockIndex = bestIdx
        LogLine "J BLOCK best match: " & parts(bestIdx).componentName

    End If

    Exit Function

ErrHandler:
    LogLine "FindBestJBlockIndex error: " & Err.Description
    FindBestJBlockIndex = 0
End Function

Private Sub LogJBlockCandidates()
On Error Resume Next

    Dim i As Long

    For i = 1 To PartCount

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName

        If ContainsAnyPipeKey(hay, J_BLOCK_NAME_KEYS) Or _
           IsPossibleJBlockByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then

            LogLine "  JBLOCK CAND: " & parts(i).componentName & _
                    " L=" & FormatNumberForCsv(parts(i).Length) & _
                    " W=" & FormatNumberForCsv(parts(i).Width) & _
                    " T=" & FormatNumberForCsv(parts(i).Thickness)
        End If

    Next i
End Sub

Private Function FindBestEjectorCamIndex() As Long
On Error GoTo ErrHandler

    FindBestEjectorCamIndex = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, EJECTOR_CAM_NAME_KEYS) Then score = score + 60
        If IsPossibleCamByDims(parts(i).Length, parts(i).Width, parts(i).Thickness) Then score = score + 30

        If InStr(UCase(hay), "PULLCORE") > 0 Or InStr(UCase(hay), "PULL CORE") > 0 Then score = score - 80

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

    Next i

    LogEjectorCamCandidates

    If bestScore <= 0 Then
        FindBestEjectorCamIndex = 0
    Else
        FindBestEjectorCamIndex = bestIdx
    End If

    Exit Function

ErrHandler:
    LogLine "FindBestEjectorCamIndex error: " & Err.Description
    FindBestEjectorCamIndex = 0
End Function

Private Sub LogEjectorCamCandidates()
On Error Resume Next

    Dim i As Long

    For i = 1 To PartCount

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName

        If ContainsAnyPipeKey(hay, EJECTOR_CAM_NAME_KEYS) Then
            LogLine "  EJ CAM CAND: " & parts(i).componentName
        End If

    Next i
End Sub

Private Sub ExportJBlockBaseExactPackage(ByVal jbFolder As String, ByVal jbIdx As Long)
On Error GoTo ErrHandler

    If jbIdx <= 0 Then Exit Sub

    Dim p As PartInfo
    p = parts(jbIdx)

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim jbToken As String
    jbToken = CleanQuoteTokenForFile(J_BLOCK_BASE_FILE_TOKEN)

    If jbToken = "" Then jbToken = "J BLOCK"

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = jbFolder & "\" & CurrentJobNumber & "_" & jbToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = jbFolder & "\" & CurrentJobNumber & "_" & jbToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Exporting J BLOCK base XT + DXF."

    If swModel.GetType = swDocASSEMBLY Then
        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, "J BLOCK", False
    ElseIf swModel.GetType = swDocPART Then
        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If
    End If

    If swModel.GetType = swDocASSEMBLY Then
        If CreateJBlockDxfFromAssemblyTopView(jbFolder, p, dxfPath) Then
            LogLine "J BLOCK DXF built from native assembly top-view path."
            Exit Sub
        Else
            LogLine "WARNING: J BLOCK native assembly DXF failed; falling back to part-based DXF."
        End If
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, "J BLOCK", _
                                  J_BLOCK_PARENT_VIEW_PRIMARY, J_BLOCK_PARENT_VIEW_FALLBACK, _
                                  True, False, False

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockBaseExactPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Function CreateJBlockDxfFromAssemblyTopView(ByVal jbFolder As String, _
                                                    ByRef p As PartInfo, _
                                                    ByVal dxfPath As String) As Boolean
On Error GoTo ErrHandler

    CreateJBlockDxfFromAssemblyTopView = False

    If swModel Is Nothing Then Exit Function
    If swModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim keepNames As Collection
    Set keepNames = New Collection
    AddUniqueComponentName keepNames, p.componentName

    Dim tempFolder As String
    Dim tempNativePath As String

    tempFolder = Environ$("TEMP") & "\CMS_JBLOCK_DXF_" & Format(Now, "yyyymmdd_hhnnss")
    EnsureFolderDeep tempFolder

    ' IMPORTANT:
    ' Use native SLDASM, not X_T, so CMS_TOP / standard views survive.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_JBLOCK_TOPVIEW_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "J BLOCK assembly-top DXF: could not isolate J BLOCK component."
        GoTo CleanExit
    End If

    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    LogLine "J BLOCK DXF: saving isolated native SLDASM for right-face center view:"
    LogLine "  " & tempNativePath

    If SaveModelCopyAs(swModel, tempNativePath) = False Then GoTo CleanExit

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "J BLOCK assembly-top DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    ' J BLOCK center/base view must be the right face. Projected views unfold from there.
    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, "J BLOCK", _
                                     "*Right", "*Right", _
                                     True, False, False, False, False

    CreateJBlockDxfFromAssemblyTopView = fso.FileExists(dxfPath)

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

    Dim fso2 As Object
    Set fso2 = CreateObject("Scripting.FileSystemObject")
    If Not fso2 Is Nothing Then
        If tempFolder <> "" Then
            If fso2.FolderExists(tempFolder) Then fso2.DeleteFolder tempFolder, True
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "CreateJBlockDxfFromAssemblyTopView error: " & Err.Description
    CreateJBlockDxfFromAssemblyTopView = False
    Resume CleanExit
End Function

Private Sub ExportJBlockCamDxfOnly(ByVal jbFolder As String, ByVal camIdx As Long)
On Error GoTo ErrHandler

    If camIdx <= 0 Then Exit Sub

    Dim p As PartInfo
    p = parts(camIdx)

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim camToken As String
    camToken = CleanQuoteTokenForFile(J_BLOCK_CAM_FILE_TOKEN)

    If camToken = "" Then camToken = "CAM"

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = jbFolder & "\" & CurrentJobNumber & "_" & camToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = jbFolder & "\" & CurrentJobNumber & "_" & camToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Exporting J BLOCK ejector CAM XT + DXF."

    If swModel.GetType = swDocASSEMBLY Then
        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, "CAM", False
    ElseIf swModel.GetType = swDocPART Then
        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, "CAM", _
                                  CMS_TOP_VIEW_NAME, "*Top", True, False, False

    Exit Sub

ErrHandler:
    LogLine "ExportJBlockCamDxfOnly error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ============================================================
' END OF PART 6 OF 8
' Paste Part 7 next.
' ============================================================
' ============================================================
' PULLCORE CAM AND KEY PACKAGE
' ============================================================

Private Sub ExportPullcoreCamKeyPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    If PullcoreMatchCount = 0 Then
        LogLine "PULLCORE package: no pullcore matches."
        Exit Sub
    End If

    Dim pcFolder As String
    pcFolder = outputFolder & "\" & CurrentJobNumber & " " & PULLCORE_CAM_KEY_FOLDER_NAME

    EnsureFolderDeep pcFolder

    Dim usedNames As Object
    Set usedNames = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex > 0 Then
            ExportPullcoreCandidateSet pcFolder, i, usedNames
        Else
            LogLine "PULLCORE skipped (no CAD): " & PullcoreMatches(i).quoteName
        End If

    Next i

    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreCamKeyPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Sub ExportPullcoreCandidateSet(ByVal pcFolder As String, _
                                       ByVal matchIdx As Long, _
                                       ByVal usedNames As Object)
On Error GoTo ErrHandler

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    If cadIdx <= 0 Then Exit Sub

    Dim exportName As String
    exportName = PullcoreMatches(matchIdx).quoteName

    If exportName = "" Then exportName = CleanPullcoreDisplayName(PullcoreMatches(matchIdx).Description)

    If exportName = "" Then
        If PullcoreMatches(matchIdx).isCam Then
            exportName = "PULLCORE CAM"
        Else
            exportName = "PULLCORE KEY"
        End If
    End If

    Dim baseKey As String
    baseKey = UCase(exportName)

    If Not usedNames Is Nothing Then

        If usedNames.Exists(baseKey) Then

            Dim n As Long
            n = CLng(usedNames(baseKey)) + 1

            usedNames(baseKey) = n
            exportName = exportName & " " & CStr(n)

        Else

            usedNames(baseKey) = 1

        End If

    End If

    LogLine "PULLCORE export name: " & exportName

    ExportOneSpecialNamedPartToFolder pcFolder, cadIdx, exportName, "PULLCORE", True
    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreCandidateSet error: " & Err.Description
End Sub

' ============================================================
' PULLCORE STOP / FLIPPER CAM COVER PACKAGE
' ============================================================

Private Sub ExportPullcoreStopPackage(ByVal outputFolder As String)
On Error GoTo ErrHandler

    If swModel Is Nothing Then Exit Sub

    Dim stopIdx As Long
    Dim coverIdx As Long

    stopIdx = FindBestPartByNameAndDims(PULLCORE_STOP_NAME_KEYS, 0, 0, 0, 0)

    coverIdx = FindBestPartByNameAndDims(FLIPPER_CAM_COVER_NAME_KEYS, _
                                         FLIPPER_CAM_COVER_TARGET_THICKNESS, _
                                         FLIPPER_CAM_COVER_TARGET_WIDTH, _
                                         FLIPPER_CAM_COVER_TARGET_LENGTH, _
                                         FLIPPER_CAM_COVER_DIM_TOL)

    If stopIdx <= 0 And coverIdx <= 0 Then
        LogLine "PULLCORE STOP package: no stop or cover found."
        Exit Sub
    End If

    Dim stopFolder As String
    stopFolder = outputFolder & "\" & CurrentJobNumber & " " & PULLCORE_STOP_FOLDER_NAME

    EnsureFolderDeep stopFolder

    If stopIdx > 0 Then
        LogLine "Pullcore STOP CAD: " & parts(stopIdx).componentName
        ExportOneSpecialNamedPartToFolder stopFolder, stopIdx, "PULLCORE STOP", "PULLCORE STOP", True
    End If

    If coverIdx > 0 Then
        LogLine "Flipper CAM COVER CAD: " & parts(coverIdx).componentName
        ExportOneSpecialNamedPartToFolder stopFolder, coverIdx, "FLIPPER CAM COVER PLATE", "FLIPPER CAM COVER PLATE", True
    End If

    Exit Sub

ErrHandler:
    LogLine "ExportPullcoreStopPackage error: " & Err.Description
    On Error Resume Next
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

Private Function FindBestPartByNameAndDims(ByVal pipeKeys As String, _
                                           ByVal targetT As Double, _
                                           ByVal targetW As Double, _
                                           ByVal targetL As Double, _
                                           ByVal tol As Double) As Long
On Error GoTo ErrHandler

    FindBestPartByNameAndDims = 0

    Dim bestIdx As Long
    Dim bestScore As Double

    bestIdx = 0
    bestScore = -1

    Dim i As Long

    For i = 1 To PartCount

        Dim score As Double
        score = 0

        Dim hay As String
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath

        If ContainsAnyPipeKey(hay, pipeKeys) Then score = score + 100

        If targetT > 0 And targetW > 0 And targetL > 0 And tol > 0 Then
            If Abs(parts(i).Thickness - targetT) <= tol Then
                If Abs(parts(i).Width - targetW) <= tol Then
                    If Abs(parts(i).Length - targetL) <= tol Then
                        score = score + 40
                    End If
                End If
            End If
        End If

        If score > bestScore Then
            bestScore = score
            bestIdx = i
        End If

    Next i

    If bestScore <= 0 Then
        FindBestPartByNameAndDims = 0
    Else
        FindBestPartByNameAndDims = bestIdx
    End If

    Exit Function

ErrHandler:
    LogLine "FindBestPartByNameAndDims error: " & Err.Description
    FindBestPartByNameAndDims = 0
End Function

Private Sub ExportOneSpecialNamedPartToFolder(ByVal folderPath As String, _
                                              ByVal cadIdx As Long, _
                                              ByVal exportName As String, _
                                              ByVal dxfQuoteName As String, _
                                              ByVal makeWireframe As Boolean)
On Error GoTo ErrHandler

    If cadIdx <= 0 Then Exit Sub

    EnsureFolderDeep folderPath

    Dim p As PartInfo
    p = parts(cadIdx)

    Dim quoteToken As String
    quoteToken = CleanQuoteTokenForFile(exportName)

    If quoteToken = "" Then quoteToken = "ITEM"

    Dim custToken As String
    Dim dateToken As String

    custToken = CleanFileName(CustomerNumber)
    dateToken = CleanFileName(DateCode)

    If custToken = "" Then custToken = "UNKNOWN"
    If dateToken = "" Then dateToken = Format(Date, "mm-dd-yyyy")

    Dim xtPath As String
    Dim dxfPath As String

    xtPath = folderPath & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".x_t"
    dxfPath = folderPath & "\" & CurrentJobNumber & "_" & quoteToken & "_" & custToken & "_" & dateToken & ".dxf"

    xtPath = GetUniqueFilePath(xtPath)
    dxfPath = GetUniqueFilePath(dxfPath)

    LogLine "Special export: " & exportName & " -> " & xtPath

    If swModel.GetType = swDocASSEMBLY Then

        ExportAssemblyComponentBySuppressRestWithOptionalDxf swModel, p, xtPath, dxfPath, dxfQuoteName, False

    ElseIf swModel.GetType = swDocPART Then

        If p.isBodyOnly Then
            ShowOnlyPartBody swModel, p.bodyName
            SaveModelAs swModel, xtPath
            ShowAllPartBodies swModel
        Else
            SaveModelAs swModel, xtPath
        End If

    End If

    Dim oldStraightenAngle As Double
    Dim oldPullcoreCadIdx As Long

    oldStraightenAngle = CurrentDxfStraightenAngleRad
    oldPullcoreCadIdx = CurrentPullcoreStraightenCadIndex

    CurrentDxfStraightenAngleRad = 0#
    CurrentPullcoreStraightenCadIndex = 0

    If STRAIGHTEN_PULLCORE_DXF And UCase(Trim(dxfQuoteName)) = "PULLCORE" Then
        CurrentPullcoreStraightenCadIndex = cadIdx

        LogLine "PULLCORE: using matched CAD component for straightening: " & _
                parts(cadIdx).componentName
    End If

    CreateProjectedDxfFromXtPath xtPath, dxfPath, dxfQuoteName, _
                                  "*Front", "*Front", makeWireframe, False, False, False, False

    CurrentDxfStraightenAngleRad = oldStraightenAngle
    CurrentPullcoreStraightenCadIndex = oldPullcoreCadIdx

    Exit Sub

ErrHandler:
    LogLine "ExportOneSpecialNamedPartToFolder error (" & exportName & "): " & Err.Description
    On Error Resume Next
    CurrentDxfStraightenAngleRad = 0#
    CurrentPullcoreStraightenCadIndex = 0
    UnsuppressAllAssemblyComponents swModel
    ShowAllAssemblyComponents swModel
End Sub

' ============================================================
' PULLCORE BEST-FIT BBOX / STRAIGHTENING
' ============================================================

Private Function TryGetPullcoreBestFitDims(ByVal cadIdx As Long, _
                                           ByRef outL As Double, _
                                           ByRef outW As Double, _
                                           ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPullcoreBestFitDims = False

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If PullcoreBestFitDimCache Is Nothing Then
        Set PullcoreBestFitDimCache = CreateObject("Scripting.Dictionary")
    End If

    Dim cacheKey As String
    cacheKey = CStr(cadIdx)

    If PullcoreBestFitDimCache.Exists(cacheKey) Then
        Dim cached() As String
        cached = Split(CStr(PullcoreBestFitDimCache(cacheKey)), "|")

        If UBound(cached) >= 2 Then
            outL = CDbl(cached(0))
            outW = CDbl(cached(1))
            outT = CDbl(cached(2))

            TryGetPullcoreBestFitDims = (outL > 0 And outW > 0 And outT > 0)
            Exit Function
        End If
    End If

    Dim partModel As Object
    Set partModel = GetModelDocForCadIndex(cadIdx)

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    If TryCreateAndReadBestFitBoundingBox(partModel, outL, outW, outT) Then

        SortThreeDimensions outL, outW, outT, outL, outW, outT

        outL = Round(outL, DIM_DECIMALS)
        outW = Round(outW, DIM_DECIMALS)
        outT = Round(outT, DIM_DECIMALS)

        PullcoreBestFitDimCache(cacheKey) = CStr(outL) & "|" & CStr(outW) & "|" & CStr(outT)

        LogLine "PULLCORE best-fit bbox for " & parts(cadIdx).componentName & " = " & _
                FormatNumberForCsv(outL) & "/" & FormatNumberForCsv(outW) & "/" & FormatNumberForCsv(outT)

        TryGetPullcoreBestFitDims = True

    End If

    Exit Function

ErrHandler:
    LogLine "TryGetPullcoreBestFitDims error (cadIdx " & cadIdx & "): " & Err.Description
    TryGetPullcoreBestFitDims = False
End Function

Private Function TryComputePullcoreStraightenAngleFromMatchedCad(ByVal cadIdx As Long, _
                                                                 ByRef detectedDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreStraightenAngleFromMatchedCad = False
    detectedDegOut = 0#

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    Dim matchIdx As Long
    matchIdx = FindPullcoreMatchIndexByCadIndex(cadIdx)

    If matchIdx > 0 Then

        If Abs(PullcoreMatches(matchIdx).DetectedAngleDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            detectedDegOut = PullcoreMatches(matchIdx).DetectedAngleDeg
            LogLine "PULLCORE angle from stored original/fitted bbox report: " & _
                    Format(detectedDegOut, "0.00") & " deg"
            TryComputePullcoreStraightenAngleFromMatchedCad = True
            Exit Function
        End If

        If PullcoreMatches(matchIdx).FittedLength > 0 And _
           PullcoreMatches(matchIdx).OriginalLength > 0 Then

            detectedDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg( _
                PullcoreMatches(matchIdx).FittedLength, _
                PullcoreMatches(matchIdx).FittedWidth, _
                PullcoreMatches(matchIdx).FittedThickness, _
                PullcoreMatches(matchIdx).OriginalLength, _
                PullcoreMatches(matchIdx).OriginalWidth, _
                PullcoreMatches(matchIdx).OriginalThickness)

            If Abs(detectedDegOut) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                TryComputePullcoreStraightenAngleFromMatchedCad = True
                Exit Function
            End If

        End If

    End If

    Dim partModel As Object
    Set partModel = GetModelDocForCadIndex(cadIdx)

    Dim modelEdgeDeg As Double

    If Not partModel Is Nothing Then
        If partModel.GetType = swDocPART Then

            If TryComputePullcoreLongestModelEdgeAngleFromModel(partModel, modelEdgeDeg) Then
                If Abs(modelEdgeDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
                    detectedDegOut = modelEdgeDeg
                    TryComputePullcoreStraightenAngleFromMatchedCad = True
                    Exit Function
                End If
            End If

        End If
    End If

    Dim bestL As Double
    Dim bestW As Double
    Dim bestT As Double

    If TryGetPullcoreMatchedReportDims(cadIdx, bestL, bestW, bestT) Then
        detectedDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg(bestL, bestW, bestT, _
            parts(cadIdx).OriginalAsmLength, parts(cadIdx).OriginalAsmWidth, parts(cadIdx).OriginalAsmThickness)

        If Abs(detectedDegOut) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            TryComputePullcoreStraightenAngleFromMatchedCad = True
            Exit Function
        End If
    End If

    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreStraightenAngleFromMatchedCad error: " & Err.Description
    TryComputePullcoreStraightenAngleFromMatchedCad = False
End Function

Private Function FindPullcoreMatchIndexByCadIndex(ByVal cadIdx As Long) As Long
On Error GoTo ErrHandler

    FindPullcoreMatchIndexByCadIndex = 0

    If cadIdx <= 0 Then Exit Function

    Dim i As Long

    For i = 1 To PullcoreMatchCount
        If PullcoreMatches(i).CadPartIndex = cadIdx Then
            FindPullcoreMatchIndexByCadIndex = i
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    FindPullcoreMatchIndexByCadIndex = 0
End Function

Private Function TryGetPullcoreMatchedReportDims(ByVal cadIdx As Long, _
                                                 ByRef outL As Double, _
                                                 ByRef outW As Double, _
                                                 ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryGetPullcoreMatchedReportDims = False

    outL = 0#
    outW = 0#
    outT = 0#

    If cadIdx <= 0 Then Exit Function

    Dim i As Long

    For i = 1 To PullcoreMatchCount

        If PullcoreMatches(i).CadPartIndex = cadIdx Then

            If PullcoreMatches(i).FittedLength > 0 And _
               PullcoreMatches(i).FittedWidth > 0 And _
               PullcoreMatches(i).FittedThickness > 0 Then

                outL = PullcoreMatches(i).FittedLength
                outW = PullcoreMatches(i).FittedWidth
                outT = PullcoreMatches(i).FittedThickness

                TryGetPullcoreMatchedReportDims = True
                Exit Function

            End If

        End If

    Next i

    Exit Function

ErrHandler:
    TryGetPullcoreMatchedReportDims = False
End Function

Private Function TryComputePullcoreBestFitStraightenAngleFromNativePath(ByVal nativePath As String, _
                                                                        ByRef detectedDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreBestFitStraightenAngleFromNativePath = False
    detectedDegOut = 0#

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(nativePath) = False Then Exit Function

    Dim ext As String
    ext = LCase(fso.GetExtensionName(nativePath))

    Dim errs As Long
    Dim warns As Long
    Dim mdl As Object

    If ext = "sldasm" Then
        Set mdl = swApp.OpenDoc6(nativePath, swDocASSEMBLY, swOpenDocOptions_Silent, "", errs, warns)
    Else
        Set mdl = swApp.OpenDoc6(nativePath, swDocPART, swOpenDocOptions_Silent, "", errs, warns)
    End If

    If mdl Is Nothing Then Exit Function

    swApp.ActivateDoc3 mdl.GetTitle, False, 0, errs
    EnsureSwHidden

    Dim modelEdgeDeg As Double
    Dim mathDeg As Double

    If TryComputePullcoreLongestModelEdgeAngleFromModel(mdl, modelEdgeDeg) Then
        If Abs(modelEdgeDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            detectedDegOut = modelEdgeDeg
            TryComputePullcoreBestFitStraightenAngleFromNativePath = True
            GoTo CleanExit
        End If
    End If

    If TryComputeBBoxDimensionMathAngleFromModel(mdl, mathDeg) Then
        If Abs(mathDeg) >= PULLCORE_STRAIGHTEN_MIN_DEG Then
            detectedDegOut = mathDeg
            TryComputePullcoreBestFitStraightenAngleFromNativePath = True
            GoTo CleanExit
        End If
    End If

CleanExit:
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreBestFitStraightenAngleFromNativePath error: " & Err.Description
    On Error Resume Next
    If Not mdl Is Nothing Then swApp.CloseDoc mdl.GetTitle
    TryComputePullcoreBestFitStraightenAngleFromNativePath = False
End Function

Private Function TryComputePullcoreLongestModelEdgeAngleFromModel(ByVal mdl As Object, _
                                                                  ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputePullcoreLongestModelEdgeAngleFromModel = False
    angleDegOut = 0#

    If mdl Is Nothing Then Exit Function

    Dim bodyList As Collection
    Set bodyList = New Collection

    CollectSolidBodiesFromModel mdl, bodyList

    If bodyList.count = 0 Then Exit Function

    Dim bestInPlane As Double
    Dim bestDx As Double
    Dim bestDy As Double

    bestInPlane = -1#
    bestDx = 0#
    bestDy = 0#

    Dim bi As Long
    Dim ei As Long
    Dim body As Object
    Dim vEdges As Variant
    Dim edge As Object
    Dim curve As Object
    Dim vCP As Variant

    Dim sx As Double, sy As Double, sz As Double
    Dim ex As Double, ey As Double, ez As Double
    Dim dx As Double, dy As Double, dz As Double
    Dim inPlane As Double, fullLen As Double

    For bi = 1 To bodyList.count

        Set body = bodyList(bi)

        If Not body Is Nothing Then

            vEdges = body.GetEdges

            If IsArray(vEdges) Then

                For ei = 0 To UBound(vEdges)

                    Set edge = vEdges(ei)

                    If Not edge Is Nothing Then

                        Set curve = Nothing
                        On Error Resume Next
                        Set curve = edge.GetCurve
                        On Error GoTo ErrHandler

                        If Not curve Is Nothing Then
                            If curve.IsLine Then

                                vCP = edge.GetCurveParams2

                                If IsArray(vCP) Then
                                    If UBound(vCP) >= 5 Then

                                        sx = CDbl(vCP(0))
                                        sy = CDbl(vCP(1))
                                        sz = CDbl(vCP(2))
                                        ex = CDbl(vCP(3))
                                        ey = CDbl(vCP(4))
                                        ez = CDbl(vCP(5))

                                        dx = ex - sx
                                        dy = ey - sy
                                        dz = ez - sz

                                        inPlane = Sqr(dx * dx + dy * dy)
                                        fullLen = Sqr(dx * dx + dy * dy + dz * dz)

                                        If fullLen > 0.000001 Then
                                            If inPlane / fullLen >= 0.5 Then
                                                If inPlane > bestInPlane Then
                                                    bestInPlane = inPlane
                                                    bestDx = dx
                                                    bestDy = dy
                                                End If
                                            End If
                                        End If

                                    End If
                                End If

                            End If
                        End If

                    End If

                Next ei

            End If

        End If

    Next bi

    If bestInPlane <= 0.000001 Then Exit Function

    Dim ang As Double
    ang = Atan2Safe(bestDy, bestDx)
    ang = NormalizeAngleToPlusMinus90(ang)

    angleDegOut = RadToDeg(ang)

    TryComputePullcoreLongestModelEdgeAngleFromModel = True
    Exit Function

ErrHandler:
    LogLine "TryComputePullcoreLongestModelEdgeAngleFromModel error: " & Err.Description
    TryComputePullcoreLongestModelEdgeAngleFromModel = False
End Function

Private Sub CollectSolidBodiesFromModel(ByVal mdl As Object, ByVal outBodies As Collection)
On Error Resume Next

    If mdl Is Nothing Then Exit Sub
    If outBodies Is Nothing Then Exit Sub

    If mdl.GetType = swDocPART Then

        Dim vB As Variant
        vB = mdl.GetBodies2(swSolidBody, False)

        If IsArray(vB) Then
            Dim i As Long
            For i = 0 To UBound(vB)
                If Not vB(i) Is Nothing Then outBodies.Add vB(i)
            Next i
        End If

    ElseIf mdl.GetType = swDocASSEMBLY Then

        Dim vComps As Variant
        vComps = mdl.GetComponents(False)

        If IsArray(vComps) Then

            Dim c As Long
            Dim comp As Object
            Dim cm As Object
            Dim vCB As Variant
            Dim j As Long

            For c = 0 To UBound(vComps)

                Set comp = vComps(c)

                If Not comp Is Nothing Then
                    If comp.IsSuppressed = False Then

                        Set cm = comp.GetModelDoc2

                        If Not cm Is Nothing Then
                            If cm.GetType = swDocPART Then

                                vCB = cm.GetBodies2(swSolidBody, False)

                                If IsArray(vCB) Then
                                    For j = 0 To UBound(vCB)
                                        If Not vCB(j) Is Nothing Then outBodies.Add vCB(j)
                                    Next j
                                End If

                            End If
                        End If

                    End If
                End If

            Next c

        End If

    End If
End Sub

Private Function TryComputeBBoxDimensionMathAngleFromModel(ByVal mdl As Object, _
                                                           ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputeBBoxDimensionMathAngleFromModel = False
    angleDegOut = 0#

    If mdl Is Nothing Then Exit Function

    If mdl.GetType = swDocPART Then
        TryComputeBBoxDimensionMathAngleFromModel = _
            TryComputeBBoxDimensionMathAngleForPart(mdl, angleDegOut)
        Exit Function
    End If

    If mdl.GetType <> swDocASSEMBLY Then Exit Function

    Dim vComps As Variant
    vComps = mdl.GetComponents(False)

    If IsEmpty(vComps) Then Exit Function

    Dim bestVol As Double
    Dim bestAng As Double

    bestVol = -1#
    bestAng = 0#

    Dim i As Long
    Dim comp As Object
    Dim cm As Object
    Dim oneAng As Double
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    For i = 0 To UBound(vComps)

        Set comp = vComps(i)

        If Not comp Is Nothing Then
            If comp.IsSuppressed = False Then

                Set cm = comp.GetModelDoc2

                If Not cm Is Nothing Then
                    If cm.GetType = swDocPART Then

                        If TryComputeBBoxDimensionMathAngleForPart(cm, oneAng) Then

                            If GetPartBoundingBoxInches(cm, dx, dy, dz) Then

                                Dim vol As Double
                                vol = dx * dy * dz

                                If vol > bestVol Then
                                    bestVol = vol
                                    bestAng = oneAng
                                End If

                            End If

                        End If

                    End If
                End If

            End If
        End If

    Next i

    If bestVol > 0# Then
        angleDegOut = bestAng
        TryComputeBBoxDimensionMathAngleFromModel = True
    End If

    Exit Function

ErrHandler:
    TryComputeBBoxDimensionMathAngleFromModel = False
End Function

Private Function TryComputeBBoxDimensionMathAngleForPart(ByVal partModel As Object, _
                                                         ByRef angleDegOut As Double) As Boolean
On Error GoTo ErrHandler

    TryComputeBBoxDimensionMathAngleForPart = False
    angleDegOut = 0#

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim dx As Double
    Dim dy As Double
    Dim dz As Double

    If GetPartBoundingBoxInches(partModel, dx, dy, dz) = False Then Exit Function

    Dim normL As Double
    Dim normW As Double
    Dim normT As Double

    SortThreeDimensions dx, dy, dz, normL, normW, normT

    Dim bestL As Double
    Dim bestW As Double
    Dim bestT As Double

    If TryCreateAndReadBestFitBoundingBox(partModel, bestL, bestW, bestT) = False Then Exit Function

    SortThreeDimensions bestL, bestW, bestT, bestL, bestW, bestT

    angleDegOut = EstimatePullcoreAngleFromFittedAndOriginalDeg(bestL, bestW, bestT, normL, normW, normT)

    TryComputeBBoxDimensionMathAngleForPart = (Abs(angleDegOut) > 0#)
    Exit Function

ErrHandler:
    TryComputeBBoxDimensionMathAngleForPart = False
End Function

Private Function SearchBoxRotationDeg(ByVal bestW As Double, ByVal bestT As Double, _
                                      ByVal normW As Double, ByVal normT As Double) As Double
On Error GoTo ErrHandler

    Dim bestAng As Double
    Dim bestErr As Double

    bestAng = 0#
    bestErr = 1E+99

    Dim a As Double
    Dim rad As Double
    Dim pW As Double
    Dim pT As Double
    Dim e As Double

    a = 0#

    Do While a <= 90#

        rad = a * PI_VALUE / 180#

        pW = bestW * Cos(rad) + bestT * Sin(rad)
        pT = bestW * Sin(rad) + bestT * Cos(rad)

        e = Abs(pW - normW) + Abs(pT - normT)

        If e < bestErr Then
            bestErr = e
            bestAng = a
        End If

        a = a + 1#

    Loop

    Dim lo As Double
    Dim hi As Double

    lo = bestAng - 1#
    hi = bestAng + 1#

    If lo < 0# Then lo = 0#
    If hi > 90# Then hi = 90#

    a = lo

    Do While a <= hi

        rad = a * PI_VALUE / 180#

        pW = bestW * Cos(rad) + bestT * Sin(rad)
        pT = bestW * Sin(rad) + bestT * Cos(rad)

        e = Abs(pW - normW) + Abs(pT - normT)

        If e < bestErr Then
            bestErr = e
            bestAng = a
        End If

        a = a + 0.1

    Loop

    SearchBoxRotationDeg = bestAng
    Exit Function

ErrHandler:
    SearchBoxRotationDeg = 0#
End Function

Private Function GetModelDocForCadIndex(ByVal cadIdx As Long) As Object
On Error GoTo ErrHandler

    Set GetModelDocForCadIndex = Nothing

    If cadIdx <= 0 Or cadIdx > PartCount Then Exit Function

    If Not swModel Is Nothing Then

        If swModel.GetType = swDocASSEMBLY Then

            Dim swComp As Object
            Set swComp = FindAssemblyComponentByName(swModel, parts(cadIdx).componentName)

            If Not swComp Is Nothing Then
                Set GetModelDocForCadIndex = swComp.GetModelDoc2
                If Not GetModelDocForCadIndex Is Nothing Then Exit Function
            End If

        ElseIf swModel.GetType = swDocPART Then

            Set GetModelDocForCadIndex = swModel
            Exit Function

        End If

    End If

    If parts(cadIdx).filePath <> "" Then

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")

        If fso.FileExists(parts(cadIdx).filePath) Then

            Dim errs As Long
            Dim warns As Long

            Set GetModelDocForCadIndex = swApp.OpenDoc6(parts(cadIdx).filePath, swDocPART, _
                                          swOpenDocOptions_Silent + swOpenDocOptions_ReadOnly, _
                                          parts(cadIdx).configName, errs, warns)

        End If

    End If

    Exit Function

ErrHandler:
    LogLine "GetModelDocForCadIndex error: " & Err.Description
    Set GetModelDocForCadIndex = Nothing
End Function

Private Function TryCreateAndReadBestFitBoundingBox(ByVal partModel As Object, _
                                                    ByRef outL As Double, _
                                                    ByRef outW As Double, _
                                                    ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryCreateAndReadBestFitBoundingBox = False

    outL = 0#
    outW = 0#
    outT = 0#

    If partModel Is Nothing Then Exit Function
    If partModel.GetType <> swDocPART Then Exit Function

    Dim bboxFeat As Object
    Set bboxFeat = Nothing

    partModel.ClearSelection2 True

    Dim featMgr As Object
    Set featMgr = partModel.FeatureManager

    If Not featMgr Is Nothing Then
        On Error Resume Next
        Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False, Nothing)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False, Nothing, False)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox(0, False, False)
        If bboxFeat Is Nothing Then Set bboxFeat = featMgr.InsertGlobalBoundingBox2(0, False, False, Nothing)
        On Error GoTo ErrHandler
    End If

    Dim gotDims As Boolean
    gotDims = False

    If Not bboxFeat Is Nothing Then

        partModel.ForceRebuild3 False

        gotDims = TryReadBoundingBoxCustomProps(partModel, outL, outW, outT)

        If gotDims = False Then
            gotDims = TryReadBoundingBoxFeatureSketchDims(bboxFeat, outL, outW, outT)
        End If

        On Error Resume Next
        partModel.ClearSelection2 True
        bboxFeat.Select2 False, 0
        partModel.EditDelete
        partModel.ClearSelection2 True
        On Error GoTo ErrHandler

    End If

    If gotDims = False Then

        Dim dx As Double
        Dim dy As Double
        Dim dz As Double

        If GetPartBoundingBoxInches(partModel, dx, dy, dz) Then
            SortThreeDimensions dx, dy, dz, outL, outW, outT
            gotDims = True
        End If

    End If

    TryCreateAndReadBestFitBoundingBox = (gotDims And outL > 0 And outW > 0 And outT > 0)
    Exit Function

ErrHandler:
    LogLine "TryCreateAndReadBestFitBoundingBox error: " & Err.Description

    On Error Resume Next
    If Not bboxFeat Is Nothing Then
        partModel.ClearSelection2 True
        bboxFeat.Select2 False, 0
        partModel.EditDelete
        partModel.ClearSelection2 True
    End If

    TryCreateAndReadBestFitBoundingBox = False
End Function

Private Function TryReadBoundingBoxCustomProps(ByVal partModel As Object, _
                                               ByRef outL As Double, _
                                               ByRef outW As Double, _
                                               ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryReadBoundingBoxCustomProps = False

    If partModel Is Nothing Then Exit Function

    Dim ext As Object
    Set ext = partModel.Extension

    If ext Is Nothing Then Exit Function

    Dim cpm As Object
    Set cpm = ext.CustomPropertyManager("")

    If cpm Is Nothing Then Exit Function

    Dim L As Double
    Dim W As Double
    Dim T As Double

    L = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxLength", "Total Bounding Box Length", "BoundingBoxLength"))
    W = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxWidth", "Total Bounding Box Width", "BoundingBoxWidth"))
    T = ReadBoundingBoxPropInches(cpm, Array("SW-BoundingBoxThickness", "Total Bounding Box Thickness", "BoundingBoxThickness", "SW-BoundingBoxHeight"))

    If L > 0 And W > 0 And T > 0 Then
        outL = L
        outW = W
        outT = T

        TryReadBoundingBoxCustomProps = True
    End If

    Exit Function

ErrHandler:
    TryReadBoundingBoxCustomProps = False
End Function

Private Function ReadBoundingBoxPropInches(ByVal cpm As Object, ByVal names As Variant) As Double
On Error Resume Next

    ReadBoundingBoxPropInches = 0#

    Dim nm As Variant
    Dim valOut As String
    Dim resolved As String

    For Each nm In names

        valOut = ""
        resolved = ""

        cpm.Get4 CStr(nm), False, valOut, resolved

        Dim useStr As String

        useStr = resolved
        If Trim(useStr) = "" Then useStr = valOut

        If Trim(useStr) <> "" Then

            Dim v As Double
            v = ParseLeadingNumber(useStr)

            If v > 0 Then
                If v < 3# Then
                    ReadBoundingBoxPropInches = v * INCHES_PER_METER
                Else
                    ReadBoundingBoxPropInches = v
                End If

                Exit Function
            End If

        End If

    Next nm
End Function

Private Function ParseLeadingNumber(ByVal s As String) As Double
On Error Resume Next

    Dim i As Long
    Dim ch As String
    Dim token As String

    token = ""

    For i = 1 To Len(s)

        ch = Mid(s, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            token = token & ch
        ElseIf token <> "" Then
            Exit For
        End If

    Next i

    If token <> "" And IsNumeric(token) Then ParseLeadingNumber = CDbl(token)
End Function

Private Function TryReadBoundingBoxFeatureSketchDims(ByVal bboxFeat As Object, _
                                                     ByRef outL As Double, _
                                                     ByRef outW As Double, _
                                                     ByRef outT As Double) As Boolean
On Error GoTo ErrHandler

    TryReadBoundingBoxFeatureSketchDims = False

    outL = 0#
    outW = 0#
    outT = 0#

    If bboxFeat Is Nothing Then Exit Function

    Dim lengths As Collection
    Set lengths = New Collection

    CollectSketchSegmentLengthsFromFeature bboxFeat, lengths

    If lengths.count < 3 Then Exit Function

    Dim a As Double
    Dim b As Double
    Dim c As Double

    If GetThreeUniqueBoxEdgeLengths(lengths, a, b, c) = False Then Exit Function

    SortThreeDimensions a, b, c, outL, outW, outT

    TryReadBoundingBoxFeatureSketchDims = (outL > 0 And outW > 0 And outT > 0)
    Exit Function

ErrHandler:
    TryReadBoundingBoxFeatureSketchDims = False
End Function

Private Sub CollectSketchSegmentLengthsFromFeature(ByVal feat As Object, ByVal lengths As Collection)
On Error Resume Next

    If feat Is Nothing Then Exit Sub

    Dim spec As Object
    Set spec = feat.GetSpecificFeature2

    If Not spec Is Nothing Then

        Dim segs As Variant
        segs = spec.GetSketchSegments

        If IsEmpty(segs) = False Then

            Dim i As Long
            Dim seg As Object

            For i = 0 To UBound(segs)

                Set seg = segs(i)

                If Not seg Is Nothing Then

                    Dim p1 As Object
                    Dim p2 As Object

                    Set p1 = seg.GetStartPoint2
                    Set p2 = seg.GetEndPoint2

                    If Not p1 Is Nothing And Not p2 Is Nothing Then

                        Dim dx As Double
                        Dim dy As Double
                        Dim dz As Double
                        Dim distIn As Double

                        dx = CDbl(p2.x) - CDbl(p1.x)
                        dy = CDbl(p2.y) - CDbl(p1.y)
                        dz = CDbl(p2.z) - CDbl(p1.z)

                        distIn = Sqr(dx * dx + dy * dy + dz * dz) * INCHES_PER_METER

                        If distIn > 0.01 Then lengths.Add distIn

                    End If

                End If

            Next i

        End If

    End If

    Dim subFeat As Object
    Set subFeat = feat.GetFirstSubFeature

    Do While Not subFeat Is Nothing
        CollectSketchSegmentLengthsFromFeature subFeat, lengths
        Set subFeat = subFeat.GetNextSubFeature
    Loop
End Sub

Private Function GetThreeUniqueBoxEdgeLengths(ByVal lengths As Collection, _
                                              ByRef a As Double, ByRef b As Double, ByRef c As Double) As Boolean
On Error GoTo ErrHandler

    GetThreeUniqueBoxEdgeLengths = False

    a = 0#
    b = 0#
    c = 0#

    If lengths Is Nothing Then Exit Function
    If lengths.count < 3 Then Exit Function

    Dim uniques As Collection
    Set uniques = New Collection

    Dim i As Long

    For i = 1 To lengths.count
        AddUniqueLength uniques, CDbl(lengths(i)), 0.02
    Next i

    If uniques.count < 3 Then Exit Function

    Dim arr() As Double
    ReDim arr(1 To uniques.count)

    For i = 1 To uniques.count
        arr(i) = CDbl(uniques(i))
    Next i

    Dim j As Long
    Dim tmp As Double

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) > arr(i) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i

    a = arr(1)
    b = arr(2)
    c = arr(3)

    GetThreeUniqueBoxEdgeLengths = True
    Exit Function

ErrHandler:
    GetThreeUniqueBoxEdgeLengths = False
End Function

Private Sub AddUniqueLength(ByVal uniques As Collection, ByVal val As Double, ByVal tol As Double)
On Error Resume Next

    If val <= 0.01 Then Exit Sub

    Dim i As Long

    For i = 1 To uniques.count
        If Abs(CDbl(uniques(i)) - val) <= tol Then Exit Sub
    Next i

    uniques.Add val
End Sub

' ============================================================
' CSV REPORTS
' ============================================================

Private Sub WritePartDimensionCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "Index,Component,CleanName,Length,Width,Thickness,OriginalAsmLength,OriginalAsmWidth,OriginalAsmThickness,HasOriginalAsmBBox,BBoxVolume,Mass,Qty,BodyOnly"

    Dim i As Long

    For i = 1 To PartCount
        Print #f, CsvText(CStr(i)) & "," & _
                  CsvText(parts(i).componentName) & "," & _
                  CsvText(parts(i).cleanName) & "," & _
                  FormatNumberForCsv(parts(i).Length) & "," & _
                  FormatNumberForCsv(parts(i).Width) & "," & _
                  FormatNumberForCsv(parts(i).Thickness) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmLength) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmWidth) & "," & _
                  FormatNumberForCsv(parts(i).OriginalAsmThickness) & "," & _
                  CStr(parts(i).hasOriginalAsmBBox) & "," & _
                  FormatNumberForCsv(parts(i).BBoxVolume) & "," & _
                  FormatNumberForCsv(parts(i).massValue) & "," & _
                  CStr(parts(i).Quantity) & "," & _
                  CStr(parts(i).isBodyOnly)
    Next i

    Close #f

    LogLine "Wrote CAD dimension CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WritePartDimensionCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobSignatureCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "JobNumber,CustomerNumber,DateCode,ComponentRole,QuoteName,CadComponent,CleanName,Length,Width,Thickness,Mass,CenterX,CenterY,CenterZ,HasCenter,Status"

    WriteJobSignatureRow f, "TCP", "TCP", "TCP|TOP SMED|TOP CLAMPING|TOP CLAMPING PLATE|ID SMED"
    WriteJobSignatureRow f, "BCP", "BCP", "BCP|BOTTOM SMED|BOT SMED|BOTTOM CLAMPING|BOTTOM CLAMPING PLATE|OD SMED"
    WriteJobSignatureRow f, "ID HOLDER", "ID HOLDER", ID_HOLDER_KEYS
    WriteJobSignatureRow f, "OD HOLDER", "OD HOLDER", OD_HOLDER_KEYS
    WriteJobSignatureRow f, "ID POT", "ID POT BLOCK", "ID POT BLOCK|ID POT|TOP POT BLOCK|TOP POT|TCP POT BLOCK|TCP POT"
    WriteJobSignatureRow f, "OD POT", "OD POT BLOCK", "OD POT BLOCK|OD POT|BOTTOM POT BLOCK|BOT POT BLOCK|BOTTOM POT|BOT POT|BCP POT BLOCK|BCP POT"

    Close #f

    LogLine "Wrote job signature CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WriteJobSignatureCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WriteJobSignatureRow(ByVal f As Integer, _
                                 ByVal roleName As String, _
                                 ByVal quoteName As String, _
                                 ByVal fallbackKeys As String)
On Error Resume Next

    Dim cadIdx As Long
    cadIdx = FindCadIndexFromExportQuote(quoteName)

    If cadIdx <= 0 Then
        cadIdx = FindCadPartIndexByQuoteOrKeys(quoteName, fallbackKeys)
    End If

    Dim cadComponent As String
    Dim cleanName As String
    Dim statusText As String
    Dim L As Double
    Dim W As Double
    Dim T As Double
    Dim massVal As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim centerZ As Double
    Dim hasCenter As Boolean

    cadComponent = ""
    cleanName = ""
    statusText = "NO CAD MATCH"
    L = 0#: W = 0#: T = 0#
    massVal = 0#
    centerX = 0#: centerY = 0#: centerZ = 0#
    hasCenter = False

    If cadIdx > 0 And cadIdx <= PartCount Then
        cadComponent = parts(cadIdx).componentName
        cleanName = parts(cadIdx).cleanName
        L = parts(cadIdx).Length
        W = parts(cadIdx).Width
        T = parts(cadIdx).Thickness
        massVal = parts(cadIdx).massValue
        hasCenter = parts(cadIdx).hasAsmCenter
        centerX = parts(cadIdx).AsmCenterX
        centerY = parts(cadIdx).AsmCenterY
        centerZ = parts(cadIdx).AsmCenterZ
        statusText = "OK"
    End If

    Print #f, CsvText(CurrentJobNumber) & "," & _
              CsvText(CustomerNumber) & "," & _
              CsvText(DateCode) & "," & _
              CsvText(roleName) & "," & _
              CsvText(quoteName) & "," & _
              CsvText(cadComponent) & "," & _
              CsvText(cleanName) & "," & _
              FormatNumberForCsv(L) & "," & _
              FormatNumberForCsv(W) & "," & _
              FormatNumberForCsv(T) & "," & _
              FormatNumberForCsv(massVal) & "," & _
              FormatNumberForCsv(centerX) & "," & _
              FormatNumberForCsv(centerY) & "," & _
              FormatNumberForCsv(centerZ) & "," & _
              CStr(hasCenter) & "," & _
              CsvText(statusText)
End Sub

Private Sub WriteExportCheckCsv(ByVal csvPath As String)
On Error GoTo ErrHandler

    csvPath = GetWritableCsvPath(csvPath)

    Dim f As Integer
    f = FreeFile

    Open csvPath For Output As #f

    Print #f, "Section,QuoteName,Material,Qty,BomL,BomW,BomT,CadL,CadW,CadT,CadComponent,Status"

    Dim i As Long

    For i = 1 To ExportCount

        Dim cadName As String
        Dim cadL As String
        Dim cadW As String
        Dim cadT As String

        If ExportRows(i).HasCad Then
            cadName = parts(ExportRows(i).CadPartIndex).componentName
            cadL = FormatNumberForCsv(ExportRows(i).Length)
            cadW = FormatNumberForCsv(ExportRows(i).Width)
            cadT = FormatNumberForCsv(ExportRows(i).Thickness)
        Else
            cadName = ""
            cadL = ""
            cadW = ""
            cadT = ""
        End If

        Print #f, "EXPORT," & _
                  CsvText(ExportRows(i).quoteName) & "," & _
                  CsvText(ExportRows(i).material) & "," & _
                  CStr(ExportRows(i).Quantity) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomLength) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomWidth) & "," & _
                  FormatNumberForCsv(ExportRows(i).BomThickness) & "," & _
                  cadL & "," & cadW & "," & cadT & "," & _
                  CsvText(cadName) & "," & _
                  CsvText(ExportRows(i).Status)
    Next i

    For i = 1 To PullcoreMatchCount

        Dim pcCadName As String
        Dim pcL As String
        Dim pcW As String
        Dim pcT As String

        If PullcoreMatches(i).CadPartIndex > 0 Then
            pcCadName = parts(PullcoreMatches(i).CadPartIndex).componentName
            pcL = FormatNumberForCsv(PullcoreMatches(i).FittedLength)
            pcW = FormatNumberForCsv(PullcoreMatches(i).FittedWidth)
            pcT = FormatNumberForCsv(PullcoreMatches(i).FittedThickness)
        Else
            pcCadName = ""
            pcL = ""
            pcW = ""
            pcT = ""
        End If

        Print #f, "PULLCORE," & _
                  CsvText(PullcoreMatches(i).quoteName) & "," & _
                  CsvText(PullcoreMatches(i).material) & "," & _
                  CStr(PullcoreMatches(i).Quantity) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomLength) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomWidth) & "," & _
                  FormatNumberForCsv(PullcoreMatches(i).BomThickness) & "," & _
                  pcL & "," & pcW & "," & pcT & "," & _
                  CsvText(pcCadName) & "," & _
                  CsvText(PullcoreMatches(i).Status)
    Next i

    WritePackageReportRow f, "J BLOCK", FindBestJBlockIndex()
    WritePackageReportRow f, "EJECTOR CAM", FindBestEjectorCamIndex()
    WritePackageReportRow f, "PULLCORE STOP", FindBestPartByNameAndDims(PULLCORE_STOP_NAME_KEYS, 0, 0, 0, 0)
    WritePackageReportRow f, "FLIPPER CAM COVER PLATE", _
        FindBestPartByNameAndDims(FLIPPER_CAM_COVER_NAME_KEYS, _
                                  FLIPPER_CAM_COVER_TARGET_THICKNESS, _
                                  FLIPPER_CAM_COVER_TARGET_WIDTH, _
                                  FLIPPER_CAM_COVER_TARGET_LENGTH, _
                                  FLIPPER_CAM_COVER_DIM_TOL)

    Print #f, "RAWBOM,---,---,---,---,---,---,---,---,---,---,parsed " & CStr(BomCount) & " rows"

    For i = 1 To BomCount
        Print #f, "RAWBOM," & _
                  CsvText(BomRows(i).Description) & "," & _
                  CsvText(BomRows(i).material) & "," & _
                  CStr(BomRows(i).Quantity) & "," & _
                  FormatNumberForCsv(BomRows(i).BomLength) & "," & _
                  FormatNumberForCsv(BomRows(i).BomWidth) & "," & _
                  FormatNumberForCsv(BomRows(i).BomThickness) & ",,,," & _
                  CsvText(BomRows(i).quoteName) & "," & _
                  CsvText("isPullcore=" & CStr(IsPullcoreBomRow(BomRows(i))) & " hasDims=" & CStr(BomRows(i).hasDims))
    Next i

    Close #f

    LogLine "Wrote BOM match report CSV: " & csvPath
    Exit Sub

ErrHandler:
    LogLine "WriteExportCheckCsv error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Private Sub WritePackageReportRow(ByVal f As Integer, ByVal label As String, ByVal cadIdx As Long)
On Error Resume Next

    Dim nm As String
    Dim L As String
    Dim W As String
    Dim T As String
    Dim st As String

    If cadIdx > 0 And cadIdx <= PartCount Then
        nm = parts(cadIdx).componentName
        L = FormatNumberForCsv(parts(cadIdx).Length)
        W = FormatNumberForCsv(parts(cadIdx).Width)
        T = FormatNumberForCsv(parts(cadIdx).Thickness)
        st = "CAD MATCH FOUND"
    Else
        nm = ""
        L = ""
        W = ""
        T = ""
        st = "NO CAD MATCH"
    End If

    Print #f, "PACKAGE," & CsvText(label) & ",,,,,," & _
              L & "," & W & "," & T & "," & CsvText(nm) & "," & CsvText(st)
End Sub

' ============================================================
' PULL CORE DIMENSIONS EXCEL REPORT
' ============================================================

Private Sub WritePullCoreDimensionsExcel(ByVal reportPath As String)
On Error GoTo ErrHandler

    If CREATE_PULLCORE_DIMENSIONS_EXCEL = False Then Exit Sub

    If reportPath = "" Then
        reportPath = CurrentJobFolder & "\" & PULLCORE_DIMENSIONS_REPORT_FILE
    End If

    LogLine "Creating Pull Core Dimensions Excel report: " & reportPath

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.ScreenUpdating = False
    xlApp.EnableEvents = False

    Set wb = xlApp.Workbooks.Add
    Set ws = wb.Worksheets(1)

    On Error Resume Next
    ws.Name = "Pull Core Dimensions"
    On Error GoTo ErrHandler

    WritePullCoreDimensionsExcelHeader ws

    Dim r As Long
    r = 2

    Dim i As Long

    For i = 1 To PullcoreMatchCount
        WritePullCoreDimensionsExcelRow ws, r, i
        r = r + 1
    Next i

    With ws
        .Columns("A:AH").AutoFit
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(220, 230, 241)
        .Rows(1).AutoFilter
    End With

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    EnsureFolderDeep fso.GetParentFolderName(reportPath)

    If fso.FileExists(reportPath) Then
        On Error Resume Next
        fso.DeleteFile reportPath, True
        On Error GoTo ErrHandler
    End If

    wb.SaveAs reportPath, xlOpenXMLWorkbook
    wb.Close False

    xlApp.EnableEvents = True
    xlApp.ScreenUpdating = True
    xlApp.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    LogLine "Pull Core Dimensions Excel report saved: " & reportPath

    Exit Sub

ErrHandler:
    LogLine "WritePullCoreDimensionsExcel error: " & Err.Description

    On Error Resume Next

    If Not wb Is Nothing Then wb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Sub

Private Sub WritePullCoreDimensionsExcelHeader(ByVal ws As Object)
On Error Resume Next

    If ws Is Nothing Then Exit Sub

    ws.Cells(1, 1).Value = "Job"
    ws.Cells(1, 2).Value = "Quote Name"
    ws.Cells(1, 3).Value = "Description"
    ws.Cells(1, 4).Value = "Material"
    ws.Cells(1, 5).Value = "Qty"
    ws.Cells(1, 6).Value = "CAD Component"
    ws.Cells(1, 7).Value = "Clean Name"

    ws.Cells(1, 8).Value = "BOM Length"
    ws.Cells(1, 9).Value = "BOM Width"
    ws.Cells(1, 10).Value = "BOM Thickness"

    ws.Cells(1, 11).Value = "Original Bounding Box Dimensions"
    ws.Cells(1, 12).Value = "Original Length"
    ws.Cells(1, 13).Value = "Original Width"
    ws.Cells(1, 14).Value = "Original Thickness"

    ws.Cells(1, 15).Value = "Fitted Bounding Box Dimensions"
    ws.Cells(1, 16).Value = "Fitted Length"
    ws.Cells(1, 17).Value = "Fitted Width"
    ws.Cells(1, 18).Value = "Fitted Thickness"

    ws.Cells(1, 19).Value = "Delta Length Original-Fitted"
    ws.Cells(1, 20).Value = "Delta Width Original-Fitted"
    ws.Cells(1, 21).Value = "Delta Thickness Original-Fitted"

    ws.Cells(1, 22).Value = "Detected Pullcore Angle Deg"
    ws.Cells(1, 23).Value = "DXF Rotation Deg"

    ws.Cells(1, 24).Value = "Status"
    ws.Cells(1, 25).Value = "Is Cam"
    ws.Cells(1, 26).Value = "CAD Index"
    ws.Cells(1, 27).Value = "Mass"
    ws.Cells(1, 28).Value = "BBox Volume"
    ws.Cells(1, 29).Value = "Notes"

    ws.Cells(1, 30).Value = "Assembly Center X"
    ws.Cells(1, 31).Value = "Assembly Center Y"
    ws.Cells(1, 32).Value = "Assembly Center Z"
    ws.Cells(1, 33).Value = "ID/OD Height Axis Used"
    ws.Cells(1, 34).Value = "ID/OD Height Value"
End Sub

Private Sub WritePullCoreDimensionsExcelRow(ByVal ws As Object, ByVal r As Long, ByVal matchIdx As Long)
On Error Resume Next

    If ws Is Nothing Then Exit Sub
    If matchIdx <= 0 Or matchIdx > PullcoreMatchCount Then Exit Sub

    Dim cadIdx As Long
    cadIdx = PullcoreMatches(matchIdx).CadPartIndex

    Dim compName As String
    Dim cleanName As String
    Dim massVal As Double
    Dim volVal As Double

    compName = ""
    cleanName = ""
    massVal = 0#
    volVal = 0#

    If cadIdx > 0 And cadIdx <= PartCount Then
        compName = parts(cadIdx).componentName
        cleanName = parts(cadIdx).cleanName
        massVal = parts(cadIdx).massValue
        volVal = parts(cadIdx).BBoxVolume
    End If

    ws.Cells(r, 1).Value = CurrentJobNumber
    ws.Cells(r, 2).Value = PullcoreMatches(matchIdx).quoteName
    ws.Cells(r, 3).Value = PullcoreMatches(matchIdx).Description
    ws.Cells(r, 4).Value = PullcoreMatches(matchIdx).material
    ws.Cells(r, 5).Value = PullcoreMatches(matchIdx).Quantity

    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).Value = compName

    ws.Cells(r, 7).NumberFormat = "@"
    ws.Cells(r, 7).Value = cleanName

    ws.Cells(r, 8).Value = PullcoreMatches(matchIdx).BomLength
    ws.Cells(r, 9).Value = PullcoreMatches(matchIdx).BomWidth
    ws.Cells(r, 10).Value = PullcoreMatches(matchIdx).BomThickness

    ws.Cells(r, 11).Value = "Original Bounding Box Dimensions"
    ws.Cells(r, 12).Value = PullcoreMatches(matchIdx).OriginalLength
    ws.Cells(r, 13).Value = PullcoreMatches(matchIdx).OriginalWidth
    ws.Cells(r, 14).Value = PullcoreMatches(matchIdx).OriginalThickness

    ws.Cells(r, 15).Value = "Fitted Bounding Box Dimensions"
    ws.Cells(r, 16).Value = PullcoreMatches(matchIdx).FittedLength
    ws.Cells(r, 17).Value = PullcoreMatches(matchIdx).FittedWidth
    ws.Cells(r, 18).Value = PullcoreMatches(matchIdx).FittedThickness

    ws.Cells(r, 19).Value = PullcoreMatches(matchIdx).OriginalLength - PullcoreMatches(matchIdx).FittedLength
    ws.Cells(r, 20).Value = PullcoreMatches(matchIdx).OriginalWidth - PullcoreMatches(matchIdx).FittedWidth
    ws.Cells(r, 21).Value = PullcoreMatches(matchIdx).OriginalThickness - PullcoreMatches(matchIdx).FittedThickness

    ws.Cells(r, 22).Value = PullcoreMatches(matchIdx).DetectedAngleDeg
    ws.Cells(r, 23).Value = PullcoreMatches(matchIdx).DxfRotationDeg

    ws.Cells(r, 24).Value = PullcoreMatches(matchIdx).Status
    ws.Cells(r, 25).Value = CStr(PullcoreMatches(matchIdx).isCam)
    ws.Cells(r, 26).Value = cadIdx
    ws.Cells(r, 27).Value = massVal
    ws.Cells(r, 28).Value = volVal

    If cadIdx <= 0 Then
        ws.Cells(r, 29).Value = "No CAD match"
    ElseIf Abs(PullcoreMatches(matchIdx).DetectedAngleDeg) < PULLCORE_STRAIGHTEN_MIN_DEG Then
        ws.Cells(r, 29).Value = "Angle below threshold or original/fitted boxes nearly identical"
    Else
        ws.Cells(r, 29).Value = "Angle calculated from original assembly/world bbox vs fitted bbox"
    End If

    If cadIdx > 0 And cadIdx <= PartCount Then
        If parts(cadIdx).hasAsmCenter Then
            ws.Cells(r, 30).Value = parts(cadIdx).AsmCenterX
            ws.Cells(r, 31).Value = parts(cadIdx).AsmCenterY
            ws.Cells(r, 32).Value = parts(cadIdx).AsmCenterZ
            ws.Cells(r, 33).Value = PullcoreIdOdHeightAxisUsed

            If PullcoreIdOdHeightAxisUsed <> "" Then
                ws.Cells(r, 34).Value = GetPullcoreMatchHeightValue(matchIdx, PullcoreIdOdHeightAxisUsed)
            End If
        End If
    End If
End Sub

' ============================================================
' END OF PART 7 OF 8
' Paste Part 8 next.
' ============================================================
' ============================================================
' FINDERS: JOB FOLDER + CAD
' ============================================================

Private Function FindJobFolderByText(ByVal rootPath As String, ByVal searchText As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootPath) Then
        LogLine "Root path does not exist: " & rootPath
        Exit Function
    End If

    Dim root As Object
    Set root = fso.GetFolder(rootPath)

    Dim wantUpper As String
    wantUpper = UCase(Trim(searchText))

    Dim bestPath As String
    Dim bestScore As Long

    bestPath = ""
    bestScore = -1

    Dim subFolder As Object
    Dim nameUpper As String
    Dim score As Long

    For Each subFolder In root.SubFolders

        nameUpper = UCase(subFolder.Name)

        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextTop
        If InStr(nameUpper, " PRINTS") > 0 Then GoTo NextTop
        If InStr(nameUpper, "PULLCORE") > 0 Then GoTo NextTop
        If InStr(nameUpper, "PYROPEL") > 0 Then GoTo NextTop
        If InStr(nameUpper, "J BLOCK") > 0 Then GoTo NextTop

        score = -1

        If nameUpper = wantUpper Then
            score = 1000
        ElseIf InStr(nameUpper, wantUpper) > 0 Then
            score = 500 - Abs(Len(nameUpper) - Len(wantUpper))
        End If

        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If

NextTop:
    Next subFolder

    If bestPath <> "" Then
        FindJobFolderByText = bestPath
        Exit Function
    End If

    For Each subFolder In root.SubFolders
        nameUpper = UCase(subFolder.Name)
        If nameUpper <> UCase(EXTRACT_FOLDER_NAME) Then
            SearchJobFolderRecursive subFolder, wantUpper, 1, bestPath, bestScore
        End If
    Next subFolder

    If bestPath = "" Then
        If InStr(UCase(root.Name), wantUpper) > 0 Then bestPath = rootPath
    End If

    FindJobFolderByText = bestPath
    Exit Function

ErrHandler:
    LogLine "FindJobFolderByText error: " & Err.Description
    FindJobFolderByText = ""
End Function

Private Sub SearchJobFolderRecursive(ByVal folder As Object, _
                                     ByVal wantUpper As String, _
                                     ByVal depth As Long, _
                                     ByRef bestPath As String, _
                                     ByRef bestScore As Long)
On Error Resume Next

    If folder Is Nothing Then Exit Sub
    If depth > 3 Then Exit Sub

    Dim subFolder As Object
    Dim nameUpper As String
    Dim score As Long

    For Each subFolder In folder.SubFolders

        nameUpper = UCase(subFolder.Name)

        If nameUpper = UCase(EXTRACT_FOLDER_NAME) Then GoTo NextSub
        If InStr(nameUpper, " PRINTS") > 0 Then GoTo NextSub
        If InStr(nameUpper, "PULLCORE") > 0 Then GoTo NextSub
        If InStr(nameUpper, "PYROPEL") > 0 Then GoTo NextSub
        If InStr(nameUpper, "J BLOCK") > 0 Then GoTo NextSub

        score = -1

        If nameUpper = wantUpper Then
            score = 1000 - depth
        ElseIf InStr(nameUpper, wantUpper) > 0 Then
            score = 500 - (depth * 20) - Abs(Len(nameUpper) - Len(wantUpper))
        End If

        If score > bestScore Then
            bestScore = score
            bestPath = subFolder.path
        End If

        SearchJobFolderRecursive subFolder, wantUpper, depth + 1, bestPath, bestScore

NextSub:
    Next subFolder
End Sub

Private Function FindAllCadModelsRanked(ByVal folderPath As String) As Collection
On Error GoTo ErrHandler

    Dim result As New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Set FindAllCadModelsRanked = result
        Exit Function
    End If

    Dim paths As Collection
    Dim scores As Collection

    Set paths = New Collection
    Set scores = New Collection

    CollectCadModelsInFolder fso.GetFolder(folderPath), paths, scores

    Dim used() As Boolean

    If paths.count > 0 Then

        ReDim used(1 To paths.count)

        Dim k As Long
        Dim i As Long
        Dim bestI As Long
        Dim bestS As Long

        For k = 1 To paths.count

            bestI = 0
            bestS = -2147483647

            For i = 1 To paths.count
                If used(i) = False Then
                    If CLng(scores(i)) > bestS Then
                        bestS = CLng(scores(i))
                        bestI = i
                    End If
                End If
            Next i

            If bestI > 0 Then
                used(bestI) = True
                result.Add CStr(paths(bestI))
            End If

        Next k

    End If

    Set FindAllCadModelsRanked = result
    Exit Function

ErrHandler:
    LogLine "FindAllCadModelsRanked error: " & Err.Description
    Set FindAllCadModelsRanked = New Collection
End Function

Private Sub CollectCadModelsInFolder(ByVal folder As Object, ByVal paths As Collection, ByVal scores As Collection)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderName As String
    folderName = UCase(folder.Name)

    If InStr(folderName, " PRINTS") > 0 Then Exit Sub
    If folderName = "BASE" Or InStr(folderName, " BASE") > 0 Then Exit Sub

    Dim file As Object
    Dim ext As String
    Dim score As Long

    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.path))
        score = CadFilePriority(ext, file.Name)

        If score > 0 Then
            paths.Add file.path
            scores.Add score
        End If

    Next file

    Dim subFolder As Object

    For Each subFolder In folder.SubFolders
        CollectCadModelsInFolder subFolder, paths, scores
    Next subFolder
End Sub

Private Sub AppendCadCandidates(ByVal target As Collection, ByVal extra As Collection)
On Error Resume Next

    If target Is Nothing Or extra Is Nothing Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim p As String
    Dim dup As Boolean

    For i = 1 To extra.count

        p = CStr(extra(i))
        dup = False

        For j = 1 To target.count
            If LCase(CStr(target(j))) = LCase(p) Then
                dup = True
                Exit For
            End If
        Next j

        If dup = False Then target.Add p

    Next i
End Sub

Private Function CadFilePriority(ByVal ext As String, ByVal fileName As String) As Long

    Dim nameUpper As String
    nameUpper = UCase(fileName)

    If Left(fileName, 2) = "~$" Then CadFilePriority = 0: Exit Function
    If Left(nameUpper, 2) = "~$" Then CadFilePriority = 0: Exit Function

    Dim bonus As Long
    bonus = 0

    If InStr(nameUpper, CurrentJobNumber) > 0 Then bonus = bonus + 5
    If InStr(nameUpper, "ASSEM") > 0 Or InStr(nameUpper, "ASSY") > 0 Or InStr(nameUpper, "BASE") > 0 Then bonus = bonus + 3

    Select Case ext
        Case "sldasm"
            CadFilePriority = 100 + bonus
        Case "easm"
            CadFilePriority = 90 + bonus
        Case "asm"
            CadFilePriority = 85 + bonus
        Case "step", "stp"
            CadFilePriority = 80 + bonus
        Case "x_t", "x_b"
            CadFilePriority = 70 + bonus
        Case "igs", "iges"
            CadFilePriority = 60 + bonus
        Case "sldprt"
            CadFilePriority = 50 + bonus
        Case "prt"
            CadFilePriority = 45 + bonus
        Case Else
            CadFilePriority = 0
    End Select
End Function

' ============================================================
' OUTPUT NAMING
' ============================================================

Private Sub DetermineOutputNamingInfo(ByVal jobFolder As String)
On Error GoTo ErrHandler

    CustomerNumber = ""
    DateCode = ""
    NamingSourceText = ""

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(jobFolder) Then Exit Sub

    SearchNamingInfoRecursive fso.GetFolder(jobFolder)

    If DateCode = "" Then DateCode = Format(Date, "mm-dd-yyyy")
    If CustomerNumber = "" Then CustomerNumber = CurrentJobNumber

    Exit Sub

ErrHandler:
    LogLine "DetermineOutputNamingInfo error: " & Err.Description
End Sub

Private Sub SearchNamingInfoRecursive(ByVal folder As Object)
On Error Resume Next

    Dim file As Object

    For Each file In folder.Files
        TryExtractNamingFromText file.Name
        If CustomerNumber <> "" And DateCode <> "" Then Exit Sub
    Next file

    TryExtractNamingFromText folder.Name

    Dim subFolder As Object

    For Each subFolder In folder.SubFolders
        SearchNamingInfoRecursive subFolder
        If CustomerNumber <> "" And DateCode <> "" Then Exit Sub
    Next subFolder
End Sub

Private Sub TryExtractNamingFromText(ByVal text As String)
On Error Resume Next

    If text = "" Then Exit Sub

    If CustomerNumber = "" Then

        Dim c As String
        c = ExtractCustomerNumberFromText(text)

        If c <> "" Then
            CustomerNumber = c
            If NamingSourceText = "" Then NamingSourceText = text
        End If

    End If

    If DateCode = "" Then

        Dim d As String
        d = ExtractDateCodeFromText(text)

        If d <> "" Then
            DateCode = d
            If NamingSourceText = "" Then NamingSourceText = text
        End If

    End If
End Sub

Private Function ExtractCustomerNumberFromText(ByVal text As String) As String
On Error GoTo ErrHandler

    Dim upper As String
    upper = UCase(text)

    Dim i As Long
    Dim runStart As Long
    Dim runLen As Long
    Dim ch As String
    Dim candidate As String

    runStart = 0
    runLen = 0

    For i = 1 To Len(upper) + 1

        If i <= Len(upper) Then
            ch = Mid(upper, i, 1)
        Else
            ch = "X"
        End If

        If ch >= "0" And ch <= "9" Then

            If runLen = 0 Then runStart = i
            runLen = runLen + 1

        Else

            If runLen >= 8 Then
                candidate = Mid(upper, runStart, runLen)

                If IsLikelyEightDigitDate(candidate) = False And IsMostlyZeros(candidate) = False Then
                    ExtractCustomerNumberFromText = candidate
                    Exit Function
                End If
            End If

            runLen = 0

        End If

    Next i

    Exit Function

ErrHandler:
    ExtractCustomerNumberFromText = ""
End Function

Private Function ExtractDateCodeFromText(ByVal text As String) As String
On Error GoTo ErrHandler

    Dim sep As Variant
    Dim guess As String

    For Each sep In Array("-", "_", ".", "/")

        guess = ExtractDateWithSeparator(text, CStr(sep))

        If guess <> "" Then
            ExtractDateCodeFromText = guess
            Exit Function
        End If

    Next sep

    Exit Function

ErrHandler:
    ExtractDateCodeFromText = ""
End Function

Private Function ExtractDateWithSeparator(ByVal text As String, ByVal sep As String) As String
On Error GoTo ErrHandler

    Dim cleaned As String
    cleaned = Replace(text, " ", "")

    Dim i As Long
    Dim n As Long
    Dim chunk As String
    Dim mm As String
    Dim dd As String
    Dim yyyy As String

    n = Len(cleaned)

    For i = 1 To n - 7

        chunk = Mid(cleaned, i, 10)

        If Len(chunk) = 10 Then
            If Mid(chunk, 3, 1) = sep And Mid(chunk, 6, 1) = sep Then

                mm = Left(chunk, 2)
                dd = Mid(chunk, 4, 2)
                yyyy = Right(chunk, 4)

                If IsNumeric(mm) And IsNumeric(dd) And IsNumeric(yyyy) Then
                    If IsValidMonthDayYear(CLng(mm), CLng(dd), CLng(yyyy)) Then
                        ExtractDateWithSeparator = mm & "-" & dd & "-" & yyyy
                        Exit Function
                    End If
                End If

            End If
        End If

    Next i

    Exit Function

ErrHandler:
    ExtractDateWithSeparator = ""
End Function

Private Function IsLikelyEightDigitDate(ByVal s As String) As Boolean
    If Len(s) <> 8 Then Exit Function
    If IsNumeric(s) = False Then Exit Function

    Dim mm As Long
    Dim dd As Long
    Dim yyyy As Long

    mm = CLng(Left(s, 2))
    dd = CLng(Mid(s, 3, 2))
    yyyy = CLng(Right(s, 4))

    If IsValidMonthDayYear(mm, dd, yyyy) Then
        IsLikelyEightDigitDate = True
        Exit Function
    End If

    yyyy = CLng(Left(s, 4))
    mm = CLng(Mid(s, 5, 2))
    dd = CLng(Right(s, 2))

    If IsValidMonthDayYear(mm, dd, yyyy) Then IsLikelyEightDigitDate = True
End Function

Private Function IsValidMonthDayYear(ByVal mm As Long, ByVal dd As Long, ByVal yyyy As Long) As Boolean
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function
    If yyyy < 2000 Or yyyy > 2100 Then Exit Function

    IsValidMonthDayYear = True
End Function

Private Function IsMostlyZeros(ByVal s As String) As Boolean
    Dim i As Long
    Dim zeros As Long

    For i = 1 To Len(s)
        If Mid(s, i, 1) = "0" Then zeros = zeros + 1
    Next i

    If Len(s) > 0 Then
        If zeros / Len(s) >= 0.7 Then IsMostlyZeros = True
    End If
End Function

' ============================================================
' ARRAY / HEADER HELPERS
' ============================================================

Private Function GetArrayValue(ByVal data As Variant, ByVal r As Long, ByVal c As Long) As String
On Error Resume Next

    If c <= 0 Then Exit Function

    Dim v As Variant
    v = data(r, c)

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function

    GetArrayValue = Trim(CStr(v))
End Function

Private Function FindBomHeaderLikeInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long, ByVal options As Variant) As Long
    Dim opt As Variant
    Dim c As Long
    Dim cellUpper As String

    For c = 1 To colCount

        cellUpper = UCase(GetArrayValue(data, r, c))

        If cellUpper <> "" Then
            For Each opt In options
                If cellUpper = UCase(CStr(opt)) Or InStr(cellUpper, UCase(CStr(opt))) > 0 Then
                    FindBomHeaderLikeInArrayRow = c
                    Exit Function
                End If
            Next opt
        End If

    Next c
End Function

Private Function FindBomQtyColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long) As Long
    FindBomQtyColumnInArrayRow = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("QTY", "QUANTITY", "QNTY", "QUAN"))
End Function

Private Function FindBomMaterialColumnInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long) As Long
    FindBomMaterialColumnInArrayRow = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("MATERIAL", "MAT", "MTL", "STEEL"))
End Function

Private Sub FindBomDimensionColumnsInArrayRow(ByVal data As Variant, ByVal r As Long, ByVal colCount As Long, _
                                              ByRef lenCol As Long, ByRef widCol As Long, ByRef thkCol As Long)

    lenCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("LENGTH", "LTH", "LEN"))
    widCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("WIDTH", "WTH", "WID"))
    thkCol = FindBomHeaderLikeInArrayRow(data, r, colCount, Array("THICKNESS", "THICK", "THK", "HGT", "HEIGHT"))
End Sub

Private Function IsLikelyBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next

    Dim n As String
    n = UCase(ws.Name)

    If InStr(n, "BOM") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "BILL") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "MATERIAL") > 0 Then IsLikelyBomWorksheet = True: Exit Function
    If InStr(n, "QUOTE") > 0 Then IsLikelyBomWorksheet = True: Exit Function

    IsLikelyBomWorksheet = True
End Function

Private Function ShouldSkipBomWorksheet(ByVal ws As Object) As Boolean
On Error Resume Next

    Dim n As String
    n = UCase(ws.Name)

    If InStr(n, "INSTRUCTION") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
    If InStr(n, "NOTES") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
    If InStr(n, "COVER") > 0 Then ShouldSkipBomWorksheet = True: Exit Function
End Function

' ============================================================
' NAME STANDARDIZATION
' ============================================================

Private Function StandardPlateName(ByVal desc As String) As String
On Error GoTo ErrHandler

    Dim s As String
    s = NormalizeText(desc)

    If InStr(s, "EJECTOR J-BLOCK") > 0 Or InStr(s, "EJ J-BLOCK") > 0 Or _
       InStr(s, "EJ J BLOCK") > 0 Or InStr(s, "J-BLOCK") > 0 Or InStr(s, "J BLOCK") > 0 Then
        StandardPlateName = "EJECTOR J-BLOCK"
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then

        Dim loc As String
        Dim kind As String

        loc = GetPullcoreLocationCode(s)

        If InStr(s, "CAM") > 0 Then
            kind = "PULLCORE CAM"
        Else
            kind = "PULLCORE KEY"
        End If

        If loc <> "" Then
            StandardPlateName = loc & " " & kind
        Else
            StandardPlateName = kind
        End If

        Exit Function
    End If

    If InStr(s, "ID HOLDER") > 0 Or InStr(s, "IDTE HOLDER") > 0 Or InStr(s, "TOP HOLDER") > 0 Or InStr(s, "TOP HOLDER BLOCK") > 0 Then
        StandardPlateName = "ID HOLDER"
        Exit Function
    End If

    If InStr(s, "OD HOLDER") > 0 Or InStr(s, "ODTE HOLDER") > 0 Or InStr(s, "BOTTOM HOLDER") > 0 Or _
       InStr(s, "BOT HOLDER") > 0 Or InStr(s, "BOTTOM HOLDER BLOCK") > 0 Then
        StandardPlateName = "OD HOLDER"
        Exit Function
    End If

    If InStr(s, "POT BLOCK") > 0 Or InStr(s, "POT BLK") > 0 Or InStr(s, " POT ") > 0 Or Right(s, 4) = " POT" Then

        If IsLikelyOdSideName(s) Then
            StandardPlateName = "OD POT BLOCK"
            Exit Function
        End If

        If IsLikelyIdSideName(s) Then
            StandardPlateName = "ID POT BLOCK"
            Exit Function
        End If

        StandardPlateName = "POT BLOCK"
        Exit Function
    End If

    If InStr(s, "SMED") > 0 Then

        If IsLikelyIdSideName(s) Then
            StandardPlateName = "TCP"
            Exit Function
        End If

        If IsLikelyOdSideName(s) Then
            StandardPlateName = "BCP"
            Exit Function
        End If

        StandardPlateName = "SMED PLATE"
        Exit Function
    End If

    If InStr(s, "TOP CLAMP") > 0 Or InStr(s, "TOP CLAMPING") > 0 Then
        StandardPlateName = "TCP"
        Exit Function
    End If

    If InStr(s, "BOTTOM CLAMP") > 0 Or InStr(s, "BOT CLAMP") > 0 Or _
       InStr(s, "BOTTOM CLAMPING") > 0 Or InStr(s, "BOT CLAMPING") > 0 Then
        StandardPlateName = "BCP"
        Exit Function
    End If

    If InStr(s, "TCP") > 0 Then
        StandardPlateName = "TCP"
        Exit Function
    End If

    If InStr(s, "BCP") > 0 Then
        StandardPlateName = "BCP"
        Exit Function
    End If

    If InStr(s, "TOP INS") > 0 Or InStr(s, "TOP INSULATION") > 0 Then
        StandardPlateName = "TOP INS"
        Exit Function
    End If

    If InStr(s, "BOT INS") > 0 Or InStr(s, "BOTTOM INS") > 0 Or InStr(s, "BOTTOM INSULATION") > 0 Then
        StandardPlateName = "BOT INS"
        Exit Function
    End If

    If InStr(s, "PULLCORE STOP") > 0 Or InStr(s, "PULL CORE STOP") > 0 Then
        StandardPlateName = "PULLCORE STOP"
        Exit Function
    End If

    If InStr(s, "FLIPPER CAM COVER") > 0 Or InStr(s, "CAM COVER PLATE") > 0 Then
        StandardPlateName = "FLIPPER CAM COVER PLATE"
        Exit Function
    End If

    StandardPlateName = ProperCaseText(Trim(desc))
    Exit Function

ErrHandler:
    StandardPlateName = Trim(desc)
End Function

Private Function IsLikelyIdSideName(ByVal s As String) As Boolean
    If InStr(s, "ID ") > 0 Or Left(s, 2) = "ID" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, " TOP") > 0 Or Left(s, 3) = "TOP" Then IsLikelyIdSideName = True: Exit Function
    If InStr(s, "TCP") > 0 Then IsLikelyIdSideName = True
End Function

Private Function IsLikelyOdSideName(ByVal s As String) As Boolean
    If InStr(s, "OD ") > 0 Or Left(s, 2) = "OD" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BOTTOM") > 0 Or InStr(s, "BOT ") > 0 Or Left(s, 3) = "BOT" Then IsLikelyOdSideName = True: Exit Function
    If InStr(s, "BCP") > 0 Then IsLikelyOdSideName = True
End Function

Private Function IsMainPrintsQuote(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "IDHOLDER", "ODHOLDER", "TCP", "BCP", "TOPINS", "BOTINS", _
             "SMEDPLATE", "IDPOTBLOCK", "ODPOTBLOCK", "POTBLOCK"
            IsMainPrintsQuote = True
    End Select
End Function

Private Function GetOutputFolderForRegularExport(ByVal quoteName As String, ByVal outputFolder As String) As String
    If IsMainPrintsQuote(quoteName) Then
        GetOutputFolderForRegularExport = outputFolder
    Else
        GetOutputFolderForRegularExport = CurrentJobFolder & "\" & CurrentJobNumber & " MISC DETAILS"
    End If
End Function

Private Function IsInsertQuoteName(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "TOPINS" Or k = "BOTINS" Then IsInsertQuoteName = True: Exit Function
    If InStr(k, "INSERT") > 0 Then IsInsertQuoteName = True: Exit Function
    If InStr(k, "INS") > 0 And Len(k) <= 16 Then IsInsertQuoteName = True
End Function

Private Function ShouldSkipExportQuoteName(ByVal quoteName As String) As Boolean
    Dim s As String
    s = NormalizeText(quoteName)

    If InStr(s, "J-BLOCK") > 0 Or InStr(s, "J BLOCK") > 0 Or InStr(s, "JBLOCK") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If

    If InStr(s, "PULLCORE") > 0 Or InStr(s, "PULL CORE") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If

    If InStr(s, "FLIPPER CAM COVER") > 0 Then
        ShouldSkipExportQuoteName = True
        Exit Function
    End If
End Function

Private Function IsNameMatch(ByVal cadName As String, ByVal quoteName As String) As Boolean
    Dim a As String
    Dim b As String

    a = NormalizeKey(cadName)
    b = NormalizeKey(quoteName)

    If a = "" Or b = "" Then Exit Function

    If InStr(a, b) > 0 Or InStr(b, a) > 0 Then IsNameMatch = True
End Function

' ============================================================
' TCP / BCP MASS HELPERS
' ============================================================

Private Function IsTcpBcpQuoteName(ByVal quoteName As String) As Boolean
    Dim k As String
    k = NormalizeKey(quoteName)

    Select Case k
        Case "TCP", "BCP", _
             "TOPCLAMPINGPLATE", "BOTTOMCLAMPINGPLATE", _
             "TOPSMEDPLATE", "BOTTOMSMEDPLATE", _
             "TOPSMED", "BOTTOMSMED"
            IsTcpBcpQuoteName = True
    End Select
End Function

Private Function ShouldUseTcpBcpMassPairRule(ByRef b As BomInfo) As Boolean
On Error GoTo ErrHandler

    ShouldUseTcpBcpMassPairRule = False

    If b.hasDims = False Then Exit Function
    If IsTcpBcpQuoteName(b.quoteName) = False Then Exit Function

    Dim counterpart As String
    counterpart = GetCounterpartQuoteName(b.quoteName)

    If counterpart = "" Then Exit Function

    Dim i As Long

    For i = 1 To BomCount

        If NormalizeKey(BomRows(i).quoteName) = NormalizeKey(counterpart) Then
            If BomRows(i).hasDims Then
                If Abs(BomRows(i).BomLength - b.BomLength) <= SAME_SIZE_PAIR_TOL Then
                    If Abs(BomRows(i).BomWidth - b.BomWidth) <= SAME_SIZE_PAIR_TOL Then
                        If Abs(BomRows(i).BomThickness - b.BomThickness) <= SAME_SIZE_PAIR_TOL Then
                            ShouldUseTcpBcpMassPairRule = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

    Next i

    Exit Function

ErrHandler:
    ShouldUseTcpBcpMassPairRule = False
End Function

Private Function GetMassPreferenceForQuoteName(ByVal quoteName As String) As String
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "TOPINS" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BOTINS" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function

    If k = "TCP" Then GetMassPreferenceForQuoteName = "LIGHT": Exit Function
    If k = "BCP" Then GetMassPreferenceForQuoteName = "HEAVY": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetMassPreferenceForQuoteName = "LIGHT"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTTOMSMEDPLATE" Or k = "BOTTOMSMED" Then
        GetMassPreferenceForQuoteName = "HEAVY"
        Exit Function
    End If

    GetMassPreferenceForQuoteName = ""
End Function

Private Function GetCounterpartQuoteName(ByVal quoteName As String) As String
    Dim k As String
    k = NormalizeKey(quoteName)

    If k = "IDHOLDER" Then GetCounterpartQuoteName = "OD HOLDER": Exit Function
    If k = "ODHOLDER" Then GetCounterpartQuoteName = "ID HOLDER": Exit Function
    If k = "TOPINS" Then GetCounterpartQuoteName = "BOT INS": Exit Function
    If k = "BOTINS" Then GetCounterpartQuoteName = "TOP INS": Exit Function
    If k = "TCP" Then GetCounterpartQuoteName = "BCP": Exit Function
    If k = "BCP" Then GetCounterpartQuoteName = "TCP": Exit Function

    If k = "TOPCLAMPINGPLATE" Or k = "TOPSMEDPLATE" Or k = "TOPSMED" Then
        GetCounterpartQuoteName = "BCP"
        Exit Function
    End If

    If k = "BOTTOMCLAMPINGPLATE" Or k = "BOTTOMSMEDPLATE" Or k = "BOTTOMSMED" Then
        GetCounterpartQuoteName = "TCP"
        Exit Function
    End If
End Function

Private Function NormalizeSteelType(ByVal matText As String) As String
    Dim s As String
    s = UCase(Trim(matText))

    If s = "" Then NormalizeSteelType = DEFAULT_STEEL_TYPE: Exit Function

    If InStr(s, "PYROPEL") > 0 Then NormalizeSteelType = "Pyropel": Exit Function
    If InStr(s, "4140") > 0 Then NormalizeSteelType = "4140": Exit Function
    If InStr(s, "H-13") > 0 Or InStr(s, "H13") > 0 Then NormalizeSteelType = "H13": Exit Function
    If InStr(s, "A-2") > 0 Or InStr(s, "A2") > 0 Then NormalizeSteelType = "A-2": Exit Function
    If InStr(s, "D-2") > 0 Or InStr(s, "D2") > 0 Then NormalizeSteelType = "D-2": Exit Function
    If InStr(s, "DRILL ROD") > 0 Then NormalizeSteelType = "Drill Rod": Exit Function
    If InStr(s, "STAINLESS") > 0 Or InStr(s, "SS") > 0 Then NormalizeSteelType = "Stainless": Exit Function

    NormalizeSteelType = Trim(matText)
End Function

Private Function IsHardwareName(ByVal d As String) As Boolean
    If InStr(d, "SCREW") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "BOLT") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "WASHER") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "DOWEL") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "NUT ") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SHCS") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SOCKET HEAD") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "SPRING") > 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "PIN") > 0 And InStr(d, "PINION") = 0 Then IsHardwareName = True: Exit Function
    If InStr(d, "O-RING") > 0 Or InStr(d, "ORING") > 0 Then IsHardwareName = True: Exit Function
End Function

Private Function FindCadPartIndexByQuoteOrKeys(ByVal quoteName As String, ByVal pipeKeys As String) As Long
On Error GoTo ErrHandler

    FindCadPartIndexByQuoteOrKeys = 0

    Dim i As Long
    Dim hay As String

    For i = 1 To PartCount
        hay = parts(i).cleanName & " " & parts(i).componentName & " " & parts(i).filePath
        If ContainsAnyPipeKey(hay, pipeKeys) Then
            FindCadPartIndexByQuoteOrKeys = i
            Exit Function
        End If
    Next i

    Dim k As String
    k = NormalizeKey(quoteName)

    For i = 1 To ExportCount
        If NormalizeKey(ExportRows(i).quoteName) = k And ExportRows(i).HasCad Then
            FindCadPartIndexByQuoteOrKeys = ExportRows(i).CadPartIndex
            Exit Function
        End If
    Next i

    Exit Function

ErrHandler:
    FindCadPartIndexByQuoteOrKeys = 0
End Function

Private Function ContainsAnyPipeKey(ByVal haystack As String, ByVal pipeKeys As String) As Boolean
On Error GoTo ErrHandler

    ContainsAnyPipeKey = False

    If haystack = "" Then Exit Function
    If pipeKeys = "" Then Exit Function

    Dim hayUpper As String
    hayUpper = NormalizeText(haystack)

    Dim keys() As String
    keys = Split(pipeKeys, "|")

    Dim i As Long
    Dim k1 As String

    For i = LBound(keys) To UBound(keys)
        k1 = NormalizeText(keys(i))
        If k1 <> "" Then
            If InStr(hayUpper, k1) > 0 Then
                ContainsAnyPipeKey = True
                Exit Function
            End If
        End If
    Next i

    Dim hayKey As String
    hayKey = NormalizeKey(haystack)

    Dim k2 As String

    For i = LBound(keys) To UBound(keys)
        k2 = NormalizeKey(keys(i))
        If k2 <> "" Then
            If InStr(hayKey, k2) > 0 Then
                ContainsAnyPipeKey = True
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    ContainsAnyPipeKey = False
End Function

' ============================================================
' QUARTER-INCH PHYSICAL HIDE
' ============================================================

Private Sub HideQuarterInchThicknessItems(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    HideQuarterInchAssemblyComponents model
    Exit Sub

ErrHandler:
    LogLine "HideQuarterInchThicknessItems error: " & Err.Description
End Sub

Private Sub HideQuarterInchAssemblyComponents(ByVal assyModel As Object)
On Error Resume Next

    Dim vComps As Variant
    vComps = assyModel.GetComponents(False)

    If IsEmpty(vComps) Then Exit Sub

    Dim i As Long
    Dim swComp As Object
    Dim swCompModel As Object
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim L As Double
    Dim W As Double
    Dim T As Double

    For i = 0 To UBound(vComps)

        Set swComp = vComps(i)

        If Not swComp Is Nothing Then
            If swComp.IsSuppressed = False Then

                Set swCompModel = swComp.GetModelDoc2

                If Not swCompModel Is Nothing Then
                    If swCompModel.GetType = swDocPART Then

                        If GetPartBoundingBoxInches(swCompModel, dx, dy, dz) Then
                            SortThreeDimensions dx, dy, dz, L, W, T

                            If IsQuarterInchThickness(T) Then
                                swComp.Visible = swComponentHidden
                            End If
                        End If

                    End If
                End If

            End If
        End If

    Next i
End Sub

' ============================================================
' GENERAL HELPERS
' ============================================================

Private Function NormalizeText(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    NormalizeText = NormalizeSpaces(s)
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeSpaces = Trim(s)
End Function

Private Function NormalizeKey(ByVal s As String) As String
    s = UCase(Trim(s))
    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, ".", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    NormalizeKey = s
End Function

Private Function ProperCaseText(ByVal s As String) As String
On Error Resume Next

    ProperCaseText = StrConv(s, vbProperCase)

    If ProperCaseText = "" Then ProperCaseText = s
End Function

Private Function IsQuarterInchThickness(ByVal T As Double) As Boolean
    IsQuarterInchThickness = (Abs(T - QUARTER_INCH_THICKNESS) <= QUARTER_INCH_TOLERANCE)
End Function

Private Function IsComponentHidden(ByVal swComp As Object) As Boolean
On Error Resume Next

    If swComp Is Nothing Then Exit Function

    IsComponentHidden = (swComp.Visible = swComponentHidden)
End Function

Private Function GetFileBaseName(ByVal path As String) As String
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    GetFileBaseName = fso.GetBaseName(path)
End Function

Private Function GetFileExtension(ByVal path As String) As String
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    GetFileExtension = fso.GetExtensionName(path)
End Function

Private Function CleanDisplayName(ByVal compName As String, ByVal compPath As String) As String
On Error Resume Next

    Dim s As String
    s = compName

    Dim dashPos As Long
    dashPos = InStrRev(s, "-")

    If dashPos > 1 Then
        Dim tail As String
        tail = Mid(s, dashPos + 1)

        If IsNumeric(tail) Then s = Left(s, dashPos - 1)
    End If

    If s = "" And compPath <> "" Then s = GetFileBaseName(compPath)

    CleanDisplayName = Trim(s)
End Function

Private Function CleanFileName(ByVal s As String) As String
    s = Trim(s)
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, ":", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, Chr(34), "_")
    s = Replace(s, "<", "_")
    s = Replace(s, ">", "_")
    s = Replace(s, "|", "_")

    CleanFileName = Trim(s)
End Function

Private Function CleanQuoteTokenForFile(ByVal s As String) As String
    s = CleanFileName(s)
    s = NormalizeSpaces(s)

    CleanQuoteTokenForFile = Trim(s)
End Function

Private Sub EnsureFolder(ByVal folderPath As String)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) = False Then fso.CreateFolder folderPath
End Sub

Private Sub EnsureFolderDeep(ByVal folderPath As String)
On Error Resume Next

    If folderPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)

    If parent <> "" And fso.FolderExists(parent) = False Then
        EnsureFolderDeep parent
    End If

    If fso.FolderExists(folderPath) = False Then fso.CreateFolder folderPath
End Sub

Private Sub CollectJobPdfsAndReports()
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim destFolder As String
    destFolder = CurrentJobFolder & "\" & CurrentJobNumber & " pdfs"

    EnsureFolderDeep destFolder

    Dim copied As Long
    copied = 0

    copied = copied + CopyFilesByExtension(CurrentJobFolder, destFolder, "pdf", destFolder, True)

    Dim names As Variant
    names = Array("XT_Export_BOM_Match_Report.csv", _
                  "XT_Export_CAD_Dimensions.csv", _
                  "XT_Export_BOM_PDF_Text.txt", _
                  JOB_SIGNATURE_REPORT_FILE, _
                  PULLCORE_DIMENSIONS_REPORT_FILE)

    Dim i As Long
    Dim src As String

    For i = LBound(names) To UBound(names)

        src = CurrentJobFolder & "\" & CStr(names(i))

        If fso.FileExists(src) Then
            fso.CopyFile src, destFolder & "\" & CStr(names(i)), True
            copied = copied + 1
        End If

    Next i

    LogLine "Collected " & copied & " PDF/report file(s) into: " & destFolder
    Exit Sub

ErrHandler:
    LogLine "CollectJobPdfsAndReports error: " & Err.Description
End Sub

Private Function CopyFilesByExtension(ByVal srcFolder As String, ByVal destFolder As String, _
                                      ByVal ext As String, ByVal skipFolder As String, _
                                      Optional ByVal deleteSource As Boolean = False) As Long
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(srcFolder) Then Exit Function

    Dim folder As Object
    Set folder = fso.GetFolder(srcFolder)

    Dim count As Long
    count = 0

    Dim f As Object
    Dim srcPath As String

    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.path)) = LCase(ext) Then
            srcPath = f.path
            fso.CopyFile srcPath, destFolder & "\" & f.Name, True
            count = count + 1

            If deleteSource Then fso.DeleteFile srcPath, True
        End If
    Next f

    Dim sub_ As Object

    For Each sub_ In folder.SubFolders
        If LCase(sub_.path) <> LCase(skipFolder) Then
            count = count + CopyFilesByExtension(sub_.path, destFolder, ext, skipFolder, deleteSource)
        End If
    Next sub_

    CopyFilesByExtension = count
End Function

Private Function GetWritableCsvPath(ByVal csvPath As String) As String
On Error GoTo ErrHandler

    Dim f As Integer
    f = FreeFile

    On Error Resume Next
    Err.Clear
    Open csvPath For Output As #f

    If Err.Number = 0 Then
        Close #f
        GetWritableCsvPath = csvPath
        Exit Function
    End If

    Err.Clear

    Dim alt As String
    alt = AppendBeforeExtension(csvPath, "_" & Format(Now, "yyyymmdd_hhnnss"))

    LogLine "CSV locked, writing to alternate file instead: " & alt

    GetWritableCsvPath = alt
    Exit Function

ErrHandler:
    GetWritableCsvPath = csvPath
End Function

Private Function AppendBeforeExtension(ByVal path As String, ByVal suffix As String) As String
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As String
    Dim base As String
    Dim ext As String

    folder = fso.GetParentFolderName(path)
    base = fso.GetBaseName(path)
    ext = fso.GetExtensionName(path)

    If ext = "" Then
        AppendBeforeExtension = folder & "\" & base & suffix
    Else
        AppendBeforeExtension = folder & "\" & base & suffix & "." & ext
    End If
End Function

Private Function GetUniqueFilePath(ByVal path As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) = False Then
        GetUniqueFilePath = path
        Exit Function
    End If

    Dim folder As String
    Dim base As String
    Dim ext As String

    folder = fso.GetParentFolderName(path)
    base = fso.GetBaseName(path)
    ext = fso.GetExtensionName(path)

    Dim n As Long
    Dim candidate As String

    n = 2

    Do
        candidate = folder & "\" & base & "_" & n & "." & ext

        If fso.FileExists(candidate) = False Then
            GetUniqueFilePath = candidate
            Exit Function
        End If

        n = n + 1
    Loop While n < 1000

    GetUniqueFilePath = path
    Exit Function

ErrHandler:
    GetUniqueFilePath = path
End Function

Private Function PowerShellQuote(ByVal s As String) As String
    PowerShellQuote = "'" & Replace(s, "'", "''") & "'"
End Function

Private Sub WaitMilliseconds(ByVal ms As Long)
On Error Resume Next

    If ms <= 0 Then Exit Sub

    Dim startT As Double
    Dim target As Double

    startT = Timer
    target = ms / 1000#

    Do
        DoEvents

        If Timer < startT Then Exit Do
        If (Timer - startT) >= target Then Exit Do
    Loop
End Sub

Private Function CsvText(ByVal s As String) As String
    s = Replace(s, Chr(34), Chr(34) & Chr(34))
    CsvText = Chr(34) & s & Chr(34)
End Function

Private Function FormatNumberForCsv(ByVal v As Double) As String
    FormatNumberForCsv = Format(v, "0.000")
End Function

Private Function ReadAllTextFile(ByVal path As String) As String
On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(path) = False Then Exit Function

    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1, False)

    If Not ts.AtEndOfStream Then
        ReadAllTextFile = ts.ReadAll
    End If

    ts.Close
    Exit Function

ErrHandler:
    ReadAllTextFile = ""
End Function

Private Sub ExtractDecimalNumbers(ByVal line As String, ByRef nums() As Double, ByRef numCount As Long)
On Error GoTo ErrHandler

    numCount = 0
    ReDim nums(0 To 50)

    Dim i As Long
    Dim ch As String
    Dim token As String

    token = ""

    For i = 1 To Len(line) + 1

        If i <= Len(line) Then
            ch = Mid(line, i, 1)
        Else
            ch = " "
        End If

        If (ch >= "0" And ch <= "9") Or ch = "." Then

            token = token & ch

        Else

            If token <> "" Then

                If IsNumeric(token) Then
                    If InStr(token, ".") > 0 Then

                        nums(numCount) = CDbl(token)
                        numCount = numCount + 1

                        If numCount > 50 Then Exit For

                    End If
                End If

                token = ""

            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    numCount = 0
End Sub

Private Function ExtractQtyAfterOutsource(ByVal upperLine As String) As Long
On Error GoTo ErrHandler

    ExtractQtyAfterOutsource = 1

    Dim kw As Variant
    Dim p As Long

    For Each kw In Array("OUTSOURCE", "PURCHASE", "MAKE", "BUY", "STOCK")

        p = InStr(upperLine, CStr(kw))

        If p > 0 Then

            Dim tail As String
            tail = Mid(upperLine, p + Len(CStr(kw)))

            Dim i As Long
            Dim ch As String
            Dim token As String

            token = ""

            For i = 1 To Len(tail)

                ch = Mid(tail, i, 1)

                If ch >= "0" And ch <= "9" Then
                    token = token & ch
                ElseIf token <> "" Then
                    Exit For
                End If

            Next i

            If token <> "" Then
                ExtractQtyAfterOutsource = CLng(token)
                Exit Function
            End If

        End If

    Next kw

    Exit Function

ErrHandler:
    ExtractQtyAfterOutsource = 1
End Function

Private Function RemoveLeadingItemNumber(ByVal line As String) As String
On Error Resume Next

    Dim s As String
    s = Trim(line)

    Dim i As Long
    Dim ch As String
    Dim digits As String

    digits = ""

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            Exit For
        End If
    Next i

    If digits <> "" And Len(digits) <= 4 Then
        s = Trim(Mid(s, Len(digits) + 1))
    End If

    RemoveLeadingItemNumber = s
End Function

' ============================================================
' LOGGING
' ============================================================

Private Sub LogLine(ByVal msg As String)
On Error Resume Next

    If ENABLE_EXPORT_LOG = False Then Exit Sub

    Dim f As Integer
    f = FreeFile

    Dim path As String
    path = RunLogPath

    If path = "" Then path = StartupLogPath
    If path = "" Then path = Environ$("USERPROFILE") & "\Desktop\CMS_XT_Export_STARTUP_Log.txt"

    Open path For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub

Private Sub LogStart(ByVal stepName As String)
    CurrentStepName = stepName
    StepStartTime = Now
    LogLine ">>> START: " & stepName
End Sub

Private Sub LogDone(ByVal stepName As String)
    LogLine "<<< DONE : " & stepName & " (" & DateDiff("s", StepStartTime, Now) & "s)"
End Sub

Private Sub LogProgress(ByVal msg As String)
    If VERBOSE_PROGRESS_LOG Then LogLine "    ... " & msg
End Sub

Private Sub LogErrorText(ByVal msg As String)
    LogLine "ERROR: " & msg
    LastJobFailReason = msg

    If SHOW_ERROR_MESSAGES Then
        ' Batch-safe: no blocking message box here.
    End If
End Sub

' ============================================================
' END OF MODULE
' ============================================================

