Attribute VB_Name = "MODULE6121"
Option Explicit

' ============================================================
' MODULE6121 - Recent-month job search + CAD naming library
' ------------------------------------------------------------
' Job folder search uses the current month plus JOB_SEARCH_MONTHS_BACK
' prior months (default 2 => June + May + April when run in June).
'
' Also seeds CMS_Block_Naming_Library.csv from recent macro exports:
'   XT_Export_BOM_Match_Report.csv
'   XT_Export_Job_Signature.csv
'   XT_Export_Job_Signature_J####.csv
' ============================================================

Public Const JOB_SEARCH_LIMIT_TO_RECENT_MONTHS As Boolean = True
Public Const JOB_SEARCH_MONTHS_BACK As Long = 2

Public Const M6121_SEED_LIBRARY_FROM_RECENT_JOBS As Boolean = True
Public Const M6121_CAD_NAMING_LIBRARY_FILE As String = "CMS_Block_Naming_Library.csv"

Private Const M6121_PUBLIC_DOWNLOADS_PATH As String = "\\Mycloudex2ultra\mexico\Downloads"
Private Const M6121_LOCAL_DOWNLOADS_FALLBACK As String = "C:\Users\lenovo\Desktop"
Private Const M6121_LOCAL_WORKSPACE_ROOT As String = "C:\CMS_Local_Workspace"
Private Const M6121_PUBLIC_DATA_ROOT As String = "\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software"
Private Const M6121_PRIVATE_DATA_ROOT As String = "C:\CMS_Local_Workspace\Matching"
Private Const M6121_COMPANY_WIFI_SSID As String = "NETGEAR"
Private Const M6121_MATCH_REPORT_FILE As String = "XT_Export_BOM_Match_Report.csv"
Private Const M6121_SIGNATURE_FILE As String = "XT_Export_Job_Signature.csv"

Private M6121_LogPath As String

' ============================================================
' PUBLIC ENTRY POINTS
' ============================================================

Public Sub BuildCadNamingLibrary()
' Run standalone from the macro list to rebuild/extend the naming library.
On Error GoTo ErrHandler

    M6121_InitLog
    M6121_LogLine "========================================"
    M6121_LogLine "MODULE6121 BuildCadNamingLibrary started"
    M6121_LogLine "Allowed months: " & M6121_AllowedMonthLabelsText()
    M6121_LogLine "========================================"

    SeedCadNamingLibraryFromRecentJobs

    M6121_LogLine "MODULE6121 BuildCadNamingLibrary finished."
    MsgBox "CAD naming library seed finished." & vbCrLf & _
           "Log: " & M6121_LogPath, vbInformation, "MODULE6121"
    Exit Sub

ErrHandler:
    M6121_LogLine "BuildCadNamingLibrary error: " & Err.Description
    MsgBox "MODULE6121 library build error: " & Err.Description, vbExclamation
End Sub

Public Sub SeedCadNamingLibraryFromRecentJobs()
On Error GoTo ErrHandler

    If M6121_SEED_LIBRARY_FROM_RECENT_JOBS = False Then Exit Sub

    M6121_InitLog
    M6121_LogLine "MODULE6121 seeding CAD naming library from recent job exports."
    M6121_LogLine "Month window: " & M6121_AllowedMonthLabelsText()

    Dim libPath As String
    libPath = M6121_GetCadNamingLibraryPath()

    M6121_EnsureCadNamingLibraryHeader libPath

    Dim existing As Object
    Set existing = M6121_LoadCadNamingLibraryKeyDict(libPath)

    Dim f As Integer
    f = FreeFile

    Open libPath For Append As #f

    Dim addedCount As Long
    addedCount = 0

    Dim root As String

    root = M6121_ResolveRootJobPath()
    If root <> "" Then
        M6121_LogLine "Scanning downloads root: " & root
        M6121_ScanDownloadsRootForLibraryRows root, existing, f, addedCount
    End If

    root = M6121_ResolveMatchingRoot()
    If root <> "" Then
        M6121_LogLine "Scanning matching root: " & root
        M6121_ScanMatchingRootForLibraryRows root, existing, f, addedCount
    End If

    If M6121_LOCAL_WORKSPACE_ROOT <> "" Then
        M6121_LogLine "Scanning local workspace: " & M6121_LOCAL_WORKSPACE_ROOT
        M6121_ScanLocalWorkspaceForLibraryRows M6121_LOCAL_WORKSPACE_ROOT, existing, f, addedCount
    End If

    Close #f

    M6121_LogLine "MODULE6121 library seed added rows=" & addedCount
    M6121_LogLine "MODULE6121 library path: " & libPath
    Exit Sub

ErrHandler:
    M6121_LogLine "SeedCadNamingLibraryFromRecentJobs error: " & Err.Description
    On Error Resume Next
    Close #f
End Sub

Public Function JobSearchTopFolderIsAllowed(ByVal folderName As String, _
                                            ByVal wantUpper As String) As Boolean
On Error GoTo ErrHandler

    JobSearchTopFolderIsAllowed = False

    If JOB_SEARCH_LIMIT_TO_RECENT_MONTHS = False Then
        JobSearchTopFolderIsAllowed = True
        Exit Function
    End If

    Dim n As String
    n = UCase(Trim(folderName))

    If wantUpper <> "" Then
        If InStr(n, UCase(wantUpper)) > 0 Then
            JobSearchTopFolderIsAllowed = True
            Exit Function
        End If
    End If

    If M6121_FolderNameMatchesAllowedMonth(n) Then
        JobSearchTopFolderIsAllowed = True
        Exit Function
    End If

    Exit Function

ErrHandler:
    JobSearchTopFolderIsAllowed = True
End Function

' ============================================================
' MONTH WINDOW
' ============================================================

Private Function M6121_FolderNameMatchesAllowedMonth(ByVal folderUpper As String) As Boolean
On Error Resume Next

    Dim i As Long
    Dim monthLabel As String

    For i = 0 To JOB_SEARCH_MONTHS_BACK
        monthLabel = UCase(Format(DateAdd("m", -i, Date), "mmmm yyyy"))
        If InStr(folderUpper, monthLabel) > 0 Then
            M6121_FolderNameMatchesAllowedMonth = True
            Exit Function
        End If
    Next i

    M6121_FolderNameMatchesAllowedMonth = False
End Function

Private Function M6121_AllowedMonthLabelsText() As String
On Error Resume Next

    Dim i As Long
    Dim parts As String

    parts = ""

    For i = 0 To JOB_SEARCH_MONTHS_BACK
        If parts <> "" Then parts = parts & ", "
        parts = parts & Format(DateAdd("m", -i, Date), "mmmm yyyy")
    Next i

    M6121_AllowedMonthLabelsText = parts
End Function

' ============================================================
' LIBRARY SCANNING
' ============================================================

Private Sub M6121_ScanDownloadsRootForLibraryRows(ByVal rootPath As String, _
                                                  ByVal existing As Object, _
                                                  ByVal f As Integer, _
                                                  ByRef addedCount As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootPath) Then Exit Sub

    Dim topFolder As Object
    For Each topFolder In fso.GetFolder(rootPath).SubFolders

        If M6121_FolderNameMatchesAllowedMonth(UCase(topFolder.name)) Then
            M6121_ScanFolderTreeForLibraryCsv topFolder, existing, f, addedCount, 0
        End If

    Next topFolder
End Sub

Private Sub M6121_ScanMatchingRootForLibraryRows(ByVal rootPath As String, _
                                                 ByVal existing As Object, _
                                                 ByVal f As Integer, _
                                                 ByRef addedCount As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootPath) Then Exit Sub

    M6121_TryImportLibraryCsvFile rootPath & "\" & M6121_SIGNATURE_FILE, existing, f, addedCount

    Dim fn As String
    fn = Dir(rootPath & "\XT_Export_Job_Signature_*.csv")
    Do While fn <> ""
        M6121_TryImportLibraryCsvFile rootPath & "\" & fn, existing, f, addedCount
        fn = Dir()
    Loop

    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(rootPath).SubFolders
        M6121_TryImportLibraryCsvFile subFolder.path & "\" & M6121_SIGNATURE_FILE, existing, f, addedCount
        M6121_TryImportLibraryCsvFile subFolder.path & "\" & M6121_MATCH_REPORT_FILE, existing, f, addedCount
    Next subFolder
End Sub

Private Sub M6121_ScanLocalWorkspaceForLibraryRows(ByVal rootPath As String, _
                                                   ByVal existing As Object, _
                                                   ByVal f As Integer, _
                                                   ByRef addedCount As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootPath) Then Exit Sub

    M6121_ScanFolderTreeForLibraryCsv fso.GetFolder(rootPath), existing, f, addedCount, 0
End Sub

Private Sub M6121_ScanFolderTreeForLibraryCsv(ByVal folder As Object, _
                                              ByVal existing As Object, _
                                              ByVal f As Integer, _
                                              ByRef addedCount As Long, _
                                              ByVal depth As Long)
On Error Resume Next

    If folder Is Nothing Then Exit Sub
    If depth > 4 Then Exit Sub

    M6121_TryImportLibraryCsvFile folder.path & "\" & M6121_MATCH_REPORT_FILE, existing, f, addedCount
    M6121_TryImportLibraryCsvFile folder.path & "\" & M6121_SIGNATURE_FILE, existing, f, addedCount

    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        M6121_ScanFolderTreeForLibraryCsv subFolder, existing, f, addedCount, depth + 1
    Next subFolder
End Sub

Private Sub M6121_TryImportLibraryCsvFile(ByVal csvPath As String, _
                                          ByVal existing As Object, _
                                          ByVal f As Integer, _
                                          ByRef addedCount As Long)
On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(csvPath) Then Exit Sub

    Dim low As String
    low = LCase(csvPath)

    If InStr(low, "job_signature") > 0 Then
        M6121_ImportSignatureCsv csvPath, existing, f, addedCount
    ElseIf InStr(low, "bom_match_report") > 0 Then
        M6121_ImportMatchReportCsv csvPath, existing, f, addedCount
    End If
End Sub

Private Sub M6121_ImportMatchReportCsv(ByVal csvPath As String, _
                                       ByVal existing As Object, _
                                       ByVal f As Integer, _
                                       ByRef addedCount As Long)
On Error Resume Next

    Dim jobNum As String
    jobNum = M6121_TryExtractJobNumberFromPath(csvPath)

    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(csvPath, 1, False)

    Dim line As String
    Dim fields() As String

    If Not ts.AtEndOfStream Then line = ts.ReadLine

    Do While Not ts.AtEndOfStream

        line = ts.ReadLine
        If Trim(line) = "" Then GoTo NextLine

        fields = M6121_ParseCsvLine(line)

        If UBound(fields) >= 11 Then

            Dim quoteName As String
            Dim cadName As String
            Dim L As Double
            Dim W As Double
            Dim T As Double

            quoteName = Trim(fields(1))
            L = CDbl(Val(fields(7)))
            W = CDbl(Val(fields(8)))
            T = CDbl(Val(fields(9)))
            cadName = Trim(fields(10))

            If M6121_ShouldLearnQuote(quoteName) Then
                If cadName <> "" And L > 0 And W > 0 And T > 0 Then
                    If M6121_AppendLibraryRow(f, existing, jobNum, quoteName, cadName, _
                                              L, W, T, 0#, 0#, 0#, 0.5, 0.5, 0.5, _
                                              "match-report", addedCount) Then
                        M6121_LogLine "Library seed match-report: " & jobNum & " " & quoteName & " -> " & cadName
                    End If
                End If
            End If

        End If

NextLine:
    Loop

    ts.Close
End Sub

Private Sub M6121_ImportSignatureCsv(ByVal csvPath As String, _
                                     ByVal existing As Object, _
                                     ByVal f As Integer, _
                                     ByRef addedCount As Long)
On Error Resume Next

    Dim fallbackJob As String
    fallbackJob = M6121_TryExtractJobNumberFromPath(csvPath)

    Dim ts As Object
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(csvPath, 1, False)

    Dim line As String
    Dim fields() As String

    If Not ts.AtEndOfStream Then line = ts.ReadLine

    Do While Not ts.AtEndOfStream

        line = ts.ReadLine
        If Trim(line) = "" Then GoTo NextSigLine

        fields = M6121_ParseCsvLine(line)

        If UBound(fields) >= 14 Then

            Dim jobNum As String
            Dim quoteName As String
            Dim cadName As String
            Dim L As Double
            Dim W As Double
            Dim T As Double
            Dim cx As Double
            Dim cy As Double
            Dim cz As Double

            jobNum = Trim(fields(0))
            If jobNum = "" Then jobNum = fallbackJob

            quoteName = Trim(fields(4))
            If quoteName = "" Then quoteName = Trim(fields(3))

            cadName = Trim(fields(5))
            L = CDbl(Val(fields(7)))
            W = CDbl(Val(fields(8)))
            T = CDbl(Val(fields(9)))
            cx = CDbl(Val(fields(11)))
            cy = CDbl(Val(fields(12)))
            cz = CDbl(Val(fields(13)))

            If M6121_ShouldLearnQuote(quoteName) Then
                If cadName <> "" And L > 0 And W > 0 And T > 0 Then
                    If M6121_AppendLibraryRow(f, existing, jobNum, quoteName, cadName, _
                                              L, W, T, cx, cy, cz, 0.5, 0.5, 0.5, _
                                              "signature", addedCount) Then
                        M6121_LogLine "Library seed signature: " & jobNum & " " & quoteName & " -> " & cadName
                    End If
                End If
            End If

        End If

NextSigLine:
    Loop

    ts.Close
End Sub

Private Function M6121_ShouldLearnQuote(ByVal quoteName As String) As Boolean
On Error Resume Next

    Dim q As String
    q = UCase(Trim(quoteName))

    If q = "" Then Exit Function
    If q = "MAIN ASSEMBLY" Then Exit Function
    If q = "HOLDERS" Then Exit Function

    M6121_ShouldLearnQuote = True
End Function

Private Function M6121_AppendLibraryRow(ByVal f As Integer, _
                                          ByVal existing As Object, _
                                          ByVal jobNum As String, _
                                          ByVal quoteName As String, _
                                          ByVal componentName As String, _
                                          ByVal L As Double, _
                                          ByVal W As Double, _
                                          ByVal T As Double, _
                                          ByVal cx As Double, _
                                          ByVal cy As Double, _
                                          ByVal cz As Double, _
                                          ByVal nx As Double, _
                                          ByVal ny As Double, _
                                          ByVal nz As Double, _
                                          ByVal sourceTag As String, _
                                          ByRef addedCount As Long) As Boolean
On Error GoTo ErrHandler

    M6121_AppendLibraryRow = False

    Dim j As String
    j = M6121_NormalizeJobNumber(jobNum)

    If j = "" Then Exit Function
    If Trim(quoteName) = "" Then Exit Function
    If Trim(componentName) = "" Then Exit Function

    Dim key As String
    key = j & "|" & M6121_NormalizeKey(quoteName) & "|" & M6121_NormalizeKey(componentName)

    If Not existing Is Nothing Then
        If existing.Exists(key) Then Exit Function
    End If

    Print #f, M6121_CsvText("1") & "," & _
              M6121_CsvText(j) & "," & _
              M6121_CsvText(quoteName) & "," & _
              M6121_FormatNum(L) & "," & _
              M6121_FormatNum(W) & "," & _
              M6121_FormatNum(T) & "," & _
              M6121_FormatNum(cx) & "," & _
              M6121_FormatNum(cy) & "," & _
              M6121_FormatNum(cz) & "," & _
              M6121_FormatNum(nx) & "," & _
              M6121_FormatNum(ny) & "," & _
              M6121_FormatNum(nz) & "," & _
              M6121_CsvText(componentName) & "," & _
              M6121_CsvText(Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & sourceTag)

    If Not existing Is Nothing Then existing(key) = True

    addedCount = addedCount + 1
    M6121_AppendLibraryRow = True
    Exit Function

ErrHandler:
    M6121_AppendLibraryRow = False
End Function

' ============================================================
' LIBRARY FILE HELPERS
' ============================================================

Private Function M6121_GetCadNamingLibraryPath() As String
On Error Resume Next

    Dim root As String
    root = M6121_ResolveMatchingRoot()

    If root = "" Then root = M6121_LOCAL_WORKSPACE_ROOT

    M6121_EnsureFolderDeep root

    M6121_GetCadNamingLibraryPath = root & "\" & M6121_CAD_NAMING_LIBRARY_FILE
End Function

Private Sub M6121_EnsureCadNamingLibraryHeader(ByVal libPath As String)
On Error Resume Next

    If libPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    M6121_EnsureFolderDeep fso.GetParentFolderName(libPath)

    If fso.FileExists(libPath) Then Exit Sub

    Dim outF As Integer
    outF = FreeFile

    Open libPath For Output As #outF
    Print #outF, "Version,JobNumber,QuoteName,Length,Width,Thickness,CenterX,CenterY,CenterZ,NormX,NormY,NormZ,ComponentName,LearnedOn"
    Close #outF

    M6121_LogLine "Created CAD naming library: " & libPath
End Sub

Private Function M6121_LoadCadNamingLibraryKeyDict(ByVal libPath As String) As Object
On Error Resume Next

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(libPath) Then
        Set M6121_LoadCadNamingLibraryKeyDict = dict
        Exit Function
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(libPath, 1, False)

    Dim line As String
    Dim fields() As String
    Dim key As String

    If Not ts.AtEndOfStream Then line = ts.ReadLine

    Do While Not ts.AtEndOfStream

        line = ts.ReadLine
        If Trim(line) = "" Then GoTo NextKeyLine

        fields = M6121_ParseCsvLine(line)

        If UBound(fields) >= 12 Then
            key = Trim(fields(1)) & "|" & M6121_NormalizeKey(fields(2)) & "|" & M6121_NormalizeKey(fields(12))
            dict(key) = True
        End If

NextKeyLine:
    Loop

    ts.Close
    Set M6121_LoadCadNamingLibraryKeyDict = dict
End Function

' ============================================================
' PATH RESOLUTION (mirrors gemini2 network-aware roots)
' ============================================================

Private Function M6121_ResolveRootJobPath() As String
On Error Resume Next

    Dim root As String

    If M6121_IsOnCompanyWifi() Then
        root = M6121_PUBLIC_DOWNLOADS_PATH
    Else
        root = M6121_LOCAL_DOWNLOADS_FALLBACK
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(M6121_PUBLIC_DOWNLOADS_PATH) Then root = M6121_PUBLIC_DOWNLOADS_PATH
    End If

    M6121_ResolveRootJobPath = root
End Function

Private Function M6121_ResolveMatchingRoot() As String
On Error Resume Next

    Dim root As String

    If M6121_IsOnCompanyWifi() Then
        root = M6121_PUBLIC_DATA_ROOT
    Else
        root = M6121_PRIVATE_DATA_ROOT
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FolderExists(M6121_PUBLIC_DATA_ROOT) Then root = M6121_PUBLIC_DATA_ROOT
    End If

    If root = "" Then root = M6121_PUBLIC_DATA_ROOT

    M6121_EnsureFolderDeep root
    M6121_ResolveMatchingRoot = root
End Function

Private Function M6121_IsOnCompanyWifi() As Boolean
On Error Resume Next

    Dim ssid As String
    ssid = UCase$(M6121_GetWifiSsid())

    If ssid <> "" Then
        If InStr(ssid, UCase$(M6121_COMPANY_WIFI_SSID)) > 0 Then
            M6121_IsOnCompanyWifi = True
            Exit Function
        End If
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(M6121_PUBLIC_DATA_ROOT) Then M6121_IsOnCompanyWifi = True
End Function

Private Function M6121_GetWifiSsid() As String
On Error Resume Next

    Dim sh As Object
    Dim ex As Object
    Dim out As String

    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec("netsh wlan show interfaces")

    If ex Is Nothing Then Exit Function

    out = ex.StdOut.ReadAll

    Dim arr() As String
    Dim i As Long
    Dim s As String
    Dim p As Long

    arr = Split(out, vbCrLf)

    For i = LBound(arr) To UBound(arr)
        s = Trim(arr(i))
        If InStr(1, s, "SSID", vbTextCompare) = 1 Then
            p = InStr(s, ":")
            If p > 0 Then
                M6121_GetWifiSsid = Trim(Mid(s, p + 1))
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub M6121_EnsureFolderDeep(ByVal folderPath As String)
On Error Resume Next

    If folderPath = "" Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)

    If parent <> "" And parent <> folderPath Then
        M6121_EnsureFolderDeep parent
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

' ============================================================
' SMALL UTILITIES
' ============================================================

Private Function M6121_TryExtractJobNumberFromPath(ByVal pathText As String) As String
On Error Resume Next

    Dim parts() As String
    Dim i As Long
    Dim seg As String
    Dim digits As String

    parts = Split(pathText, "\")

    For i = UBound(parts) To LBound(parts) Step -1

        seg = UCase(Trim(parts(i)))

        If Len(seg) >= 4 Then

            If Left(seg, 1) = "J" And IsNumeric(Mid(seg, 2, 4)) Then
                M6121_TryExtractJobNumberFromPath = "J" & Mid(seg, 2, 4)
                Exit Function
            End If

            digits = ""
            If IsNumeric(Left(seg, 4)) Then
                M6121_TryExtractJobNumberFromPath = Left(seg, 4)
                Exit Function
            End If

            If InStr(seg, "JOB_SIGNATURE_") > 0 Then
                digits = Replace(seg, "XT_EXPORT_JOB_SIGNATURE_", "")
                digits = Replace(digits, ".CSV", "")
                If IsNumeric(digits) Then
                    M6121_TryExtractJobNumberFromPath = digits
                    Exit Function
                End If
            End If

        End If

    Next i
End Function

Private Function M6121_NormalizeJobNumber(ByVal jobText As String) As String
On Error Resume Next

    Dim s As String
    s = UCase(Trim(jobText))

    If s = "" Then Exit Function

    If Left(s, 1) = "J" And Len(s) >= 5 Then
        M6121_NormalizeJobNumber = Left(s, 5)
        Exit Function
    End If

    If IsNumeric(Left(s, 4)) Then
        M6121_NormalizeJobNumber = Left(s, 4)
        Exit Function
    End If

    M6121_NormalizeJobNumber = s
End Function

Private Function M6121_NormalizeKey(ByVal s As String) As String
On Error Resume Next

    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim src As String

    out = ""
    src = UCase(Trim(s))

    For i = 1 To Len(src)
        ch = Mid(src, i, 1)
        If ch >= "A" And ch <= "Z" Then
            out = out & ch
        ElseIf ch >= "0" And ch <= "9" Then
            out = out & ch
        End If
    Next i

    M6121_NormalizeKey = out
End Function

Private Function M6121_ParseCsvLine(ByVal line As String) As String()
On Error Resume Next

    Dim fields As Collection
    Set fields = New Collection

    Dim i As Long
    Dim ch As String
    Dim cur As String
    Dim inQuote As Boolean

    cur = ""
    inQuote = False

    For i = 1 To Len(line)

        ch = Mid(line, i, 1)

        If ch = """" Then
            If inQuote Then
                If i < Len(line) And Mid(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                inQuote = True
            End If
        ElseIf ch = "," And inQuote = False Then
            fields.Add cur
            cur = ""
        Else
            cur = cur & ch
        End If

    Next i

    fields.Add cur

    Dim arr() As String
    ReDim arr(0 To fields.count - 1)

    Dim k As Long
    For k = 1 To fields.count
        arr(k - 1) = CStr(fields(k))
    Next k

    M6121_ParseCsvLine = arr
End Function

Private Function M6121_CsvText(ByVal s As String) As String
On Error Resume Next

    Dim t As String
    t = Replace(s, """", """""")

    If InStr(t, ",") > 0 Or InStr(t, """") > 0 Or InStr(t, vbCr) > 0 Or InStr(t, vbLf) > 0 Then
        M6121_CsvText = """" & t & """"
    Else
        M6121_CsvText = t
    End If
End Function

Private Function M6121_FormatNum(ByVal v As Double) As String
On Error Resume Next
    M6121_FormatNum = Replace(Format(v, "0.000"), ",", "")
End Function

Private Sub M6121_InitLog()
On Error Resume Next
    M6121_LogPath = Environ$("USERPROFILE") & "\Desktop\CMS_MODULE6121_Log.txt"
End Sub

Private Sub M6121_LogLine(ByVal msg As String)
On Error Resume Next

    If M6121_LogPath = "" Then M6121_InitLog

    Dim f As Integer
    f = FreeFile

    Open M6121_LogPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub
