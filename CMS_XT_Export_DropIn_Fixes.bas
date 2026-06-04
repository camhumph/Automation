Option Explicit

' ============================================================
' CMS XT Export macro drop-in fixes
' ------------------------------------------------------------
' This repository did not contain the original SolidWorks macro
' module, so this file contains the replacement VBA sections to
' paste into the existing CMS XT Export macro.
'
' Fixes included:
'   1. J BLOCK DXF uses a temporary native SLDASM instead of X_T.
'   2. All DXF views are forced to 1:1, including BASE/HOLDERS.
'   3. Drawing sheet scale is forced to 1:1.
'   4. Every drawing view scale setter honors the global 1:1 flag.
' ============================================================

' ============================================================
' REPLACE CreateJBlockDxfFromAssemblyTopView WITH THIS VERSION
' ============================================================

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
    ' Use native SLDASM, not X_T. X_T does not preserve CMS_TOP /
    ' redefined SolidWorks standard views, which is why the J BLOCK
    ' could still arrive incorrectly oriented when reopened for DXF.
    tempNativePath = tempFolder & "\" & CurrentJobNumber & "_JBLOCK_TOPVIEW_TEMP.sldasm"

    Dim hiddenNames As Collection
    Set hiddenNames = New Collection

    If HideAllExceptComponentNamesOnce(swModel, keepNames, hiddenNames) = False Then
        LogLine "J BLOCK assembly-top DXF: could not isolate J BLOCK component."
        GoTo CleanExit
    End If

    ' Put the live assembly into the corrected TCP/top orientation before
    ' saving the isolated native assembly used as the drawing source.
    ApplyCmsTopView swModel
    StabilizeActiveView swModel, 100

    LogLine "J BLOCK DXF: saving isolated native SLDASM from corrected CMS_TOP orientation:"
    LogLine "  " & tempNativePath

    SaveModelAs swModel, tempNativePath

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(tempNativePath) = False Then
        LogLine "J BLOCK assembly-top DXF: temp native SLDASM was not created."
        GoTo CleanExit
    End If

    ' True assembly top-down J BLOCK DXF:
    ' the center view matches the corrected BASE / CMS_TOP orientation.
    CreateProjectedDxfFromNativePath tempNativePath, dxfPath, "J BLOCK", _
                                     CMS_TOP_VIEW_NAME, "*Top", _
                                     True, False, False, False, False

    ' If the old J BLOCK side-profile layout is preferred instead, use
    ' this call in place of the CMS_TOP call above. It still keeps the
    ' native assembly path, so standard views are preserved:
    '
    'CreateProjectedDxfFromNativePath tempNativePath, dxfPath, "J BLOCK", _
    '                                 "*Right", "*Right", _
    '                                 True, False, False, False, False

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

' ============================================================
' PATCH CreateProjectedDxfFromNativePath SCALE SECTION
' ============================================================
'
' In CreateProjectedDxfFromNativePath, replace this old block:
'
'     CurrentDxfForce1to1 = FORCE_ALL_DXF_VIEWS_1_TO_1
'
'     If IsMainBaseDxfQuote(quoteName) Or IsHoldersDxfQuote(quoteName) Then
'         CurrentDxfForce1to1 = False
'     End If
'
' and replace the following Dim scaleVal block with the code below.
' This makes FORCE_ALL_DXF_VIEWS_1_TO_1 absolute for every DXF,
' including BASE, HOLDERS, J BLOCK, TCP, BCP, PULLCORE, and PYROPEL.
'
'     Dim scaleVal As Double
'
'     If FORCE_ALL_DXF_VIEWS_1_TO_1 Then
'
'         CurrentDxfForce1to1 = True
'         scaleVal = 1#
'
'         LogLine "DXF scale forced 1:1 for " & quoteName
'
'     Else
'
'         CurrentDxfForce1to1 = False
'
'         scaleVal = CalculateProjectedFourViewDxfScale(partL, partW, partT)
'         scaleVal = scaleVal * MULTIVIEW_FIT_SAFETY
'
'         If scaleVal <= 0 Then scaleVal = 0.1
'
'         LogLine "DXF auto-fit scale for " & quoteName & " = " & Format(scaleVal, "0.0000")
'
'     End If
'
'     If scaleVal = 1# Then
'         Dim layoutWCheck As Double
'         Dim layoutHCheck As Double
'
'         layoutWCheck = partL + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
'         layoutHCheck = partW + (2# * partT) + (2# * DXF_PROJECTED_VIEW_GAP_IN)
'
'         If layoutWCheck > (E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) Or _
'            layoutHCheck > (E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN)) Then
'
'             LogLine "WARNING: 1:1 DXF layout may not fit E-size sheet for " & quoteName & _
'                     ". Required W/H=" & FormatNumberForCsv(layoutWCheck) & "/" & _
'                     FormatNumberForCsv(layoutHCheck) & _
'                     ", usable W/H=" & FormatNumberForCsv(E_SHEET_WIDTH_IN - (2# * DXF_MARGIN_IN)) & "/" & _
'                     FormatNumberForCsv(E_SHEET_HEIGHT_IN - (2# * DXF_MARGIN_IN))
'
'         End If
'     End If

' ============================================================
' REPLACE SetupDrawingAsESize WITH THIS VERSION
' ============================================================

Private Sub SetupDrawingAsESize(ByVal swDraw As Object)
On Error Resume Next

    If swDraw Is Nothing Then Exit Sub

    Dim swSheet As Object
    Set swSheet = swDraw.GetCurrentSheet

    If Not swSheet Is Nothing Then

        ' E-size sheet.
        swSheet.SetSize 12, E_SHEET_WIDTH_IN / INCHES_PER_METER, E_SHEET_HEIGHT_IN / INCHES_PER_METER

        ' Force sheet scale 1:1. Late-bound because some SolidWorks
        ' versions expose slightly different signatures.
        Err.Clear
        swSheet.SetScale 1#, 1#, True, True
        Err.Clear

    End If

    swDraw.GraphicsRedraw2
End Sub

' ============================================================
' REPLACE SetDrawingViewScale WITH THIS VERSION
' ============================================================

Private Sub SetDrawingViewScale(ByVal swView As Object, ByVal scaleVal As Double)
On Error Resume Next

    If swView Is Nothing Then Exit Sub

    If FORCE_ALL_DXF_VIEWS_1_TO_1 Or CurrentDxfForce1to1 Then
        scaleVal = 1#
    End If

    If scaleVal <= 0 Then scaleVal = 1#

    swView.UseSheetScale = False
    swView.ScaleDecimal = scaleVal

    If Abs(scaleVal - 1#) < 0.000001 Then
        swView.ScaleRatio = "1:1"
    End If
End Sub
