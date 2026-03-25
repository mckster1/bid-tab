Attribute VB_Name = "TSC_v10_GenerateTradeTabs"
Option Explicit

Public Sub GenerateTradeTabs_v10()
    Dim wb As Workbook: Set wb = ThisWorkbook
    If wb.ProtectStructure Then
        MsgBox "Workbook STRUCTURE is protected. Unprotect first.", vbExclamation
        Exit Sub
    End If

    Dim cfg As Worksheet, tpl As Worksheet
    On Error Resume Next
    Set cfg = wb.Worksheets(SHEET_CONFIG)
    Set tpl = wb.Worksheets(SHEET_TEMPLATE)
    On Error GoTo 0
    If cfg Is Nothing Or tpl Is Nothing Then
        MsgBox "Missing Config_CSI or TradeTemplate.", vbExclamation
        Exit Sub
    End If

    ' Ensure anchors exist on template
    tpl.Range("B17").Value = KEY_ADD_SCOPE
    tpl.Range("B23").Value = KEY_ADD_ALT

    ' Ensure evaluation row labels exist on template
    WriteEvalLabels tpl

    Dim colTrade As Long, colCSI As Long, colShort As Long, colCreate As Long, colDefaults As Long
    colTrade = FindHeaderCol(cfg, "TradeName")
    colCSI = FindHeaderCol(cfg, "CSI_Code")
    colShort = FindHeaderCol(cfg, "ShortName")
    colCreate = FindHeaderCol(cfg, "Create")
    colDefaults = FindHeaderCol(cfg, "DefaultScopeLines")
    If colTrade * colCSI * colShort * colCreate * colDefaults = 0 Then
        MsgBox "Config_CSI headers required: TradeName, CSI_Code, ShortName, Create, DefaultScopeLines", vbExclamation
        Exit Sub
    End If

    Dim overwriteDefaults As VbMsgBoxResult
    overwriteDefaults = MsgBox("Overwrite default scope lines on existing tabs?" & vbCrLf & _
                              "Yes = clear seeded + re-seed" & vbCrLf & _
                              "No = seed only on NEW tabs", vbYesNoCancel + vbQuestion)
    If overwriteDefaults = vbCancel Then Exit Sub

    Dim jobTitle As String, estimator As String, bidDate As String, jobGsf As String, tradeSf As String
    jobTitle = Trim$(InputBox("Job Title (" & CELL_JOBTITLE & "). Leave blank to skip.", "Setup – Job Info"))
    estimator = Trim$(InputBox("Estimator (" & CELL_ESTIMATOR & "). Leave blank to skip.", "Setup – Job Info"))
    bidDate = Trim$(InputBox("Bid Date (" & CELL_BIDDATE & "). Leave blank to skip.", "Setup – Job Info"))
    jobGsf = Trim$(InputBox("Job GSF (" & CELL_JOB_GSF & "). Leave blank to skip.", "Setup – Job Info"))
    tradeSf = Trim$(InputBox("Trade SF (" & CELL_TRADE_SF & "). Leave blank to skip.", "Setup – Job Info"))

    Dim lastRow As Long: lastRow = cfg.Cells(cfg.Rows.Count, colTrade).End(xlUp).Row
    Dim r As Long, created As Long, updated As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 2 To lastRow
        Dim tradeName As String, csiRaw As Variant, shortName As String, createFlag As Variant, defaultsPipe As String
        tradeName = Trim$(CStr(cfg.Cells(r, colTrade).Value))
        csiRaw = cfg.Cells(r, colCSI).Value
        shortName = Trim$(CStr(cfg.Cells(r, colShort).Value))
        createFlag = cfg.Cells(r, colCreate).Value
        defaultsPipe = Trim$(CStr(cfg.Cells(r, colDefaults).Value))

        If Not IsTruthy(createFlag) Then GoTo NextRow
        If Len(Trim$(CStr(csiRaw))) = 0 Or Len(shortName) = 0 Then GoTo NextRow

        Dim csi As String: csi = NormalizeCSI(csiRaw)
        Dim safeShort As String: safeShort = UCase$(SanitizeShortName(shortName))
        Dim desired As String: desired = Left$(csi & " " & safeShort, 31)

        Dim wsT As Worksheet
        If SheetExists(wb, desired) Then
            Set wsT = wb.Worksheets(desired)
            WriteIdentity wsT, tradeName, csi, jobTitle, estimator, bidDate, jobGsf, tradeSf
            WriteEvalLabels wsT
            If overwriteDefaults = vbYes And Len(defaultsPipe) > 0 Then
                ClearSeededScope wsT
                SeedDefaultScope wsT, defaultsPipe
            End If
            updated = updated + 1
        Else
            tpl.Copy After:=wb.Worksheets(wb.Worksheets.Count)
            Set wsT = wb.Worksheets(wb.Worksheets.Count)
            wsT.Name = desired
            WriteIdentity wsT, tradeName, csi, jobTitle, estimator, bidDate, jobGsf, tradeSf
            WriteEvalLabels wsT
            If Len(defaultsPipe) > 0 Then SeedDefaultScope wsT, defaultsPipe
            created = created + 1
        End If

NextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "GenerateTradeTabs_v10 complete." & vbCrLf & "Created: " & created & vbCrLf & "Updated: " & updated, vbInformation
End Sub

Private Sub WriteIdentity(ByVal ws As Worksheet, ByVal tradeName As String, ByVal csi As String, _
                          ByVal jobTitle As String, ByVal estimator As String, ByVal bidDate As String, _
                          ByVal jobGsf As String, ByVal tradeSf As String)
    ws.Range(CELL_TRADE_NAME).Value = tradeName
    ws.Range(CELL_CSI).Value = csi
    If Len(jobTitle) > 0 Then ws.Range(CELL_JOBTITLE).Value = jobTitle
    If Len(estimator) > 0 Then ws.Range(CELL_ESTIMATOR).Value = estimator
    If Len(bidDate) > 0 Then ws.Range(CELL_BIDDATE).Value = bidDate
    If Len(jobGsf) > 0 Then ws.Range(CELL_JOB_GSF).Value = jobGsf
    If Len(tradeSf) > 0 Then ws.Range(CELL_TRADE_SF).Value = tradeSf
End Sub

Private Sub WriteEvalLabels(ByVal ws As Worksheet)
    ' Write evaluation row labels in col G (HDR) — safe to call repeatedly
    ws.Cells(ROW_LOWEST_BIDDER, COL_HDR).Value = TXT_LBL_LOWEST
    ws.Cells(ROW_AVG_BASE, COL_HDR).Value = TXT_LBL_AVG_BASE
    ws.Cells(ROW_AVG_ADJ, COL_HDR).Value = TXT_LBL_AVG_ADJ
End Sub

Private Sub SeedDefaultScope(ByVal ws As Worksheet, ByVal defaultsPipe As String)
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then anchor = 17: ws.Range("B17").Value = KEY_ADD_SCOPE

    Dim items() As String: items = Split(defaultsPipe, "|")
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim s As String: s = Trim$(items(i))
        If Len(s) > 0 Then
            ws.Rows(anchor).Insert Shift:=xlDown
            ws.Cells(anchor, COL_LINE_NO).Value = "SEED"
            ws.Cells(anchor, COL_DESC).Value = s
            ws.Cells(anchor, COL_ADJ_FLAG).Value = "NO"
            anchor = anchor + 1
        End If
    Next i
End Sub

Private Sub ClearSeededScope(ByVal ws As Worksheet)
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then Exit Sub
    Dim r As Long: r = anchor - 1
    Do While r >= 1
        If UCase$(Trim$(CStr(ws.Cells(r, COL_LINE_NO).Value))) = "SEED" Then
            ws.Rows(r).Delete
            anchor = anchor - 1
            r = anchor - 1
        Else
            Exit Do
        End If
    Loop
End Sub

' =========================
' TOGGLE ALL CREATE CHECKBOXES
' =========================
' Assigns to a single button on Config_CSI.
' If all rows are checked (TRUE) → unchecks all.
' If any row is unchecked → checks all.
Public Sub ToggleAllCreate_v10()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim cfg As Worksheet
    On Error Resume Next
    Set cfg = wb.Worksheets(SHEET_CONFIG)
    On Error GoTo 0
    If cfg Is Nothing Then MsgBox "Missing Config_CSI sheet.", vbExclamation: Exit Sub

    Dim colCreate As Long: colCreate = FindHeaderCol(cfg, "Create")
    If colCreate = 0 Then MsgBox "Can't find 'Create' column in Config_CSI.", vbExclamation: Exit Sub

    ' Anchor lastRow to TradeName column — more reliable than Create column with checkboxes
    Dim colTrade As Long: colTrade = FindHeaderCol(cfg, "TradeName")
    Dim lastRow As Long
    If colTrade > 0 Then
        lastRow = cfg.Cells(cfg.Rows.Count, colTrade).End(xlUp).Row
    Else
        lastRow = cfg.Cells(cfg.Rows.Count, colCreate).End(xlUp).Row
    End If
    If lastRow < 2 Then MsgBox "No data rows found in Config_CSI.", vbInformation: Exit Sub

    ' Use .Value2 — avoids Type Mismatch from native Excel 365 checkbox type coercion
    Dim allChecked As Boolean: allChecked = True
    Dim r As Long
    For r = 2 To lastRow
        If Not IsTruthy(cfg.Cells(r, colCreate).Value2) Then
            allChecked = False
            Exit For
        End If
    Next r

    ' Toggle: if all checked → uncheck all; otherwise → check all
    Dim newValue As Boolean: newValue = Not allChecked
    For r = 2 To lastRow
        cfg.Cells(r, colCreate).Value2 = newValue
    Next r
End Sub
