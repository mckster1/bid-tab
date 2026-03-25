Attribute VB_Name = "TSC_v9_1_GenerateTradeTabs"
Option Explicit

Public Sub GenerateTradeTabs_v9_1()
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
            If Len(defaultsPipe) > 0 Then SeedDefaultScope wsT, defaultsPipe
            created = created + 1
        End If

NextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "GenerateTradeTabs_v9_1 complete." & vbCrLf & "Created: " & created & vbCrLf & "Updated: " & updated, vbInformation
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

Private Sub SeedDefaultScope(ByVal ws As Worksheet, ByVal defaultsPipe As String)
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then anchor = 17: ws.Range("B17").Value = KEY_ADD_SCOPE

    Dim items() As String: items = Split(defaultsPipe, "|")
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim s As String: s = Trim$(items(i))
        If Len(s) > 0 Then
            ws.Rows(anchor).Insert Shift:=xlDown
            ws.Cells(anchor, COL_LINE_NO).Value = ""  ' leave numbering to user
            ws.Cells(anchor, COL_DESC).Value = s
            ws.Cells(anchor, 1).Value = "SEED"
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
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "SEED" Then
            ws.Rows(r).Delete
            anchor = anchor - 1
            r = anchor - 1
        Else
            Exit Do
        End If
    Loop
End Sub
