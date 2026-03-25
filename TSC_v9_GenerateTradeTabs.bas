Attribute VB_Name = "TSC_v9_GenerateTradeTabs"
Option Explicit

Public Sub GenerateTradeTabs_v9()
    Dim wb As Workbook: Set wb = ThisWorkbook

    If wb.ProtectStructure Then
        MsgBox "Workbook STRUCTURE is protected. Unprotect first:" & vbCrLf & _
               "Review -> Protect Workbook -> Unprotect (Structure).", vbExclamation
        Exit Sub
    End If

    Dim cfg As Worksheet, tpl As Worksheet
    On Error Resume Next
    Set cfg = wb.Worksheets(TSC_CONFIG_SHEET)
    Set tpl = wb.Worksheets(TSC_TEMPLATE_SHEET)
    On Error GoTo 0
    If cfg Is Nothing Then MsgBox "Missing sheet: " & TSC_CONFIG_SHEET, vbExclamation: Exit Sub
    If tpl Is Nothing Then MsgBox "Missing sheet: " & TSC_TEMPLATE_SHEET, vbExclamation: Exit Sub

    If UCase$(Trim$(CStr(tpl.Range("B17").Value))) <> TSC_KEY_ADD_SCOPE Then tpl.Range("B17").Value = TSC_KEY_ADD_SCOPE
    If UCase$(Trim$(CStr(tpl.Range("B22").Value))) <> TSC_KEY_ADD_ALT Then tpl.Range("B22").Value = TSC_KEY_ADD_ALT

    Dim colTrade As Long, colCSI As Long, colShort As Long, colCreate As Long, colDefaults As Long
    colTrade = TSC_FindHeaderCol(cfg, "TradeName")
    colCSI = TSC_FindHeaderCol(cfg, "CSI_Code")
    colShort = TSC_FindHeaderCol(cfg, "ShortName")
    colCreate = TSC_FindHeaderCol(cfg, "Create")
    colDefaults = TSC_FindHeaderCol(cfg, "DefaultScopeLines")

    If colTrade * colCSI * colShort * colCreate = 0 Then
        MsgBox "Config_CSI missing headers: TradeName, CSI_Code, ShortName, Create, DefaultScopeLines", vbExclamation
        Exit Sub
    End If
    If colDefaults = 0 Then colDefaults = colCreate + 1

    Dim overwriteDefaults As VbMsgBoxResult
    overwriteDefaults = MsgBox("Overwrite default scope lines on existing tabs?" & vbCrLf & _
                              "Yes = clear seeded defaults + re-seed" & vbCrLf & _
                              "No = seed only on NEW tabs", vbYesNoCancel + vbQuestion, "Default Scope Overwrite")
    If overwriteDefaults = vbCancel Then Exit Sub

    Dim jobTitle As String, estimator As String, bidDate As String, gsf As String
    jobTitle = Trim$(InputBox("Job Title (C4). Leave blank to skip.", "Setup – Job Info"))
    estimator = Trim$(InputBox("Estimator (C5). Leave blank to skip.", "Setup – Job Info"))
    bidDate = Trim$(InputBox("Bid Date (C6). Leave blank to skip.", "Setup – Job Info"))
    gsf = Trim$(InputBox("Job GSF (C7). Leave blank to skip.", "Setup – Job Info"))

    Dim lastRow As Long
    lastRow = cfg.Cells(cfg.Rows.Count, colTrade).End(xlUp).Row
    If lastRow < 2 Then MsgBox "Config_CSI has no trade rows.", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim created As Long, updated As Long, r As Long
    For r = 2 To lastRow
        Dim tradeName As String, csiRaw As Variant, shortName As String, createFlag As Variant, defaults As String
        tradeName = Trim$(CStr(cfg.Cells(r, colTrade).Value))
        csiRaw = cfg.Cells(r, colCSI).Value
        shortName = Trim$(CStr(cfg.Cells(r, colShort).Value))
        createFlag = cfg.Cells(r, colCreate).Value
        defaults = Trim$(CStr(cfg.Cells(r, colDefaults).Value))

        If Len(Trim$(CStr(csiRaw))) = 0 Or Len(shortName) = 0 Then GoTo NextRow
        If Not TSC_IsTruthy(createFlag) Then GoTo NextRow

        Dim csi As String: csi = TSC_NormalizeCSI(csiRaw)
        Dim safeShort As String: safeShort = UCase$(TSC_SanitizeShortName(shortName))
        Dim desired As String: desired = Left$(csi & " " & safeShort, 31)

        Dim wsT As Worksheet
        If TSC_SheetExists(wb, desired) Then
            Set wsT = wb.Worksheets(desired)
            WriteTradeIdentity wsT, tradeName, csi, jobTitle, estimator, bidDate, gsf
            If overwriteDefaults = vbYes And Len(defaults) > 0 Then
                ClearSeededScope wsT
                SeedDefaultScope wsT, defaults
            End If
            updated = updated + 1
        Else
            tpl.Copy After:=wb.Worksheets(wb.Worksheets.Count)
            Set wsT = wb.Worksheets(wb.Worksheets.Count)
            wsT.Name = desired
            WriteTradeIdentity wsT, tradeName, csi, jobTitle, estimator, bidDate, gsf
            If Len(defaults) > 0 Then SeedDefaultScope wsT, defaults
            created = created + 1
        End If

NextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "GenerateTradeTabs_v9 complete." & vbCrLf & _
           "Created: " & created & vbCrLf & _
           "Updated: " & updated, vbInformation
End Sub

Private Sub WriteTradeIdentity(ByVal ws As Worksheet, ByVal tradeName As String, ByVal csi As String, _
                               ByVal jobTitle As String, ByVal estimator As String, ByVal bidDate As String, ByVal gsf As String)
    ws.Range(TSC_CELL_TRADENAME).Value = tradeName
    ws.Range(TSC_CELL_CSI).Value = csi
    If Len(jobTitle) > 0 Then ws.Range(TSC_CELL_JOBTITLE).Value = jobTitle
    If Len(estimator) > 0 Then ws.Range(TSC_CELL_ESTIMATOR).Value = estimator
    If Len(bidDate) > 0 Then ws.Range(TSC_CELL_BIDDATE).Value = bidDate
    If Len(gsf) > 0 Then ws.Range(TSC_CELL_GSF).Value = gsf
End Sub

Private Sub SeedDefaultScope(ByVal ws As Worksheet, ByVal defaultsPipe As String)
    Dim anchor As Long: anchor = TSC_FindRowByExactKeyInColB(ws, TSC_KEY_ADD_SCOPE)
    If anchor = 0 Then anchor = 17: ws.Range("B17").Value = TSC_KEY_ADD_SCOPE

    Dim items() As String: items = Split(defaultsPipe, "|")
    Dim i As Long
    For i = LBound(items) To UBound(items)
        Dim s As String: s = Trim$(items(i))
        If Len(s) > 0 Then
            ws.Rows(anchor).Insert Shift:=xlDown
            ws.Cells(anchor, 2).Value = s
            ws.Cells(anchor, 1).Value = "SEED"
            anchor = anchor + 1
        End If
    Next i
End Sub

Private Sub ClearSeededScope(ByVal ws As Worksheet)
    Dim anchor As Long: anchor = TSC_FindRowByExactKeyInColB(ws, TSC_KEY_ADD_SCOPE)
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
