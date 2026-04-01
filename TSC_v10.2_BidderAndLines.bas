Attribute VB_Name = "TSC_v10_BidderAndLines"
Option Explicit

' =========================
' ADD BIDDER
' =========================
Public Sub AddBidder_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then
        MsgBox "Use Add Bidder on a trade tab (not TradeTemplate).", vbExclamation
        Exit Sub
    End If

    ' Collect all fields first — Cancel on any field aborts with no changes written
    Dim company As String: company = InputBox("Company name:", "Add Bidder")
    If StrPtr(company) = 0 Then Exit Sub

    Dim contact As String: contact = InputBox("Contact name:", "Add Bidder")
    If StrPtr(contact) = 0 Then Exit Sub

    Dim phone As String: phone = InputBox("Phone:", "Add Bidder")
    If StrPtr(phone) = 0 Then Exit Sub

    Dim email As String: email = InputBox("Email:", "Add Bidder")
    If StrPtr(email) = 0 Then Exit Sub

    Dim dateRcvd As String: dateRcvd = InputBox("Date Received:", "Add Bidder")
    If StrPtr(dateRcvd) = 0 Then Exit Sub

    Dim notes As String: notes = InputBox("Notes:", "Add Bidder")
    If StrPtr(notes) = 0 Then Exit Sub

    Dim baseBid As String: baseBid = InputBox("Base Bid (number or expression like 100000+5000):", "Add Bidder")
    If StrPtr(baseBid) = 0 Then Exit Sub

    Dim doScope As Boolean
    doScope = (MsgBox("Re-enter scope + alternates for this bidder now?", _
                      vbYesNo + vbQuestion, "Add Bidder") = vbYes)

    ' All input collected — now write to sheet
    Dim newCol As Long: newCol = NextBidderCol(ws)

    CopyBidderColFormat ws, COL_BIDDER_START, newCol

    ws.Cells(ROW_WIZ_ACTION, newCol).Value = BIDDER_INCLUDE
    ApplyIncludeExcludeValidation ws, newCol

    ws.Cells(2, newCol).Value = company
    ws.Cells(3, newCol).Value = contact
    ws.Cells(4, newCol).Value = phone
    ws.Cells(5, newCol).Value = email
    ws.Cells(6, newCol).Value = dateRcvd
    ws.Cells(7, newCol).Value = notes

    If Len(Trim$(baseBid)) > 0 Then WriteAmountOrFormula ws.Cells(ROW_BASE_BID, newCol), baseBid

    If doScope Then ReEnterScopeAndAlternates_ForCol ws, newCol, True

    RefreshHighlights_v10
End Sub

' =========================
' RE-ENTER SCOPE (button-assignable — uses active cell column)
' =========================
Public Sub ReEnterScope_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then
        MsgBox "Run on a trade tab (not TradeTemplate).", vbExclamation
        Exit Sub
    End If

    Dim col As Long: col = ActiveCell.Column
    If col < COL_BIDDER_START Then
        MsgBox "Select any cell in a bidder column first, then run this macro.", vbExclamation
        Exit Sub
    End If
    If IsBidderColEmpty(ws, col) Then
        MsgBox "No bidder data found in this column.", vbExclamation
        Exit Sub
    End If

    Dim bidName As String: bidName = Trim$(CStr(ws.Cells(2, col).Value))
    If Len(bidName) = 0 Then bidName = "Column " & ColLetter(col)

    If MsgBox("Re-enter scope + alternates for:" & vbCrLf & bidName & "?", _
              vbYesNo + vbQuestion, "Re-enter Scope") = vbNo Then Exit Sub

    ReEnterScopeAndAlternates_ForCol ws, col, False
    RefreshHighlights_v10
End Sub

' =========================
' RE-ENTER SCOPE + ALTS for a given bidder column
' blankIsUnconf = True  → blank response writes UNCONFIRMED (AddBidder, AddAlternate contexts)
' blankIsUnconf = False → blank response keeps the existing cell value (ReEnterScope context)
' =========================
Public Sub ReEnterScopeAndAlternates_ForCol(ByVal ws As Worksheet, ByVal bidderCol As Long, ByVal blankIsUnconf As Boolean)
    Dim scopeAnchor As Long: scopeAnchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If scopeAnchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub
    Dim scopeLast As Long: scopeLast = scopeAnchor - 1

    Dim altAnchor As Long: altAnchor = FindRowKeyInColB(ws, KEY_ADD_ALT)

    Dim bidName As String: bidName = Trim$(CStr(ws.Cells(2, bidderCol).Value))
    If Len(bidName) = 0 Then bidName = "Column " & ColLetter(bidderCol)

    Dim blankLabel As String
    blankLabel = IIf(blankIsUnconf, "Unconfirmed / Skip", "Keep existing")

    ' Scope prompts — Cancel exits immediately
    Dim r As Long, resp As String, respRaw As String, respUC As String
    For r = ROW_SCOPE_START To scopeLast
        Dim desc As String: desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextScope
        If UCase$(desc) = UCase$(KEY_ADD_SCOPE) Then GoTo NextScope

        resp = InputBox(bidName & vbCrLf & vbCrLf & desc & vbCrLf & vbCrLf & _
                        "[number] = Dollar amount" & vbCrLf & _
                        "blank    = " & blankLabel & vbCrLf & _
                        "I        = Included" & vbCrLf & _
                        "E        = Excluded" & vbCrLf & _
                        "(Cancel to stop)", "Scope Entry")
        If StrPtr(resp) = 0 Then Exit Sub

        respRaw = Trim$(resp)
        respUC = UCase$(respRaw)

        If respUC = "I" Then
            ws.Cells(r, bidderCol).Value = TXT_INCLUDED
            ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
        ElseIf respUC = "E" Then
            ws.Cells(r, bidderCol).Value = TXT_EXCLUDED
            ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
        ElseIf Len(respRaw) > 0 Then
            WriteAmountOrFormula ws.Cells(r, bidderCol), respRaw
            ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
        ElseIf blankIsUnconf Then
            ws.Cells(r, bidderCol).Value = TXT_UNCONF
            ws.Cells(r, bidderCol).Interior.Color = RGB_LIGHT_YELLOW
        End If
        ' else blankIsUnconf=False: keep existing (do nothing)
NextScope:
    Next r

    ' Alternates — auto-formula for "+Adjusted Base Bid" rows; prompt for "Alternate N:" rows
    If altAnchor > 0 Then
        Dim altDataLast As Long
        altDataLast = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row

        Dim aDesc As String, aAmt As String, aAmtRaw As String, aAmtUC As String
        Dim adjRef As String, altRef As String
        For r = altAnchor + 1 To altDataLast
            aDesc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
            If Len(aDesc) = 0 Then GoTo NextAlt

            ' Auto-write formula for "+Adjusted Base Bid" rows — no prompt
            If InStr(UCase$(aDesc), "ADJUSTED BASE BID") > 0 Then
                adjRef = ws.Cells(ROW_ADJ_BASE, bidderCol).Address(False, False)
                altRef = ws.Cells(r - 1, bidderCol).Address(False, False)
                ws.Cells(r, bidderCol).Formula = "=IFERROR(" & adjRef & "+IF(ISNUMBER(" & altRef & ")," & altRef & ",0),"""")"
                GoTo NextAlt
            End If

            ' Only prompt for "Alternate N:" rows
            If UCase$(Left$(aDesc, 9)) <> "ALTERNATE" Then GoTo NextAlt

            aAmt = InputBox(bidName & " — " & aDesc & vbCrLf & vbCrLf & _
                            "[number] = Dollar amount" & vbCrLf & _
                            "blank    = " & blankLabel & vbCrLf & _
                            "E        = Excluded" & vbCrLf & _
                            "I        = Included" & vbCrLf & _
                            "(Cancel to stop)", "Alternate Entry")
            If StrPtr(aAmt) = 0 Then Exit Sub

            aAmtRaw = Trim$(aAmt)
            aAmtUC = UCase$(aAmtRaw)

            If aAmtUC = "E" Then
                ws.Cells(r, bidderCol).Value = TXT_EXCLUDED
                ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
            ElseIf aAmtUC = "I" Then
                ws.Cells(r, bidderCol).Value = TXT_INCLUDED
                ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
            ElseIf Len(aAmtRaw) > 0 Then
                WriteAmountOrFormula ws.Cells(r, bidderCol), aAmtRaw
                ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
            ElseIf blankIsUnconf Then
                ws.Cells(r, bidderCol).Value = TXT_UNCONF
                ws.Cells(r, bidderCol).Interior.Color = RGB_LIGHT_YELLOW
            End If
NextAlt:
        Next r
    End If
End Sub

' =========================
' ADD SCOPE LINE
' =========================
Public Sub AddScopeLine_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub

    Dim desc As String
    desc = Trim$(InputBox("New Scope Description:", "Add Scope Line"))
    If StrPtr(desc) = 0 Or Len(desc) = 0 Then Exit Sub

    Dim typeChoice As VbMsgBoxResult
    typeChoice = MsgBox("Add """ & desc & """ as an exception / exclusion line (red row)?" & vbCrLf & vbCrLf & _
                        "Yes = Exception / Exclusion (red row)" & vbCrLf & _
                        "No  = Normal scope line" & vbCrLf & _
                        "Cancel = Don't add this line", _
                        vbYesNoCancel + vbQuestion, "Scope Line Type")
    If typeChoice = vbCancel Then Exit Sub

    Dim isException As Boolean: isException = (typeChoice = vbYes)

    ws.Rows(anchor).Insert Shift:=xlDown
    ws.Cells(anchor, COL_DESC).Value = IIf(isException, "Exception: " & desc, desc)
    ws.Cells(anchor, COL_ADJ_FLAG).Value = "NO"

    If isException Then
        ws.Range(ws.Cells(anchor, 1), ws.Cells(anchor, COL_NOTES)).Interior.Color = RGB_LIGHT_RED
    End If

    ' Prompt each existing bidder — Cancel stops the loop (row already added)
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, resp As String, respRaw As String, respUC As String, bidName As String
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            bidName = Trim$(CStr(ws.Cells(2, c).Value))
            If Len(bidName) = 0 Then bidName = "Column " & ColLetter(c)

            resp = InputBox(bidName & vbCrLf & vbCrLf & desc & vbCrLf & vbCrLf & _
                            "[number] = Dollar amount" & vbCrLf & _
                            "blank    = Unconfirmed / Skip" & vbCrLf & _
                            "I        = Included" & vbCrLf & _
                            "E        = Excluded" & vbCrLf & _
                            "(Cancel to stop entering)", "Scope Entry")
            If StrPtr(resp) = 0 Then Exit For

            respRaw = Trim$(resp)
            respUC = UCase$(respRaw)

            If respUC = "I" Then
                ws.Cells(anchor, c).Value = TXT_INCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            ElseIf respUC = "E" Then
                ws.Cells(anchor, c).Value = TXT_EXCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            ElseIf Len(respRaw) > 0 Then
                WriteAmountOrFormula ws.Cells(anchor, c), respRaw
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            Else
                ws.Cells(anchor, c).Value = TXT_UNCONF
                ws.Cells(anchor, c).Interior.Color = RGB_LIGHT_YELLOW
            End If
        End If
    Next c

    RefreshHighlights_v10
End Sub

' =========================
' ADD ALTERNATE
' =========================
Public Sub AddAlternate_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_ALT)
    If anchor = 0 Then MsgBox "Can't find '" & KEY_ADD_ALT & "' in col B.", vbExclamation: Exit Sub

    Dim n As Long: n = NextAlternateNumber(ws, anchor)
    ws.Rows(anchor).Insert Shift:=xlDown
    ws.Rows(anchor).Insert Shift:=xlDown

    ws.Cells(anchor, COL_LINE_NO).Value = n
    ws.Cells(anchor, COL_DESC).Value = "Alternate " & n & ":"
    ws.Cells(anchor + 1, COL_DESC).Value = "Alternate " & n & " + Adjusted Base Bid"

    ' Write Alt+Adjusted formula for budget col and all bidder cols
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, adjRef As String, altRef As String
    For c = COL_BUDGET To lastCol
        If c = COL_BUDGET Or Not IsBidderColEmpty(ws, c) Then
            adjRef = ws.Cells(ROW_ADJ_BASE, c).Address(False, False)
            altRef = ws.Cells(anchor, c).Address(False, False)
            ws.Cells(anchor + 1, c).Formula = "=IFERROR(" & adjRef & "+IF(ISNUMBER(" & altRef & ")," & altRef & ",0),"""")"
        End If
    Next c

    ' Prompt each bidder for their alternate amount
    Dim altDesc As String: altDesc = CStr(ws.Cells(anchor, COL_DESC).Value)
    Dim bidName As String, aAmt As String, aAmtRaw As String, aAmtUC As String
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            bidName = Trim$(CStr(ws.Cells(2, c).Value))
            If Len(bidName) = 0 Then bidName = "Column " & ColLetter(c)

            aAmt = InputBox(bidName & " — " & altDesc & vbCrLf & vbCrLf & _
                            "[number] = Dollar amount" & vbCrLf & _
                            "blank    = Unconfirmed / Skip" & vbCrLf & _
                            "E        = Excluded" & vbCrLf & _
                            "I        = Included" & vbCrLf & _
                            "(Cancel to stop)", "Alternate Entry")
            If StrPtr(aAmt) = 0 Then Exit For

            aAmtRaw = Trim$(aAmt)
            aAmtUC = UCase$(aAmtRaw)

            If aAmtUC = "E" Then
                ws.Cells(anchor, c).Value = TXT_EXCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            ElseIf aAmtUC = "I" Then
                ws.Cells(anchor, c).Value = TXT_INCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            ElseIf Len(aAmtRaw) > 0 Then
                WriteAmountOrFormula ws.Cells(anchor, c), aAmtRaw
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            Else
                ws.Cells(anchor, c).Value = TXT_UNCONF
                ws.Cells(anchor, c).Interior.Color = RGB_LIGHT_YELLOW
            End If
        End If
    Next c

    RefreshHighlights_v10
End Sub

' =========================
' MOVE EXCLUSIONS TO BOTTOM
' =========================
Public Sub MoveExclusionsToBottom_v10()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then
        MsgBox "Run on a trade tab (not TradeTemplate).", vbExclamation
        Exit Sub
    End If

    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If anchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub

    Dim scopeFirst As Long: scopeFirst = ROW_SCOPE_START
    Dim scopeLast As Long: scopeLast = anchor - 1
    If scopeLast < scopeFirst Then MsgBox "No scope lines found.", vbInformation: Exit Sub

    Dim normRows() As Long, excRows() As Long
    Dim nNorm As Long, nExc As Long
    Dim r As Long
    For r = scopeFirst To scopeLast
        If Len(Trim$(CStr(ws.Cells(r, COL_DESC).Value))) = 0 Then GoTo NextScopeRow
        If IsExceptionScopeLine(ws, r) Then
            nExc = nExc + 1
            ReDim Preserve excRows(1 To nExc)
            excRows(nExc) = r
        Else
            nNorm = nNorm + 1
            ReDim Preserve normRows(1 To nNorm)
            normRows(nNorm) = r
        End If
NextScopeRow:
    Next r

    If nExc = 0 Then MsgBox "No exception lines found — nothing to move.", vbInformation: Exit Sub
    If nNorm = 0 Then MsgBox "All scope lines are exceptions — nothing to reorder.", vbInformation: Exit Sub

    Dim alreadySorted As Boolean: alreadySorted = True
    If nNorm > 0 And nExc > 0 Then
        If normRows(nNorm) > excRows(1) Then alreadySorted = False
    End If
    If alreadySorted Then MsgBox "Exception lines are already at the bottom.", vbInformation: Exit Sub

    Dim wb As Workbook: Set wb = ws.Parent
    Dim tmp As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("zz_tmpScope").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set tmp = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    tmp.Name = "zz_tmpScope"
    tmp.Visible = xlSheetVeryHidden

    Dim destRow As Long: destRow = 1
    Dim i As Long
    For i = 1 To nNorm
        ws.Rows(normRows(i)).Copy
        tmp.Cells(destRow, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        destRow = destRow + 1
    Next i
    For i = 1 To nExc
        ws.Rows(excRows(i)).Copy
        tmp.Cells(destRow, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        destRow = destRow + 1
    Next i

    ws.Rows(scopeFirst & ":" & scopeLast).Clear

    Dim totalRows As Long: totalRows = nNorm + nExc
    For i = 1 To totalRows
        tmp.Rows(i).Copy
        ws.Cells(scopeFirst + i - 1, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    Next i

    Application.DisplayAlerts = False
    tmp.Delete
    Application.DisplayAlerts = True

    RefreshHighlights_v10

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox nExc & " exception line(s) moved to bottom of scope section.", vbInformation
End Sub

Private Function IsExceptionScopeLine(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If ws.Cells(r, COL_LINE_NO).Interior.Color = RGB_LIGHT_RED Then
        IsExceptionScopeLine = True: Exit Function
    End If
    Dim desc As String: desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
    If Left$(UCase$(desc), 10) = "EXCEPTION:" Then IsExceptionScopeLine = True
End Function

' =========================
' PRIVATE HELPERS
' =========================
Private Sub CopyBidderColFormat(ByVal ws As Worksheet, ByVal srcCol As Long, ByVal dstCol As Long)
    If dstCol = srcCol Then Exit Sub
    ws.Columns(srcCol).Copy
    ws.Columns(dstCol).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub ApplyIncludeExcludeValidation(ByVal ws As Worksheet, ByVal col As Long)
    On Error Resume Next
    With ws.Cells(ROW_WIZ_ACTION, col).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=BIDDER_INCLUDE & "," & BIDDER_EXCLUDE
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Function NextAlternateNumber(ByVal ws As Worksheet, ByVal altAnchor As Long) As Long
    Dim r As Long, maxN As Long
    For r = 1 To altAnchor - 1
        Dim v As Variant: v = ws.Cells(r, COL_LINE_NO).Value
        If IsNumeric(v) Then
            If CLng(v) > maxN Then maxN = CLng(v)
        End If
    Next r
    NextAlternateNumber = maxN + 1
End Function
