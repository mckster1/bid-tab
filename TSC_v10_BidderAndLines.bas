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

    Dim newCol As Long: newCol = NextBidderCol(ws)

    ' Copy formatting from the first bidder column template (Column I)
    CopyBidderColFormat ws, COL_BIDDER_START, newCol

    ' Set Wizard Action dropdown default to INCLUDE
    ws.Cells(ROW_WIZ_ACTION, newCol).Value = BIDDER_INCLUDE
    ApplyIncludeExcludeValidation ws, newCol

    ' Prompt header fields
    ws.Cells(2, newCol).Value = Trim$(InputBox("Company name:", "Add Bidder"))
    ws.Cells(3, newCol).Value = Trim$(InputBox("Contact:", "Add Bidder"))
    ws.Cells(4, newCol).Value = Trim$(InputBox("Phone:", "Add Bidder"))
    ws.Cells(5, newCol).Value = Trim$(InputBox("Email:", "Add Bidder"))
    ws.Cells(6, newCol).Value = Trim$(InputBox("Date Received:", "Add Bidder"))
    ws.Cells(7, newCol).Value = Trim$(InputBox("Notes:", "Add Bidder"))

    ' Base bid supports expressions like 1000+200
    Dim bb As String
    bb = Trim$(InputBox("Base Bid (can be 1234 or 1000+200):", "Add Bidder"))
    If Len(bb) > 0 Then WriteAmountOrFormula ws.Cells(ROW_BASE_BID, newCol), bb

    ' Offer to re-enter scope for this new bidder
    If MsgBox("Re-enter Scope + Alternates for this bidder now?", vbYesNo + vbQuestion, "Add Bidder") = vbYes Then
        ReEnterScopeAndAlternates_ForCol ws, newCol
    End If

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

    ReEnterScopeAndAlternates_ForCol ws, col
    RefreshHighlights_v10
End Sub

' =========================
' RE-ENTER SCOPE + ALTS for a given bidder column
' =========================
Public Sub ReEnterScopeAndAlternates_ForCol(ByVal ws As Worksheet, ByVal bidderCol As Long)
    Dim scopeAnchor As Long: scopeAnchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If scopeAnchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub
    Dim scopeLast As Long: scopeLast = scopeAnchor - 1

    Dim altAnchor As Long: altAnchor = FindRowKeyInColB(ws, KEY_ADD_ALT)
    Dim altLast As Long
    If altAnchor > 0 Then altLast = altAnchor - 1 Else altLast = 0

    Dim bidName As String: bidName = Trim$(CStr(ws.Cells(2, bidderCol).Value))
    If Len(bidName) = 0 Then bidName = "Column " & ColLetter(bidderCol)

    Dim r As Long
    ' Scope prompts
    For r = ROW_SCOPE_START To scopeLast
        Dim desc As String: desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextScope
        If UCase$(desc) = UCase$(KEY_ADD_SCOPE) Then GoTo NextScope

        Dim resp As String
        resp = Trim$(InputBox(bidName & " — does this bidder include:" & vbCrLf & desc & vbCrLf & vbCrLf & _
                              "Type:" & vbCrLf & _
                              " I = Included" & vbCrLf & _
                              " E = Excluded" & vbCrLf & _
                              " $ = Enter dollar amount" & vbCrLf & _
                              " (blank) = Unconfirmed (skip)", "Scope Entry"))
        If StrPtr(resp) = 0 Then Exit Sub
        If resp = "" Then
            ws.Cells(r, bidderCol).Value = TXT_UNCONF
            ws.Cells(r, bidderCol).Interior.Color = RGB_LIGHT_YELLOW
        ElseIf UCase$(resp) = "I" Then
            ws.Cells(r, bidderCol).Value = TXT_INCLUDED
            ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
        ElseIf UCase$(resp) = "E" Then
            ws.Cells(r, bidderCol).Value = TXT_EXCLUDED
            ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
        Else
            Dim amt As String
            amt = Trim$(InputBox("Enter amount (can be 1234 or 1000+200):" & vbCrLf & desc, "Amount"))
            If StrPtr(amt) = 0 Then Exit Sub
            If Len(amt) = 0 Then
                ws.Cells(r, bidderCol).Value = TXT_UNCONF
                ws.Cells(r, bidderCol).Interior.Color = RGB_LIGHT_YELLOW
            Else
                WriteAmountOrFormula ws.Cells(r, bidderCol), amt
                ws.Cells(r, bidderCol).Interior.ColorIndex = xlNone
            End If
        End If
NextScope:
    Next r

    ' Alternates prompts
    If altLast > 0 Then
        For r = scopeAnchor + 3 To altLast
            Dim aDesc As String: aDesc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
            If UCase$(Left$(aDesc, 9)) <> "ALTERNATE" Then GoTo NextAlt

            Dim aAmt As String
            aAmt = Trim$(InputBox("Alternate amount for:" & vbCrLf & aDesc & vbCrLf & "(blank = Unconfirmed)", "Alternate Entry"))
            If StrPtr(aAmt) = 0 Then Exit Sub
            If aAmt = "" Then
                ws.Cells(r, bidderCol).Value = TXT_UNCONF
            Else
                WriteAmountOrFormula ws.Cells(r, bidderCol), aAmt
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

    Dim normalPrompt As VbMsgBoxResult
    normalPrompt = MsgBox("Enter this new line as a normal Scope Entry?" & vbCrLf & _
                          "YES = normal scope" & vbCrLf & _
                          "NO  = Exception/Exclusion (light red)", vbYesNoCancel + vbQuestion, "Scope Type")
    If normalPrompt = vbCancel Then Exit Sub

    ws.Rows(anchor).Insert Shift:=xlDown
    ws.Cells(anchor, COL_DESC).Value = IIf(normalPrompt = vbYes, desc, "Exception: " & desc)
    ws.Cells(anchor, COL_ADJ_FLAG).Value = "NO"

    If normalPrompt = vbNo Then
        ws.Range(ws.Cells(anchor, 1), ws.Cells(anchor, COL_NOTES)).Interior.Color = RGB_LIGHT_RED
    End If

    ' Prompt each existing bidder — show company name, not just column letter
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            Dim bidName As String: bidName = Trim$(CStr(ws.Cells(2, c).Value))
            If Len(bidName) = 0 Then bidName = "Column " & ColLetter(c)

            Dim resp As String
            resp = Trim$(InputBox("For " & bidName & ":" & vbCrLf & _
                                  "Enter $ amount OR type I / E OR leave blank to skip (marks Unconfirmed).", "Scope Entry"))
            If StrPtr(resp) = 0 Then Exit For
            If resp = "" Then
                ws.Cells(anchor, c).Value = TXT_UNCONF
                ws.Cells(anchor, c).Interior.Color = RGB_LIGHT_YELLOW
            ElseIf UCase$(resp) = "I" Then
                ws.Cells(anchor, c).Value = TXT_INCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            ElseIf UCase$(resp) = "E" Then
                ws.Cells(anchor, c).Value = TXT_EXCLUDED
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
            Else
                WriteAmountOrFormula ws.Cells(anchor, c), resp
                ws.Cells(anchor, c).Interior.ColorIndex = xlNone
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

    ' Write Alt+Adjusted formula for each bidder column
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = COL_BUDGET To lastCol
        If c = COL_BUDGET Or Not IsBidderColEmpty(ws, c) Then
            ws.Cells(anchor + 1, c).Formula = "=IFERROR(" & ws.Cells(ROW_ADJ_BASE, c).Address(False, False) & "+" & ws.Cells(anchor, c).Address(False, False) & ","""")"
        End If
    Next c

    RefreshHighlights_v10
End Sub

' =========================
' MOVE EXCLUSIONS TO BOTTOM
' =========================
' Reorders the scope section so all exception/exclusion lines (light red rows)
' sink to the bottom, just above the ADD SCOPE LINE anchor.
' Normal scope lines keep their relative order.
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

    ' Collect normal and exception row indices separately
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

    ' Check if already sorted (all exceptions already at the bottom)
    Dim alreadySorted As Boolean: alreadySorted = True
    If nNorm > 0 And nExc > 0 Then
        If normRows(nNorm) > excRows(1) Then alreadySorted = False
    End If
    If alreadySorted Then MsgBox "Exception lines are already at the bottom.", vbInformation: Exit Sub

    ' Use a temp sheet to rebuild the section without row-shift corruption
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

    ' Copy normal rows to temp sheet first, then exception rows
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

    ' Clear the scope section on the trade tab
    ws.Rows(scopeFirst & ":" & scopeLast).Clear

    ' Copy back in sorted order
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
    ' Primary check: light red background set by AddScopeLine
    If ws.Cells(r, COL_LINE_NO).Interior.Color = RGB_LIGHT_RED Then
        IsExceptionScopeLine = True: Exit Function
    End If
    ' Fallback: description prefix
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
