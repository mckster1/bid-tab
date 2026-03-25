Attribute VB_Name = "TSC_v9_1_BidderAndLines"
Option Explicit

' =========================
' ADD BIDDER
' =========================
Public Sub AddBidder_v9_1()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then
        MsgBox "Use Add Bidder on a trade tab (not TradeTemplate).", vbExclamation
        Exit Sub
    End If

    Dim newCol As Long: newCol = NextBidderCol(ws)

    ' Copy formatting from the FIRST bidder column template (Column I) if available
    CopyBidderColFormat ws, COL_BIDDER_START, newCol

    ' Set Wizard Action dropdown default INCLUDE
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

    ' Ask to re-enter scope after base bid
    Dim goScope As VbMsgBoxResult
    goScope = MsgBox("Re-enter Scope + Alternates for this bidder now?", vbYesNo + vbQuestion, "Add Bidder")
    If goScope = vbYes Then
        ReEnterScopeAndAlternates_ForCol ws, newCol
    End If
End Sub

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

' =========================
' RE-ENTER SCOPE + ALTS for a given bidder column (called by Add Bidder)
' =========================
Public Sub ReEnterScopeAndAlternates_ForCol(ByVal ws As Worksheet, ByVal bidderCol As Long)
    Dim scopeAnchor As Long: scopeAnchor = FindRowKeyInColB(ws, KEY_ADD_SCOPE)
    If scopeAnchor = 0 Then MsgBox "Can't find '" & KEY_ADD_SCOPE & "' in col B.", vbExclamation: Exit Sub
    Dim scopeLast As Long: scopeLast = scopeAnchor - 1

    Dim altAnchor As Long: altAnchor = FindRowKeyInColB(ws, KEY_ADD_ALT)
    Dim altLast As Long
    If altAnchor > 0 Then altLast = altAnchor - 1 Else altLast = 0

    Dim r As Long
    ' Scope prompts (default: Included/Excluded, then $)
    For r = 17 To scopeLast
        Dim desc As String: desc = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextScope
        If UCase$(desc) = UCase$(KEY_ADD_SCOPE) Then GoTo NextScope

        Dim resp As String
        resp = Trim$(InputBox("Does this bidder include:" & vbCrLf & desc & vbCrLf & vbCrLf & _
                              "Type:" & vbCrLf & _
                              " I = Included" & vbCrLf & _
                              " E = Excluded" & vbCrLf & _
                              " $ = Enter dollar amount" & vbCrLf & _
                              " (blank) = Unconfirmed (skip)", "Scope Entry"))
        If StrPtr(resp) = 0 Then Exit Sub
        If resp = "" Then
            ws.Cells(r, bidderCol).Value = TXT_UNCONF
        ElseIf UCase$(resp) = "I" Then
            ws.Cells(r, bidderCol).Value = TXT_INCLUDED
        ElseIf UCase$(resp) = "E" Then
            ws.Cells(r, bidderCol).Value = TXT_EXCLUDED
        Else
            ' $ or numeric/expression
            Dim amt As String
            amt = Trim$(InputBox("Enter amount (can be 1234 or 1000+200):" & vbCrLf & desc, "Amount"))
            If StrPtr(amt) = 0 Then Exit Sub
            If Len(amt) = 0 Then
                ws.Cells(r, bidderCol).Value = TXT_UNCONF
            Else
                WriteAmountOrFormula ws.Cells(r, bidderCol), amt
            End If
        End If
NextScope:
    Next r

    ' Alternates prompts (default $)
    If altLast > 0 Then
        For r = scopeAnchor + 3 To altLast ' skip header row and "Alt + Adjusted" rows are formula-driven
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
' ADD SCOPE LINE (normal or exception)
' =========================
Public Sub AddScopeLine_v9_1()
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

    ' Prompt each existing bidder column (skippable -> UNCONFIRMED)
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            Dim resp As String
            resp = Trim$(InputBox("For bidder column " & ColLetter(c) & ":" & vbCrLf & _
                                  "Enter $ amount OR type I/E OR leave blank to skip.", "Scope Entry"))
            If StrPtr(resp) = 0 Then Exit For
            If resp = "" Then
                ws.Cells(anchor, c).Value = TXT_UNCONF
            ElseIf UCase$(resp) = "I" Then
                ws.Cells(anchor, c).Value = TXT_INCLUDED
            ElseIf UCase$(resp) = "E" Then
                ws.Cells(anchor, c).Value = TXT_EXCLUDED
            Else
                WriteAmountOrFormula ws.Cells(anchor, c), resp
            End If
        End If
    Next c
End Sub

' =========================
' ADD ALTERNATE (adds 2 rows: Alternate n and Alt n + Adjusted)
' =========================
Public Sub AddAlternate_v9_1()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim anchor As Long: anchor = FindRowKeyInColB(ws, KEY_ADD_ALT)
    If anchor = 0 Then MsgBox "Can't find '" & KEY_ADD_ALT & "' in col B.", vbExclamation: Exit Sub

    Dim n As Long: n = NextAlternateNumber(ws, anchor)
    ws.Rows(anchor).Insert Shift:=xlDown
    ws.Rows(anchor).Insert Shift:=xlDown

    ws.Cells(anchor, COL_LINE_NO).Value = n
    ws.Cells(anchor, COL_DESC).Value = "Alternate " & n & ":"
    ws.Cells(anchor + 1, COL_DESC).Value = "Alternate " & n & " + Adjusted Base Bid"

    ' For each bidder col: Alt+Adjusted formula = AdjBase + Alt
    Dim lastCol As Long: lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = COL_BUDGET To lastCol
        ' Only apply where column is budget or an existing bidder
        If c = COL_BUDGET Or Not IsBidderColEmpty(ws, c) Then
            ws.Cells(anchor + 1, c).Formula = "=IFERROR(" & ws.Cells(ROW_ADJ_BASE, c).Address(False, False) & "+" & ws.Cells(anchor, c).Address(False, False) & ","""")"
        End If
    Next c
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
