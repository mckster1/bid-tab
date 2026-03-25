Attribute VB_Name = "TSC_v9_ReEnterScope"
Option Explicit

Private Const BTN_ROW As Long = 11
Private Const BTN_HEIGHT As Double = 18

Public Sub ReEnterScopeAndAlternates_ActiveColumn_v9()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim bidderCol As Long: bidderCol = ActiveCell.Column
    If bidderCol < TSC_COL_BIDDER_START Then
        MsgBox "Select a cell in a bidder column (H or later).", vbExclamation
        Exit Sub
    End If
    RunWizard ws, bidderCol
End Sub

Public Sub ReEnterScopeAndAlternates_FromButton_v9()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim caller As String: caller = CStr(Application.Caller)
    Dim col As Long: col = ParseCol(caller)
    If col = 0 Then
        MsgBox "Can't determine bidder column from button: " & caller, vbExclamation
        Exit Sub
    End If
    RunWizard ws, col
End Sub

Public Sub EnsureReEnterButton_v9(ByVal ws As Worksheet, ByVal bidderCol As Long)
    Dim nm As String: nm = "btnReEnter_col" & bidderCol
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                ws.Cells(BTN_ROW, bidderCol).Left, _
                                ws.Cells(BTN_ROW, bidderCol).Top, _
                                ws.Cells(BTN_ROW, bidderCol).Width, _
                                BTN_HEIGHT)
    With shp
        .Name = nm
        .TextFrame2.TextRange.Text = "Re-Enter Scope"
        .OnAction = "ReEnterScopeAndAlternates_FromButton_v9"
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub RunWizard(ByVal ws As Worksheet, ByVal bidderCol As Long)
    Dim scopeLast As Long: scopeLast = FindLastRowBeforeKey(ws, TSC_KEY_ADD_SCOPE, 17)
    If scopeLast <= 0 Then
        MsgBox "Missing '" & TSC_KEY_ADD_SCOPE & "' anchor on this trade tab.", vbExclamation
        Exit Sub
    End If
    Dim altLast As Long: altLast = FindLastRowBeforeKey(ws, TSC_KEY_ADD_ALT, 22)

    Dim r As Long
    For r = 1 To scopeLast
        Dim desc As String: desc = Trim$(CStr(ws.Cells(r, TSC_COL_DESC).Value))
        If Len(desc) = 0 Then GoTo NextScope
        If UCase$(desc) = UCase$(TSC_KEY_ADD_SCOPE) Then GoTo NextScope
        PromptAndWrite ws, r, bidderCol, desc, False
NextScope:
    Next r

    If altLast > 0 Then
        For r = scopeLast + 1 To altLast
            Dim ad As String: ad = Trim$(CStr(ws.Cells(r, TSC_COL_DESC).Value))
            If UCase$(Left$(ad, 9)) = "ALTERNATE" Then
                PromptAndWrite ws, r, bidderCol, ad, True
            End If
        Next r
    End If

    MsgBox "Re-Enter complete for column " & TSC_ColLetter(bidderCol) & ".", vbInformation
End Sub

Private Sub PromptAndWrite(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal bidderCol As Long, _
                           ByVal desc As String, ByVal preferDollar As Boolean)
    Dim cur As String: cur = Trim$(CStr(ws.Cells(rowNum, bidderCol).Value))

    Dim resp As String
    resp = InputBox("Bidder: " & TSC_ColLetter(bidderCol) & vbCrLf & _
                    "Item: " & desc & vbCrLf & vbCrLf & _
                    "Enter:" & vbCrLf & _
                    " I = Included" & vbCrLf & _
                    " E = Excluded" & vbCrLf & _
                    " $ = Dollar amount (or type 1234 / 123+45)" & vbCrLf & _
                    " (blank) = Unconfirmed" & vbCrLf & vbCrLf & _
                    "Current: " & cur, _
                    IIf(preferDollar, "Alternates (prefer $)", "Scope"))
    If StrPtr(resp) = 0 Then Exit Sub
    resp = Trim$(resp)

    If resp = "" Then
        ws.Cells(rowNum, bidderCol).Value = TSC_TEXT_UNCONF
    ElseIf UCase$(resp) = "I" Then
        ws.Cells(rowNum, bidderCol).Value = TSC_TEXT_INCLUDED
    ElseIf UCase$(resp) = "E" Then
        ws.Cells(rowNum, bidderCol).Value = TSC_TEXT_EXCLUDED
    ElseIf UCase$(resp) = "$" Then
        Dim amt As String
        amt = InputBox("Enter amount for:" & vbCrLf & desc, "Dollar Amount")
        If StrPtr(amt) = 0 Then Exit Sub
        WriteAmount ws.Cells(rowNum, bidderCol), Trim$(amt)
    Else
        WriteAmount ws.Cells(rowNum, bidderCol), resp
    End If
End Sub

Private Sub WriteAmount(ByVal target As Range, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    If InStr(s, "+") > 0 Or InStr(s, "-") > 0 Or InStr(s, "*") > 0 Or InStr(s, "/") > 0 Then
        If Left$(s, 1) <> "=" Then s = "=" & s
        target.Formula = s
    ElseIf IsNumeric(s) Then
        target.Value = CDbl(s)
    Else
        target.Value = s
    End If
End Sub

Private Function FindLastRowBeforeKey(ByVal ws As Worksheet, ByVal key As String, ByVal fallbackAnchor As Long) As Long
    Dim anchor As Long: anchor = TSC_FindRowByExactKeyInColB(ws, key)
    If anchor = 0 Then anchor = fallbackAnchor
    FindLastRowBeforeKey = anchor - 1
End Function

Private Function ParseCol(ByVal shpName As String) As Long
    Dim p As Long: p = InStr(1, shpName, "col", vbTextCompare)
    If p = 0 Then Exit Function
    Dim n As String: n = Mid$(shpName, p + 3)
    If IsNumeric(n) Then ParseCol = CLng(n)
End Function
