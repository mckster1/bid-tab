Attribute VB_Name = "TSC_v9_1_Utils"
Option Explicit

Public Function SheetExists(ByVal wb As Workbook, ByVal name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, name, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

Public Function NormalizeCSI(ByVal v As Variant) As String
    Dim s As String: s = Trim$(Replace(CStr(v), " ", ""))
    s = Replace(s, "'", "")
    If Len(s) = 0 Then NormalizeCSI = "00.0000": Exit Function

    Dim a() As String
    If InStr(1, s, ".") > 0 Then
        a = Split(s, ".")
        Dim leftP As String, rightP As String
        leftP = Right$("00" & DigitsOnly(a(0)), 2)
        rightP = Left$(DigitsOnly(IIf(UBound(a) >= 1, a(1), "0")) & "0000", 4)
        NormalizeCSI = leftP & "." & rightP
    Else
        Dim d As String: d = DigitsOnly(s)
        d = Right$("00" & d, 2) & "0000"
        NormalizeCSI = Left$(d, 2) & "." & Mid$(d, 3, 4)
    End If
End Function

Public Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Public Function SanitizeShortName(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, " ", "")
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "'", "")
    s = Replace(s, Chr$(34), "")
    If Len(s) = 0 Then s = "TRADE"
    SanitizeShortName = s
End Function

Public Function IsTruthy(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTruthy = CBool(v)
    Else
        Dim s As String: s = UCase$(Trim$(CStr(v)))
        IsTruthy = (s = "TRUE" Or s = "YES" Or s = "1" Or s = "Y")
    End If
End Function

Public Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If UCase$(Trim$(CStr(ws.Cells(1, c).Value))) = UCase$(headerName) Then
            FindHeaderCol = c: Exit Function
        End If
    Next c
End Function

Public Function FindRowKeyInColB(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastUsed As Long: lastUsed = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row
    Dim r As Long, v As String
    For r = 1 To lastUsed
        v = UCase$(Trim$(CStr(ws.Cells(r, COL_DESC).Value)))
        If v = UCase$(Trim$(key)) Then FindRowKeyInColB = r: Exit Function
    Next r
End Function

Public Function ColLetter(ByVal col As Long) As String
    ColLetter = Split(Cells(1, col).Address(True, False), "$")(0)
End Function

Public Function NextBidderCol(ByVal ws As Worksheet) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_BIDDER_START Then lastCol = COL_BIDDER_START
    For c = COL_BIDDER_START To lastCol + 10
        If IsBidderColEmpty(ws, c) Then NextBidderCol = c: Exit Function
    Next c
    NextBidderCol = lastCol + 1
End Function

Public Function IsBidderColEmpty(ByVal ws As Worksheet, ByVal c As Long) As Boolean
    ' User request: check 2-8 AND 11 for "empty enough"
    Dim r As Long
    For r = 2 To 8
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then Exit Function
    Next r
    If Len(Trim$(CStr(ws.Cells(ROW_BASE_BID, c).Value))) > 0 Then Exit Function
    IsBidderColEmpty = True
End Function

Public Sub WriteAmountOrFormula(ByVal target As Range, ByVal s As String)
    s = Trim$(s)
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
