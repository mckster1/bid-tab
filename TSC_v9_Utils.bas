Attribute VB_Name = "TSC_v9_Utils"
Option Explicit

Public Function TSC_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            TSC_SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Public Function TSC_NormalizeCSI(ByVal v As Variant) As String
    Dim s As String: s = CStr(v)
    s = Trim$(Replace(Replace(s, "'", ""), " ", ""))
    If Len(s) = 0 Then TSC_NormalizeCSI = "00.0000": Exit Function

    If InStr(1, s, ".") > 0 Then
        Dim parts() As String: parts = Split(s, ".")
        Dim leftP As String, rightP As String
        leftP = Right$("00" & TSC_DigitsOnly(parts(0)), 2)
        rightP = Left$(TSC_DigitsOnly(IIf(UBound(parts) >= 1, parts(1), "0")) & "0000", 4)
        TSC_NormalizeCSI = leftP & "." & rightP
    Else
        Dim d As String: d = TSC_DigitsOnly(s)
        d = Right$("00" & d, 2) & "0000"
        TSC_NormalizeCSI = Left$(d, 2) & "." & Mid$(d, 3, 4)
    End If
End Function

Public Function TSC_DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    TSC_DigitsOnly = out
End Function

Public Function TSC_SanitizeShortName(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "'", "")
    s = Replace(s, Chr$(34), "")
    s = Replace(s, " ", "")
    If Len(s) = 0 Then s = "TRADE"
    TSC_SanitizeShortName = s
End Function

Public Function TSC_IsTruthy(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        TSC_IsTruthy = CBool(v)
    Else
        Dim s As String: s = UCase$(Trim$(CStr(v)))
        TSC_IsTruthy = (s = "TRUE" Or s = "YES" Or s = "1" Or s = "Y")
    End If
End Function

Public Function TSC_FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If UCase$(Trim$(CStr(ws.Cells(1, c).Value))) = UCase$(headerName) Then
            TSC_FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

Public Function TSC_FindRowByExactKeyInColB(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim lastUsed As Long: lastUsed = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim r As Long, v As String, keyU As String
    keyU = UCase$(Trim$(key))
    For r = 1 To lastUsed
        v = UCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
        If v = keyU Then TSC_FindRowByExactKeyInColB = r: Exit Function
    Next r
End Function

Public Function TSC_ColLetter(ByVal col As Long) As String
    TSC_ColLetter = Split(Cells(1, col).Address(True, False), "$")(0)
End Function
