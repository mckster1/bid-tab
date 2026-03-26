Attribute VB_Name = "TSC_v10_SanityCheck"
Option Explicit

Public Sub SanityCheck_v10()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim issues As Long, msg As String
    Dim ws As Worksheet, hidden As Long

    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ' zz_tmpSort is a working sheet created/deleted by the sort macro — not a real issue
            If StrComp(ws.Name, "zz_tmpSort", vbTextCompare) <> 0 Then
                hidden = hidden + 1
            End If
        End If
    Next ws
    If hidden > 0 Then
        issues = issues + 1
        msg = msg & "- " & hidden & " unexpected hidden sheet(s) detected. " & _
              "These can cause tab generation bugs — unhide or delete them." & vbCrLf
    End If

    If Not SheetExists(wb, SHEET_CONFIG) Then
        issues = issues + 1
        msg = msg & "- Missing Config_CSI sheet." & vbCrLf
    End If
    If Not SheetExists(wb, SHEET_TEMPLATE) Then
        issues = issues + 1
        msg = msg & "- Missing TradeTemplate sheet." & vbCrLf
    End If

    If issues = 0 Then
        MsgBox "SanityCheck_v10: All OK", vbInformation
    Else
        MsgBox "SanityCheck_v10 found " & issues & " issue(s):" & vbCrLf & vbCrLf & msg, vbExclamation
    End If
End Sub
