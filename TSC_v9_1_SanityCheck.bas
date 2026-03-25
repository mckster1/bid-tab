Attribute VB_Name = "TSC_v9_1_SanityCheck"
Option Explicit

Public Sub SanityCheck_v9_1()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim issues As Long, msg As String
    Dim ws As Worksheet, hidden As Long
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then hidden = hidden + 1
    Next ws
    If hidden > 0 Then issues = issues + 1: msg = msg & "- Hidden sheets detected (" & hidden & "). Remove/unhide to avoid generator bugs." & vbCrLf

    If Not SheetExists(wb, SHEET_CONFIG) Then issues = issues + 1: msg = msg & "- Missing Config_CSI." & vbCrLf
    If Not SheetExists(wb, SHEET_TEMPLATE) Then issues = issues + 1: msg = msg & "- Missing TradeTemplate." & vbCrLf

    If issues = 0 Then
        MsgBox "SanityCheck_v9_1: OK", vbInformation
    Else
        MsgBox "SanityCheck_v9_1 issues:" & vbCrLf & msg, vbExclamation
    End If
End Sub
