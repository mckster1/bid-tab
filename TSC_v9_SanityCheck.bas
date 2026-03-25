Attribute VB_Name = "TSC_v9_SanityCheck"
Option Explicit

Public Sub SanityCheck_v9()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim msg As String, issues As Long

    Dim ws As Worksheet, countTpl As Long
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, TSC_TEMPLATE_SHEET, vbTextCompare) = 0 Then countTpl = countTpl + 1
        If ws.Visible <> xlSheetVisible Then
            issues = issues + 1
            msg = msg & "- Hidden sheet found: " & ws.Name & " (remove or make visible)" & vbCrLf
        End If
    Next ws
    If countTpl <> 1 Then
        issues = issues + 1
        msg = msg & "- Expected exactly 1 '" & TSC_TEMPLATE_SHEET & "'. Found: " & countTpl & vbCrLf
    End If

    Dim cfg As Worksheet
    On Error Resume Next
    Set cfg = wb.Worksheets(TSC_CONFIG_SHEET)
    On Error GoTo 0
    If cfg Is Nothing Then
        issues = issues + 1
        msg = msg & "- Missing '" & TSC_CONFIG_SHEET & "'." & vbCrLf
    Else
        Dim need: need = Array("TradeName","CSI_Code","ShortName","Create","DefaultScopeLines")
        Dim i As Long
        For i = 0 To UBound(need)
            If TSC_FindHeaderCol(cfg, CStr(need(i))) = 0 Then
                issues = issues + 1
                msg = msg & "- Config_CSI missing header: " & need(i) & vbCrLf
            End If
        Next i
    End If

    If issues = 0 Then
        MsgBox "SanityCheck_v9: OK", vbInformation
    Else
        MsgBox "SanityCheck_v9 issues:" & vbCrLf & msg, vbExclamation
    End If
End Sub
