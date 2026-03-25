Attribute VB_Name = "TSC_v9_1_SortBidders"
Option Explicit

'===========================
' PUBLIC MACROS (button-bind)
'===========================
' These TWO procedures are what you should see in Assign Macro:
'  - SortBidders_ByBaseBid_v9_1
'  - SortBidders_ByAdjustedBid_v9_1

Private Const HUGE As Double = 9.9E+307

Public Sub SortBidders_ByBaseBid_v9_1()
    SortBidders_Internal True
End Sub

Public Sub SortBidders_ByAdjustedBid_v9_1()
    SortBidders_Internal False
End Sub

'===========================
' INTERNAL IMPLEMENTATION
'===========================
Private Sub SortBidders_Internal(ByVal sortOnBaseBid As Boolean)
    Dim ws As Worksheet: Set ws = ActiveSheet

    If StrComp(ws.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then
        MsgBox "Run sort on a trade tab, not TradeTemplate.", vbExclamation
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = ws.Cells(ROW_WIZ_ACTION, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_BIDDER_START Then
        MsgBox "No bidder columns found to sort.", vbInformation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_DESC).End(xlUp).Row
    If lastRow < 50 Then lastRow = 200

    Dim cols() As Long, scores() As Double, n As Long, c As Long
    ReDim cols(1 To lastCol - COL_BIDDER_START + 1)
    ReDim scores(1 To lastCol - COL_BIDDER_START + 1)

    For c = COL_BIDDER_START To lastCol
        If Not IsBidderColEmpty(ws, c) Then
            n = n + 1
            cols(n) = c
            scores(n) = ScoreBidder(ws, c, sortOnBaseBid)
        End If
    Next c

    If n <= 1 Then
        MsgBox "Nothing to sort (need 2+ bidders).", vbInformation
        Exit Sub
    End If

    ' Sort by score ascending
    Dim i As Long, j As Long, tmpC As Long, tmpS As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If scores(j) < scores(i) Then
                tmpS = scores(i): scores(i) = scores(j): scores(j) = tmpS
                tmpC = cols(i): cols(i) = cols(j): cols(j) = tmpC
            End If
        Next j
    Next i

    ' Temp sheet approach avoids Excel column-shift chaos
    Dim wb As Workbook: Set wb = ws.Parent
    Dim tmp As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("zz_tmpSort").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set tmp = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    tmp.Name = "zz_tmpSort"
    tmp.Visible = xlSheetVeryHidden

    For i = 1 To n
        ws.Range(ws.Cells(1, cols(i)), ws.Cells(lastRow, cols(i))).Copy
        tmp.Cells(1, i).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    Next i

    For i = 1 To n
        ws.Range(ws.Cells(1, cols(i)), ws.Cells(lastRow, cols(i))).Clear
    Next i

    Dim destCol As Long: destCol = COL_BIDDER_START
    For i = 1 To n
        tmp.Range(tmp.Cells(1, i), tmp.Cells(lastRow, i)).Copy
        ws.Cells(1, destCol).PasteSpecial xlPasteAll
        Application.CutCopyMode = False

        EnsureWizardActionValidation ws, destCol
        destCol = destCol + 1
    Next i

    Application.DisplayAlerts = False
    tmp.Delete
    Application.DisplayAlerts = True

    MsgBox IIf(sortOnBaseBid, "Sorted by Base Bid.", "Sorted by Adjusted Base Bid."), vbInformation
End Sub

Private Function ScoreBidder(ByVal ws As Worksheet, ByVal col As Long, ByVal sortOnBaseBid As Boolean) As Double
    ' Excluded bidders go last
    Dim action As String
    action = UCase$(Trim$(CStr(ws.Cells(ROW_WIZ_ACTION, col).Value)))
    If action = BIDDER_EXCLUDE Then
        ScoreBidder = HUGE
        Exit Function
    End If

    Dim v As Variant
    If sortOnBaseBid Then
        v = ws.Cells(ROW_BASE_BID, col).Value
    Else
        v = ws.Cells(ROW_ADJ_BASE, col).Value
    End If

    If IsNumeric(v) Then
        ScoreBidder = CDbl(v)
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        ScoreBidder = HUGE
    Else
        On Error Resume Next
        ScoreBidder = CDbl(v)
        If Err.Number <> 0 Then
            Err.Clear
            ScoreBidder = HUGE
        End If
        On Error GoTo 0
    End If
End Function

Private Sub EnsureWizardActionValidation(ByVal ws As Worksheet, ByVal col As Long)
    On Error Resume Next
    With ws.Cells(ROW_WIZ_ACTION, col).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:=BIDDER_INCLUDE & "," & BIDDER_EXCLUDE
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

    If Len(Trim$(CStr(ws.Cells(ROW_WIZ_ACTION, col).Value))) = 0 Then
        ws.Cells(ROW_WIZ_ACTION, col).Value = BIDDER_INCLUDE
    End If
End Sub
