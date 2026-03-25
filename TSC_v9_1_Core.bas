Attribute VB_Name = "TSC_v9_1_Core"
Option Explicit

' ===== Core sheet names =====
Public Const SHEET_TEMPLATE As String = "TradeTemplate"
Public Const SHEET_CONFIG As String = "Config_CSI"

' ===== Anchor keys (exact match in Col B) =====
Public Const KEY_ADD_SCOPE As String = "ADD SCOPE LINE"
Public Const KEY_ADD_ALT As String = "ADD ALTERNATE"

' ===== Template layout (v8-compatible) =====
' Identity cells (left block)
Public Const CELL_TRADE_NAME As String = "C2"
Public Const CELL_CSI As String = "C3"
Public Const CELL_JOBTITLE As String = "C4"
Public Const CELL_ESTIMATOR As String = "C5"
Public Const CELL_BIDDATE As String = "C6"
Public Const CELL_JOB_GSF As String = "C7"
Public Const CELL_TRADE_SF As String = "C8"

' Columns
Public Const COL_LINE_NO As Long = 1          ' A
Public Const COL_DESC As Long = 2             ' B
Public Const COL_QTY As Long = 3              ' C
Public Const COL_UNIT As Long = 4             ' D
Public Const COL_ADJ_FLAG As Long = 5         ' E  (Yes/No addback)
Public Const COL_NOTES As Long = 6            ' F
Public Const COL_HDR As Long = 7              ' G  (labels)
Public Const COL_BUDGET As Long = 8           ' H  (budget)
Public Const COL_BIDDER_START As Long = 9     ' I  (first bidder col)

' Key rows (labels live in col G)
Public Const ROW_BASE_BID As Long = 11        ' Base Bid value row
Public Const ROW_ADJ_BASE As Long = 12        ' Adjusted Base Bid value row
Public Const ROW_PSF As Long = 13             ' $/SF value row
Public Const ROW_WIZ_ACTION As Long = 8       ' "Wizard Action" row (Include/Exclude dropdown)

' Text markers
Public Const TXT_INCLUDED As String = "INCLUDED"
Public Const TXT_EXCLUDED As String = "EXCLUDED"
Public Const TXT_UNCONF As String = "UNCONFIRMED"

' Include/Exclude dropdown values
Public Const BIDDER_INCLUDE As String = "INCLUDE"
Public Const BIDDER_EXCLUDE As String = "EXCLUDE"

' Color helpers (RGB)
Public Const RGB_LIGHT_YELLOW As Long = 13434879  ' #FFF2CC-ish
Public Const RGB_LIGHT_RED As Long = 14474460     ' #F8CBAD-ish
Public Const RGB_GREY As Long = 15132390          ' #E6E6E6-ish
Public Const RGB_LIGHT_GREEN As Long = 13421823   ' #C6EFCE-ish
