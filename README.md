# Tab Sheet Creator — v10

## What's New in v10

| Feature | Details |
|---|---|
| **Evaluation summary** | Rows 14–16 on each trade tab show Lowest Adj Bidder, Avg Base Bid, and Avg Adj Base Bid (included bidders only) |
| **Lowest bid highlight** | The column with the lowest adjusted base bid is highlighted green (header rows 2–8) |
| **Anomaly highlighting** | UNCONFIRMED cells stay yellow; EXCLUDED cells turn orange when ≥60% of other bidders included that scope item |
| **Re-enter Scope button** | `ReEnterScope_v10` — select any cell in a bidder column and run; it detects the column automatically |
| **Bidder names in prompts** | Scope entry dialogs now show the company name instead of just the column letter |
| **Auto-highlights on sort** | Sort macros call `RefreshHighlights_v10` automatically after reordering |
| **SanityCheck fix** | No longer false-flags `zz_tmpSort` (the sort helper sheet) as a hidden sheet problem |

---

## Module Inventory

| File | Public Macros |
|---|---|
| `TSC_v10_Core.bas` | *(constants only)* |
| `TSC_v10_Utils.bas` | *(shared helpers)* |
| `TSC_v10_BidderAndLines.bas` | `AddBidder_v10`, `ReEnterScope_v10`, `AddScopeLine_v10`, `AddAlternate_v10`, `MoveExclusionsToBottom_v10` |
| `TSC_v10_GenerateTradeTabs.bas` | `GenerateTradeTabs_v10`, `ToggleAllCreate_v10` |
| `TSC_v10_SortBidders.bas` | `SortBidders_ByBaseBid_v10`, `SortBidders_ByAdjustedBid_v10` |
| `TSC_v10_Evaluation.bas` | `RefreshHighlights_v10` |
| `TSC_v10_SanityCheck.bas` | `SanityCheck_v10` |

---

## Installation (upgrade from v9.1)

### Step 1 — Open the VBA Editor
Press `ALT + F11`.

### Step 2 — Remove old modules
In the Project Explorer (left panel), expand your workbook.
Delete every module named `TSC_v9_*` or `TSC_v9_1_*`:
- Right-click each module → **Remove** → when prompted to export, click **No**.

### Step 3 — Import v10 modules
**File → Import File** — import all 7 `.bas` files:
```
TSC_v10_Core.bas
TSC_v10_Utils.bas
TSC_v10_BidderAndLines.bas
TSC_v10_GenerateTradeTabs.bas
TSC_v10_SortBidders.bas
TSC_v10_Evaluation.bas
TSC_v10_SanityCheck.bas
```

### Step 4 — Re-assign macro buttons
Right-click each button → **Assign Macro** → select the matching `_v10` version.

| Old macro | New macro |
|---|---|
| `GenerateTradeTabs_v9_1` | `GenerateTradeTabs_v10` |
| `AddBidder_v9_1` | `AddBidder_v10` |
| `AddScopeLine_v9_1` | `AddScopeLine_v10` |
| `AddAlternate_v9_1` | `AddAlternate_v10` |
| `SortBidders_ByBaseBid_v9_1` | `SortBidders_ByBaseBid_v10` |
| `SortBidders_ByAdjustedBid_v9_1` | `SortBidders_ByAdjustedBid_v10` |
| `SanityCheck_v9_1` | `SanityCheck_v10` |
| *(new — add a button)* | `ReEnterScope_v10` |
| *(new — add a button)* | `RefreshHighlights_v10` |

### Step 5 — Compile check
In the VBA editor: **Debug → Compile VBAProject**. Fix any errors before closing the editor.

### Step 6 — Set up auto-highlights (optional but recommended)

This makes highlights refresh automatically whenever you edit a cell on a trade tab.

In the VBA Project Explorer, double-click **ThisWorkbook**.
Add the following — paste it alongside any existing code, do **not** replace the whole module:

```vb
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If StrComp(Sh.Name, SHEET_TEMPLATE, vbTextCompare) = 0 Then Exit Sub
    If StrComp(Sh.Name, SHEET_CONFIG, vbTextCompare) = 0 Then Exit Sub
    If Left$(Sh.Name, 3) = "zz_" Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub
    On Error Resume Next
    Application.EnableEvents = False
    RefreshHighlights_v10
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
```

If you skip this step, assign `RefreshHighlights_v10` to a button and run it manually instead.

---

## Evaluation Row Layout (rows 14–16)

Written automatically by `RefreshHighlights_v10` and `GenerateTradeTabs_v10`.

| Row | Col G label | Col H value |
|---|---|---|
| 14 | Lowest Adj Bidder | Company name |
| 15 | Avg Base Bid | Dollar average |
| 16 | Avg Adj Base Bid | Dollar average |
| 17 | ADD SCOPE LINE | *(anchor — do not edit)* |

Only **INCLUDE** bidders with numeric values count toward the averages.

---

## Color Key

| Color | Meaning |
|---|---|
| Light green (header rows 2–8) | Lowest adjusted base bid bidder |
| Light yellow (scope cell) | UNCONFIRMED — bidder not yet asked about this scope item |
| Orange (scope cell) | Exclusion anomaly — this bidder excluded something ≥60% of others included |
| Light red (entire row) | Exception / Exclusion line |

---

## Debugging

### After every import
`ALT+F11` → **Debug → Compile VBAProject**

If anything is wrong (bad reference, typo, mismatched `End Sub`) this surfaces it immediately. Fix all errors before closing the editor.

### Stepping through a macro
1. In the VBA editor, click anywhere inside the sub you want to test
2. Press **F8** to execute one line at a time (yellow highlight shows current line)
3. Hover over any variable to see its current value in a tooltip
4. **View → Immediate Window** (`CTRL+G`) — type `?variableName` and Enter to inspect mid-run, or add `Debug.Print someValue` to the code for tracing

### If a macro errors mid-run
Click **Debug** in the error dialog. The editor opens at the failing line. Check variable values, then trace backward.

### Common symptoms

| Symptom | Likely cause |
|---|---|
| "Sub or Function not defined" | A button is still assigned to an old `_v9_1` macro name — re-assign to `_v10` |
| Highlights don't update on cell edit | `Workbook_SheetChange` snippet not yet added to `ThisWorkbook` |
| Eval rows 14–16 blank | No bidders have numeric values in the Adjusted Base Bid row yet |
| Orange anomaly never appears | Needs 2+ bidders where ≥60% include something one excludes |
| Green highlight on wrong column after sort | Row 12 (Adj Base Bid) has a non-numeric value for that bidder |
