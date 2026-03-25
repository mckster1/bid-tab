# CLAUDE.md — Tab Sheet Creator Project Context

## Current Version: v10 (Stable)

The workbook is on a clean, stable VBA foundation. All v9.0 and v9.1 modules have been
removed. Seven v10 `.bas` modules are the source of truth. The user imports them manually
into the Excel VBA editor. The xlsm file lives locally on the user's device.

**Branch:** `claude/improve-bid-tab-project-Bo0Ya`

---

## What Is This Tool

An Excel VBA bid tabulation workbook for construction estimating. It is NOT a rigid
database — it is a structured visual decision tool designed to:
- Reduce estimator mistakes by surfacing scope omissions and anomalies
- Encourage apples-to-apples subcontractor comparisons
- Allow scope to evolve across multiple bids
- Remain human-first while being AI-agent compatible

---

## Repository File Structure

```
TabSheetCreator_v10.xlsm   ← lives locally on user's device (not in repo)
TSC_v10_Core.bas            ← constants (sheet names, row/col positions, colors)
TSC_v10_Utils.bas           ← shared helpers (CSI normalization, anchor finders, etc.)
TSC_v10_BidderAndLines.bas  ← scope, bidder, alternate, and exclusion macros
TSC_v10_GenerateTradeTabs.bas ← tab generation from Config_CSI + toggle all create
TSC_v10_SortBidders.bas     ← sort bidder columns by base or adjusted bid
TSC_v10_Evaluation.bas      ← highlights, eval summary, anomaly detection
TSC_v10_SanityCheck.bas     ← workbook integrity check
README.md                   ← install instructions, color key, debugging guide
CLAUDE.md                   ← this file
```

---

## Workbook Structure

**Fixed sheets:**
- `Dashboard` — reserved for future expansion
- `Config_CSI` — trade configuration table (headers: TradeName, CSI_Code, ShortName, Create, DefaultScopeLines)
- `TradeTemplate` — master template copied when generating trade tabs

**Generated trade tabs** — named `XX.XXXX SHORTNAME` (e.g., `03.3000 CONCRETE`)

**Temp sheets** (created and deleted automatically by macros):
- `zz_tmpSort` — used by sort macros
- `zz_tmpScope` — used by MoveExclusionsToBottom

---

## Template Layout (each trade tab)

| Rows | Col G label | Notes |
|---|---|---|
| 2–8 | Identity + Wizard Action | C2=Trade Name, C3=CSI, C4=Job Title, C5=Estimator, C6=Bid Date, C7=Job GSF, C8=Trade SF |
| 8 | Wizard Action | INCLUDE/EXCLUDE dropdown per bidder (col I+) |
| 11 | Base Bid | Numeric per bidder |
| 12 | Adjusted Base Bid | Numeric per bidder |
| 13 | $/SF | Numeric per bidder |
| 14 | Lowest Adj Bidder | Written by RefreshHighlights — company name in col H |
| 15 | Avg Base Bid | Written by RefreshHighlights — dollar average in col H |
| 16 | Avg Adj Base Bid | Written by RefreshHighlights — dollar average in col H |
| 17 | ADD SCOPE LINE | Anchor row — do not move or rename |
| 17→anchor-1 | Scope items | Inserted above anchor by AddScopeLine / seeded by GenerateTradeTabs |
| anchor+n | ADD ALTERNATE | Second anchor |
| anchor+n→ | Alternate rows | Pairs: "Alternate N:" + "Alternate N + Adjusted Base Bid" |

**Columns:**
- A=Line No, B=Description, C=Qty, D=Unit, E=Adj Flag, F=Notes, G=Labels, H=Budget, I+=Bidders

---

## Macro Inventory

### TSC_v10_BidderAndLines.bas
| Macro | Description |
|---|---|
| `AddBidder_v10` | Adds a new bidder column; prompts for company, contact, phone, email, date, notes, base bid; offers to re-enter scope |
| `ReEnterScope_v10` | Re-walks scope + alternates for the bidder column containing the active cell; detects column automatically |
| `AddScopeLine_v10` | Inserts a new scope line above the ADD SCOPE LINE anchor; prompts type (normal or exception/red); walks existing bidders by name for I/E/$/skip |
| `AddAlternate_v10` | Inserts two rows (Alternate N + Alternate N + Adjusted Base Bid) with auto-incrementing N; writes IFERROR formula for adjusted row |
| `MoveExclusionsToBottom_v10` | Moves all exception/exclusion lines (light red rows) to the bottom of the scope section; normal lines keep relative order; uses temp sheet pattern |

### TSC_v10_GenerateTradeTabs.bas
| Macro | Description |
|---|---|
| `GenerateTradeTabs_v10` | Reads Config_CSI, creates or updates trade tabs from template; prompts for job info; seeds default scope from pipe-delimited list |
| `ToggleAllCreate_v10` | Toggles all checkboxes in the Create column of Config_CSI — if all checked → uncheck all; otherwise → check all |

### TSC_v10_SortBidders.bas
| Macro | Description |
|---|---|
| `SortBidders_ByBaseBid_v10` | Sorts bidder columns left→right by Base Bid (lowest first); EXCLUDE bidders pushed right; calls RefreshHighlights after |
| `SortBidders_ByAdjustedBid_v10` | Same as above but sorts by Adjusted Base Bid |

### TSC_v10_Evaluation.bas
| Macro | Description |
|---|---|
| `RefreshHighlights_v10` | Writes eval summary (rows 14–16), highlights lowest adj bid column green, flags exclusion anomalies orange; runs on active sheet; called by sort macros and optionally auto-triggered by Workbook_SheetChange |

### TSC_v10_SanityCheck.bas
| Macro | Description |
|---|---|
| `SanityCheck_v10` | Checks for unexpected hidden sheets, missing Config_CSI, missing TradeTemplate; ignores zz_tmpSort |

---

## Color Key

| Color | Meaning |
|---|---|
| Light green (header rows 2–8) | Lowest adjusted base bid bidder |
| Light yellow (scope cell) | UNCONFIRMED — bidder not yet asked about this scope item |
| Orange (scope cell) | Exclusion anomaly — bidder excluded something ≥60% of others included |
| Light red (entire row) | Exception / Exclusion line |

---

## Completed Work (do not re-propose)

- [x] Migrated from v9.1 to v10 — all legacy modules removed
- [x] Evaluation summary rows 14–16 (Lowest Adj Bidder, Avg Base Bid, Avg Adj Base Bid)
- [x] Lowest bid column green highlight (header rows 2–8)
- [x] Anomaly highlighting: UNCONFIRMED=yellow, exclusion outlier=orange (≥60% threshold)
- [x] `ReEnterScope_v10` — active cell column detection (no picker needed)
- [x] Bidder names shown in scope entry prompts (not column letters)
- [x] Sort macros auto-call `RefreshHighlights_v10`
- [x] SanityCheck false-positive fix (zz_tmpSort no longer flagged)
- [x] `MoveExclusionsToBottom_v10` — temp-sheet-safe row reordering
- [x] `ToggleAllCreate_v10` — select/deselect all checkboxes on Config_CSI
- [x] README with install steps, color key, debugging guide
- [x] CLAUDE.md (this file)
- [x] Workbook_SheetChange snippet for auto-highlights (documented in README, one-time paste into ThisWorkbook)

---

## Prioritized Next Steps

These are known feature gaps, ordered by estimator value. None have been started.

### High priority
1. **Adjusted Base Bid formula / row** — Currently row 12 is manually entered. A macro or template formula to compute Adjusted Base Bid from Base Bid ± scope line adjustments (using the Adj Flag column E) would eliminate manual re-entry and make highlights more reliable.
2. **$/SF auto-formula** — Row 13 is blank unless manually entered. Formula: `=Adjusted Base Bid / Trade SF`. Needs Trade SF to be numeric.
3. **Per-bidder scope completion status** — A visual indicator (e.g., a count or color in the header block) showing how many scope lines are still UNCONFIRMED for each bidder.

### Medium priority
4. **Scope line numbering** — Auto-number scope lines in col A when lines are added/moved so the estimator can reference them in conversation with subs.
5. **Copy scope from one trade tab to another** — When a later bid adds scope items not in an earlier tab, a macro to propagate new lines (as UNCONFIRMED) across related tabs.
6. **Print / export layout** — Clean print range and formatting for sending bid tab summaries to project managers.

### Low priority / future
7. **Dashboard summary** — The Dashboard sheet is reserved; a macro to write a cross-trade summary (lowest adjusted bid per trade, total, etc.) would complete the package.
8. **Config_CSI validation** — Warn on duplicate CSI codes or missing ShortName before generating tabs.

---

## Key Design Constraints (never violate)

- This is NOT a rigid database. Scope can evolve. Don't enforce strict data entry.
- The tool is human-first. Highlight anomalies; do not block the estimator.
- Comparing to the MOST COMPLETE bid, not the first bid.
- `.bas` files are the source of truth. The xlsm is maintained locally by the user.
- All public macros must be button-assignable (no required arguments).
- Use the temp-sheet pattern (`zz_tmp*`) for any row/column reordering to avoid Excel shift corruption.
- `RefreshHighlights_v10` must remain safe to call from both a button and from `Workbook_SheetChange` (idempotent, EnableEvents-safe).
