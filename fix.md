# Production Dashboard — Issues & Fixes Log

---

## Excel File Structure (Expected)

**Sheet1** — Summary data:
- Column A: Dates
- Column B–D: Status breakdowns (Assigned, Available, Held)
- Column E: Inflow counts (Grand Total)

**Sheet2 (report name)** — Detailed task data:
- Columns A–Y: Full task details
- Column F: Remaining Days (critical for due calculations)
- Status, Team, Stage, Process, Assigned To, Held Reason
- Available Date (for inflow tracking)
- Assigned Date (for held article aging)

**Due Categories based on Remaining Days:**
- Overdue: Remaining Days < 0
- Due Today: Remaining Days = 0
- Due Tomorrow: Remaining Days = 1
- Due in 2+ Days: Remaining Days ≥ 2

---

## Issues Solved

### Bug 1 — Column Mapping Keys Mismatch (Critical)
**Problem:** `column_mapping` dict used keys with spaces (e.g. `'Remaining Days'`, `'Available Date'`,
`'Assigned Date'`, `'Task Status'`, `'Held Reason'`, `'Assigned To'`), but all downstream code
checked for underscore versions (`'Remaining_Days'`, `'Available_Date'`, etc.).
This caused ALL due calculations, inflow tracking, and held analysis to silently produce zeros.

**Fix:** Renamed all `column_mapping` keys to use underscores matching downstream checks:
`'Remaining_Days'`, `'Available_Date'`, `'Assigned_Date'`, `'Task_Status'`,
`'Held_Reason'`, `'Assigned_To'`, `'Actual_Date'`.

---

### Bug 2 — Datetime fillna Corruption
**Problem:** The fillna loop treated `datetime64[ns]` columns the same as `object` columns,
filling `NaT` with the string `'Unknown'`. This converted datetime columns to `object` dtype,
breaking all subsequent `.dt` accessor operations (`.dt.days`, `.dt.normalize()`, `.dt.strftime()`).

**Fix:** Skip datetime columns in the fillna loop — keep `NaT` intact.

---

### Bug 3 — Hardcoded Today's Date in New Inflow Calculation
**Problem:** `new_inflow_today` compared `Available_Date` against the hardcoded
`datetime(2026, 2, 23)` instead of the actual current date.

**Fix:** Replaced with `pd.Timestamp.today().normalize()` for dynamic date comparison.

---

### Bug 4 — Available_Date Display Crash on NaT
**Problem:** `.dt.strftime('%Y-%m-%d')` on a datetime column containing `NaT` values
produced literal `'NaT'` strings in the display table.

**Fix:** Added `.replace({'NaT': ''})` with a try/except fallback for safe formatting.

---


### UI Improvements
**Problem:** Multiple UI areas had poor contrast, no visual differentiation, and flat styling:
- KPI cards had a plain grey background with no colour cues
- Stage-wise due columns were plain markdown text with no visual grouping
- User cards were flat light-blue boxes
- Sidebar was plain white
- Page background was default Streamlit grey-white

**Fix:** Full CSS overhaul:
- Page background → soft blue-grey `#F0F4F8`
- Sidebar → deep navy gradient with white text
- KPI cards → white cards with colour-coded top borders per category
- Stage-wise columns → coloured panel cards (red / yellow / orange / green)
- Stage rows → clean flex rows with dividers and bold counts
- User cards → white card with navy left border and subtle shadow
 
 expected outcomes : 
  The Excel file should contain two sheets:

Sheet1: Summary data with:

    Dates in column A
    Inflow counts (Grand Total) in column E
    Status breakdowns in columns B-D

report name: Detailed task data with:

    Task details in columns A through Y
    Remaining Days in column F (critical for due calculations)
    Status information
    Team assignments
    Available dates for inflow tracking
    Assigned dates for held article aging

Due Categories based on Remaining Days:

    Overdue: Remaining Days < 0
    Due Today: Remaining Days = 0
    Due Tomorrow: Remaining Days = 1
    Due in 2+ Days: Remaining Days ≥ 2
  it have issues 
  