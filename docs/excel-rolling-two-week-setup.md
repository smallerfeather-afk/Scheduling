# Excel Setup: Master Two-Week Template + Rolling Two-Week Calendar

This setup gives you:

- A reusable **master template** for 14 days of staffing logic.
- A **rolling calendar** where each row is tied to a real date.
- Automatic mapping so the correct master day (Day 1..14) is used for each calendar date.

---

## 1) Build the `MasterTemplate` sheet

Create this structure:

- `A1`: `TemplateDay`
- `B1`: `Shift`
- `C1`: `Role`
- `D1`: `Employee`
- `E1`: `StartTime`
- `F1`: `EndTime`
- `G1`: `Notes`

In `A2:A15`, enter numbers `1` through `14`.

For each template day, enter as many staffing rows as you need. Example:

- Day 1 has 4 rows (RN, CNA, Clerk, Charge)
- Day 2 has 5 rows
- etc.

You can keep this as a flat table where each row is one shift assignment and `TemplateDay` identifies which of the 14-day slots it belongs to.

---

## 2) Build the `RollingCalendar` sheet

Use this layout:

- `A1`: `AnchorStartDate`
- `B1`: Enter your start date for the current two-week window (example: `6/3/2026`)

Headers in row 3:

- `A3`: `Date`
- `B3`: `TemplateDay`
- `C3`: `Shift`
- `D3`: `Role`
- `E3`: `Employee`
- `F3`: `StartTime`
- `G3`: `EndTime`
- `H3`: `Notes`

### Date sequence for two weeks

In `A4`, put:

```excel
=$B$1
```

In `A5`, put and copy down through `A17`:

```excel
=A4+1
```

This creates 14 consecutive dates.

### Template day mapping (date -> 1..14)

In `B4`, put and copy down through `B17`:

```excel
=MOD(A4-$B$1,14)+1
```

This ensures each date maps to the correct master template day.

---

## 3) Pull staffing rows from `MasterTemplate`

If each day has exactly one row, use simple lookups.

If each day has multiple staffing rows (most common), use **Power Query** (recommended) or VBA.

### Recommended no-code approach (Power Query)

1. Convert both ranges to Tables:
   - On `MasterTemplate`, create table `tblMaster`.
   - On `RollingCalendar`, create table `tblDates` using only `Date` + `TemplateDay` rows (`A3:B17`).
2. Data -> Get Data -> From Table/Range for each table.
3. In Power Query:
   - Merge `tblDates` with `tblMaster` on `TemplateDay`.
   - Expand merged columns: `Shift`, `Role`, `Employee`, `StartTime`, `EndTime`, `Notes`.
4. Load result to a new sheet (`RollingOutput`).
5. When you change `AnchorStartDate` or staffing assignments, click **Refresh All**.

This generates all rows for each date while preserving multiple shifts per day.

---

## 4) Optional formula-only approach (Excel 365 dynamic arrays)

If you have Microsoft 365, use this in `RollingCalendar!C4` to return all matching rows for each date/day:

```excel
=LET(
  d,$B4,
  t,$B4,
  src,MasterTemplate!$A$2:$G$1000,
  rows,FILTER(src,INDEX(src,,1)=t,""),
  HSTACK(MAKEARRAY(ROWS(rows),1,LAMBDA(r,c,d)),DROP(rows,,1))
)
```

Notes:

- This spills per date/day block.
- You may prefer Power Query for easier maintenance.

---

## 5) Implement `RefreshRollingSchedule.bas` in Excel (step-by-step)

Use this when you want a one-click way to rebuild a dated schedule output sheet from your 14-day template.

1. Open your Excel workbook (`.xlsm` format is required for macros).
2. Press `Alt + F11` to open the VBA editor.
3. In the Project pane, right-click your workbook -> **Import File...**
4. Select `excel/RefreshRollingSchedule.bas`.
5. Confirm your sheet names match exactly:
   - `MasterTemplate`
   - `RollingCalendar`
6. Confirm your layouts match the macro assumptions:
   - `MasterTemplate` columns `A:G` are `TemplateDay, Shift, Role, Employee, StartTime, EndTime, Notes`
   - `RollingCalendar` has dates in `A4:A17` and template day formula in `B4:B17`
7. Close the VBA editor and save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**.

### Run the macro

- In Excel: **Developer -> Macros -> `RefreshRollingScheduleOutput` -> Run**.
- The macro creates (or replaces contents of) a `RollingOutput` sheet.

### Optional: add a button

1. Developer -> Insert -> **Button (Form Control)**.
2. Draw the button on `RollingCalendar`.
3. Assign macro: `RefreshRollingScheduleOutput`.
4. Rename the button label to `Refresh Rolling Schedule`.

### What happens when it runs

- Clears old data in `RollingOutput`.
- Writes headers in row 1.
- Loops each rolling date (`A4:A17`) and finds all matching `TemplateDay` rows from `MasterTemplate`.
- Writes one dated row per matching shift/role row.

### Troubleshooting

- **Macro not visible**: save as `.xlsm`, then reopen workbook.
- **Compile or runtime error**: verify sheet names are exact (`MasterTemplate`, `RollingCalendar`).
- **No output rows**: check `TemplateDay` values are numeric 1..14 in both sheets.
- **Date issues**: ensure `RollingCalendar!A4:A17` are true Excel dates, not text.

---

## 6) Rolling forward week to week

To roll the calendar forward:

1. Change `RollingCalendar!B1` to the new window start date.
2. Dates and day mapping update automatically.
3. Refresh Power Query output (or run VBA macro if using VBA).

---

## 7) Validation checks

Use these checks to confirm alignment:

- `B4` should always equal `1` when `A4 = AnchorStartDate`.
- `B17` should always equal `14`.
- If you set anchor + 14 days later, mapping should restart at day 1.

Validation formula example:

```excel
=MOD((A4-$B$1),14)+1
```

---

## 8) Common pitfalls

- Date cells stored as text instead of real dates.
- Forgetting absolute refs (`$B$1`) in formulas.
- Mismatched day numbering (master starts at 0 instead of 1).
- Not refreshing Power Query after template edits.

---

## 9) Suggested workbook naming convention

- Template workbook: `Staffing_Master_Template.xlsx`
- Live rolling workbook: `Staffing_Rolling_Calendar.xlsx`

If both live in the same workbook, use sheet protection to lock formula columns and prevent accidental edits.
