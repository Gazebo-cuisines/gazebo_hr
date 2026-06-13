# Weekly Report — How to use it

A short procedure for running the weekly payroll report on
[gazeboo.cloud](https://gazeboo.cloud).

## Weekly report vs Daily report

| | **Weekly report** | **Daily report** |
| --- | --- | --- |
| Use for | Extra hours + **additional holiday pay** (0.12 × extra) | Overtime only |
| Export file | `gazebo_weekly_report_*.xlsx` | `gazebo_daily_report_*.xlsx` |
| Key columns | Actual hours, Contracted hours, Extra hours, Additional Holiday pay | Overtime |
| Monthly report input | No — monthly uses **daily** exports | Yes |

Use **Weekly report** when HR needs Amit’s additional holiday pay calculation. Use **Daily report** for the existing overtime workflow and monthly roll-up.

## 1. Sign in

1. Open <https://gazeboo.cloud/login/>.
2. Enter your **Username** and **Password**.
   - Forgot it? Click **Forgot password?** → contact Utsav for a reset.
3. After signing in you land on the **Dashboard**.

## 2. Get the two input files from ClockRite

You need **two** Excel files exported from ClockRite. Both are
**Excel 7.0 (.xls)** — the same files as the daily report.

### File A — Employee hours file

1. Open **ClockRite**.
2. Top-left → **Print Report**.
3. From the dropdown, pick **Paid Hour (Incl. Absence) Summary**.
4. **Export** → **Excel 7.0 format (.xls)**.

### File B — Contract hours file

1. Open **ClockRite**.
2. Top-left → **Print Report**.
3. Pick **Employee Details (Advanced)**.
4. **Export** → **Excel 7.0 format (.xls)**.

> The same **Pay ID / Payroll number** must appear in both files so the
> system can match employees.

## 3. Upload and process

1. From the dashboard, click **Weekly report** (not Daily report).
2. **Employee hours file** → choose File A.
3. **Contract hours file** → choose File B.
4. Click **Process files**.
5. A **"Processing your files…"** overlay appears.
   Wait until it disappears (usually a few seconds).

If a file is wrong or missing a column, an error banner appears at the
top — fix the file and try again.

## 4. Weekly calculations

For each employee row:

- **Actual hours** = total paid hours from the time file (never below zero).
- **Contracted hours** = from the contract export (never below zero).
- **Extra hours** = actual minus contracted (minimum zero).
- **Additional Holiday pay** = **0.12 × extra hours** (decimal hours, e.g. 1.2).

Examples (matching the director’s test sheet):

| Actual | Contract | Extra | Additional Holiday pay |
| ------ | -------- | ----- | ---------------------- |
| 50 | 40 | 10 | 1.2 |
| 16 | 8 | 8 | 0.96 |
| 20 | 10 | 10 | 1.2 |
| 30 | 40 | 0 | 0 |

This is **separate** from normal **Annual holiday** (28-day entitlement HR manages in ClockRite).

## 5. Review the results

Once processing finishes you'll see, top-down:

- A **summary toolbar** with: total rows, agency, Gazebo, total paid hours.
- **View graphs** button — click to expand charts.
- **Export data** button — drop-down with **Excel / CSV / PDF**.
- A **scrollable data table** with **ExtraHours** and
  **AdditionalHolidayPay** columns (preview of first 200 rows).

The full result is **always** in the export.

## 6. Export

Click **Export data** and pick a format. Confirm the filename starts with
**`gazebo_weekly_report_`** (not `gazebo_daily_report_`).

| Format | Best for | Includes |
| ------ | -------- | -------- |
| **Excel** | Internal HR use, Sage prep | Cover sheet + **All Data** with Actual hours, Contracted hours, Extra hours, Additional Holiday pay |
| **CSV** | Importing to other systems | HR-friendly headers + all rows |
| **PDF** | Sharing / printing | Branded landscape page, full table |

In Excel, open sheet **All Data** and scroll right past **Contracted hours**
to see **Extra hours** and **Additional Holiday pay**.

Filenames are stamped with date + time, e.g.
`gazebo_weekly_report_20260501-1530.xlsx`.

> **Monthly report** still uses **Daily report** exports
> (`gazebo_daily_report_*.xlsx`), not weekly exports. Sum weekly
> **Additional Holiday pay** manually for monthly Sage payment until a
> monthly rollup is added.

## 7. Run again

To process a new week, upload two new files and click **Process files**
again — the previous result is replaced.

To sign out, use **Sign out** in the top-right corner.

---

## Troubleshooting

| Problem | What to do |
| ------- | ---------- |
| "No processed data available" when clicking Export | Upload and process the two files first. |
| No **Additional Holiday pay** column in export | You exported from **Daily report** — use **Weekly report** instead. |
| Wrong people in the result | Check that **Pay ID** matches **Payroll number** in the contract file. |
| **Contract match** = **No** | Pay ID missing from contract export — see **Case studies** in the app nav. |
| Extra hours look wrong | Check actual and contracted hours; both clamp to zero if negative. |
| Additional Holiday pay is zero | Actual hours ≤ contracted (e.g. sick/absence week) — formula uses Max(0, …). |
| Page isn't updating | Hard-refresh the browser (Cmd-Shift-R / Ctrl-Shift-R). |
