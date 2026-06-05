# Weekly Report — How to use it

A short procedure for running the weekly payroll report on
[gazeboo.cloud](https://gazeboo.cloud).

## 1. Sign in

1. Open <https://gazeboo.cloud/login/>.
2. Enter your **Username** and **Password**.
   - Forgot it? Click **Forgot password?** → contact Utsav for a reset.
3. After signing in you land on the **Dashboard**.

## 2. Get the two input files from ClockRite

You need **two** Excel files exported from ClockRite. Both are
**Excel 7.0 (.xls)**.

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

1. From the dashboard, click **Weekly report**.
2. **Employee hours file** → choose File A.
3. **Contract hours file** → choose File B.
4. Click **Process files**.
5. A **"Processing your files…"** overlay appears.
   Wait until it disappears (usually a few seconds).

If a file is wrong or missing a column, an error banner appears at the
top — fix the file and try again.

## 4. Review the results

Once processing finishes you'll see, top-down:

- A **summary toolbar** with: total rows, agency, Gazebo, total paid hours.
- **View graphs** button — click to expand charts:
  hours distribution, 60+ hours table, EMP vs Agency, etc.
- **Export data** button — drop-down with **Excel / CSV / PDF**.
- A **scrollable data table** with filter (Category) and search.

The full result is **always** in the export — the on-screen table only
previews the first 200 rows.

## 5. Export

Click **Export data** and pick a format:

| Format | Best for | Includes |
| ------ | -------- | -------- |
| **Excel** | Internal HR use, further calcs | Branded "Cover" sheet + All Data, Agency, Gazebo, Analysis, EMP Agency Total, Category summary, Hours over 60 |
| **CSV** | Importing to other systems | Brand + generated date metadata, then headers + all rows |
| **PDF** | Sharing / printing | Branded landscape page, summary line, full table, footer with brand + page number |

Filenames are stamped with date + time, e.g.
`gazebo_weekly_report_20260501-1530.xlsx`.

## 6. Run again

To process a new week, simply upload two new files and click
**Process files** again — the previous result is replaced.

To sign out, use **Sign out** in the top-right corner.

---

## Troubleshooting

| Problem | What to do |
| ------- | ---------- |
| "No processed data available" when clicking Export | Upload and process the two files first. |
| Wrong people in the result | Check that **Pay ID** in the employee file matches **Payroll number** in the contract file. |
| **Contract match** = **No** | That Pay ID is **not** in the contract export (often new joiners — HR may hold setup). Re-export **Employee Details (Advanced)** after HR releases them. See **Case studies** in the app nav. Contact **HR**, then the **Developer** if both exports contain the Pay ID but match stays No. Column **ContractMatchReason** shows e.g. `Pay ID not in contract export`. |
| Numbers look off for one person | Check that ClockRite export type is the right one (see step 2). Re-export and try again. |
| Page isn't updating | Hard-refresh the browser (Cmd-Shift-R / Ctrl-Shift-R). |
| Can't sign in | Contact Utsav to reset the password. |

## Employee directory (hour contracts)

The **Employee hour contracts** page lists staff imported from the same **Employee Details (Advanced)** export used for weekly contract matching.

### First import or refresh after HR updates ClockRite

On **Employee hour contracts**, use **Import / update** and upload the `.xls` or `.xlsx` file from ClockRite (**Employee Details (Advanced)**).

Or on the server (or locally with MySQL running):

```bash
python manage.py import_employees path/to/Employee_Details_export.xls
```

- Re-importing **updates** existing rows by payroll number; it does not create duplicates.
- Optional checkbox on the page (or `--deactivate-missing` on the command): mark employees not in the file as inactive — use only for a full company export.

After import, search by name or group and click a card to open the employee profile. Holiday forms and correspondence documents are planned for a later release.
