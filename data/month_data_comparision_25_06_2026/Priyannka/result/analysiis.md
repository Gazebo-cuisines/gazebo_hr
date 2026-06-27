e-audit: monthly_report_output (7).xlsx vs legacy
Calculations — all pass
Check	Result
5 weeks included
Week1–Week5 + Summary
Band sum (Basic+MF+SS+Annual=Total)
0 errors all 5 weeks
Extra hours
0 errors
Holiday pay (0.1207 × extra)
0 errors
Weekly vs legacy blocks
Output week	Legacy block	Match
Week4
R296 — 5,817.75
Exact
Week5
R300 — 5,920.75
0.33h off (Sage 205 rounding)
Week1–3
—
Different EMP totals (see below)
Week1 agency 1,753.25 matches legacy week1 agency exactly. EMP differs by ~110h — likely different weekly source files for weeks 1–3, not a formula bug.

Monthly totals
Legacy	Output (7)	Diff
Total
28,181.25
27,888.83
-292.42h
AGENCY
10,305.50
10,305.50
Exact
EMP
17,875.75
17,583.33
-292.42h
Per-employee (by Sage ID)
98 shared Sage IDs
97/98 exact on all hour bands
1 diff: Sage 205 (Surendran Subramani) — -0.33h in Week5
292h gap = legacy-only employees not in your upload:

P LEITE (504) — 122h
R VIRA (728) — 108h
Y JAGDISH (745) — 62.75h
H SONAWANE RESPONSE (sage 0 in legacy) — 114h
Output has RESPONSE staff with real Sage IDs (e.g. 1660); legacy monthly used sage 0 for some of them.

Verdict
Calculations are correct. With week 5 included, results align closely with legacy:

Agency monthly total: exact
Shared employees: 97/98 exact
Week4: exact
Week5: 0.33h rounding only
Remaining 292h gap is missing/different source employees (Leite, Vira, Jagdish + sage-0 RESPONSE rows), not formula errors.

Still open: contracted hours only set for 3 employees — contract file matching still needs checking for extra/holiday pay accuracy.

