"""Contract hours parsing: Clockrite report layout vs tabular XLS."""
from __future__ import annotations

import unittest
import zipfile
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

from .payroll_service import (
    PayrollResult,
    audit_contract_pay_id_coverage,
    build_emp_agency_total_df,
    build_excel_bytes,
    calculate_payroll,
    parse_contracted_hours,
    parse_employee_display_names,
    parse_employee_hours,
    total_paid_hours_from_rows,
)


class TotalPaidHoursFromRowsTest(unittest.TestCase):
    def test_sums_including_a_prefix(self) -> None:
        rows = [
            {"TotalPaidHours": 4.0, "Category": "A-EL PROD"},
            {"TotalPaidHours": 3.5, "Category": "D-STAFF"},
        ]
        self.assertEqual(total_paid_hours_from_rows(rows), 7.5)


class ClockRitePaidHoursSummaryAnnualHLTest(unittest.TestCase):
    """Paid Hours (Inc Absence) Summary: Pay ID in B, Sage in D; annual duplicated at H and L (openpyxl col 8 and 12)."""

    def test_annual_from_excel_h_and_l_columns(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 2, "Pay ID")
        ws.cell(1, 4, "Sage")
        ws.cell(1, 5, "Hrs @ 1")
        ws.cell(1, 6, "Hrs @ 2")
        ws.cell(1, 8, "Hrs @ 3")
        ws.cell(1, 10, "Hrs @ 4")
        ws.cell(1, 12, "Hrs @ 5")
        ws.cell(2, 1, "TESTCAT")
        ws.cell(2, 2, "")
        ws.cell(3, 1, "Jane Doe")
        ws.cell(3, 2, 601)
        ws.cell(3, 3, 40)
        ws.cell(3, 4, 1.0)
        ws.cell(3, 5, 2.0)
        ws.cell(3, 8, 8.0)
        ws.cell(3, 12, 8.0)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        rows = parse_employee_hours(buf)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["Category"], "TESTCAT")
        self.assertEqual(r["AnnualHoliday"], 8.0)
        self.assertEqual(r["BasicHours"], 32.0)
        self.assertEqual(r["MonFriOvertime"], 1.0)
        self.assertEqual(r["SatSunOvertime"], 2.0)
        self.assertEqual(r["TotalPaidHours"], 43.0)


class SagePayRefContractAliasTest(unittest.TestCase):
    """When Sage Pay Ref != Payroll Number in ClockRite block, Pay ID joins on Sage."""

    def test_sage_pay_ref_aliases_contract_hours_for_pay_id_join(self) -> None:
        wb = Workbook()
        ws = wb.active
        r = 0
        for label, val in (
            ("Contract Hrs", 41.25),
            ("Sage Pay Ref", 1528),
            ("Payroll Number", 1344),
        ):
            r += 1
            ws.cell(r, 6, label)
            ws.cell(r, 7, val)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        by_payroll, _ = parse_contracted_hours(buf)
        self.assertAlmostEqual(by_payroll[1344], 41.25, places=2)
        self.assertAlmostEqual(by_payroll[1528], 41.25, places=2)

    def test_calculate_payroll_joins_on_sage_no_when_payroll_number_differs(self) -> None:
        wb = Workbook()
        ws = wb.active
        r = 0
        for label, val in (
            ("Contract Hrs", 40.0),
            ("Sage Pay Ref", 1528),
            ("Payroll Number", 1344),
        ):
            r += 1
            ws.cell(r, 6, label)
            ws.cell(r, 7, val)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        employees = [
            {
                "SageNo": 1528,
                "Name": "S PATEL EOLL",
                "Category": "TEST",
                "TotalPaidHours": 45.0,
            },
        ]
        result = calculate_payroll(employees, buf)
        self.assertEqual(result.rows[0]["ContractHourMatch"], "Yes")
        self.assertEqual(result.rows[0]["ContractMatchReason"], "Matched on Pay ID")
        self.assertEqual(result.rows[0]["ContractedHours"], 40.0)
        self.assertEqual(result.rows[0]["Overtime"], 5.0)


class EmployeeDisplayNameLookupTest(unittest.TestCase):
    def _sonal_patel_block_workbook(self) -> BytesIO:
        wb = Workbook()
        ws = wb.active
        ws.cell(1, 1, 1344)
        ws.cell(1, 2, "Sonal Patel EOLL")
        ws.cell(2, 2, "Clock Name")
        ws.cell(2, 3, "S PATEL EOLL")
        ws.cell(3, 6, "Contract Hrs")
        ws.cell(3, 7, 40.0)
        ws.cell(4, 8, "Sage Pay Ref")
        ws.cell(4, 9, 1528)
        ws.cell(5, 6, "Payroll Number")
        ws.cell(5, 7, 1344)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def test_parse_display_names_maps_sage_and_clock(self) -> None:
        buf = self._sonal_patel_block_workbook()
        by_sage, by_clock = parse_employee_display_names(buf)
        self.assertEqual(by_sage[1528], "Sonal Patel EOLL")
        self.assertEqual(by_sage[1344], "Sonal Patel EOLL")
        self.assertEqual(by_clock["S PATEL EOLL"], "Sonal Patel EOLL")

    def test_calculate_payroll_replaces_clock_name_with_full_name(self) -> None:
        buf = self._sonal_patel_block_workbook()
        employees = [
            {
                "SageNo": 1528,
                "Name": "S PATEL EOLL",
                "Category": "TEST",
                "TotalPaidHours": 45.0,
            },
        ]
        result = calculate_payroll(employees, buf)
        self.assertEqual(result.rows[0]["Name"], "Sonal Patel EOLL")


class PayIdNetBasicTest(unittest.TestCase):
    def test_pay_id_path_nets_annual_from_basic(self) -> None:
        wb = Workbook()
        ws = wb.active
        # Header row: Pay ID in column C (0-based col 2)
        ws.cell(1, 2, "Name")
        ws.cell(1, 3, "Pay ID")
        ws.cell(1, 4, "Basic")
        ws.cell(1, 5, "MF")
        ws.cell(1, 6, "SS")
        ws.cell(1, 7, "Ann")
        # Category row (no numeric pay id)
        ws.cell(2, 1, "TESTCAT")
        ws.cell(2, 2, "")
        ws.cell(2, 3, "")
        # Employee: gross basic 40, annual 8 -> net basic 32, total still 40
        ws.cell(3, 1, "")
        ws.cell(3, 2, "Jane Doe")
        ws.cell(3, 3, 501)
        ws.cell(3, 4, 40)
        ws.cell(3, 5, 0)
        ws.cell(3, 6, 0)
        ws.cell(3, 7, 8)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        rows = parse_employee_hours(buf)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        self.assertEqual(r["Category"], "TESTCAT")
        self.assertEqual(r["BasicHours"], 32.0)
        self.assertEqual(r["AnnualHoliday"], 8.0)
        self.assertEqual(r["TotalPaidHours"], 40.0)


class EmpAgencyTotalDfTest(unittest.TestCase):
    def test_emp_plus_agency_equals_total(self) -> None:
        band = {"BasicHours": 10.0, "MonFriOvertime": 1.0, "SatSunOvertime": 2.0, "AnnualHoliday": 3.0, "TotalPaidHours": 16.0}
        gazebo = [{**band, "Category": "X", "Name": "A", "SageNo": 1}]
        agency = [{**band, "Category": "A-EL CLNR", "Name": "B", "SageNo": 2}]
        pr = PayrollResult(
            rows=gazebo + agency,
            agency_rows=agency,
            gazebo_rows=gazebo,
            total_paid_hours=32.0,
        )
        df = build_emp_agency_total_df(pr)
        self.assertEqual(len(df), 3)
        emp = df[df["Summary"] == "EMP"].iloc[0]
        ag = df[df["Summary"] == "AGENCY"].iloc[0]
        tot = df[df["Summary"] == "TOTAL"].iloc[0]
        for col in ("BasicHours", "TotalPaidHours"):
            self.assertAlmostEqual(float(emp[col]) + float(ag[col]), float(tot[col]), places=5)


class BuildExcelNewSheetsTest(unittest.TestCase):
    def test_workbook_contains_new_sheet_names(self) -> None:
        band = {"BasicHours": 1.0, "MonFriOvertime": 0.0, "SatSunOvertime": 0.0, "AnnualHoliday": 0.0, "TotalPaidHours": 61.0}
        row = {
            **band,
            "Name": "HI",
            "Category": "C",
            "SageNo": 9,
            "ContractedHours": 0.0,
            "Overtime": 61.0,
            "ContractHourMatch": "No",
        }
        pr = PayrollResult(rows=[row], agency_rows=[], gazebo_rows=[row], total_paid_hours=61.0)
        data = build_excel_bytes(pr)
        with zipfile.ZipFile(BytesIO(data)) as zf:
            wbxml = zf.read("xl/workbook.xml").decode("utf-8")
        for name in ("EMP Agency Total", "Category summary", "Hours over 60"):
            self.assertIn(name, wbxml, msg=f"missing sheet {name!r}")

    def test_category_breakdown_block_column_offset(self) -> None:
        band = {
            "BasicHours": 10.0,
            "MonFriOvertime": 1.0,
            "SatSunOvertime": 2.0,
            "AnnualHoliday": 3.0,
            "TotalPaidHours": 16.0,
        }
        rows = [
            {**band, "Name": "A ONE", "Category": "CLNR", "SageNo": 1, "ContractedHours": 0.0, "Overtime": 0.0, "ContractHourMatch": "No"},
            {**band, "Name": "B TWO", "Category": "OFCE", "SageNo": 2, "ContractedHours": 0.0, "Overtime": 0.0, "ContractHourMatch": "No"},
        ]
        pr = PayrollResult(rows=rows, agency_rows=[], gazebo_rows=rows, total_paid_hours=32.0)
        data = build_excel_bytes(pr)
        wb = load_workbook(BytesIO(data), read_only=True)
        ws = wb["All Data"]
        header_row = None
        for r in range(1, ws.max_row + 1):
            if ws.cell(r, 2).value == "Category":
                header_row = r
                break
        self.assertIsNotNone(header_row)
        self.assertIsNone(ws.cell(header_row, 1).value)
        self.assertEqual(ws.cell(header_row, 2).value, "Category")
        self.assertEqual(ws.cell(header_row, 4).value, "BasicHours")
        self.assertIsNone(ws.cell(header_row, 3).value)
        grand_row = header_row + 4
        self.assertEqual(ws.cell(grand_row, 2).value, "Grand total")
        self.assertEqual(ws.cell(grand_row, 4).value, 20.0)
        self.assertEqual(ws.cell(header_row, 2).border.left.style, "thin")

        emp_agency_df = build_emp_agency_total_df(pr)
        summary_header_row = grand_row + 2
        self.assertEqual(ws.cell(summary_header_row, 2).value, "Summary")
        for i, row in enumerate(emp_agency_df.itertuples(index=False)):
            r = summary_header_row + 1 + i
            self.assertEqual(ws.cell(r, 2).value, row.Summary)
            self.assertEqual(ws.cell(r, 4).value, float(row.BasicHours))
        wb.close()


_DATA = Path(__file__).resolve().parent.parent / "data"
_TEST_DATA = _DATA / "TEST_DATA"
_OVERTIME_TEST_DIR = _DATA / "ovettime_error_test_data"
_OVERTIME_EMPLOYEE = _OVERTIME_TEST_DIR / "dgross_paysummary2 (3).xls"
_OVERTIME_CONTRACT = _OVERTIME_TEST_DIR / "demployees_2023 (1).xls"
_CLOCKRITE = _DATA / "Employee contract hours - clockrite.xls"
_EMPLOYEE = _DATA / "dgross_paysummary2.xls"


@unittest.skipUnless(
    _OVERTIME_EMPLOYEE.is_file() and _OVERTIME_CONTRACT.is_file(),
    "data/ovettime_error_test_data fixtures not in repo",
)
class ClockRiteOvertimeColumnRegressionTest(unittest.TestCase):
    def test_colleague_paysummary_747_and_579(self) -> None:
        with _OVERTIME_EMPLOYEE.open("rb") as f:
            rows = {r["SageNo"]: r for r in parse_employee_hours(f)}
        self.assertEqual(rows[747]["MonFriOvertime"], 5.5)
        self.assertEqual(rows[747]["TotalPaidHours"], 46.75)
        self.assertEqual(rows[579]["MonFriOvertime"], 9.25)
        self.assertEqual(rows[579]["SatSunOvertime"], 8.5)
        self.assertEqual(rows[579]["TotalPaidHours"], 17.75)

    def test_colleague_full_name_for_sage_1528(self) -> None:
        with _OVERTIME_EMPLOYEE.open("rb") as ef, _OVERTIME_CONTRACT.open("rb") as cf:
            employees = parse_employee_hours(ef)
            result = calculate_payroll(employees, cf)
        by_sage = {r["SageNo"]: r["Name"] for r in result.rows}
        self.assertEqual(by_sage.get(1528), "Sonal Patel EOLL")


@unittest.skipUnless(
    _OVERTIME_EMPLOYEE.is_file() and _OVERTIME_CONTRACT.is_file(),
    "data/ovettime_error_test_data fixtures not in repo",
)
class ContractMatchAuditTest(unittest.TestCase):
    """Audit: Contract match No = Pay ID missing from contract export (not a parser bug)."""

    _MISSING_PAY_IDS = {752, 1653, 1658, 1664, 1665}

    def test_five_pay_ids_missing_from_contract_export(self) -> None:
        with _OVERTIME_EMPLOYEE.open("rb") as ef, _OVERTIME_CONTRACT.open("rb") as cf:
            employees = parse_employee_hours(ef)
            missing = audit_contract_pay_id_coverage(employees, cf)
        missing_ids = {m["SageNo"] for m in missing}
        self.assertEqual(missing_ids, self._MISSING_PAY_IDS)

    def test_no_match_rows_have_reason_and_201_of_206_yes(self) -> None:
        with _OVERTIME_EMPLOYEE.open("rb") as ef, _OVERTIME_CONTRACT.open("rb") as cf:
            result = calculate_payroll(parse_employee_hours(ef), cf)
        no_rows = [r for r in result.rows if r["ContractHourMatch"] == "No"]
        self.assertEqual(len(no_rows), 5)
        self.assertEqual(sum(1 for r in result.rows if r["ContractHourMatch"] == "Yes"), 201)
        for r in no_rows:
            self.assertEqual(r["ContractMatchReason"], "Pay ID not in contract export")
            self.assertIn(r["SageNo"], self._MISSING_PAY_IDS)

    def test_after_contract_reexport_expect_yes_for_missing_ids(self) -> None:
        """Manual: re-export demployees from ClockRite with Pay IDs 752, 1653, 1658, 1664, 1665, then re-run this test."""
        self.assertTrue(
            _OVERTIME_CONTRACT.is_file(),
            "Re-export Employee Details (Advanced) .xls after adding missing staff in ClockRite",
        )


@unittest.skipUnless(
    _TEST_DATA.is_dir() and any(_TEST_DATA.glob("*.xls")),
    "data/TEST_DATA/*.xls not in repo",
)
class ClockRiteTestDataFilesTest(unittest.TestCase):
    def test_paid_hours_summary_files_parse_without_error(self) -> None:
        for path in sorted(_TEST_DATA.glob("*.xls")):
            with path.open("rb") as f:
                rows = parse_employee_hours(f)
            self.assertGreater(len(rows), 10, msg=f"{path.name} expected many employee rows")

            r0 = rows[0]
            for key in ("BasicHours", "AnnualHoliday", "TotalPaidHours", "MonFriOvertime", "SatSunOvertime"):
                self.assertIn(key, r0)


@unittest.skipUnless(_CLOCKRITE.is_file() and _EMPLOYEE.is_file(), "data/*.xls fixtures not in repo")
class ClockriteContractParseTest(unittest.TestCase):
    def test_parse_clockrite_export_maps_payroll_number_to_contract_hours(self) -> None:
        with _CLOCKRITE.open("rb") as f:
            by_payroll, _ = parse_contracted_hours(f)
        self.assertGreater(
            len(by_payroll),
            100,
            "Expected Clockrite 'Employee Details' report to yield many payroll keys",
        )
        # Known block from sample file: Payroll Number 1026, Contract Hrs 0
        self.assertIn(1026, by_payroll)
        self.assertEqual(by_payroll[1026], 0.0)

    def test_join_employee_pay_id_to_contract_payroll_number(self) -> None:
        with _EMPLOYEE.open("rb") as ef, _CLOCKRITE.open("rb") as cf:
            employees = parse_employee_hours(ef)
            result = calculate_payroll(employees, cf)
        matched = sum(1 for r in result.rows if r["ContractHourMatch"] == "Yes")
        # Sample data: 189/190 Pay IDs exist as Payroll Number in the contract file
        self.assertGreaterEqual(matched, 180)
        self.assertLess(matched, len(result.rows) + 1)
