"""Contract hours parsing: Clockrite report layout vs tabular XLS."""
from __future__ import annotations

import unittest
import zipfile
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook

from .payroll_service import (
    PayrollResult,
    build_emp_agency_total_df,
    build_excel_bytes,
    calculate_payroll,
    parse_contracted_hours,
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
        ws.cell(3, 5, 1.0)
        ws.cell(3, 6, 2.0)
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
        self.assertEqual(result.rows[0]["ContractedHours"], 40.0)
        self.assertEqual(result.rows[0]["Overtime"], 5.0)


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


_DATA = Path(__file__).resolve().parent.parent / "data"
_TEST_DATA = _DATA / "TEST_DATA"
_CLOCKRITE = _DATA / "Employee contract hours - clockrite.xls"
_EMPLOYEE = _DATA / "dgross_paysummary2.xls"


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
