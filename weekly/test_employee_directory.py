from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path

from django.contrib.auth import get_user_model
from django.core.management import call_command
from django.test import Client, TestCase
from django.urls import reverse

from weekly.models import Employee
from weekly.payroll_service import parse_employee_directory

_DATA = Path(__file__).resolve().parent.parent / "data" / "bug1528"
_DEMPLOYEES = _DATA / "demployees_2023.xls"


@unittest.skipUnless(_DEMPLOYEES.is_file(), "fixture demployees_2023.xls missing")
class ParseEmployeeDirectoryTest(unittest.TestCase):
    def test_parse_count_and_sample_employee(self) -> None:
        with _DEMPLOYEES.open("rb") as f:
            records = parse_employee_directory(f)
        self.assertGreaterEqual(len(records), 200)
        by_pay = {r["payroll_number"]: r for r in records}
        self.assertIn(736, by_pay)
        rec = by_pay[736]
        self.assertEqual(rec["full_name"], "SAMJHANA ADHIKARI KARKI")
        self.assertIn("ADHIKARI", rec["clock_name"])
        self.assertEqual(rec["contract_hours"], 0.0)
        self.assertEqual(rec["status"], "Current")


@unittest.skipUnless(_DEMPLOYEES.is_file(), "fixture demployees_2023.xls missing")
class ImportEmployeesCommandTest(TestCase):
    def test_upsert_no_duplicates(self) -> None:
        path = str(_DEMPLOYEES)
        call_command("import_employees", path)
        count_first = Employee.objects.count()
        self.assertGreaterEqual(count_first, 200)
        call_command("import_employees", path)
        self.assertEqual(Employee.objects.count(), count_first)


class EmployeeViewsTest(TestCase):
    def setUp(self) -> None:
        self.client = Client()
        User = get_user_model()
        self.user = User.objects.create_user(username="testhr", password="testpass123")
        Employee.objects.create(
            payroll_number=99901,
            full_name="Test Person",
            clock_name="T PERSON",
            slug="test-person-99901",
            group="PROD TEST",
            status="Current",
            contract_hours=40,
            is_active=True,
        )

    def test_list_requires_login(self) -> None:
        resp = self.client.get(reverse("weekly:employee_hour_contracts"))
        self.assertEqual(resp.status_code, 302)
        self.assertIn("login", resp["Location"])

    def test_list_and_detail_when_logged_in(self) -> None:
        self.client.login(username="testhr", password="testpass123")
        list_resp = self.client.get(reverse("weekly:employee_hour_contracts"))
        self.assertEqual(list_resp.status_code, 200)
        self.assertContains(list_resp, "Test Person")
        self.assertContains(list_resp, "Import / update")
        detail_resp = self.client.get(reverse("weekly:employee_detail", kwargs={"slug": "test-person-99901"}))
        self.assertEqual(detail_resp.status_code, 200)
        self.assertContains(detail_resp, "Payroll number")
        bad = self.client.get(reverse("weekly:employee_detail", kwargs={"slug": "no-such-slug"}))
        self.assertEqual(bad.status_code, 404)


@unittest.skipUnless(_DEMPLOYEES.is_file(), "fixture demployees_2023.xls missing")
class EmployeeImportUploadTest(TestCase):
    def setUp(self) -> None:
        self.client = Client()
        User = get_user_model()
        self.user = User.objects.create_user(username="importhr", password="testpass123")
        self.client.login(username="importhr", password="testpass123")

    def test_post_import_updates_directory(self) -> None:
        before = Employee.objects.count()
        with _DEMPLOYEES.open("rb") as f:
            resp = self.client.post(
                reverse("weekly:employee_hour_contracts"),
                {"employee_details_file": f},
            )
        self.assertEqual(resp.status_code, 302)
        self.assertGreater(Employee.objects.count(), before)
        follow = self.client.get(reverse("weekly:employee_hour_contracts"))
        self.assertContains(follow, "Import complete")
