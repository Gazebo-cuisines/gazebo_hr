from __future__ import annotations

from pathlib import Path

from django.core.management.base import BaseCommand, CommandError

from weekly.employee_import import import_employee_records
from weekly.payroll_service import parse_employee_directory


class Command(BaseCommand):
    help = "Import or update employees from a ClockRite Employee Details (Advanced) .xls file."

    def add_arguments(self, parser) -> None:
        parser.add_argument("xls_path", type=str, help="Path to Employee Details export (.xls/.xlsx)")
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Parse and report counts without writing to the database.",
        )
        parser.add_argument(
            "--deactivate-missing",
            action="store_true",
            help="Mark employees not in this file as inactive.",
        )

    def handle(self, *args, **options) -> None:
        path = Path(options["xls_path"]).expanduser().resolve()
        if not path.is_file():
            raise CommandError(f"File not found: {path}")

        with path.open("rb") as f:
            records = parse_employee_directory(f)

        if not records:
            raise CommandError(f"No employees parsed from {path}")

        if options["dry_run"]:
            from weekly.models import Employee

            created = sum(1 for r in records if not Employee.objects.filter(payroll_number=r["payroll_number"]).exists())
            updated = len({r["payroll_number"] for r in records}) - created
            self.stdout.write(
                self.style.SUCCESS(f"[dry-run] Would parse {len(records)} blocks; created≈{created}, updated≈{updated}")
            )
            return

        result = import_employee_records(
            records,
            path.name,
            deactivate_missing=options["deactivate_missing"],
        )
        msg = (
            f"Parsed {result.parsed} blocks; "
            f"created={result.created}, updated={result.updated}, skipped_duplicate={result.skipped_duplicate}"
        )
        if options["deactivate_missing"]:
            msg += f", deactivated={result.deactivated}"
        self.stdout.write(self.style.SUCCESS(msg))
