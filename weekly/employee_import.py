from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from django.utils import timezone

from weekly.models import Employee


def derive_is_active(status: str, finish_date) -> bool:
    if (status or "").strip().lower() != "current":
        return False
    if finish_date is None:
        return True
    return finish_date >= timezone.localdate()


@dataclass
class EmployeeImportResult:
    parsed: int
    created: int
    updated: int
    skipped_duplicate: int
    deactivated: int


def import_employee_records(
    records: list[dict[str, Any]],
    source_filename: str,
    *,
    deactivate_missing: bool = False,
) -> EmployeeImportResult:
    seen_payroll: set[int] = set()
    created = updated = skipped = 0

    for rec in records:
        pay = int(rec["payroll_number"])
        if pay in seen_payroll:
            skipped += 1
            continue
        seen_payroll.add(pay)

        slug = Employee.make_slug(rec["full_name"], pay)
        defaults = {
            "prox_id": rec.get("prox_id"),
            "sage_pay_ref": rec.get("sage_pay_ref"),
            "full_name": rec["full_name"],
            "clock_name": rec.get("clock_name") or "",
            "slug": slug,
            "company": rec.get("company") or "",
            "group": rec.get("group") or "",
            "status": rec.get("status") or "",
            "contract_hours": rec.get("contract_hours") or 0,
            "basic_rate": rec.get("basic_rate") or 0,
            "holidays_left": rec.get("holidays_left"),
            "sage_pay_freq": rec.get("sage_pay_freq") or "",
            "sage_paylink": rec.get("sage_paylink") or "",
            "start_date": rec.get("start_date"),
            "finish_date": rec.get("finish_date"),
            "shift_group": rec.get("shift_group") or "",
            "ot_rule": rec.get("ot_rule") or "",
            "default_activity": rec.get("default_activity") or "",
            "default_job": rec.get("default_job") or "",
            "notes": rec.get("notes") or "",
            "is_active": derive_is_active(rec.get("status", ""), rec.get("finish_date")),
            "source_filename": source_filename,
        }

        _obj, was_created = Employee.objects.update_or_create(
            payroll_number=pay,
            defaults=defaults,
        )
        if was_created:
            created += 1
        else:
            updated += 1

    deactivated = 0
    if deactivate_missing:
        deactivated = (
            Employee.objects.exclude(payroll_number__in=seen_payroll)
            .filter(is_active=True)
            .update(is_active=False)
        )

    return EmployeeImportResult(
        parsed=len(records),
        created=created,
        updated=updated,
        skipped_duplicate=skipped,
        deactivated=deactivated,
    )
