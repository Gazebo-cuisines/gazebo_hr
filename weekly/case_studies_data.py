"""Case study content for the weekly dashboard help section."""

from __future__ import annotations

from typing import Any

CASE_STUDIES: list[dict[str, Any]] = [
    {
        "id": "contract-match-no",
        "title": "Contract match shows No",
        "tab_label": "Contract hours not matching",
        "tags": [
            "contract",
            "match",
            "no",
            "pay id",
            "new employee",
            "clockrite",
            "hr",
            "export",
        ],
        "summary": (
            "An employee appears in the hours file but Contract match is No because "
            "their Pay ID is missing from the contract export. This is usually an "
            "HR timing exception, not a system bug."
        ),
        "sections": [
            {
                "heading": "What you see",
                "body": [
                    "Contract match = No and Match reason = Pay ID not in contract export.",
                    "Contracted hours = 0; the person still has paid hours from the time file.",
                    "Only some staff are affected — often new joiners or agency temps.",
                ],
            },
            {
                "heading": "Why this happens (exception case)",
                "body": [
                    "The weekly report joins Pay ID from the hours file to Payroll Number / Sage Pay Ref in the Employee Details (Advanced) export.",
                    "When a new employee joins, HR may hold or delay setting them up in ClockRite Employee Details.",
                    "Until they appear in that export, their Pay ID will not be in the contract hours file — even if they already have hours in the Paid Hours summary.",
                    "The current implementation is correct; this is a data timing gap between the two exports, not a parsing error.",
                ],
            },
            {
                "heading": "What to do",
                "body": [
                    "Ask HR to confirm the employee exists in ClockRite with the correct Payroll Number / Sage Pay Ref.",
                    "When new employees join, use a fresh contract file exported after HR has released them — do not reuse an older Employee Details export.",
                    "Re-export Employee Details (Advanced) on the same day as the hours file, then upload both files and process again.",
                    "Contract match should change to Yes once their block appears in the new contract file.",
                ],
            },
            {
                "heading": "Example from testing",
                "body": [
                    "In a May 2026 test export, Pay IDs 752, 1653, 1658, 1664, and 1665 were in the hours file but absent from the contract file — all five showed No; the other 201 employees matched Yes.",
                ],
            },
            {
                "heading": "If you still have a problem",
                "body": [
                    "Contact HR first (employee setup / contract export).",
                    "Contact the Developer if the Pay ID is definitely in both exports but match stays No.",
                ],
            },
        ],
    },
]
