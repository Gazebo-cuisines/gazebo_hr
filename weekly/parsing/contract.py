from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .common import (
    _load_sheet,
    _normalize_header,
    _parse_decimal,
    _parse_int,
    _row_contains_date_range,
    _to_text,
)

@dataclass(frozen=True)
class EmployeeContractBlock:
    payroll_number: int | None
    sage_pay_ref: int | None
    prox_id: int | None
    title_id: int | None
    full_name: str
    clock_name: str
    contract_hours: float
    source_row: int


@dataclass
class ContractFileIndex:
    blocks: list[EmployeeContractBlock]
    by_payroll: dict[int, float]
    by_name: dict[str, float]
    by_sage_name: dict[int, str]
    by_clock_name: dict[str, str]
    pay_id_to_block: dict[int, EmployeeContractBlock]
    name_to_block: dict[str, EmployeeContractBlock]
    conflicted_pay_ids: frozenset[int]
    conflicts: list[dict[str, Any]]

def _find_header_index(headers: list[str], *candidates: str) -> int:
    for idx, cell in enumerate(headers):
        n = _normalize_header(cell)
        if any(n == c for c in candidates):
            return idx
    return -1


def _contract_hrs_value_in_row(row: list[str]) -> float | None:
    for c, cell in enumerate(row):
        n = _normalize_header(cell)
        if n in ("contracthrs", "contracthours", "contracthr", "contractedhours"):
            if c + 1 < len(row):
                return _parse_decimal(row[c + 1])
            return 0.0
    return None


def _block_start_row(rows: list[list[str]], r_payroll: int) -> int:
    """Row index of the block title row (Prox ID + full name) above Payroll Number."""
    for back in range(1, 30):
        if r_payroll - back < 0:
            break
        pr = rows[r_payroll - back]
        if not pr:
            continue
        a0 = _to_text(pr[0])
        if a0.isdigit() and len(a0) < 6 and len(pr) > 1 and _to_text(pr[1]):
            return r_payroll - back
    return max(0, r_payroll - 25)


def _block_names_near(rows: list[list[str]], r_payroll: int) -> tuple[str, str]:
    """Full name from block title row; clock name from Clock Name row above Payroll Number."""
    full_name = ""
    clock_name = ""
    for back in range(1, 30):
        if r_payroll - back < 0:
            break
        pr = rows[r_payroll - back]
        if not pr:
            continue
        for c, cell in enumerate(pr):
            if _normalize_header(cell) == "clockname" and c + 1 < len(pr):
                cn = _to_text(pr[c + 1])
                if cn:
                    clock_name = cn
        if full_name:
            continue
        a0 = _to_text(pr[0])
        if a0.isdigit() and len(a0) < 6 and len(pr) > 1:
            candidate = _to_text(pr[1])
            if candidate:
                full_name = candidate
    return full_name, clock_name

def _parse_single_clockrite_block(rows: list[list[str]], r_payroll: int, pay_no: int) -> EmployeeContractBlock:
    block_start = _block_start_row(rows, r_payroll)
    hours = 0.0
    for back in range(1, r_payroll - block_start + 1):
        got = _contract_hrs_value_in_row(rows[r_payroll - back])
        if got is not None:
            hours = got
            break

    sage_pay_ref: int | None = None
    prox_id: int | None = None
    r1 = r_payroll + 1
    for rr in range(block_start, r1):
        row = rows[rr]
        for c, cell in enumerate(row):
            n = _normalize_header(cell)
            if n == "sagepayref" and c + 1 < len(row):
                sid = _parse_int(row[c + 1])
                if sid is not None and sid > 0:
                    sage_pay_ref = sid
            elif n == "proxid" and c + 1 < len(row):
                pid = _parse_int(row[c + 1])
                if pid is not None and pid > 0:
                    prox_id = pid

    full_name, clock_name = _block_names_near(rows, r_payroll)
    title_id = _parse_int(rows[block_start][0] if block_start < len(rows) and rows[block_start] else "")
    return EmployeeContractBlock(
        payroll_number=pay_no,
        sage_pay_ref=sage_pay_ref,
        prox_id=prox_id,
        title_id=title_id,
        full_name=full_name,
        clock_name=clock_name,
        contract_hours=hours,
        source_row=r_payroll + 1,
    )


def _parse_clockrite_blocks(rows: list[list[str]]) -> list[EmployeeContractBlock]:
    blocks: list[EmployeeContractBlock] = []
    for r, row in enumerate(rows):
        for c, cell in enumerate(row):
            if _normalize_header(cell) != "payrollnumber":
                continue
            if c + 1 >= len(row):
                continue
            pay_no = _parse_int(row[c + 1])
            if pay_no is None or pay_no == 0:
                continue
            blocks.append(_parse_single_clockrite_block(rows, r, pay_no))
    return blocks


def _block_pay_ids(block: EmployeeContractBlock) -> list[int]:
    """Sage/payroll identifiers only. Prox ID and Title ID are badge/card numbers."""
    ids: list[int] = []
    for value in (block.payroll_number, block.sage_pay_ref):
        if value is not None and value > 0 and value not in ids:
            ids.append(value)
    return ids


def _build_contract_index_from_blocks(blocks: list[EmployeeContractBlock]) -> ContractFileIndex:
    by_payroll: dict[int, float] = {}
    by_name: dict[str, float] = {}
    by_sage_name: dict[int, str] = {}
    by_clock_name: dict[str, str] = {}
    pay_id_to_block: dict[int, EmployeeContractBlock] = {}
    name_to_block: dict[str, EmployeeContractBlock] = {}
    conflicts: list[dict[str, Any]] = []
    conflicted_pay_ids: set[int] = set()

    def register_pay_id(pay_id: int, block: EmployeeContractBlock) -> None:
        if pay_id in conflicted_pay_ids:
            return
        existing = pay_id_to_block.get(pay_id)
        if existing is None:
            pay_id_to_block[pay_id] = block
            by_payroll[pay_id] = block.contract_hours
            return
        if abs(existing.contract_hours - block.contract_hours) > 0.001:
            conflicted_pay_ids.add(pay_id)
            conflicts.append(
                {
                    "PayId": pay_id,
                    "HoursA": existing.contract_hours,
                    "HoursB": block.contract_hours,
                    "RowA": existing.source_row,
                    "RowB": block.source_row,
                    "NameA": existing.full_name,
                    "NameB": block.full_name,
                }
            )

    def register_name(name: str, block: EmployeeContractBlock) -> None:
        key = name.upper()
        if not key:
            return
        existing = name_to_block.get(key)
        if existing is None:
            name_to_block[key] = block
            by_name[key] = block.contract_hours
            return
        if abs(existing.contract_hours - block.contract_hours) > 0.001:
            conflicts.append(
                {
                    "Name": name,
                    "HoursA": existing.contract_hours,
                    "HoursB": block.contract_hours,
                    "RowA": existing.source_row,
                    "RowB": block.source_row,
                }
            )

    for block in blocks:
        for pay_id in _block_pay_ids(block):
            register_pay_id(pay_id, block)
            if block.full_name and pay_id not in conflicted_pay_ids:
                by_sage_name[pay_id] = block.full_name
        if block.full_name:
            register_name(block.full_name, block)
        if block.clock_name:
            by_clock_name[block.clock_name.upper()] = block.full_name

    return ContractFileIndex(
        blocks=blocks,
        by_payroll=by_payroll,
        by_name=by_name,
        by_sage_name=by_sage_name,
        by_clock_name=by_clock_name,
        pay_id_to_block=pay_id_to_block,
        name_to_block=name_to_block,
        conflicted_pay_ids=frozenset(conflicted_pay_ids),
        conflicts=conflicts,
    )


def _build_contract_index_from_tabular_rows(
    rows: list[list[str]], header_row: int, payroll_col: int, contract_col: int, name_col: int
) -> ContractFileIndex:
    blocks: list[EmployeeContractBlock] = []
    for offset, row in enumerate(rows[header_row + 1 :], start=header_row + 2):
        if _row_contains_date_range(row):
            break
        hours = _parse_decimal(row[contract_col] if contract_col < len(row) else "")
        if hours == 0.0:
            continue
        pay_no = _parse_int(row[payroll_col] if payroll_col < len(row) else "")
        name = _to_text(row[name_col] if name_col >= 0 and name_col < len(row) else "")
        blocks.append(
            EmployeeContractBlock(
                payroll_number=pay_no,
                sage_pay_ref=None,
                prox_id=None,
                title_id=None,
                full_name=name,
                clock_name=name,
                contract_hours=hours,
                source_row=offset,
            )
        )
    return _build_contract_index_from_blocks(blocks)


def load_contract_file_index(file_obj: Any) -> ContractFileIndex:
    df = _load_sheet(file_obj)
    rows = [[_to_text(v) for v in rec] for rec in df.values.tolist()]
    if not rows:
        return ContractFileIndex([], {}, {}, {}, {}, {}, {}, frozenset(), [])

    header_row = -1
    payroll_col = -1
    contract_col = -1
    name_col = -1
    for r, row in enumerate(rows[:80]):
        pc = _find_header_index(row, "payrollnumber", "payrollno", "payroll")
        cc = _find_header_index(row, "contracthrs", "contracthours", "contracthr", "contractedhours")
        if pc >= 0 and cc >= 0:
            header_row = r
            payroll_col = pc
            contract_col = cc
            name_col = _find_header_index(row, "clockname", "name", "employeename")
            break

    if header_row >= 0:
        index = _build_contract_index_from_tabular_rows(rows, header_row, payroll_col, contract_col, name_col)
        if index.blocks:
            return index

    return _build_contract_index_from_blocks(_parse_clockrite_blocks(rows))



def _parse_clockrite_display_names(rows: list[list[str]]) -> tuple[dict[int, str], dict[str, str]]:
    index = _build_contract_index_from_blocks(_parse_clockrite_blocks(rows))
    return index.by_sage_name, index.by_clock_name


def parse_employee_display_names(file_obj: Any) -> tuple[dict[int, str], dict[str, str]]:
    """Full names from contract file (ClockRite Employee Details layout)."""
    index = load_contract_file_index(file_obj)
    return index.by_sage_name, index.by_clock_name


def _parse_clockrite_contract_report(rows: list[list[str]]) -> tuple[dict[int, float], dict[str, float]]:
    index = _build_contract_index_from_blocks(_parse_clockrite_blocks(rows))
    return index.by_payroll, index.by_name


def parse_contracted_hours(file_obj: Any) -> tuple[dict[int, float], dict[str, float]]:
    index = load_contract_file_index(file_obj)
    return index.by_payroll, index.by_name

