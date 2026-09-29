"""按姓名 / 工号 / 关键字检索上传表，并把选中的人投放到核算子表。"""

from __future__ import annotations

import datetime as _dt
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field

from .ingest import SheetTable
from .normalize import clean_text, fmt_date, parse_date, person_key
from .roster import (
    ACTION_ADD,
    ACTION_MOVE,
    ACTION_REMOVE,
    ACTION_UPDATE,
    BonusFile,
    DiffItem,
    Person,
    SHEET_FRONTLINE,
    SHEET_OTHERS,
)

CATEGORY_MANUAL = "检索投放"
DEFAULT_SEARCH_LIMIT = 80

_EID_TOKENS = ("员工编号", "员工号", "工号")
_DUTY_TOKENS = ("职务", "职位", "岗位名称", "岗位")


@dataclass
class SearchHit:
    file_name: str
    sheet_name: str
    excel_row: int
    cells: dict[str, str]
    name: str
    eid: str
    matched: list[str]


@dataclass
class ManualPlacement:
    placement_id: str
    name: str
    eid: str
    duty: str
    workshop: str
    hire_date: _dt.date | None
    target_sheet: str
    source_file: str = ""
    source_sheet: str = ""
    source_row: int = 0
    cells: dict[str, str] = field(default_factory=dict)

    @property
    def key(self) -> tuple[str, str]:
        return person_key(self.name, self.eid)

    def reason(self) -> str:
        return f"检索投放：来自 {self.source_file}/{self.source_sheet} 第 {self.source_row} 行"

    def values(self) -> dict[str, object]:
        return {
            "姓名": self.name,
            "员工编号": self.eid,
            "职务": self.duty,
            "入职日期": self.hire_date,
        }


def _is_eid_header(header: str) -> bool:
    return any(token in header for token in _EID_TOKENS)


def _first_by_tokens(cells: Mapping[str, str], tokens: Sequence[str]) -> str:
    for token in tokens:
        for header, value in cells.items():
            if token in header:
                return value
    return ""


def infer_fields(cells: Mapping[str, str]) -> dict[str, object]:
    """从整行原文里抽出投放表单的默认值。"""
    name = ""
    eid = ""
    workshop = ""
    group = ""
    hire = None
    for header, value in cells.items():
        if not name and "姓名" in header:
            name = value
        if not eid and _is_eid_header(header):
            eid = value
        if hire is None and "入职" in header:
            hire = parse_date(value)
        if "目前分组" in header:
            group = value
        elif "车间" in header and not workshop:
            workshop = value
    return {
        "name": name,
        "eid": eid,
        "duty": _first_by_tokens(cells, _DUTY_TOKENS),
        "workshop": workshop or group,
        "hire_date": hire,
    }


def search_tables(
    tables: Sequence[SheetTable],
    query: str,
    *,
    limit: int = DEFAULT_SEARCH_LIMIT,
) -> list[SearchHit]:
    """在所有已入库二维表里做包含匹配，返回带整行原文的命中。"""
    needle = clean_text(query).lower()
    if not needle or limit <= 0:
        return []

    hits: list[SearchHit] = []
    for table in tables:
        for index, row in enumerate(table.rows):
            cells: dict[str, str] = {}
            for header, value in zip(table.header, row):
                text = clean_text(value)
                if text:
                    cells[header] = text
            if not cells:
                continue
            matched = [header for header, text in cells.items() if needle in text.lower()]
            if not matched:
                continue
            fields = infer_fields(cells)
            hits.append(
                SearchHit(
                    file_name=table.file_name,
                    sheet_name=table.sheet_name,
                    excel_row=table.header_row_index + 1 + index,
                    cells=cells,
                    name=str(fields["name"] or ""),
                    eid=str(fields["eid"] or ""),
                    matched=matched,
                )
            )
            if len(hits) >= limit:
                return hits
    return hits


def placement_from_dict(data: Mapping) -> ManualPlacement:
    target = clean_text(data.get("target_sheet")) or SHEET_FRONTLINE
    if target not in (SHEET_FRONTLINE, SHEET_OTHERS):
        target = SHEET_FRONTLINE
    return ManualPlacement(
        placement_id=str(data.get("placement_id") or ""),
        name=clean_text(data.get("name")),
        eid=clean_text(data.get("eid")),
        duty=clean_text(data.get("duty")),
        workshop=clean_text(data.get("workshop")),
        hire_date=parse_date(data.get("hire_date")),
        target_sheet=target,
        source_file=clean_text(data.get("source_file")),
        source_sheet=clean_text(data.get("source_sheet")),
        source_row=int(data.get("source_row") or 0),
    )


def placements_from_dicts(rows: Iterable[Mapping] | None) -> list[ManualPlacement]:
    out: list[ManualPlacement] = []
    for row in rows or ():
        placement = placement_from_dict(row)
        if placement.name or placement.eid:
            out.append(placement)
    return out


def placement_keys(
    placements: Iterable[ManualPlacement],
    *,
    target_sheet: str | None = None,
) -> set[tuple[str, str]]:
    return {
        placement.key
        for placement in placements
        if target_sheet is None or placement.target_sheet == target_sheet
    }


def _others_index(
    others: Mapping[tuple[str, str], Person] | Sequence[Person] | None,
    bonus: BonusFile,
) -> dict[tuple[str, str], Person]:
    if others is None:
        return dict(bonus.others)
    if isinstance(others, Mapping):
        return dict(others)
    return {person.key: person for person in others}


def _make_item(placement: ManualPlacement, action: str, **kwargs) -> DiffItem:
    return DiffItem(
        key=placement.key,
        name=placement.name,
        eid=placement.eid,
        category=CATEGORY_MANUAL,
        duty=placement.duty,
        duty_raw=placement.duty,
        group=placement.workshop,
        workshop=placement.workshop,
        hire_date=placement.hire_date,
        reason=placement.reason(),
        action=action,
        new_values=placement.values(),
        target_sheet=kwargs.pop("target_sheet", placement.target_sheet),
        target_workshop=placement.workshop,
        roster_row=placement.source_row,
        **kwargs,
    )


def _updates_against(current: Person, placement: ManualPlacement) -> dict[str, tuple[str, str]]:
    pairs = (
        ("姓名", current.name, placement.name),
        ("员工编号", current.eid, placement.eid),
        ("职务", current.duty, placement.duty),
        ("入职日期", fmt_date(current.hire_date), fmt_date(placement.hire_date)),
    )
    return {
        field: (old, new)
        for field, old, new in pairs
        if clean_text(old) != clean_text(new)
    }


def split_placements(
    placements: Sequence[ManualPlacement],
    bonus: BonusFile,
    others: Mapping[tuple[str, str], Person] | Sequence[Person] | None = None,
) -> tuple[list[DiffItem], list[DiffItem]]:
    """把投放清单拆成一线人员 DiffItem 和副主任表 DiffItem。"""
    index = _others_index(others, bonus)
    unique: dict[tuple[str, str], ManualPlacement] = {}
    for placement in placements:
        if placement.name or placement.eid:
            unique[placement.key] = placement

    frontline: list[DiffItem] = []
    supervisor: list[DiffItem] = []
    for placement in unique.values():
        key = placement.key
        current = bonus.frontline.get(key)
        existing_other = index.get(key)
        if placement.target_sheet == SHEET_FRONTLINE:
            if current is not None:
                frontline.append(
                    _make_item(
                        placement,
                        ACTION_UPDATE,
                        frontline_row=current.row,
                        updates=_updates_against(current, placement),
                    )
                )
            else:
                frontline.append(_make_item(placement, ACTION_ADD))
            if existing_other is not None:
                supervisor.append(
                    _make_item(
                        placement,
                        ACTION_REMOVE,
                        frontline_row=existing_other.row,
                        target_sheet=SHEET_OTHERS,
                    )
                )
            continue
        if current is not None:
            frontline.append(
                _make_item(
                    placement,
                    ACTION_MOVE,
                    frontline_row=current.row,
                    target_sheet=SHEET_OTHERS,
                )
            )
        elif existing_other is None:
            supervisor.append(_make_item(placement, ACTION_ADD, target_sheet=SHEET_OTHERS))
    return frontline, supervisor
