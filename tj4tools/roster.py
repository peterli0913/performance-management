"""人员清单 ↔ 安全质量奖核算数据 的解析与对账。

四类判定（按需求原文）：
  1. 新入职员工        —— 清单"生产部"有、核算"一线人员"无，且入职时间在参照日期一个月内，
                          或出现在"人员变动说明"且离职时间为空
  2. 待定需填入人员(A) —— 清单有、核算无，但不满足上面的新入职条件
  3. 离职人员          —— 核算"一线人员"有、清单"生产部"无，且在"离职人员&调出人员"，
                          或在"人员变动说明"且有离职时间
  4. 待定需填入人员(B) —— 核算有、清单无，但不满足上面的离职条件
"""

from __future__ import annotations

import datetime as _dt
import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from .normalize import (
    clean_text,
    fmt_date,
    is_blank,
    key_label,
    months_before,
    norm_eid,
    norm_name,
    parse_date,
    date_from_filename,
)

# 需求指定的目标职务；"助理工程师" 在核算表里写作 "助工"
TARGET_DUTIES = ("助工", "助理工程师", "操作工", "工程师", "班长")
DUTY_ALIAS = {"助理工程师": "助工"}

CATEGORY_NEW = "新入职员工"
CATEGORY_PENDING_ADD = "待定需填入人员（清单有·核算无）"
CATEGORY_LEFT = "离职人员"
CATEGORY_PENDING_DEL = "待定需填入人员（核算有·清单无）"

CATEGORIES = (CATEGORY_NEW, CATEGORY_PENDING_ADD, CATEGORY_LEFT, CATEGORY_PENDING_DEL)
ADD_CATEGORIES = (CATEGORY_NEW, CATEGORY_PENDING_ADD)
REMOVE_CATEGORIES = (CATEGORY_LEFT, CATEGORY_PENDING_DEL)

SHEET_PRODUCTION = "生产部"
SHEET_EQUIPMENT = "生产设备部"
SHEET_DEPARTURE = "离职人员&调出人员"
SHEET_CHANGES = "人员变动说明"
SHEET_FRONTLINE = "一线人员"
SHEET_OTHERS = "副主任&工艺组长及其他"

NEW_BLOCK_SENTINEL = "__新增车间__"


def canon_duty(value) -> str:
    text = clean_text(value)
    return DUTY_ALIAS.get(text, text)


def is_target_duty(value) -> bool:
    return canon_duty(value) in {DUTY_ALIAS.get(d, d) for d in TARGET_DUTIES}


# --------------------------------------------------------------------------- #
# 通用定位工具
# --------------------------------------------------------------------------- #


def find_sheet(workbook, *candidates: str):
    """按名字精确/模糊定位工作表。"""
    titles = {clean_text(ws.title): ws for ws in workbook.worksheets}
    for candidate in candidates:
        target = clean_text(candidate)
        if target in titles:
            return titles[target]
    for candidate in candidates:
        target = clean_text(candidate)
        for title, sheet in titles.items():
            if target and (target in title or title in target):
                return sheet
    return None


def find_col(header: list, *candidates: str) -> int | None:
    """在表头里找列（0 基）。先精确，再前缀/包含。"""
    cleaned = [clean_text(cell) for cell in header]
    for candidate in candidates:
        target = clean_text(candidate)
        for index, cell in enumerate(cleaned):
            if cell == target:
                return index
    for candidate in candidates:
        target = clean_text(candidate)
        if not target:
            continue
        for index, cell in enumerate(cleaned):
            if cell and (cell.startswith(target) or target in cell):
                return index
    return None


def _get(row: tuple, index: int | None):
    if index is None or index >= len(row):
        return None
    return row[index]


# --------------------------------------------------------------------------- #
# 数据模型
# --------------------------------------------------------------------------- #


@dataclass
class Person:
    name: str
    eid: str
    duty: str
    duty_raw: str
    group: str
    hire_date: _dt.date | None
    remark: str = ""
    leave_date: _dt.date | None = None
    leave_raw: str = ""
    row: int = 0
    workshop: str = ""

    @property
    def key(self) -> tuple[str, str]:
        return (self.name, self.eid)


@dataclass
class RosterFile:
    """TJ4 人员清单。"""

    file_name: str
    ref_date: _dt.date | None
    production: dict[tuple[str, str], Person] = field(default_factory=dict)
    production_all: dict[tuple[str, str], Person] = field(default_factory=dict)
    equipment: dict[tuple[str, str], Person] = field(default_factory=dict)
    departures: dict[tuple[str, str], Person] = field(default_factory=dict)
    changes: dict[tuple[str, str], dict] = field(default_factory=dict)
    duty_field: str = "职位"
    notes: list[str] = field(default_factory=list)


@dataclass
class BonusFile:
    """安全质量奖核算数据。"""

    file_name: str
    frontline: dict[tuple[str, str], Person] = field(default_factory=dict)
    frontline_order: list[tuple[str, str]] = field(default_factory=list)
    blocks: list[tuple[str, int, int]] = field(default_factory=list)
    others: dict[tuple[str, str], Person] = field(default_factory=dict)
    first_data_row: int = 3
    last_data_row: int = 3
    columns: dict[str, str] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    def block_of(self, workshop: str) -> tuple[str, int, int] | None:
        for block in self.blocks:
            if block[0] == workshop:
                return block
        return None

    @property
    def workshops(self) -> list[str]:
        return [block[0] for block in self.blocks]


# --------------------------------------------------------------------------- #
# 人员清单解析
# --------------------------------------------------------------------------- #


def parse_roster(data: bytes, file_name: str, duty_field: str = "职位") -> RosterFile:
    import openpyxl

    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    result = RosterFile(file_name=file_name, ref_date=date_from_filename(file_name), duty_field=duty_field)
    try:
        production = find_sheet(workbook, SHEET_PRODUCTION)
        if production is None:
            raise ValueError(f"{file_name} 里找不到「{SHEET_PRODUCTION}」子表")
        result.production_all, result.production = _parse_staff_sheet(production, duty_field)

        equipment = find_sheet(workbook, SHEET_EQUIPMENT)
        if equipment is not None:
            result.equipment = _parse_staff_sheet(equipment, duty_field)[0]

        departure = find_sheet(workbook, SHEET_DEPARTURE)
        if departure is None:
            result.notes.append(f"未找到「{SHEET_DEPARTURE}」子表，离职判定将只依赖「{SHEET_CHANGES}」")
        else:
            result.departures = _parse_departure_sheet(departure)

        changes = find_sheet(workbook, SHEET_CHANGES)
        if changes is None:
            result.notes.append(f"未找到「{SHEET_CHANGES}」子表，新入职判定将只依赖入职时间")
        else:
            result.changes = _parse_changes_sheet(changes)
    finally:
        workbook.close()
    if result.ref_date is None:
        result.notes.append("文件名里未识别出参照日期，请在界面上手工指定")
    return result


def _staff_header(sheet):
    header = [cell.value for cell in sheet[1]]
    return header, {
        "name": find_col(header, "姓名"),
        "eid": find_col(header, "员工编号", "员工号", "工号"),
        "hire": find_col(header, "入职时间", "入职日期"),
        "post": find_col(header, "岗位"),
        "title": find_col(header, "职位", "职务"),
        "group": find_col(header, "目前分组", "目前二级分组", "分组"),
        "remark": find_col(header, "备注"),
        "leave": find_col(header, "离职时间", "离职日期"),
    }


def _parse_staff_sheet(sheet, duty_field: str):
    """返回 (全部人员, 目标职务人员)。"""
    _, cols = _staff_header(sheet)
    everyone: dict[tuple[str, str], Person] = {}
    targets: dict[tuple[str, str], Person] = {}
    duty_index = cols["title"] if duty_field == "职位" else cols["post"]
    for offset, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        name = norm_name(_get(row, cols["name"]))
        if not name:
            continue
        eid = norm_eid(_get(row, cols["eid"]))
        duty_raw = clean_text(_get(row, duty_index))
        person = Person(
            name=name,
            eid=eid,
            duty=canon_duty(duty_raw),
            duty_raw=duty_raw,
            group=clean_text(_get(row, cols["group"])),
            hire_date=parse_date(_get(row, cols["hire"])),
            remark=clean_text(_get(row, cols["remark"])),
            row=offset,
        )
        everyone[person.key] = person
        if is_target_duty(duty_raw):
            targets[person.key] = person
    return everyone, targets


def _parse_departure_sheet(sheet) -> dict[tuple[str, str], Person]:
    _, cols = _staff_header(sheet)
    out: dict[tuple[str, str], Person] = {}
    for offset, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        name = norm_name(_get(row, cols["name"]))
        if not name:
            continue
        leave_raw = clean_text(_get(row, cols["leave"]))
        person = Person(
            name=name,
            eid=norm_eid(_get(row, cols["eid"])),
            duty=canon_duty(_get(row, cols["title"]) or _get(row, cols["post"])),
            duty_raw=clean_text(_get(row, cols["title"]) or _get(row, cols["post"])),
            group=clean_text(_get(row, cols["group"])),
            hire_date=parse_date(_get(row, cols["hire"])),
            remark=clean_text(_get(row, cols["remark"])),
            leave_date=parse_date(_get(row, cols["leave"])),
            leave_raw=leave_raw,
            row=offset,
        )
        out[person.key] = person
    return out


def _parse_changes_sheet(sheet) -> dict[tuple[str, str], dict]:
    """「人员变动说明」的表头不在第 1 行（第 1 行是合并标题），需要探测。"""
    header_row = 1
    for index in range(1, min(sheet.max_row, 8) + 1):
        values = [clean_text(cell.value) for cell in sheet[index]]
        if "姓名" in values:
            header_row = index
            break
    header = [cell.value for cell in sheet[header_row]]
    cols = {
        "kind": find_col(header, "分类（入职/离职/调动）", "分类"),
        "group": find_col(header, "分组"),
        "name": find_col(header, "姓名"),
        "eid": find_col(header, "员工号", "员工编号", "工号"),
        "post": find_col(header, "岗位", "职位"),
        "hire": find_col(header, "入职时间", "入职日期"),
        "leave": find_col(header, "离职时间", "离职日期"),
        "remark": find_col(header, "备注"),
    }
    out: dict[tuple[str, str], dict] = {}
    for offset, row in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
        name = norm_name(_get(row, cols["name"]))
        if not name:
            continue
        leave_value = _get(row, cols["leave"])
        out[(name, norm_eid(_get(row, cols["eid"])))] = {
            "kind": clean_text(_get(row, cols["kind"])),
            "group": clean_text(_get(row, cols["group"])),
            "post": clean_text(_get(row, cols["post"])),
            "hire_date": parse_date(_get(row, cols["hire"])),
            "leave_date": parse_date(leave_value),
            "leave_blank": is_blank(leave_value),
            "leave_raw": clean_text(leave_value),
            "remark": clean_text(_get(row, cols["remark"])),
            "row": offset,
        }
    return out


# --------------------------------------------------------------------------- #
# 核算数据解析
# --------------------------------------------------------------------------- #


def parse_bonus(data: bytes, file_name: str) -> BonusFile:
    import openpyxl

    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    result = BonusFile(file_name=file_name)
    try:
        frontline = find_sheet(workbook, SHEET_FRONTLINE)
        if frontline is None:
            raise ValueError(f"{file_name} 里找不到「{SHEET_FRONTLINE}」子表")
        _parse_frontline(frontline, result)
        others = find_sheet(workbook, SHEET_OTHERS)
        if others is None:
            result.notes.append(f"未找到「{SHEET_OTHERS}」子表，无法排除已在该表的人员")
        else:
            result.others = _parse_others(others)
    finally:
        workbook.close()
    return result


def _frontline_header(sheet) -> tuple[int, dict[str, str]]:
    """返回 (数据起始行, {字段: 列字母})。表头可能占两行。"""
    from openpyxl.utils import get_column_letter

    best_row, best_cols, best_hits = 1, {}, -1
    for row_index in range(1, min(sheet.max_row, 4) + 1):
        header = [cell.value for cell in sheet[row_index]]
        cols = {
            "workshop": find_col(header, "车间"),
            "duty": find_col(header, "职务", "岗位"),
            "name": find_col(header, "姓名"),
            "eid": find_col(header, "员工编号", "员工号", "工号"),
            "hire": find_col(header, "入职日期", "入职时间"),
        }
        hits = sum(1 for value in cols.values() if value is not None)
        if hits > best_hits:
            best_row, best_cols, best_hits = row_index, cols, hits
    # 表头行可能是合并的两行，数据从表头下一行开始；再向下跳过空姓名行
    name_index = best_cols.get("name")
    data_row = best_row + 1
    while data_row <= sheet.max_row:
        value = sheet.cell(data_row, (name_index or 2) + 1).value
        if norm_name(value):
            break
        data_row += 1
    letters = {
        key: get_column_letter(index + 1) for key, index in best_cols.items() if index is not None
    }
    return data_row, letters


def _parse_frontline(sheet, result: BonusFile) -> None:
    from openpyxl.utils import column_index_from_string

    data_row, letters = _frontline_header(sheet)
    result.first_data_row = data_row
    result.columns = letters
    required = ("workshop", "duty", "name", "eid", "hire")
    missing = [key for key in required if key not in letters]
    if missing:
        raise ValueError(f"「{SHEET_FRONTLINE}」缺少列：{missing}")

    ws_col = column_index_from_string(letters["workshop"])
    duty_col = column_index_from_string(letters["duty"])
    name_col = column_index_from_string(letters["name"])
    eid_col = column_index_from_string(letters["eid"])
    hire_col = column_index_from_string(letters["hire"])

    # 车间列纵向合并：先建立"行 -> 车间"映射
    merged: dict[int, str] = {}
    for merge in sheet.merged_cells.ranges:
        if merge.min_col != ws_col or merge.max_col != ws_col:
            continue
        value = clean_text(sheet.cell(merge.min_row, ws_col).value)
        if not value:
            continue
        for row in range(merge.min_row, merge.max_row + 1):
            merged[row] = value

    order: list[str] = []
    spans: dict[str, list[int]] = {}
    current = ""
    for row in range(data_row, sheet.max_row + 1):
        label = merged.get(row) or clean_text(sheet.cell(row, ws_col).value)
        if label:
            current = label
        name = norm_name(sheet.cell(row, name_col).value)
        if not name:
            continue
        eid = norm_eid(sheet.cell(row, eid_col).value)
        person = Person(
            name=name,
            eid=eid,
            duty=canon_duty(sheet.cell(row, duty_col).value),
            duty_raw=clean_text(sheet.cell(row, duty_col).value),
            group="",
            hire_date=parse_date(sheet.cell(row, hire_col).value),
            row=row,
            workshop=current,
        )
        key = person.key
        if key not in result.frontline:
            result.frontline[key] = person
            result.frontline_order.append(key)
        else:
            result.notes.append(f"「{SHEET_FRONTLINE}」第 {row} 行与前面重复：{name} {eid}")
        if current not in spans:
            spans[current] = [row, row]
            order.append(current)
        else:
            spans[current][1] = row
        result.last_data_row = row

    result.blocks = [(name, spans[name][0], spans[name][1]) for name in order]


def _parse_others(sheet) -> dict[tuple[str, str], Person]:
    header = [cell.value for cell in sheet[1]]
    cols = {
        "workshop": find_col(header, "车间"),
        "duty": find_col(header, "职务", "岗位"),
        "name": find_col(header, "姓名"),
        "eid": find_col(header, "员工编号", "员工号"),
        "hire": find_col(header, "入职日期", "入职时间"),
    }
    out: dict[tuple[str, str], Person] = {}
    for offset, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        name = norm_name(_get(row, cols["name"]))
        if not name:
            continue
        person = Person(
            name=name,
            eid=norm_eid(_get(row, cols["eid"])),
            duty=canon_duty(_get(row, cols["duty"])),
            duty_raw=clean_text(_get(row, cols["duty"])),
            group="",
            hire_date=parse_date(_get(row, cols["hire"])),
            row=offset,
            workshop=clean_text(_get(row, cols["workshop"])),
        )
        out.setdefault(person.key, person)
    return out


# --------------------------------------------------------------------------- #
# 车间映射
# --------------------------------------------------------------------------- #


@dataclass
class WorkshopGuess:
    group: str
    workshop: str
    source: str
    support: int = 0
    share: float = 0.0
    headcount: int = 0

    @property
    def confidence(self) -> str:
        if self.source == "经验":
            if self.support >= 5 and self.share >= 0.7:
                return "高"
            if self.support >= 2:
                return "中"
            return "低"
        if self.source == "规则":
            return "中"
        return "低"


_PAREN = re.compile(r"[（(]([^（()）]*)[)）]")
_TRAINEE = re.compile(r"[（(]([^（()）]*?)委培[)）]")


def build_workshop_mapping(roster: RosterFile, bonus: BonusFile) -> dict[str, WorkshopGuess]:
    """先用两表已匹配人员反推经验映射，覆盖不到的再用规则兜底。"""
    pairs: dict[str, Counter] = defaultdict(Counter)
    for key, person in bonus.frontline.items():
        source = roster.production.get(key) or roster.production_all.get(key)
        if source is None or not source.group:
            continue
        pairs[source.group][person.workshop] += 1

    headcount = Counter(person.group for person in roster.production.values())
    workshops = set(bonus.workshops)
    mapping: dict[str, WorkshopGuess] = {}
    for group, counter in pairs.items():
        workshop, support = counter.most_common(1)[0]
        total = sum(counter.values())
        mapping[group] = WorkshopGuess(
            group=group,
            workshop=workshop,
            source="经验",
            support=support,
            share=support / total if total else 0.0,
            headcount=headcount.get(group, 0),
        )

    for group in headcount:
        if group in mapping:
            continue
        guessed = _rule_workshop(group, workshops, mapping)
        mapping[group] = WorkshopGuess(
            group=group,
            workshop=guessed or "",
            source="规则" if guessed else "未知",
            headcount=headcount.get(group, 0),
        )
    return dict(sorted(mapping.items(), key=lambda item: -item[1].headcount))


def _rule_workshop(group: str, workshops: set[str], mapping: dict[str, WorkshopGuess]) -> str:
    if not group:
        return ""
    if group in workshops:
        return group
    # X（Y委培）：实际在 Y 干活
    trainee = _TRAINEE.search(group)
    if trainee:
        target = clean_text(trainee.group(1))
        if target in workshops:
            return target
        guess = mapping.get(target)
        if guess and guess.workshop:
            return guess.workshop
    # 去掉括号后缀再试
    base = clean_text(_PAREN.sub("", group))
    if base and base != group:
        if base in workshops:
            return base
        guess = mapping.get(base)
        if guess and guess.workshop:
            return guess.workshop
    # 区域 / 车间 用词差异
    normalized = base.replace("区域", "车间") if base else ""
    if normalized in workshops:
        return normalized
    for workshop in workshops:
        if base and (base in workshop or workshop in base):
            return workshop
    return ""


# --------------------------------------------------------------------------- #
# 对账
# --------------------------------------------------------------------------- #


@dataclass
class DiffItem:
    key: tuple[str, str]
    name: str
    eid: str
    category: str
    duty: str
    duty_raw: str
    group: str = ""
    workshop: str = ""
    hire_date: _dt.date | None = None
    leave_date: _dt.date | None = None
    leave_raw: str = ""
    departure_remark: str = ""
    reason: str = ""
    flags: list[str] = field(default_factory=list)
    frontline_row: int = 0
    roster_row: int = 0

    @property
    def label(self) -> str:
        return key_label(self.key)

    @property
    def action(self) -> str:
        return "add" if self.category in ADD_CATEGORIES else "remove"

    def as_dict(self) -> dict:
        return {
            "分类": self.category,
            "姓名": self.name,
            "员工编号": self.eid,
            "职务": self.duty,
            "车间": self.workshop,
            "目前分组": self.group,
            "入职时间": fmt_date(self.hire_date),
            "离职时间": fmt_date(self.leave_date) or self.leave_raw,
            "离职/调出备注": self.departure_remark,
            "判定依据": self.reason,
            "提示": "；".join(self.flags),
            "_key": self.label,
        }


@dataclass
class Reconciliation:
    items: list[DiffItem]
    ref_date: _dt.date | None
    window_start: _dt.date | None
    mapping: dict[str, WorkshopGuess]
    matched: int = 0
    only_roster: int = 0
    only_bonus: int = 0
    excluded_in_others: int = 0
    notes: list[str] = field(default_factory=list)

    @property
    def counts(self) -> dict[str, int]:
        counter = Counter(item.category for item in self.items)
        return {category: counter.get(category, 0) for category in CATEGORIES}

    def by_category(self, category: str) -> list[DiffItem]:
        return [item for item in self.items if item.category == category]

    def category_of(self, name: str, eid: str) -> str | None:
        target = (norm_name(name), norm_eid(eid))
        for item in self.items:
            if item.key == target:
                return item.category
        return None


def reconcile(
    roster: RosterFile,
    bonus: BonusFile,
    *,
    ref_date: _dt.date | None = None,
    window_months: int = 1,
    exclude_in_others: bool = True,
    mapping: dict[str, WorkshopGuess] | None = None,
) -> Reconciliation:
    anchor = ref_date or roster.ref_date
    window_start = months_before(anchor, window_months) if anchor else None
    mapping = mapping or build_workshop_mapping(roster, bonus)

    roster_keys = set(roster.production)
    bonus_keys = set(bonus.frontline)
    matched = roster_keys & bonus_keys

    eid_to_bonus = defaultdict(list)
    name_to_bonus = defaultdict(list)
    for key in bonus_keys:
        eid_to_bonus[key[1]].append(key)
        name_to_bonus[key[0]].append(key)
    eid_to_roster = defaultdict(list)
    name_to_roster = defaultdict(list)
    for key in roster_keys:
        eid_to_roster[key[1]].append(key)
        name_to_roster[key[0]].append(key)

    items: list[DiffItem] = []
    excluded = 0

    # --- 清单有、核算无 ------------------------------------------------- #
    for key in sorted(roster_keys - bonus_keys, key=lambda k: (roster.production[k].group, k)):
        person = roster.production[key]
        flags: list[str] = []
        if exclude_in_others and key in bonus.others:
            excluded += 1
            continue
        if key in bonus.others:
            flags.append(f"已在「{SHEET_OTHERS}」子表")
        for other in eid_to_bonus.get(key[1], []):
            if other != key:
                flags.append(f"核算表有同编号不同名：{other[0]}")
        for other in name_to_bonus.get(key[0], []):
            if other != key:
                flags.append(f"核算表有同名不同编号：{other[1]}")

        change = roster.changes.get(key)
        reasons: list[str] = []
        is_new = False
        if person.hire_date and window_start and window_start <= person.hire_date <= anchor:
            is_new = True
            reasons.append(f"入职 {person.hire_date} 在 {window_start}~{anchor} 内")
        elif person.hire_date and anchor and person.hire_date > anchor:
            flags.append(f"入职时间 {person.hire_date} 晚于参照日期 {anchor}")
        if change is not None and change["leave_blank"]:
            is_new = True
            reasons.append(f"「{SHEET_CHANGES}」第 {change['row']} 行（{change['kind']}）且离职时间为空")
        elif change is not None:
            reasons.append(f"「{SHEET_CHANGES}」第 {change['row']} 行已有离职时间 {change['leave_raw']}")
        if not reasons:
            reasons.append(
                f"入职 {fmt_date(person.hire_date) or '未知'} 不在近 {window_months} 个月内，"
                f"且「{SHEET_CHANGES}」无在职记录"
            )
        guess = mapping.get(person.group)
        items.append(
            DiffItem(
                key=key,
                name=person.name,
                eid=person.eid,
                category=CATEGORY_NEW if is_new else CATEGORY_PENDING_ADD,
                duty=person.duty,
                duty_raw=person.duty_raw,
                group=person.group,
                workshop=guess.workshop if guess else "",
                hire_date=person.hire_date,
                departure_remark=person.remark,
                reason="；".join(reasons),
                flags=flags,
                roster_row=person.row,
            )
        )

    # --- 核算有、清单无 ------------------------------------------------- #
    for key in sorted(bonus_keys - roster_keys, key=lambda k: bonus.frontline[k].row):
        person = bonus.frontline[key]
        flags = []
        departure = roster.departures.get(key)
        change = roster.changes.get(key)
        reasons = []
        is_left = False
        if departure is not None:
            is_left = True
            reasons.append(
                f"「{SHEET_DEPARTURE}」第 {departure.row} 行"
                + (f"，离职时间 {fmt_date(departure.leave_date) or departure.leave_raw}" if departure.leave_raw else "")
            )
        if change is not None and not change["leave_blank"]:
            is_left = True
            reasons.append(f"「{SHEET_CHANGES}」第 {change['row']} 行有离职时间 {change['leave_raw']}")
        elif change is not None:
            reasons.append(f"「{SHEET_CHANGES}」第 {change['row']} 行（{change['kind']}）离职时间为空")

        still = roster.production_all.get(key)
        if still is not None:
            flags.append(f"仍在「{SHEET_PRODUCTION}」但职务为「{still.duty_raw}」（非目标职务）")
        if key in roster.equipment:
            flags.append(f"在「{SHEET_EQUIPMENT}」子表")
        for other in eid_to_roster.get(key[1], []):
            if other != key:
                flags.append(f"清单有同编号不同名：{other[0]}")
        for other in name_to_roster.get(key[0], []):
            if other != key:
                flags.append(f"清单有同名不同编号：{other[1]}")
        if not reasons:
            reasons.append(f"「{SHEET_DEPARTURE}」与「{SHEET_CHANGES}」都查不到离职记录")

        items.append(
            DiffItem(
                key=key,
                name=person.name,
                eid=person.eid,
                category=CATEGORY_LEFT if is_left else CATEGORY_PENDING_DEL,
                duty=person.duty,
                duty_raw=person.duty_raw,
                group=departure.group if departure else (still.group if still else ""),
                workshop=person.workshop,
                hire_date=person.hire_date,
                leave_date=departure.leave_date if departure else (change["leave_date"] if change else None),
                leave_raw=departure.leave_raw if departure else (change["leave_raw"] if change else ""),
                departure_remark=departure.remark if departure else (change["remark"] if change else ""),
                reason="；".join(reasons),
                flags=flags,
                frontline_row=person.row,
            )
        )

    notes = list(roster.notes) + list(bonus.notes)
    unmapped = sorted(
        guess.group
        for guess in mapping.values()
        if not guess.workshop and any(item.group == guess.group for item in items if item.action == "add")
    )
    if unmapped:
        notes.append("以下分组无法自动对应车间，需要在映射表里手工指定：" + "、".join(unmapped))

    return Reconciliation(
        items=items,
        ref_date=anchor,
        window_start=window_start,
        mapping=mapping,
        matched=len(matched),
        only_roster=len(roster_keys - bonus_keys),
        only_bonus=len(bonus_keys - roster_keys),
        excluded_in_others=excluded,
        notes=notes,
    )
