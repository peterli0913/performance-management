"""功能二：生成「副主任&工艺组长及其他」子表。

目标人群由两组构成：

  组一（管理类职务，全取）
      清单「职位」为 车间副主任 / 副主任 / 工艺组长 / 经理 的人。
      示例文件里 68 人，其中 58 人已经在这张子表里。

  组二（一线四职务，但不在一线人员子表）
      清单「职位」为 助理工程师(助工) / 工程师 / 操作工 / 班长，且不在「一线人员」子表。
      "不在一线人员子表"有两种口径，差一个数量级，所以做成可切换：

      * ``SCOPE_STRICT``（默认）额外要求"功能一也放不进一线车间"。
        这些人本来就没有一线车间可去（生产技术转移组、安全组、在其他厂区…），
        归到这张表才合理。示例文件 82 人。
      * ``SCOPE_LITERAL`` 只看是否出现在一线人员子表。示例文件 319 人——
        这批人和功能一要往一线人员里加的是同一批，两个功能会把同一个人放进两张表。

职务和车间的写法都和清单不一样（清单写「车间副主任」，这张表写「副主任」；
一线人员写「11号楼D级车间」，这张表写「11号楼车间D级区域」），
所以两套映射都用两表已匹配的人反推，而不是硬编码。
"""

from __future__ import annotations

import datetime as _dt
from collections import Counter, defaultdict

from .normalize import fmt_date
from .roster import (
    ACTION_ADD,
    ACTION_REMOVE,
    CATEGORY_LEFT,
    CATEGORY_NEW,
    CATEGORY_PENDING_ADD,
    CATEGORY_PENDING_DEL,
    SHEET_OTHERS,
    SHEET_PRODUCTION,
    BonusFile,
    DiffItem,
    Person,
    Reconciliation,
    RosterFile,
    build_others_workshop_map,
    classify_intern,
    departure_evidence,
    months_before,
    new_hire_evidence,
)

# 组一：管理类职务，不管在不在一线人员子表都归这张表
MANAGEMENT_TITLES = ("车间副主任", "副主任", "工艺组长", "经理")
# 组二：一线四职务的清单写法
FRONTLINE_TITLES = ("助理工程师", "助工", "工程师", "操作工", "班长")

SCOPE_STRICT = "不在一线人员子表，且无法归入一线车间"
SCOPE_LITERAL = "不在一线人员子表"
SCOPES = (SCOPE_STRICT, SCOPE_LITERAL)

# 经验映射覆盖不到时的兜底（都能在示例文件里找到依据）
DUTY_FALLBACK = {
    "车间副主任": "副主任",
    "工艺副主管": "工艺组长",
    "助工": "助理工程师",
    "安全员": "助工",
}


def build_duty_map(roster: RosterFile, bonus: BonusFile) -> dict[str, str]:
    """清单「职位」→ 这张子表的「职务」写法，用已匹配的人反推众数。"""
    layout = bonus.others_layout
    votes: dict[str, Counter] = defaultdict(Counter)
    if layout is not None:
        for key, person in layout.people.items():
            source = roster.production_all.get(key)
            if source is not None and source.title and person.duty_raw:
                votes[source.title][person.duty_raw] += 1
    mapping = {title: counter.most_common(1)[0][0] for title, counter in votes.items()}
    for title, duty in DUTY_FALLBACK.items():
        mapping.setdefault(title, duty)
    return mapping


def build_target(
    roster: RosterFile,
    bonus: BonusFile,
    *,
    scope: str = SCOPE_STRICT,
    placeable_keys: set[tuple[str, str]] | None = None,
) -> tuple[dict[tuple[str, str], Person], dict[tuple[str, str], str]]:
    """返回 (目标人群, 每个人属于哪一组)。"""
    frontline = set(bonus.frontline)
    placeable = placeable_keys or set()
    target: dict[tuple[str, str], Person] = {}
    groups: dict[tuple[str, str], str] = {}
    for key, person in roster.production_all.items():
        if person.title in MANAGEMENT_TITLES:
            target[key] = person
            groups[key] = "管理类职务"
            continue
        if person.title not in FRONTLINE_TITLES or key in frontline:
            continue
        if scope == SCOPE_STRICT and key in placeable:
            continue
        target[key] = person
        groups[key] = "一线四职务·不在一线人员表"
    return target, groups


def reconcile_supervisors(
    roster: RosterFile,
    bonus: BonusFile,
    *,
    ref_date: _dt.date | None = None,
    new_hire_since: _dt.date | None = None,
    intern_asof: _dt.date | None = None,
    intern_months: int = 3,
    scope: str = SCOPE_STRICT,
    placeable_keys: set[tuple[str, str]] | None = None,
) -> Reconciliation:
    """把目标人群和「副主任&工艺组长及其他」子表现有人员对账，分出同样的四类。"""
    layout = bonus.others_layout
    if layout is None:
        raise ValueError(f"核算文件里找不到「{SHEET_OTHERS}」子表")

    anchor = ref_date or roster.ref_date
    if new_hire_since is None:
        new_hire_since = months_before(anchor, 1) if anchor else None
    if intern_asof is None:
        intern_asof = anchor

    duty_map = build_duty_map(roster, bonus)
    workshop_map = build_others_workshop_map(roster, bonus)
    target, groups = build_target(roster, bonus, scope=scope, placeable_keys=placeable_keys)

    current = layout.people
    target_keys = set(target)
    current_keys = set(current)
    items: list[DiffItem] = []

    for key in sorted(target_keys - current_keys, key=lambda k: (target[k].group, k)):
        person = target[key]
        is_new, reasons, flags = new_hire_evidence(person, roster, new_hire_since, anchor)
        duty = duty_map.get(person.title, person.title)
        workshop = workshop_map.get(person.group, "")
        if not workshop:
            flags.append(f"分组「{person.group}」在本表里没有对应车间，需要手工指定")
        reasons.append(f"归入依据：{groups[key]}")
        item = DiffItem(
            key=key,
            name=person.name,
            eid=person.eid,
            category=CATEGORY_NEW if is_new else CATEGORY_PENDING_ADD,
            duty=duty,
            duty_raw=person.title,
            group=person.group,
            workshop=workshop,
            hire_date=person.hire_date,
            departure_remark=person.remark,
            reason="；".join(reasons),
            flags=flags,
            roster_row=person.row,
            is_intern=person.is_intern,
            intern_source=person.intern_source,
            action=ACTION_ADD,
            target_sheet=layout.name,
            target_workshop=workshop,
            target_workshop_source="经验映射" if workshop else "",
        )
        classify_intern(item, intern_asof, intern_months)
        items.append(item)

    for key in sorted(current_keys - target_keys, key=lambda k: current[k].row):
        person = current[key]
        is_left, reasons, departure, change = departure_evidence(key, roster)
        flags: list[str] = []
        source = roster.production_all.get(key)
        if source is None:
            flags.append(f"清单「{SHEET_PRODUCTION}」里查不到这个人")
        else:
            flags.append(f"清单职位为「{source.title}」，不在本表的取人口径内")
        if key in bonus.frontline:
            flags.append("同时出现在「一线人员」子表")
        items.append(
            DiffItem(
                key=key,
                name=person.name,
                eid=person.eid,
                category=CATEGORY_LEFT if is_left else CATEGORY_PENDING_DEL,
                duty=person.duty_raw,
                duty_raw=person.duty_raw,
                group=source.group if source else "",
                workshop=person.workshop,
                hire_date=person.hire_date,
                leave_date=departure.leave_date if departure else (change["leave_date"] if change else None),
                leave_raw=departure.leave_raw if departure else (change["leave_raw"] if change else ""),
                departure_remark=departure.remark if departure else (change["remark"] if change else ""),
                reason="；".join(reasons),
                flags=flags,
                frontline_row=person.row,
                action=ACTION_REMOVE,
                target_sheet=layout.name,
            )
        )

    notes: list[str] = []
    unmapped = Counter(
        item.group for item in items if item.action == ACTION_ADD and not item.workshop
    )
    if unmapped:
        notes.append(
            f"有 {sum(unmapped.values())} 名待新增人员的分组在本表里没有对应车间，"
            "请在下方「车间映射」里指定，否则导出时会被跳过。"
        )
    scope_counts = Counter(groups[key] for key in target_keys)
    notes.append(
        f"取人口径：{scope}；目标共 {len(target_keys)} 人（"
        + "、".join(f"{name} {count} 人" for name, count in scope_counts.items())
        + f"），本表现有 {len(current_keys)} 人。"
    )

    return Reconciliation(
        items=items,
        ref_date=anchor,
        new_hire_since=new_hire_since,
        mapping={},
        intern_asof=intern_asof,
        intern_months=intern_months,
        matched=len(target_keys & current_keys),
        only_roster=len(target_keys - current_keys),
        only_bonus=len(current_keys - target_keys),
        notes=notes,
        unmapped_groups=sorted(unmapped.items(), key=lambda kv: (-kv[1], kv[0])),
    )


def target_summary(roster: RosterFile, bonus: BonusFile, placeable_keys) -> dict[str, int]:
    """两种口径各自的目标人数，用来在界面上把选择的后果说清楚。"""
    out = {}
    for scope in SCOPES:
        target, _ = build_target(roster, bonus, scope=scope, placeable_keys=placeable_keys)
        out[scope] = len(target)
    return out


def describe_duty_map(mapping: dict[str, str]) -> list[dict[str, str]]:
    return [
        {"清单职位": title, "本表职务": duty}
        for title, duty in sorted(mapping.items(), key=lambda kv: kv[0])
    ]


def format_item(item: DiffItem) -> dict:
    payload = item.as_dict()
    payload["本表职务"] = item.duty
    payload["清单职位"] = item.duty_raw
    payload["入职时间"] = fmt_date(item.hire_date)
    return payload
