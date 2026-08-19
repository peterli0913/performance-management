"""生成新的「安全质量奖核算数据」xlsx。

两种输出：
  * ``mode="mark"``  对照标记版
      - 新增人员插到对应车间**该职务**的最下方，姓名与员工编号填**绿色**
      - 需删除的人员保留原行，姓名与员工编号填**红色**
      - 需按清单更新的人员就地改值，改动的单元格填**橙色**
  * ``mode="apply"`` 已应用版
      - 删除类真删行、新增类真插入，新增人员的内容用**红色字体**
      - 更新类就地改值，不着色

两版共同：清单里「职位」为实习生的人，其**职务**单元格填**黄色**底纹。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .roster import (
    ACTION_ADD,
    ACTION_REMOVE,
    ACTION_UPDATE,
    NEW_BLOCK_SENTINEL,
    BonusFile,
    DiffItem,
)
from .xlsx_surgery import (
    FILL_GREEN,
    FILL_ORANGE,
    FILL_RED,
    FILL_YELLOW,
    FONT_RED,
    Highlight,
    InsertGroup,
    NewRow,
    XlsxEditor,
)

FIELD_TO_COLUMN = {"姓名": "name", "员工编号": "eid", "职务": "duty", "入职日期": "hire"}

# 新建车间块时的职务排列顺序，跟原表各车间块的习惯一致
DEFAULT_DUTY_ORDER = ("工程师", "班长", "助工", "操作工")


def _duty_rank(order, duty: str) -> int:
    return list(order).index(duty) if duty in order else len(order)


def _count(summary: "ExportSummary", workshop: str, items: list) -> None:
    summary.added += len(items)
    summary.interns += sum(1 for item in items if item.is_intern)
    summary.per_workshop[workshop] = summary.per_workshop.get(workshop, 0) + len(items)


@dataclass
class ExportSummary:
    mode: str = "mark"
    added: int = 0
    removed: int = 0
    updated: int = 0
    interns: int = 0
    new_blocks: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    per_workshop: dict[str, int] = field(default_factory=dict)
    fallback_duties: list[str] = field(default_factory=list)

    def text(self) -> str:
        if self.mode == "apply":
            parts = [f"直接插入 {self.added} 人（红字）", f"直接删除 {self.removed} 人"]
        else:
            parts = [f"新增 {self.added} 人（标绿）", f"标记删除 {self.removed} 人（标红）"]
        if self.updated:
            parts.append(f"按清单更新 {self.updated} 人" + ("（标橙）" if self.mode == "mark" else ""))
        if self.interns:
            parts.append(f"其中实习生 {self.interns} 人（黄底）")
        if self.new_blocks:
            parts.append("新建车间分组：" + "、".join(self.new_blocks))
        if self.skipped:
            parts.append(f"跳过 {len(self.skipped)} 人")
        return "；".join(parts)


def build_workbook(
    data: bytes,
    bonus: BonusFile,
    adds: list[DiffItem],
    removes: list[DiffItem],
    updates: list[DiffItem] | None = None,
    *,
    mode: str = "mark",
    workshop_for=None,
    sheet_name: str = "一线人员",
) -> tuple[bytes, ExportSummary]:
    """返回 (xlsx 字节, 摘要)。``workshop_for(item)`` 决定每个新增人员落到哪个车间。"""
    if mode not in ("mark", "apply"):
        raise ValueError("mode 只能是 mark 或 apply")

    summary = ExportSummary(mode=mode)
    editor = XlsxEditor(data)
    columns = bonus.columns
    for required in ("duty", "name", "eid", "hire"):
        if required not in columns:
            raise ValueError(f"「{sheet_name}」缺少列：{required}")

    delete_rows: set[int] = set()
    highlights: list[Highlight] = []

    for item in removes:
        if not item.frontline_row:
            summary.skipped.append(f"{item.name}（{item.eid}）没有原始行号，已跳过")
            continue
        if mode == "apply":
            delete_rows.add(item.frontline_row)
        else:
            highlights.append(
                Highlight(
                    row=item.frontline_row,
                    cols=[columns["name"], columns["eid"]],
                    color=FILL_RED,
                )
            )
        summary.removed += 1

    for item in updates or ():
        if not item.frontline_row:
            summary.skipped.append(f"{item.name}（{item.eid}）没有原始行号，已跳过")
            continue
        if not item.updates:
            continue
        values = {
            columns[FIELD_TO_COLUMN[name]]: item.new_values.get(name)
            for name in item.updates
            if name in FIELD_TO_COLUMN
        }
        highlights.append(
            Highlight(
                row=item.frontline_row,
                cols=list(values) if mode == "mark" else [],
                color=FILL_ORANGE if mode == "mark" else None,
                values=values,
            )
        )
        if item.is_intern:
            highlights.append(
                Highlight(row=item.frontline_row, cols=[columns["duty"]], color=FILL_YELLOW)
            )
            summary.interns += 1
        summary.updated += 1

    known = {block[0]: block for block in bonus.blocks}
    # 已有车间按 (车间, 职务) 归集，每种职务插到自己那一段的最下面；
    # 全新车间必须整体成一个块，否则会被拆成几个互不相连的合并区
    buckets: dict[tuple[str, str], list[DiffItem]] = {}
    new_blocks: dict[str, list[DiffItem]] = {}
    for item in adds:
        workshop = (workshop_for(item) if workshop_for else item.workshop) or ""
        if not workshop:
            summary.skipped.append(f"{item.name}（{item.eid}）未指定车间，已跳过")
            continue
        if workshop in known:
            buckets.setdefault((workshop, item.duty), []).append(item)
        else:
            new_blocks.setdefault(workshop, []).append(item)

    groups: list[InsertGroup] = []
    ordered_keys = sorted(
        buckets,
        key=lambda pair: (
            list(known).index(pair[0]),
            _duty_rank(bonus.duty_order.get(pair[0], []), pair[1]),
            pair[1],
        ),
    )
    for workshop, duty in ordered_keys:
        items = buckets[(workshop, duty)]
        anchor, exact = bonus.anchor_for(workshop, duty)
        if not exact:
            summary.fallback_duties.append(f"{workshop}·{duty}")
        groups.append(
            InsertGroup(
                anchor_row=anchor,
                template_row=_pick_template(known[workshop][1], anchor, delete_rows),
                rows=[_new_row(item, columns, mode) for item in items],
                new_block=False,
            )
        )
        _count(summary, workshop, items)

    for workshop, items in new_blocks.items():
        items = sorted(items, key=lambda item: _duty_rank(DEFAULT_DUTY_ORDER, item.duty))
        anchor = bonus.last_data_row
        groups.append(
            InsertGroup(
                anchor_row=anchor,
                template_row=_pick_template(bonus.first_data_row, anchor, delete_rows),
                rows=[_new_row(item, columns, mode) for item in items],
                new_block=True,
                block_col=columns.get("workshop", "A"),
                block_label="" if workshop == NEW_BLOCK_SENTINEL else workshop,
            )
        )
        summary.new_blocks.append(workshop)
        summary.warnings.append(
            f"「{workshop}」在原表「{sheet_name}」中不存在，已在最下方新建分组块；"
            "分块统计公式（如 a/b/c 占比）不会自动覆盖新块，请人工确认。"
        )
        _count(summary, workshop, items)

    if summary.fallback_duties:
        summary.warnings.append(
            "以下车间原本没有这个职务，新增人员放在了该车间最下方："
            + "、".join(sorted(set(summary.fallback_duties)))
        )

    editor.edit_rows(
        sheet_name,
        deletes=delete_rows,
        inserts=groups,
        highlights=highlights,
        first_data_row=bonus.first_data_row,
    )
    summary.warnings.extend(editor.warnings)
    return editor.to_bytes(), summary


def _new_row(item: DiffItem, columns: dict[str, str], mode: str) -> NewRow:
    values = {
        columns["duty"]: item.duty,
        columns["name"]: item.name,
        columns["eid"]: item.eid,
        columns["hire"]: item.hire_date,
    }
    fills: dict[str, str] = {}
    fonts: dict[str, str] = {}
    if mode == "mark":
        fills = {columns["name"]: FILL_GREEN, columns["eid"]: FILL_GREEN}
    else:
        fonts = {column: FONT_RED for column in values}
    if item.is_intern:
        fills[columns["duty"]] = FILL_YELLOW
    return NewRow(values=values, fills=fills, font_colors=fonts)


def _pick_template(start: int, end: int, deletes: set[int]) -> int:
    """挑一个不会被删掉的模板行，优先取锚点行本身。"""
    for row in range(end, start - 1, -1):
        if row not in deletes:
            return row
    return start


def split_by_action(items: list[DiffItem]):
    """按动作把待处理人员拆成三组。"""
    adds = [item for item in items if item.action == ACTION_ADD]
    removes = [item for item in items if item.action == ACTION_REMOVE]
    updates = [item for item in items if item.action == ACTION_UPDATE]
    return adds, removes, updates
