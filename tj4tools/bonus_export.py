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
    ACTION_MOVE,
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
    moved: int = 0
    interns: int = 0
    new_blocks: list[str] = field(default_factory=list)
    new_other_workshops: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    per_workshop: dict[str, int] = field(default_factory=dict)
    fallback_duties: list[str] = field(default_factory=list)

    def text(self) -> str:
        if self.mode == "apply":
            parts = [f"直接插入 {self.added} 人（红字）", f"直接删除 {self.removed} 人"]
        else:
            parts = [f"新增 {self.added} 人（标绿）", f"标记删除 {self.removed} 人（标红）"]
        if self.moved:
            parts.append(f"移到「副主任&工艺组长及其他」{self.moved} 人")
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
    moves: list[DiffItem] | None = None,
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

    # 移出「一线人员」的人在这张表上的处理和删除一样
    for item in list(removes) + list(moves or ()):
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
        if item.action == ACTION_MOVE:
            summary.moved += 1
        else:
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
    if moves:
        _insert_into_others(editor, bonus, moves, mode, summary)
    summary.warnings.extend(editor.warnings)
    return editor.to_bytes(), summary


def insert_into_layout(
    editor: XlsxEditor,
    layout,
    items: list[DiffItem],
    mode: str,
    summary: ExportSummary,
    *,
    duty_of=None,
    workshop_of=None,
    deletes: set[int] | None = None,
    highlights: list[Highlight] | None = None,
) -> None:
    """把人员插进一张"按车间分组、车间内按职务分段"的子表。

    「副主任&工艺组长及其他」的车间列**不合并**（每行都写车间名），所以不建合并区，
    直接把车间当成普通一列写进去。功能一的"移到该表"和功能二的"新增到该表"共用这里。
    """
    columns = layout.columns
    duty_of = duty_of or (lambda item: str(item.new_values.get("职务") or item.duty))
    workshop_of = workshop_of or (lambda item: item.target_workshop or item.workshop)

    buckets: dict[tuple[str, str], list[DiffItem]] = {}
    for item in items:
        workshop = workshop_of(item)
        if not workshop:
            summary.skipped.append(f"{item.name}（{item.eid}）未指定车间，已跳过")
            continue
        buckets.setdefault((workshop, duty_of(item)), []).append(item)

    groups: list[InsertGroup] = []
    known = layout.workshops
    ordered = sorted(
        buckets,
        key=lambda pair: (
            known.index(pair[0]) if pair[0] in known else len(known),
            _duty_rank(layout.duty_order.get(pair[0], []), pair[1]),
            pair[0],
            pair[1],
        ),
    )
    for workshop, duty in ordered:
        bucket = buckets[(workshop, duty)]
        anchor, precision = layout.anchor_for(workshop, duty)
        if precision == "表尾" and workshop not in summary.new_other_workshops:
            summary.new_other_workshops.append(workshop)
        # 写回用单元格里的原始文本，避免全角/半角括号差异把车间块拆成两段
        label = layout.raw_workshop(workshop)
        groups.append(
            InsertGroup(
                anchor_row=anchor,
                template_row=_pick_template(layout.first_data_row, anchor, deletes or set()),
                rows=[
                    _new_row(item, columns, mode, workshop=label, duty=duty) for item in bucket
                ],
                new_block=False,
            )
        )
        summary.interns += sum(1 for item in bucket if item.is_intern)
    if summary.new_other_workshops:
        summary.warnings.append(
            f"「{layout.name}」里原本没有这些车间，相关人员已追加在该表人员区最下方："
            + "、".join(summary.new_other_workshops)
        )
    editor.edit_rows(
        layout.name,
        deletes=deletes or set(),
        inserts=groups,
        highlights=highlights or [],
        first_data_row=layout.first_data_row,
    )


def _insert_into_others(
    editor: XlsxEditor, bonus: BonusFile, moves: list[DiffItem], mode: str, summary: ExportSummary
) -> None:
    layout = bonus.others_layout
    if layout is None:
        summary.warnings.append("核算文件里没有「副主任&工艺组长及其他」子表，移动动作已跳过")
        return
    insert_into_layout(editor, layout, moves, mode, summary)


def _new_row(
    item: DiffItem,
    columns: dict[str, str],
    mode: str,
    workshop: str | None = None,
    duty: str | None = None,
) -> NewRow:
    values = {
        columns["duty"]: duty or item.duty,
        columns["name"]: str(item.new_values.get("姓名") or item.name),
        columns["eid"]: str(item.new_values.get("员工编号") or item.eid),
        columns["hire"]: item.new_values.get("入职日期") or item.hire_date,
    }
    # 车间列不合并的子表（副主任&工艺组长及其他）需要每行都写车间名
    if workshop is not None and "workshop" in columns:
        values[columns["workshop"]] = workshop
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
    """按动作把待处理人员拆成四组：新增 / 删除 / 就地更新 / 移到另一子表。"""
    adds = [item for item in items if item.action == ACTION_ADD]
    removes = [item for item in items if item.action == ACTION_REMOVE]
    updates = [item for item in items if item.action == ACTION_UPDATE]
    moves = [item for item in items if item.action == ACTION_MOVE]
    return adds, removes, updates, moves
