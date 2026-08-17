"""生成新的「安全质量奖核算数据」xlsx。

两种输出：
  * ``mode="mark"``  对照标记版：新增人员插到对应车间最下方并把姓名/员工编号填绿，
                     需删除的人员保留原行、姓名/员工编号填红。
  * ``mode="apply"`` 已应用版：删除类真删行、新增类真插入，不做任何着色。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .roster import NEW_BLOCK_SENTINEL, BonusFile, DiffItem
from .xlsx_surgery import (
    FILL_GREEN,
    FILL_RED,
    Highlight,
    InsertGroup,
    NewRow,
    XlsxEditor,
)


@dataclass
class ExportSummary:
    mode: str = "mark"
    added: int = 0
    removed: int = 0
    highlighted: int = 0
    new_blocks: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    per_workshop: dict[str, int] = field(default_factory=dict)

    def text(self) -> str:
        if self.mode == "apply":
            parts = [f"直接插入 {self.added} 人", f"直接删除 {self.removed} 人"]
        else:
            parts = [f"新增 {self.added} 人（标绿）", f"标记删除 {self.removed} 人（标红）"]
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
            summary.removed += 1
        else:
            highlights.append(
                Highlight(
                    row=item.frontline_row,
                    cols=[columns["name"], columns["eid"]],
                    color=FILL_RED,
                )
            )
            summary.removed += 1
            summary.highlighted += 1

    # 按车间归集新增人员
    buckets: dict[str, list[DiffItem]] = {}
    for item in adds:
        workshop = (workshop_for(item) if workshop_for else item.workshop) or ""
        if not workshop:
            summary.skipped.append(f"{item.name}（{item.eid}）未指定车间，已跳过")
            continue
        buckets.setdefault(workshop, []).append(item)

    known = {block[0]: block for block in bonus.blocks}
    groups: list[InsertGroup] = []
    for workshop in list(known) + [w for w in buckets if w not in known]:
        items = buckets.get(workshop)
        if not items:
            continue
        rows = [_new_row(item, columns, highlight=(mode == "mark")) for item in items]
        block = known.get(workshop)
        if block is not None:
            _, start, end = block
            template = _pick_template(start, end, delete_rows)
            groups.append(
                InsertGroup(anchor_row=end, template_row=template, rows=rows, new_block=False)
            )
        else:
            anchor = bonus.last_data_row
            template = _pick_template(bonus.first_data_row, anchor, delete_rows)
            label = "" if workshop == NEW_BLOCK_SENTINEL else workshop
            groups.append(
                InsertGroup(
                    anchor_row=anchor,
                    template_row=template,
                    rows=rows,
                    new_block=True,
                    block_col=columns.get("workshop", "A"),
                    block_label=label,
                )
            )
            summary.new_blocks.append(workshop)
            summary.warnings.append(
                f"「{workshop}」在原表「{sheet_name}」中不存在，已在最下方新建分组块；"
                "分块统计公式（如 a/b/c 占比）不会自动覆盖新块，请人工确认。"
            )
        summary.added += len(rows)
        summary.per_workshop[workshop] = len(rows)

    editor.edit_rows(
        sheet_name,
        deletes=delete_rows,
        inserts=groups,
        highlights=highlights,
        first_data_row=bonus.first_data_row,
    )
    summary.warnings.extend(editor.warnings)
    return editor.to_bytes(), summary


def _new_row(item: DiffItem, columns: dict[str, str], *, highlight: bool) -> NewRow:
    values = {
        columns["duty"]: item.duty,
        columns["name"]: item.name,
        columns["eid"]: item.eid,
        columns["hire"]: item.hire_date,
    }
    fills = {}
    if highlight:
        fills = {columns["name"]: FILL_GREEN, columns["eid"]: FILL_GREEN}
    return NewRow(values=values, fills=fills)


def _pick_template(start: int, end: int, deletes: set[int]) -> int:
    """挑一个不会被删掉的模板行，优先取块内最后一行。"""
    for row in range(end, start - 1, -1):
        if row not in deletes:
            return row
    return start
