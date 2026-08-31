"""生成「副主任&工艺组长及其他」子表更新后的 xlsx。

和功能一的两种输出对齐：
  * ``mode="mark"``  对照标记版：新增人员插到对应车间该职务的最下方、姓名与员工编号填绿色；
                     需删除的人员保留原行、填红色。
  * ``mode="apply"`` 已应用版：删除类真删行、新增类真插入且内容用红色字体。
两版共同：实习生的职务单元格填黄色底纹。
"""

from __future__ import annotations

from .bonus_export import ExportSummary, insert_into_layout
from .roster import ACTION_ADD, ACTION_REMOVE, BonusFile, DiffItem
from .xlsx_surgery import FILL_RED, Highlight, XlsxEditor


def build_supervisor_workbook(
    data: bytes,
    bonus: BonusFile,
    adds: list[DiffItem],
    removes: list[DiffItem],
    *,
    mode: str = "mark",
    workshop_for=None,
) -> tuple[bytes, ExportSummary]:
    if mode not in ("mark", "apply"):
        raise ValueError("mode 只能是 mark 或 apply")
    layout = bonus.others_layout
    if layout is None:
        raise ValueError("核算文件里找不到「副主任&工艺组长及其他」子表")

    summary = ExportSummary(mode=mode)
    editor = XlsxEditor(data)
    columns = layout.columns

    deletes: set[int] = set()
    highlights: list[Highlight] = []
    for item in removes:
        if not item.frontline_row:
            summary.skipped.append(f"{item.name}（{item.eid}）没有原始行号，已跳过")
            continue
        if mode == "apply":
            deletes.add(item.frontline_row)
        else:
            highlights.append(
                Highlight(
                    row=item.frontline_row,
                    cols=[columns["name"], columns["eid"]],
                    color=FILL_RED,
                )
            )
        summary.removed += 1

    before = len(summary.skipped)
    insert_into_layout(
        editor,
        layout,
        adds,
        mode,
        summary,
        duty_of=lambda item: item.duty,
        workshop_of=workshop_for or (lambda item: item.target_workshop or item.workshop),
        deletes=deletes,
        highlights=highlights,
    )
    summary.added = len(adds) - (len(summary.skipped) - before)
    for item in adds:
        workshop = (workshop_for(item) if workshop_for else item.target_workshop) or item.workshop
        if workshop:
            summary.per_workshop[workshop] = summary.per_workshop.get(workshop, 0) + 1
    summary.warnings.extend(editor.warnings)
    return editor.to_bytes(), summary


def split_by_action(items: list[DiffItem]):
    adds = [item for item in items if item.action == ACTION_ADD]
    removes = [item for item in items if item.action == ACTION_REMOVE]
    return adds, removes
