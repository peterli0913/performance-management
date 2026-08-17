import io
import zipfile

import openpyxl
import pytest

from tj4tools.bonus_export import build_workbook
from tj4tools.roster import ADD_CATEGORIES, CATEGORY_LEFT, CATEGORY_NEW

SHEET = "一线人员"


def _mapped(result, categories=None):
    """取有明确车间的人员，避免未映射项影响断言。"""
    items = [i for i in result.items if i.action == "add" and i.workshop]
    if categories:
        items = [i for i in items if i.category in categories]
    return items


@pytest.fixture(scope="module")
def marked(bonus_bytes, bonus, result):
    adds = _mapped(result)
    removes = [i for i in result.items if i.action == "remove"]
    data, summary = build_workbook(bonus_bytes, bonus, adds, removes, mode="mark")
    return data, summary, adds, removes


def test_no_parts_lost_except_calc_chain(bonus_bytes, marked):
    data = marked[0]
    before = set(zipfile.ZipFile(io.BytesIO(bonus_bytes)).namelist())
    after = set(zipfile.ZipFile(io.BytesIO(data)).namelist())
    assert before - after == {"xl/calcChain.xml"}
    assert after - before == set()
    # 外部链接必须保留，否则副主任表的 VLOOKUP 全废
    assert "xl/externalLinks/externalLink1.xml" in after
    assert "xl/printerSettings/printerSettings1.bin" in after
    assert "xl/customProperty1.bin" in after
    assert "xl/comments1.xml" in after


def test_unmodified_parts_are_byte_identical(bonus_bytes, marked):
    source = zipfile.ZipFile(io.BytesIO(bonus_bytes))
    output = zipfile.ZipFile(io.BytesIO(marked[0]))
    changed = {"xl/worksheets/sheet2.xml", "xl/styles.xml", "xl/workbook.xml",
               "[Content_Types].xml", "xl/_rels/workbook.xml.rels"}
    for name in source.namelist():
        if name in changed or name == "xl/calcChain.xml":
            continue
        assert source.read(name) == output.read(name), name


def test_row_count_and_values(marked):
    data, summary, adds, removes = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    assert sheet.max_row == 676 + len(adds)
    assert summary.added == len(adds)
    assert summary.removed == len(removes)
    names = {
        (str(sheet.cell(row, 3).value or ""), str(sheet.cell(row, 4).value or ""))
        for row in range(3, sheet.max_row + 1)
    }
    for item in adds:
        assert (item.name, item.eid) in names, item


def test_inserted_rows_sit_at_bottom_of_their_workshop(bonus, marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    merged = {}
    for merge in sheet.merged_cells.ranges:
        if merge.min_col == 1 == merge.max_col:
            label = sheet.cell(merge.min_row, 1).value
            if label:
                merged[str(label).strip()] = (merge.min_row, merge.max_row)
    per_workshop = {}
    for item in adds:
        per_workshop.setdefault(item.workshop, []).append(item)
    for workshop, items in per_workshop.items():
        start, end = merged[workshop]
        # 合并区终点必须顺延，新增人员正好占据块尾
        tail = [
            (str(sheet.cell(row, 3).value), str(sheet.cell(row, 4).value))
            for row in range(end - len(items) + 1, end + 1)
        ]
        assert tail == [(i.name, i.eid) for i in items], workshop
        assert start == merged[workshop][0]


def test_inserted_row_formulas_reference_their_own_row(marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    sheet = workbook[SHEET]
    checked = 0
    for row in range(3, sheet.max_row + 1):
        key = (str(sheet.cell(row, 3).value or ""), str(sheet.cell(row, 4).value or ""))
        if key not in {(i.name, i.eid) for i in adds}:
            continue
        assert sheet.cell(row, 6).value == f"=(K{row}+G{row}+H{row}+I{row}-J{row})"
        assert f"C{row}" in sheet.cell(row, 7).value
        assert f"L{row}:AP{row}" in sheet.cell(row, 11).value
        checked += 1
    assert checked == len(adds)


def test_existing_row_formulas_follow_their_new_row(marked):
    data, _, _, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    sheet = workbook[SHEET]
    for row in range(3, sheet.max_row):
        formula = sheet.cell(row, 6).value
        if isinstance(formula, str) and formula.startswith("="):
            assert formula == f"=(K{row}+G{row}+H{row}+I{row}-J{row})", row


def test_workshop_blocks_grow_by_their_own_additions(bonus, marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    blocks = sorted(
        (merge.min_row, merge.max_row)
        for merge in sheet.merged_cells.ranges
        if merge.min_col == 1 == merge.max_col and merge.min_row >= 3
    )
    assert len(blocks) == len(bonus.blocks)
    per_workshop = {}
    for item in adds:
        per_workshop[item.workshop] = per_workshop.get(item.workshop, 0) + 1
    for (start, end), (workshop, old_start, old_end) in zip(blocks, bonus.blocks):
        assert str(sheet.cell(start, 1).value).strip() == workshop
        expected = (old_end - old_start + 1) + per_workshop.get(workshop, 0)
        assert end - start + 1 == expected, workshop


def test_block_statistics_ranges_cover_new_rows(bonus, marked):
    """AQ/AR/AS 的 a/b/c 分布按块统计，区间必须把落进该区间的新增人员包进来。

    注意原表 AQ525 的区间是 $L$523:$AP$630，横跨 DCS/计算机化/培训 三个车间块，
    所以断言写成"按原区间覆盖的车间累加新增人数"，而不是假设区间等于单个块。
    """
    import re

    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    sheet = workbook[SHEET]

    per_workshop = {}
    for item in adds:
        per_workshop[item.workshop] = per_workshop.get(item.workshop, 0) + 1
    # 原表每个车间块末行收到的新增人数
    added_at = {end: per_workshop.get(name, 0) for name, _, end in bonus.blocks}
    ends = sorted(added_at)

    def shift(row, inclusive):
        return row + sum(added_at[e] for e in ends if (e <= row if inclusive else e < row))

    original_ranges = [
        (3, 73), (74, 164), (165, 226), (227, 298), (299, 349), (350, 522),
        (523, 630), (624, 630), (631, 643), (644, 656), (657, 675),
    ]
    expected = {(shift(a, False), shift(b, True)) for a, b in original_ranges}

    pattern = re.compile(r"\$?L\$?(\d+):\$?AP\$?(\d+)")
    found = 0
    for row in range(1, sheet.max_row + 1):
        for column in (43, 44, 45):
            formula = sheet.cell(row, column).value
            if not isinstance(formula, str) or "COUNTIF" not in formula:
                continue
            for start, end in pattern.findall(formula):
                assert (int(start), int(end)) in expected, (row, column, formula)
                found += 1
    assert found == len(original_ranges) * 3


def test_total_row_sum_extends(marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    sheet = workbook[SHEET]
    last = sheet.max_row
    assert sheet.cell(last, 6).value == f"=SUM(F3:J{last - 1})"


def test_autofilter_and_defined_names_updated(marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    assert sheet.auto_filter.ref == f"A1:AU{sheet.max_row}"
    text = zipfile.ZipFile(io.BytesIO(data)).read("xl/workbook.xml").decode("utf-8")
    assert f"一线人员!$A$1:$AU${sheet.max_row}" in text
    assert f"一线人员!$A$1:$F${sheet.max_row - 1}" in text
    assert 'fullCalcOnLoad="1"' in text


def test_new_rows_are_green_and_removed_rows_red(marked):
    data, _, adds, removes = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    add_keys = {(i.name, i.eid) for i in adds}
    remove_keys = {(i.name, i.eid) for i in removes}
    greens = reds = 0
    for row in range(3, sheet.max_row + 1):
        name_cell = sheet.cell(row, 3)
        eid_cell = sheet.cell(row, 4)
        key = (str(name_cell.value or ""), str(eid_cell.value or ""))
        colors = {name_cell.fill.fgColor.rgb, eid_cell.fill.fgColor.rgb}
        if key in add_keys:
            assert colors == {"FF92D050"}, key
            greens += 1
        elif key in remove_keys or _strip(key) in remove_keys:
            assert colors == {"FFFF0000"}, key
            reds += 1
    assert greens == len(adds)
    assert reds == len(removes)


def _strip(key):
    from tj4tools.normalize import norm_eid, norm_name

    return (norm_name(key[0]), norm_eid(key[1]))


def test_formatting_of_inserted_rows_matches_template(marked):
    data, _, adds, _ = marked
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    reference = sheet.cell(4, 2)
    add_keys = {(i.name, i.eid) for i in adds}
    for row in range(3, sheet.max_row + 1):
        key = (str(sheet.cell(row, 3).value or ""), str(sheet.cell(row, 4).value or ""))
        if key not in add_keys:
            continue
        cell = sheet.cell(row, 2)
        assert cell.font.name == reference.font.name
        assert cell.font.sz == reference.font.sz
        assert cell.alignment.horizontal == reference.alignment.horizontal
        assert sheet.row_dimensions[row].height == sheet.row_dimensions[4].height
        # 入职日期应写成真正的日期而不是文本
        assert sheet.cell(row, 5).is_date
        assert sheet.cell(row, 5).number_format == sheet.cell(4, 5).number_format


def test_other_sheets_untouched(bonus_bytes, marked):
    before = openpyxl.load_workbook(io.BytesIO(bonus_bytes), data_only=False)
    after = openpyxl.load_workbook(io.BytesIO(marked[0]), data_only=False)
    assert before.sheetnames == after.sheetnames
    for name in before.sheetnames:
        if name == SHEET:
            continue
        source, target = before[name], after[name]
        assert source.max_row == target.max_row, name
        assert source.sheet_state == target.sheet_state, name
        assert source.auto_filter.ref == target.auto_filter.ref, name
        for row in source.iter_rows():
            for cell in row:
                assert target[cell.coordinate].value == cell.value, (name, cell.coordinate)


def test_apply_mode_deletes_and_inserts_without_color(bonus_bytes, bonus, result):
    adds = _mapped(result, ADD_CATEGORIES)[:5]
    removes = [i for i in result.items if i.category == CATEGORY_LEFT][:4]
    data, summary = build_workbook(bonus_bytes, bonus, adds, removes, mode="apply")
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    assert summary.added == len(adds)
    assert summary.removed == len(removes)
    assert sheet.max_row == 676 + len(adds) - len(removes)
    present = {
        (str(sheet.cell(row, 3).value or ""), str(sheet.cell(row, 4).value or ""))
        for row in range(3, sheet.max_row + 1)
    }
    for item in removes:
        assert (item.name, item.eid) not in present
    for item in adds:
        assert (item.name, item.eid) in present
        row = next(r for r in range(3, sheet.max_row + 1)
                   if str(sheet.cell(r, 3).value or "") == item.name)
        assert sheet.cell(row, 3).fill.fgColor.rgb not in ("FF92D050", "FFFF0000")
        assert sheet.cell(row, 4).fill.fgColor.rgb not in ("FF92D050", "FFFF0000")


def test_apply_mode_keeps_formulas_consistent(bonus_bytes, bonus, result):
    adds = _mapped(result, [CATEGORY_NEW])[:6]
    removes = [i for i in result.items if i.category == CATEGORY_LEFT][:6]
    data, _ = build_workbook(bonus_bytes, bonus, adds, removes, mode="apply")
    workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    sheet = workbook[SHEET]
    for row in range(3, sheet.max_row):
        formula = sheet.cell(row, 6).value
        if isinstance(formula, str) and formula.startswith("="):
            assert formula == f"=(K{row}+G{row}+H{row}+I{row}-J{row})", row
    merged = [m for m in sheet.merged_cells.ranges if m.min_col == 1 == m.max_col and m.min_row >= 3]
    merged.sort(key=lambda m: m.min_row)
    assert merged[0].min_row == 3
    for previous, current in zip(merged, merged[1:]):
        assert current.min_row == previous.max_row + 1


def test_new_workshop_block_is_created(bonus_bytes, bonus, result):
    items = [i for i in result.items if i.action == "add"][:3]
    for item in items:
        item.workshop = "13号楼"
    data, summary = build_workbook(bonus_bytes, bonus, items, [], mode="mark")
    assert summary.new_blocks == ["13号楼"]
    assert any("新建分组块" in w for w in summary.warnings)
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    sheet = workbook[SHEET]
    # 新块紧跟在最后一个数据行之后，原本的合计行被顺延
    assert sheet.cell(676, 1).value == "13号楼"
    assert any(str(m) == "A676:A678" for m in sheet.merged_cells.ranges)
    # 原最后一个车间块不能被新块吞掉
    assert any(str(m) == "A657:A675" for m in sheet.merged_cells.ranges)
    assert sheet.cell(679, 6).value == "=SUM(F3:J678)"


def test_items_without_workshop_are_skipped(bonus_bytes, bonus, result):
    orphans = [i for i in result.items if i.action == "add" and not i.workshop]
    if not orphans:
        pytest.skip("样本里没有未映射人员")
    data, summary = build_workbook(bonus_bytes, bonus, orphans[:3], [], mode="mark")
    assert summary.added == 0
    assert len(summary.skipped) == 3
    workbook = openpyxl.load_workbook(io.BytesIO(data))
    assert workbook[SHEET].max_row == 676
