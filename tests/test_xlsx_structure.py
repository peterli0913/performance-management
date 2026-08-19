"""对生成的 xlsx 做结构性体检——Excel 报"需要修复"通常就是这几类问题。"""

import io
import re
import xml.etree.ElementTree as ET
import zipfile

import pytest

from tj4tools.bonus_export import build_workbook
from tj4tools.xlsx_surgery import MAIN_NS, col_to_index, split_coord

Q = lambda tag: f"{{{MAIN_NS}}}{tag}"  # noqa: E731


@pytest.fixture(scope="module")
def generated(bonus_bytes, bonus, result):
    adds = [i for i in result.items if i.action == "add" and i.workshop]
    removes = [i for i in result.items if i.action == "remove"]
    updates = [i for i in result.items if i.action == "update"]
    mark, _ = build_workbook(bonus_bytes, bonus, adds, removes, updates, mode="mark")
    apply_, _ = build_workbook(
        bonus_bytes, bonus, adds[:20], removes[:10], updates[:5], mode="apply"
    )
    return {"mark": mark, "apply": apply_}


def _sheet_xml(data, path="xl/worksheets/sheet2.xml"):
    return ET.fromstring(zipfile.ZipFile(io.BytesIO(data)).read(path))


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_rows_are_unique_and_ascending(generated, mode):
    sheet_data = _sheet_xml(generated[mode]).find(Q("sheetData"))
    rows = [int(row.get("r")) for row in sheet_data.findall(Q("row"))]
    assert rows == sorted(rows)
    assert len(rows) == len(set(rows))


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_cell_refs_match_their_row_and_ascend(generated, mode):
    sheet_data = _sheet_xml(generated[mode]).find(Q("sheetData"))
    for row in sheet_data.findall(Q("row")):
        expected = int(row.get("r"))
        previous = 0
        for cell in row.findall(Q("c")):
            column, row_index = split_coord(cell.get("r"))
            assert row_index == expected, cell.get("r")
            index = col_to_index(column)
            assert index > previous, cell.get("r")
            previous = index


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_no_shared_formula_leftovers(generated, mode):
    """共享公式必须全部展开；残留 si 会在行号变化后指向错误的主单元格。"""
    text = zipfile.ZipFile(io.BytesIO(generated[mode])).read("xl/worksheets/sheet2.xml").decode()
    assert 't="shared"' not in text
    assert "<f/>" not in text
    assert "<f />" not in text


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_no_formula_references_out_of_range(generated, mode):
    root = _sheet_xml(generated[mode])
    last = max(int(row.get("r")) for row in root.find(Q("sheetData")).findall(Q("row")))
    pattern = re.compile(r"(?<![A-Za-z0-9_$.!])\$?[A-Z]{1,3}\$?(\d{1,7})(?![\d(])")
    for row in root.find(Q("sheetData")).findall(Q("row")):
        for cell in row.findall(Q("c")):
            formula = cell.find(Q("f"))
            if formula is None or not formula.text:
                continue
            for number in pattern.findall(formula.text):
                assert 1 <= int(number) <= max(last, 1048576)


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_merges_do_not_overlap(generated, mode):
    root = _sheet_xml(generated[mode])
    ranges = []
    for merge in root.find(Q("mergeCells")).findall(Q("mergeCell")):
        left, right = merge.get("ref").split(":")
        c1, r1 = split_coord(left)
        c2, r2 = split_coord(right)
        assert r1 <= r2 and col_to_index(c1) <= col_to_index(c2)
        ranges.append((col_to_index(c1), r1, col_to_index(c2), r2))
    for index, first in enumerate(ranges):
        for second in ranges[index + 1 :]:
            overlap_col = first[0] <= second[2] and second[0] <= first[2]
            overlap_row = first[1] <= second[3] and second[1] <= first[3]
            assert not (overlap_col and overlap_row), (first, second)


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_merge_count_attribute_is_correct(generated, mode):
    merge = _sheet_xml(generated[mode]).find(Q("mergeCells"))
    assert int(merge.get("count")) == len(merge.findall(Q("mergeCell")))


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_shared_strings_untouched(bonus_bytes, generated, mode):
    """新值用 inlineStr 写入，sharedStrings 不应该被改动。"""
    before = zipfile.ZipFile(io.BytesIO(bonus_bytes)).read("xl/sharedStrings.xml")
    after = zipfile.ZipFile(io.BytesIO(generated[mode])).read("xl/sharedStrings.xml")
    assert before == after


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_styles_only_grow(bonus_bytes, generated, mode):
    before = zipfile.ZipFile(io.BytesIO(bonus_bytes)).read("xl/styles.xml").decode()
    after = zipfile.ZipFile(io.BytesIO(generated[mode])).read("xl/styles.xml").decode()
    counts = lambda text, tag: int(re.search(rf"<{tag} count=\"(\d+)\"", text).group(1))  # noqa: E731
    # 这几张表不能动：dxfs 被条件格式按索引引用，动了就串色
    for tag in ("borders", "dxfs", "cellStyleXfs"):
        assert counts(before, tag) == counts(after, tag), tag
    # 填充色、字体（红字）、cellXfs 只允许追加
    for tag in ("cellXfs", "fills", "fonts"):
        assert counts(after, tag) >= counts(before, tag), tag
    # 声明的 count 必须和实际元素数一致
    root = ET.fromstring(after.encode())
    for tag in ("fills", "fonts", "cellXfs"):
        element = root.find(Q(tag))
        assert int(element.get("count")) == len(list(element))


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_calc_chain_removed_and_full_calc_requested(generated, mode):
    archive = zipfile.ZipFile(io.BytesIO(generated[mode]))
    assert "xl/calcChain.xml" not in archive.namelist()
    assert b"calcChain" not in archive.read("[Content_Types].xml")
    assert b"calcChain" not in archive.read("xl/_rels/workbook.xml.rels")
    assert b'fullCalcOnLoad="1"' in archive.read("xl/workbook.xml")


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_all_parts_are_well_formed_xml(generated, mode):
    archive = zipfile.ZipFile(io.BytesIO(generated[mode]))
    for name in archive.namelist():
        if not name.endswith((".xml", ".rels")):
            continue
        ET.fromstring(archive.read(name))


@pytest.mark.parametrize("mode", ["mark", "apply"])
def test_namespace_declarations_preserved(bonus_bytes, generated, mode):
    for path in ("xl/workbook.xml", "xl/worksheets/sheet2.xml"):
        before = zipfile.ZipFile(io.BytesIO(bonus_bytes)).read(path).decode()
        after = zipfile.ZipFile(io.BytesIO(generated[mode])).read(path).decode()
        pattern = re.compile(r'xmlns(?::[\w.-]+)?="[^"]+"')
        assert set(pattern.findall(before.split(">", 2)[1])) <= set(pattern.findall(after))
