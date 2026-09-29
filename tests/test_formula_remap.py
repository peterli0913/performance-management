from tj4tools.xlsx_surgery import (
    RowMap,
    remap_formula,
    remap_sqref,
    translate_formula,
)

SHEET = "一线人员"


def make_map(deletes=(), extend=None, blocks=None, threshold=3):
    return RowMap(set(deletes), extend or {}, blocks or {}, threshold)


def test_same_row_formula_shifts_with_row():
    rowmap = make_map(extend={73: 2})
    assert remap_formula("=(K100+G100+H100+I100-J100)", rowmap, SHEET) == "=(K102+G102+H102+I102-J102)"


def test_other_sheet_and_external_refs_untouched():
    rowmap = make_map(extend={73: 2})
    formula = "=COUNTIF(L100:AP100,标准数据!$A$2)*标准数据!$B$2"
    assert remap_formula(formula, rowmap, SHEET) == "=COUNTIF(L102:AP102,标准数据!$A$2)*标准数据!$B$2"
    external = "=IFERROR(VLOOKUP(C100,[1]生产部!$C:$C,1,0),\"\")"
    assert remap_formula(external, rowmap, SHEET) == "=IFERROR(VLOOKUP(C102,[1]生产部!$C:$C,1,0),\"\")"


def test_block_range_end_extends_at_anchor():
    # 车间块 L3:AP73 末尾追加 2 人后，统计区间必须延伸到 75
    rowmap = make_map(extend={73: 2})
    assert remap_formula('=COUNTIF($L$3:$AP$73,"a")', rowmap, SHEET) == '=COUNTIF($L$3:$AP$75,"a")'


def test_block_local_range_ignores_new_block_rows():
    # 在锚点 675 新建车间块时，块内区间只顺延本块新增，跨全表区间连新块一起顺延
    rowmap = make_map(extend={675: 1}, blocks={675: 3})
    assert remap_formula('=COUNTIF($L$657:$AP$675,"a")', rowmap, SHEET) == '=COUNTIF($L$657:$AP$676,"a")'
    assert remap_formula("=SUM(F3:J675)", rowmap, SHEET) == "=SUM(F3:J679)"


def test_function_names_are_not_mistaken_for_refs():
    rowmap = make_map(extend={1: 5})
    for formula in ("=LOG10(A100)", "=SUM(A100)/COUNT(A100)"):
        remapped = remap_formula(formula, rowmap, SHEET)
        assert "LOG105" not in remapped
        assert remapped.count("A105") >= 1


def test_string_literals_are_not_rewritten():
    rowmap = make_map(extend={1: 5})
    assert remap_formula('=IF(A10="B10","B10",A10)', rowmap, SHEET) == '=IF(A15="B10","B10",A15)'


def test_deleted_row_clamps_range():
    rowmap = make_map(deletes={50, 51})
    # 50、51 被删；区间 A49:A51 的终点回退到最后一个存活行
    assert remap_formula("=SUM(A49:A51)", rowmap, SHEET) == "=SUM(A49:A49)"


def test_self_sheet_qualified_ref_is_remapped():
    rowmap = make_map(extend={10: 3})
    assert remap_formula("=一线人员!$C$20", rowmap, SHEET) == "=一线人员!$C$23"


def test_translate_formula_shifts_relative_only():
    assert translate_formula("=(K3+G3-$J$3)", 5) == "=(K8+G8-$J$3)"
    assert translate_formula("=COUNTIF(L3:AP3,标准数据!$A$2)", 2) == "=COUNTIF(L5:AP5,标准数据!$A$2)"


def test_remap_sqref_multi_range():
    rowmap = make_map(extend={73: 2})
    assert remap_sqref("A3:A73 D80 D100:D110", rowmap) == "A3:A75 D82 D102:D112"
