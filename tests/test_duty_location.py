"""09 月起：核算表 A 列（车间）要跟人员清单 J 列「实际履职属地」对应。

清单表头改成了「二级部门 / 三级分组 / 实际履职属地」，没有「目前分组」了；
「三级分组」写的是 `CNC区域`、`D级区域` 这种跨厂房重名的值，不能拿来落车间。
"""

import io
from collections import Counter
from dataclasses import replace

import openpyxl

from tj4tools.roster import (
    ROUTE_TO_OTHERS,
    build_workshop_mapping,
    find_relocations,
    placeable_keys,
    reconcile,
    resolved_group_workshop,
)

FRONTLINE = "一线人员"
DONG = ("董浩波", "ALS12806")  # 一线「多肽厂房1D级区域」助工，第 287 行


def _moved(roster, key, group):
    """复制一份清单，把某人的实际履职属地改掉；不碰 session 级 fixture。"""
    production_all = dict(roster.production_all)
    production = dict(roster.production)
    production_all[key] = replace(production_all[key], group=group)
    if key in production:
        production[key] = replace(production[key], group=group)
    return replace(roster, production_all=production_all, production=production)


def test_group_comes_from_duty_location_column(sep_roster):
    assert sep_roster.group_field == "实际履职属地"
    groups = Counter(person.group for person in sep_roster.production_all.values())
    # 三级分组的写法不能混进来
    assert "CNC区域" not in groups
    assert "D级区域" not in groups
    assert groups["多肽厂房2CNC区域"] == 171
    assert groups["寡核苷酸厂房1D级区域"] == 101


def test_old_roster_still_uses_current_group(roster):
    assert roster.group_field == "目前分组"


def test_existing_people_sit_in_the_block_named_by_their_location(sep_roster, sep_bonus):
    """黄金样本里已在一线的人，A 列与 J 列一致（仅 寡核苷酸厂房1 → 制备组 一处别名）。"""
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    mismatched = Counter()
    for key, person in sep_bonus.frontline.items():
        source = sep_roster.production_all.get(key)
        if source is None:
            continue
        if resolved_group_workshop(source.group, mapping) != person.workshop:
            mismatched[(source.group, person.workshop)] += 1
    assert not mismatched, mismatched.most_common(5)
    assert mapping["寡核苷酸厂房1"].workshop == "寡核苷酸厂房1制备组"


def test_exact_block_names_are_not_held_even_with_parentheses(sep_roster, sep_bonus):
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    for group in (
        "多肽厂房2(清洗组)",
        "HP厂房1(清洗组)",
        "计算机化设备保障组(多肽厂房1&多肽厂房2)",
        "多肽厂房1D级区域",
    ):
        guess = mapping[group]
        assert guess.workshop == group, group
        assert not guess.needs_manual, group
        assert not guess.route_others, group


def test_new_hires_land_in_the_block_matching_their_location(sep_roster, sep_bonus):
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    result = reconcile(sep_roster, sep_bonus, mapping=mapping)
    adds = [item for item in result.items if item.action == "add"]
    assert adds
    workshops = set(sep_bonus.workshops)
    for item in adds:
        if item.group in workshops:
            assert item.workshop == item.group, (item.name, item.group, item.workshop)
    # 旧逻辑按「三级分组」把 多肽厂房3 的人推到了 多肽厂房1D级区域
    assert not any(i.group == "多肽厂房3" and i.workshop == "多肽厂房1D级区域" for i in adds)


def test_locations_only_on_others_sheet_default_to_others(sep_roster, sep_bonus):
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    # 副主任表有同名车间块，一线只有「多肽厂房3(清洗组)」
    assert resolved_group_workshop("多肽厂房3", mapping) == ROUTE_TO_OTHERS
    # 副主任表「多肽厂房1项目组(5104)」里的人 J 列都写「多肽厂房1」
    assert resolved_group_workshop("多肽厂房1", mapping) == ROUTE_TO_OTHERS

    result = reconcile(sep_roster, sep_bonus, mapping=mapping)
    keys = placeable_keys(result)
    routed = [i for i in result.items if i.action == "add" and i.group in ("多肽厂房3", "多肽厂房1")]
    assert routed
    assert all(item.key not in keys for item in routed)


def test_ambiguous_building_name_is_left_for_manual_choice(sep_roster, sep_bonus):
    """「多肽厂房2」同时像 CNC区域 / D级区域 / 中试 / 清洗组，不能随便挑一个。"""
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    assert resolved_group_workshop("多肽厂房2", mapping) == ""


# --------------------------------------------------------------------------- #
# 已在一线的人：履职属地变了，车间块要跟着换
# --------------------------------------------------------------------------- #


def test_golden_pair_needs_no_relocation(sep_roster, sep_bonus):
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    assert find_relocations(sep_roster, sep_bonus, mapping) == []


def test_changed_location_is_detected(sep_roster, sep_bonus):
    roster = _moved(sep_roster, DONG, "HP厂房1")
    mapping = build_workshop_mapping(roster, sep_bonus)
    moves = find_relocations(roster, sep_bonus, mapping)
    assert [(i.key, i.frontline_row, i.workshop) for i in moves] == [(DONG, 287, "HP厂房1")]
    item = moves[0]
    assert item.action == "relocate"
    assert "多肽厂房1D级区域" in item.reason and "HP厂房1" in item.reason


def test_mapping_override_for_new_hires_does_not_move_existing_people(sep_roster, sep_bonus):
    """车间映射里的人工覆盖只管新增人员；已在岗的人只按同名车间块/已匹配人员挪。"""
    import inspect

    assert "group_overrides" not in inspect.signature(find_relocations).parameters


def test_september_workshop_column_is_not_merged(sep_bonus, bonus):
    assert sep_bonus.merged_workshop is False
    assert bonus.merged_workshop is True


def test_new_rows_write_workshop_name_when_column_is_not_merged(
    sep_roster, sep_bonus, sep_bonus_bytes
):
    """09 月表 A 列每行都写车间名，新增行也必须写，否则 A 列就对不上 J 列了。"""
    from tj4tools.bonus_export import build_workbook

    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    result = reconcile(sep_roster, sep_bonus, mapping=mapping)
    adds = [i for i in result.items if i.action == "add" and i.workshop]
    new_block = replace(adds[0], key=("测试甲", "T0001"), name="测试甲", eid="T0001", workshop="测试新车间")
    data, summary = build_workbook(
        sep_bonus_bytes, sep_bonus, adds + [new_block], [], [], [], mode="apply"
    )
    sheet = openpyxl.load_workbook(io.BytesIO(data))[FRONTLINE]
    written = {
        str(sheet.cell(r, 4).value): str(sheet.cell(r, 1).value or "")
        for r in range(4, sheet.max_row + 1)
    }
    from tj4tools.normalize import clean_text

    for item in adds:
        assert clean_text(written[item.eid]) == item.workshop, (item.name, item.workshop)
    assert written["T0001"] == "测试新车间"
    assert [str(m) for m in sheet.merged_cells.ranges if m.min_col == 1] == ["A2:A3"]
    assert summary.new_blocks == ["测试新车间"]


def test_september_full_export_keeps_every_row_labelled_and_blocks_contiguous(
    sep_roster, sep_bonus, sep_bonus_bytes
):
    from tj4tools.bonus_export import build_workbook
    from tj4tools.normalize import clean_text

    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    result = reconcile(sep_roster, sep_bonus, mapping=mapping)
    adds = [i for i in result.items if i.action == "add" and i.workshop]
    removes = [i for i in result.items if i.action == "remove"]
    data, summary = build_workbook(sep_bonus_bytes, sep_bonus, adds, removes, [], [], mode="apply")
    assert summary.added == len(adds) and not summary.new_blocks
    sheet = openpyxl.load_workbook(io.BytesIO(data))[FRONTLINE]
    sequence = []
    for row in range(sep_bonus.first_data_row, sheet.max_row + 1):
        if not sheet.cell(row, 3).value:
            continue
        label = clean_text(sheet.cell(row, 1).value)
        assert label, row
        sequence.append(label)
    assert len(sequence) == len(sep_bonus.frontline) + summary.added - summary.removed
    runs = [v for index, v in enumerate(sequence) if index == 0 or sequence[index - 1] != v]
    assert len(runs) == len(set(runs)), runs


def test_parenthesised_workshop_keeps_original_text(sep_roster, sep_bonus, sep_bonus_bytes):
    from tj4tools.bonus_export import build_workbook

    raw = openpyxl.load_workbook(io.BytesIO(sep_bonus_bytes), read_only=True)[FRONTLINE]
    originals = {
        str(row[0]).strip()
        for row in raw.iter_rows(min_row=4, max_col=1, values_only=True)
        if row[0] and "清洗组" in str(row[0])
    }
    mapping = build_workshop_mapping(sep_roster, sep_bonus)
    result = reconcile(sep_roster, sep_bonus, mapping=mapping)
    adds = [i for i in result.items if i.action == "add" and i.workshop == "多肽厂房2(清洗组)"]
    assert adds
    data, _ = build_workbook(sep_bonus_bytes, sep_bonus, adds, [], [], [], mode="apply")
    sheet = openpyxl.load_workbook(io.BytesIO(data))[FRONTLINE]
    eids = {i.eid for i in adds}
    labels = {str(sheet.cell(r, 1).value) for r in range(4, sheet.max_row + 1) if sheet.cell(r, 4).value in eids}
    assert len(labels) == 1 and labels <= originals


def test_old_roster_without_duty_location_is_not_relocated(roster, bonus):
    """07 月清单按「目前分组」，和车间本来就不是一一对应，不能据此挪人。"""
    mapping = build_workshop_mapping(roster, bonus)
    assert find_relocations(roster, bonus, mapping) == []


def test_relocation_export_moves_the_row_between_blocks(sep_roster, sep_bonus, sep_bonus_bytes):
    from tj4tools.bonus_export import build_workbook

    roster = _moved(sep_roster, DONG, "HP厂房1")
    mapping = build_workshop_mapping(roster, sep_bonus)
    moves = find_relocations(roster, sep_bonus, mapping)
    data, summary = build_workbook(
        sep_bonus_bytes, sep_bonus, [], [], [], [], mode="apply", relocations=moves
    )
    assert summary.relocated == 1
    assert summary.added == 0 and summary.removed == 0
    assert "按实际履职属地调整车间 1 人" in summary.text()

    sheet = openpyxl.load_workbook(io.BytesIO(data))[FRONTLINE]
    assert sheet.max_row == openpyxl.load_workbook(io.BytesIO(sep_bonus_bytes))[FRONTLINE].max_row
    rows = [r for r in range(4, sheet.max_row + 1) if sheet.cell(r, 4).value == DONG[1]]
    assert len(rows) == 1
    row = rows[0]
    block_of_row = {}
    for merge in sheet.merged_cells.ranges:
        if merge.min_col == 1 == merge.max_col:
            for r in range(merge.min_row, merge.max_row + 1):
                block_of_row[r] = sheet.cell(merge.min_row, 1).value
    assert block_of_row.get(row, sheet.cell(row, 1).value) == "HP厂房1"
    assert sheet.cell(row, 2).value == "助工"
    # 插在 HP厂房1 助工段末尾：下一行不再是助工（或已出块）
    assert sheet.cell(row - 1, 2).value == "助工"


def test_relocation_mark_mode_flags_both_rows(sep_roster, sep_bonus, sep_bonus_bytes):
    from tj4tools.bonus_export import build_workbook

    roster = _moved(sep_roster, DONG, "HP厂房1")
    mapping = build_workshop_mapping(roster, sep_bonus)
    moves = find_relocations(roster, sep_bonus, mapping)
    data, summary = build_workbook(
        sep_bonus_bytes, sep_bonus, [], [], [], [], mode="mark", relocations=moves
    )
    sheet = openpyxl.load_workbook(io.BytesIO(data))[FRONTLINE]
    rows = [r for r in range(4, sheet.max_row + 1) if sheet.cell(r, 4).value == DONG[1]]
    assert len(rows) == 2
    assert summary.relocated == 1
