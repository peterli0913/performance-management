import datetime as dt

from tj4tools.roster import (
    ADD_CATEGORIES,
    CATEGORIES,
    CATEGORY_LEFT,
    CATEGORY_NEW,
    CATEGORY_PENDING_ADD,
    CATEGORY_PENDING_DEL,
    build_workshop_mapping,
    reconcile,
)


def test_roster_parsing_shape(roster):
    assert roster.ref_date == dt.date(2026, 7, 31)
    # 生产部共 1151 人，按"职位"过滤出助工/操作工/工程师/班长 935 人（实习生被排除）
    assert len(roster.production_all) == 1151
    assert len(roster.production) == 935
    assert all(p.duty in {"助工", "操作工", "工程师", "班长"} for p in roster.production.values())
    assert len(roster.departures) > 800
    assert len(roster.changes) == 41


def test_roster_duty_field_switch(roster_bytes):
    from tj4tools.roster import parse_roster

    by_post = parse_roster(roster_bytes, "x，07-31-2026.xlsx", duty_field="岗位")
    # 按"岗位"过滤会把实习生（岗位记作助理工程师）也算进来
    assert len(by_post.production) == 980


def test_bonus_parsing_shape(bonus):
    assert len(bonus.frontline) == 673
    assert bonus.first_data_row == 3
    assert bonus.last_data_row == 675
    assert bonus.columns == {"workshop": "A", "duty": "B", "name": "C", "eid": "D", "hire": "E"}
    assert [b[0] for b in bonus.blocks][:3] == ["4号楼1&2车间", "4号楼3&4车间", "11号楼CNC车间"]
    # 车间块必须首尾相接且覆盖全部数据行
    assert bonus.blocks[0][1] == 3
    assert bonus.blocks[-1][2] == 675
    for previous, current in zip(bonus.blocks, bonus.blocks[1:]):
        assert current[1] == previous[2] + 1
    assert len(bonus.others) == 139


def test_merged_workshop_is_filled_down(bonus):
    # 一线人员 A 列是纵向合并的，未填充会导致车间全丢
    assert all(person.workshop for person in bonus.frontline.values())
    assert bonus.frontline[("陈玉慧", "ALS12990")].workshop == "4号楼1&2车间"
    assert bonus.frontline[("李硕", "ALS17098")].workshop == "清洗组"


def test_conservation_of_headcount(result):
    counts = result.counts
    assert set(counts) == set(CATEGORIES)
    assert sum(counts.values()) == (result.only_roster - result.excluded_in_others) + result.only_bonus
    assert result.matched + result.only_roster == 935
    assert result.matched + result.only_bonus == 673


def test_known_people_land_in_right_category(result):
    # 变动说明里"入职"且离职时间为空
    assert result.category_of("周圣玮", "ALS18092") == CATEGORY_NEW
    # 离职表里有明确离职时间
    assert result.category_of("王子睿", "ALS15643") == CATEGORY_LEFT
    # 核算表有、清单查不到离职记录（实际是升了工艺组长）
    assert result.category_of("胡强", "ALS10148") == CATEGORY_PENDING_DEL
    # 入职 2025-08-20，早于一个月窗口且变动说明无记录
    assert result.category_of("曹睿晟", "ALS14679") == CATEGORY_PENDING_ADD
    # 已在「副主任&工艺组长及其他」子表，默认排除
    assert result.category_of("赵凯", "ALS11919") is None


def test_new_hire_window_is_one_month(result):
    assert result.window_start == dt.date(2026, 6, 30)
    for item in result.by_category(CATEGORY_NEW):
        in_window = item.hire_date is not None and result.window_start <= item.hire_date <= result.ref_date
        has_change_row = "人员变动说明" in item.reason and "离职时间为空" in item.reason
        assert in_window or has_change_row, item


def test_pending_add_is_outside_window(result):
    for item in result.by_category(CATEGORY_PENDING_ADD):
        if item.hire_date is not None:
            assert not (result.window_start <= item.hire_date <= result.ref_date)


def test_left_people_have_evidence(result):
    for item in result.by_category(CATEGORY_LEFT):
        assert "离职人员&调出人员" in item.reason or "有离职时间" in item.reason


def test_pending_delete_has_no_evidence(result):
    for item in result.by_category(CATEGORY_PENDING_DEL):
        assert "都查不到离职记录" in item.reason


def test_every_item_carries_display_fields(result):
    for item in result.items:
        payload = item.as_dict()
        assert payload["姓名"] and payload["员工编号"]
        assert payload["分类"] in CATEGORIES
        assert payload["判定依据"]
        if item.action == "remove":
            assert item.frontline_row >= 3
        else:
            assert item.roster_row >= 2


def test_promotion_case_is_flagged(result):
    item = next(i for i in result.items if i.key == ("胡强", "ALS10148"))
    assert any("仍在「生产部」" in flag for flag in item.flags)


def test_workshop_mapping_prefers_empirical_evidence(roster, bonus):
    mapping = build_workshop_mapping(roster, bonus)
    assert mapping["4号楼1&2车间"].workshop == "4号楼1&2车间"
    assert mapping["11号楼CNC区域"].workshop == "11号楼CNC车间"
    assert mapping["12号楼CNC区域"].workshop == "12号楼"
    assert mapping["验证组"].workshop == "计算机化设备保障&验证组"
    assert mapping["外围/罐区组"].workshop == "外围/罐区/泵房"
    assert mapping["清洗组(4号楼)"].workshop == "清洗组"
    for group in ("4号楼1&2车间", "11号楼CNC区域", "12号楼CNC区域", "5号楼"):
        assert mapping[group].confidence == "高"
        assert mapping[group].source == "经验"


def test_unmappable_groups_are_reported(result):
    assert any("无法自动对应车间" in note for note in result.notes)
    unmapped = [item for item in result.items if item.action == "add" and not item.workshop]
    assert unmapped, "应当存在需要人工指定车间的人员"
    assert all("生产技术转移组" in i.group or "在其他厂区" in i.group or "安全组" in i.group
               or "计算机化设备保障组" == i.group for i in unmapped)


def test_exclude_in_others_can_be_disabled(roster, bonus):
    kept = reconcile(roster, bonus, exclude_in_others=False)
    assert kept.excluded_in_others == 0
    assert sum(kept.counts.values()) == kept.only_roster + kept.only_bonus
    item = next(i for i in kept.items if i.key == ("赵凯", "ALS11919"))
    assert any("副主任" in flag for flag in item.flags)


def test_window_months_changes_split(roster, bonus):
    wide = reconcile(roster, bonus, window_months=6)
    narrow = reconcile(roster, bonus, window_months=1)
    assert wide.counts[CATEGORY_NEW] > narrow.counts[CATEGORY_NEW]
    assert sum(wide.counts.values()) == sum(narrow.counts.values())


def test_add_items_are_only_from_roster(result):
    for item in result.items:
        if item.category in ADD_CATEGORIES:
            assert item.frontline_row == 0
        else:
            assert item.roster_row == 0
