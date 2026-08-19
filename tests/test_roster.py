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
    # 按「岗位」过滤时"职位写实习生"的人仍被排除（默认 include_interns=False）
    assert len(by_post.production) == 937
    assert not any(p.title == "实习生" for p in by_post.production.values())


def test_interns_are_excluded_by_default_and_can_be_included(roster_bytes):
    from tj4tools.roster import parse_roster

    default = parse_roster(roster_bytes, "x，07-31-2026.xlsx")
    assert len(default.production) == 935
    assert not any(p.title == "实习生" for p in default.production.values())

    with_interns = parse_roster(roster_bytes, "x，07-31-2026.xlsx", include_interns=True)
    by_title = [p for p in with_interns.production.values() if p.intern_source == "职位"]
    assert len(with_interns.production) == 935 + 43
    assert len(by_title) == 43
    # 实习生的「职位」是实习生、「岗位」是实际岗位，纳入时按岗位归入职务分组
    assert all(p.title == "实习生" for p in by_title)
    assert {p.duty for p in by_title} == {"操作工"}


def test_interns_can_also_be_detected_from_the_remark(roster_bytes):
    """有的实习生职位和岗位都写实际岗位，只有备注里写着「校招实习生」。"""
    from tj4tools.roster import parse_roster

    roster = parse_roster(roster_bytes, "x，07-31-2026.xlsx")
    by_remark = [p for p in roster.production_all.values() if p.intern_source == "备注"]
    assert len(by_remark) == 4
    assert all("实习" in p.remark for p in by_remark)
    assert all(p.title != "实习生" for p in by_remark)
    # 这些人的职位本来就是目标职务，所以不受「纳入实习生」开关影响
    assert all(p.key in roster.production for p in by_remark)
    assert {p.duty for p in by_remark} == {"助工", "操作工", "工程师"}


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
    # 只数到第 129 行的真实人员，不含第 130 行之后混排的参照表
    assert len(bonus.others) == 128


def test_merged_workshop_is_filled_down(bonus):
    # 一线人员 A 列是纵向合并的，未填充会导致车间全丢
    assert all(person.workshop for person in bonus.frontline.values())
    assert bonus.frontline[("陈玉慧", "ALS12990")].workshop == "4号楼1&2车间"
    assert bonus.frontline[("李硕", "ALS17098")].workshop == "清洗组"


def test_conservation_of_headcount(result):
    counts = result.counts
    assert set(counts) == set(CATEGORIES)
    assert sum(counts.values()) == (
        result.only_roster - result.excluded_in_others - result.paired_renames
    ) + result.only_bonus
    assert result.matched + result.only_roster == 935
    assert result.matched + result.only_bonus == 673


def test_same_eid_different_name_is_merged_not_double_counted(result):
    """清单叫「曹睿晟」、核算叫「曹静旺」，同一个编号——改名即可，不能又加又删。"""
    assert result.paired_renames == 1
    update = next(i for i in result.items if i.key == ("曹静旺", "ALS14679"))
    assert update.action == "update"
    assert update.new_values["姓名"] == "曹睿晟"
    assert any("按清单改名即可，不需要另外新增" in flag for flag in update.flags)
    # 新增那一侧必须已经被撤掉，否则导出会出现两行同一个人
    assert result.category_of("曹睿晟", "ALS14679") is None
    assert any("编号相同、姓名不同" in note for note in result.notes)


def test_known_people_land_in_right_category(result):
    # 变动说明里"入职"且离职时间为空
    assert result.category_of("周圣玮", "ALS18092") == CATEGORY_NEW
    # 离职表里有明确离职时间
    assert result.category_of("王子睿", "ALS15643") == CATEGORY_LEFT
    # 核算表有、清单查不到离职记录（实际是升了车间副主任）
    assert result.category_of("胡强", "ALS10148") == CATEGORY_PENDING_DEL
    # 已在「副主任&工艺组长及其他」子表，默认排除
    assert result.category_of("赵凯", "ALS11919") is None


def test_new_hire_since_defaults_to_one_month_before(result):
    assert result.new_hire_since == dt.date(2026, 6, 30)
    for item in result.by_category(CATEGORY_NEW):
        in_window = item.hire_date is not None and item.hire_date >= result.new_hire_since
        has_change_row = "人员变动说明" in item.reason and "离职时间为空" in item.reason
        assert in_window or has_change_row, item


def test_new_hire_since_can_be_given_explicitly(roster, bonus):
    early = reconcile(roster, bonus, new_hire_since=dt.date(2026, 1, 1))
    late = reconcile(roster, bonus, new_hire_since=dt.date(2026, 7, 20))
    assert early.new_hire_since == dt.date(2026, 1, 1)
    assert early.counts[CATEGORY_NEW] > late.counts[CATEGORY_NEW]
    assert sum(early.counts.values()) == sum(late.counts.values())
    for item in late.by_category(CATEGORY_NEW):
        if "人员变动说明" not in item.reason:
            assert item.hire_date >= dt.date(2026, 7, 20)


def test_pending_add_is_outside_window(result):
    for item in result.by_category(CATEGORY_PENDING_ADD):
        if item.hire_date is not None and "人员变动说明" not in item.reason:
            assert not (result.new_hire_since <= item.hire_date < result.ref_date)


def test_hires_on_or_after_reference_date_are_pending_not_new(result):
    """窗口有上界：入职时间不早于参照日期的算待定，并且要标出来。"""
    late = [i for i in result.items if i.hire_date and i.hire_date >= result.ref_date]
    assert late, "样本里应当有不早于参照日期入职的人"
    for item in late:
        assert item.category == CATEGORY_PENDING_ADD, item
        assert any("不早于参照日期" in flag for flag in item.flags)


def test_intern_classification(roster_with_interns, bonus):
    from tj4tools.roster import INTERN_JUNIOR, INTERN_SENIOR

    analysis = reconcile(roster_with_interns, bonus)
    interns = [i for i in analysis.items if i.is_intern]
    assert interns
    assert analysis.intern_asof == dt.date(2026, 7, 31)
    assert analysis.intern_months == 3
    cutoff = dt.date(2026, 4, 30)
    for item in interns:
        if item.intern_class == INTERN_SENIOR:
            assert item.hire_date <= cutoff, item
        elif item.intern_class == INTERN_JUNIOR:
            assert cutoff < item.hire_date < analysis.intern_asof, item
        assert any("实习生" in flag for flag in item.flags)
    assert sum(analysis.intern_counts.values()) == len(interns)
    # 判断日期往后推，原来"不到 3 个月"的会变成"超过 3 个月"
    later = reconcile(roster_with_interns, bonus, intern_asof=dt.date(2026, 12, 31))
    assert later.intern_counts[INTERN_SENIOR] > analysis.intern_counts[INTERN_SENIOR]


def test_intern_classification_flags_impossible_dates(roster_with_interns, bonus):
    """判断日期早于入职日期时算不出时长，单独归到"入职日期异常"而不是硬塞进某一类。"""
    from tj4tools.roster import INTERN_JUNIOR, INTERN_UNKNOWN

    early = reconcile(roster_with_interns, bonus, intern_asof=dt.date(2026, 5, 1))
    assert early.intern_counts[INTERN_UNKNOWN] > 40
    assert INTERN_JUNIOR not in early.intern_counts
    for item in early.items:
        if item.intern_class == INTERN_UNKNOWN:
            assert item.hire_date is None or item.hire_date >= dt.date(2026, 5, 1)
            assert any("无法计算入职时长" in flag for flag in item.flags)


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
    unmapped = [item for item in result.items if item.action == "add" and not item.workshop]
    assert unmapped, "应当存在需要人工指定车间的人员"
    assert all("生产技术转移组" in i.group or "在其他厂区" in i.group or "安全组" in i.group
               or "计算机化设备保障组" == i.group for i in unmapped)
    # 待指定分组要按人数汇总报出来，供界面渲染成清单
    assert result.unmapped_groups
    assert sum(count for _, count in result.unmapped_groups) == len(unmapped)
    assert dict(result.unmapped_groups)["生产技术转移组(多肽)"] == 7
    counts = [count for _, count in result.unmapped_groups]
    assert counts == sorted(counts, reverse=True)


def test_exclude_in_others_can_be_disabled(roster, bonus):
    kept = reconcile(roster, bonus, exclude_in_others=False)
    assert kept.excluded_in_others == 0
    assert sum(kept.counts.values()) == (
        kept.only_roster - kept.paired_renames
    ) + kept.only_bonus
    item = next(i for i in kept.items if i.key == ("赵凯", "ALS11919"))
    assert any("副主任" in flag for flag in item.flags)


def test_add_items_carry_roster_rows_and_removes_carry_frontline_rows(result):
    for item in result.items:
        if item.category in ADD_CATEGORIES:
            assert item.action == "add"
            assert item.frontline_row == 0
            assert item.roster_row >= 2
        else:
            assert item.frontline_row >= 3


def test_pending_delete_splits_into_keep_and_move(result):
    """核算有清单无的人都能在「生产部」里找到本人：职务还是一线四种就保留，否则移到副主任表。"""
    from tj4tools.roster import SHEET_OTHERS, is_target_duty

    items = result.by_category(CATEGORY_PENDING_DEL)
    assert len(items) == 16
    assert all(item.action in ("update", "move") for item in items)
    assert all(item.roster_row >= 2 for item in items)
    for item in items:
        new_duty = item.new_values["职务"]
        if is_target_duty(new_duty):
            assert item.action == "update", item
            assert "保留" in item.action_text
        else:
            assert item.action == "move", item
            assert item.target_sheet == SHEET_OTHERS
            assert item.target_workshop, item
    assert len(result.by_action("move")) == 15
    assert len(result.by_action("update")) == 1
    # 离职人员仍然是删除
    assert all(item.action == "remove" for item in result.by_category(CATEGORY_LEFT))


def test_move_targets_use_the_other_sheet_naming(result):
    """两张子表的车间叫法不同，移动目标要用副主任表的叫法。"""
    张睿尧 = next(i for i in result.by_action("move") if i.name == "张睿尧")
    assert 张睿尧.workshop == "11号楼D级车间"  # 一线人员的叫法
    assert 张睿尧.target_workshop == "11号楼车间D级区域"  # 副主任表的叫法
    assert 张睿尧.target_workshop_source == "经验映射"

    # 副主任表里没有清洗组，沿用一线车间名，导出时会新建
    李硕 = next(i for i in result.by_action("move") if i.name == "李硕")
    assert 李硕.target_workshop == "清洗组"
    assert 李硕.target_workshop_source == "沿用一线车间"


def test_others_sheet_layout_stops_before_the_reference_tables(bonus):
    """副主任表第 130 行之后是混排在 A~D 列的参照表，不能当成人员。"""
    layout = bonus.others_layout
    assert layout is not None
    assert layout.first_data_row == 2
    assert layout.last_data_row == 129
    assert len(layout.people) == 128
    assert not layout.merged_workshop
    assert layout.workshops[:2] == ["4号楼1&2车间", "4号楼3&4车间"]
    assert layout.workshops[-1] == "安全组"
    assert "产能利用率" not in layout.workshops
    assert "A" not in layout.workshops
    # 车间块首尾相接
    for previous, current in zip(layout.blocks, layout.blocks[1:]):
        assert current[1] == previous[2] + 1
    assert layout.anchor_for("12号楼", "副主任")[1] == "职务"
    assert layout.anchor_for("12号楼", "清洗工")[1] == "车间"
    assert layout.anchor_for("清洗组", "清洗工") == (129, "表尾")


def test_update_values_come_from_the_roster(result):
    promoted = next(i for i in result.items if i.key == ("胡强", "ALS10148"))
    assert promoted.updates == {"职务": ("班长", "车间副主任")}
    assert promoted.new_values["职务"] == "车间副主任"
    assert promoted.new_values["姓名"] == "胡强"

    # 编号相同、姓名不同（清单里名字带零宽字符）——改的是姓名
    renamed = next(i for i in result.items if i.key == ("曹静旺", "ALS14679"))
    assert renamed.updates == {"姓名": ("曹静旺", "曹睿晟")}
    assert renamed.new_values["姓名"] == "曹睿晟"
    assert renamed.update_text == "姓名 曹静旺→曹睿晟"
