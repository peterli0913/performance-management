"""用 Streamlit 的 AppTest 做端到端冒烟测试：不起浏览器、不截图，直接断言。"""

import datetime as dt
import os

import pytest

pytest.importorskip("streamlit")
from streamlit.testing.v1 import AppTest  # noqa: E402

APP = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "app.py")


@pytest.fixture(scope="module")
def app():
    instance = AppTest.from_file(APP, default_timeout=180).run()
    assert not instance.exception, [e.value for e in instance.exception]
    return instance


def test_app_boots_and_analyzes_sample_files(app):
    text = " ".join(block.value for block in app.markdown) + " ".join(
        block.value for block in app.caption
    )
    assert "TJ4 安全质量奖核算表生成" in app.title[0].value
    labels = {metric.label for metric in app.metric}
    assert {"两表一致", "新入职员工", "离职人员", "已在其他子表"} <= labels
    values = {metric.label: metric.value for metric in app.metric}
    assert values["两表一致"] == "616"
    assert values["新入职员工"] == "249"
    assert values["离职人员"] == "41"
    assert values["待定·需核实"] == "16"
    assert values["已在其他子表"] == "59"
    assert "参照日期 2026-07-31" in text
    assert "入职时间在 2026-06-30 ~ 2026-07-31 之间的算新入职" in text


def test_new_hire_window_is_a_date_input(app):
    labels = [widget.label for widget in app.sidebar.date_input]
    assert "新入职判定窗口日期" in labels
    assert not [w for w in app.sidebar.number_input if "窗口" in w.label]
    window = next(w for w in app.sidebar.date_input if w.label == "新入职判定窗口日期")
    assert str(window.value) == "2026-06-30"


def test_interns_are_included_and_counted(app):
    intern_box = next(c for c in app.sidebar.checkbox if "纳入实习生" in c.label)
    assert intern_box.value is True
    assert "47 人" in intern_box.label


def test_unmapped_groups_are_summarised_not_dumped(app):
    warnings = [block.value for block in app.warning]
    assert any("23 名待新增人员" in text for text in warnings)
    # 长长的分组清单要收进折叠区，而不是塞在警告文字里
    assert all(len(text) < 120 for text in warnings), warnings


def test_same_person_under_two_names_is_reported_once(app):
    warnings = " ".join(block.value for block in app.warning)
    assert "有 1 人在两表里编号相同、姓名不同" in warnings


def test_sidebar_defaults_pick_the_right_files(app):
    roster = app.sidebar.selectbox[0]
    bonus = app.sidebar.selectbox[1]
    assert "人员清单" in roster.value
    assert "安全质量奖" in bonus.value
    assert app.sidebar.radio[0].value == "职位"


def test_tabs_are_present(app):
    labels = [tab.label for tab in app.tabs] if hasattr(app, "tabs") else []
    if not labels:
        pytest.skip("当前 Streamlit 版本不暴露 tabs 元素")
    assert any("新入职员工" in label for label in labels)
    assert any("生成核算表" in label for label in labels)


def test_generate_marked_workbook_end_to_end(app):
    button = next(b for b in app.button if b.label == "生成对照标记版")
    button.click().run()
    assert not app.exception, [e.value for e in app.exception]
    downloads = [d for d in app.get("download_button")]
    assert any("对照标记版" in d.label for d in downloads)
    success = " ".join(block.value for block in app.success)
    assert "标绿" in success and "标红" in success
    assert "按清单更新 1 人" in success
    assert "移到「副主任&工艺组长及其他」15 人" in success
    assert "实习生" in success
    # 23 名分组无法自动对应车间的人员应被跳过而不是硬塞进去
    assert "跳过 23 人" in success


def test_window_date_change_reruns_and_shifts_the_split(app):
    window = next(w for w in app.sidebar.date_input if w.label == "新入职判定窗口日期")
    baseline = {metric.label: metric.value for metric in app.metric}
    window.set_value(dt.date(2026, 1, 1)).run()
    assert not app.exception, [e.value for e in app.exception]
    widened = {metric.label: metric.value for metric in app.metric}
    assert int(widened["新入职员工"]) > int(baseline["新入职员工"])
    assert int(widened["待定·需新增"]) < int(baseline["待定·需新增"])
    window.set_value(dt.date(2026, 6, 30)).run()
    assert not app.exception


def test_intern_classification_is_shown_and_configurable(app):
    captions = " ".join(block.value for block in app.caption)
    assert "实习生 46 人（按 2026-07-31 判断）" in captions
    assert "入职超过3个月 1 人" in captions and "入职不到3个月 45 人" in captions
    asof = next(w for w in app.sidebar.date_input if w.label == "实习生判断日期")
    assert str(asof.value) == "2026-07-31"
    asof.set_value(dt.date(2026, 12, 31)).run()
    assert not app.exception, [e.value for e in app.exception]
    captions = " ".join(block.value for block in app.caption)
    assert "入职超过3个月 46 人" in captions
    asof.set_value(dt.date(2026, 7, 31)).run()


def test_pending_delete_tab_is_read_only_and_lists_actions(app):
    text = " ".join(block.value for block in app.markdown) + " ".join(
        block.value for block in app.info
    )
    assert "全部按人员清单的信息处理，无需复核" in text
    assert "移到「副主任&工艺组长及其他」子表" in text
    labels = {metric.label for metric in app.metric}
    assert {"保留并更新", "移到副主任表"} <= labels
    values = {metric.label: metric.value for metric in app.metric}
    assert values["保留并更新"] == "1"
    assert values["移到副主任表"] == "15"


def test_turning_interns_off_shrinks_the_target_list(app):
    intern_box = next(c for c in app.sidebar.checkbox if "纳入实习生" in c.label)
    intern_box.set_value(False).run()
    assert not app.exception, [e.value for e in app.exception]
    values = {metric.label: metric.value for metric in app.metric}
    assert values["新入职员工"] == "213"
    intern_box.set_value(True).run()
    assert {m.label: m.value for m in app.metric}["新入职员工"] == "249"


def test_duty_field_switch_reruns_without_error(app):
    app.sidebar.radio[0].set_value("岗位").run()
    assert not app.exception, [e.value for e in app.exception]
    app.sidebar.radio[0].set_value("职位").run()
    assert not app.exception


FEATURE_SUPERVISOR = "② 副主任&工艺组长及其他"


@pytest.fixture(scope="module")
def supervisor_app():
    instance = AppTest.from_file(APP, default_timeout=250).run()
    assert not instance.exception, [e.value for e in instance.exception]
    instance.radio(key="feature").set_value(FEATURE_SUPERVISOR).run()
    assert not instance.exception, [e.value for e in instance.exception]
    return instance


def test_switching_to_the_supervisor_feature_works(supervisor_app):
    values = {metric.label: metric.value for metric in supervisor_app.metric}
    assert values["已在本表"] == "117"
    assert values["新入职员工"] == "20"
    assert values["待定·需新增"] == "13"
    assert values["离职人员"] == "6"
    assert values["待定·需核实"] == "5"
    text = " ".join(block.value for block in supervisor_app.caption)
    assert "副主任&工艺组长及其他" in text
    assert "管理类职务" in text or "管理类职务" in " ".join(
        block.value for block in supervisor_app.info
    )


def test_supervisor_scope_choice_is_explicit(supervisor_app):
    from tj4tools.supervisor import SCOPE_STRICT

    scope = supervisor_app.radio(key="sup_scope")
    assert scope.value == SCOPE_STRICT, "默认应当是严格口径"
    # 两种口径的目标人数要直接标在选项上，否则用户看不出选择的后果
    labels = " ".join(str(option) for option in scope.options)
    assert "目标 150 人" in labels
    assert "目标 387 人" in labels


def test_supervisor_literal_scope_warns_about_double_placement(supervisor_app):
    from tj4tools.supervisor import SCOPE_LITERAL, SCOPE_STRICT

    supervisor_app.radio(key="sup_scope").set_value(SCOPE_LITERAL).run()
    assert not supervisor_app.exception, [e.value for e in supervisor_app.exception]
    warnings = " ".join(block.value for block in supervisor_app.warning)
    assert "同一个人会同时进两张子表" in warnings
    values = {metric.label: metric.value for metric in supervisor_app.metric}
    assert int(values["新入职员工"]) > 200
    supervisor_app.radio(key="sup_scope").set_value(SCOPE_STRICT).run()
    assert {m.label: m.value for m in supervisor_app.metric}["新入职员工"] == "20"


def test_supervisor_generates_and_offers_download(supervisor_app):
    button = next(b for b in supervisor_app.button if b.label == "生成对照标记版")
    button.click().run()
    assert not supervisor_app.exception, [e.value for e in supervisor_app.exception]
    success = " ".join(block.value for block in supervisor_app.success)
    assert "新增" in success and "标记删除 11 人" in success
    labels = [d.label for d in supervisor_app.get("download_button")]
    assert any("副主任表·对照标记版" in label for label in labels)


def test_two_features_keep_separate_decisions():
    """同一个人可能出现在两个功能里，决策不能串。"""
    instance = AppTest.from_file(APP, default_timeout=250).run()
    _tick(instance, "main_editor_新入职员工_0", [0])
    assert len(instance.session_state["decisions"]) == 1
    instance.radio(key="feature").set_value(FEATURE_SUPERVISOR).run()
    assert not instance.exception, [e.value for e in instance.exception]
    # 功能二有自己的命名空间，功能一的决策不会漏过去
    assert instance.session_state["sup__decisions"] == {}
    assert len(instance.session_state["decisions"]) == 1


def _tick(instance, key, rows, column="应用"):
    """模拟在 data_editor 里勾选若干行。"""
    instance.session_state[key] = {
        "edited_rows": {row: {column: True} for row in rows},
        "added_rows": [],
        "deleted_rows": [],
    }
    return instance.run()


def test_table_checkboxes_survive_reruns_and_reach_the_export():
    """回归：勾选「应用」后切页签、点其他按钮，决策不能丢。"""
    instance = AppTest.from_file(APP, default_timeout=200).run()
    key = "main_editor_新入职员工_0"
    _tick(instance, key, [0, 1, 2])
    assert instance.session_state["decisions"], "勾选应当写进决策状态"
    assert list(instance.session_state["decisions"].values()) == ["应用"] * 3

    # 多跑几轮空 rerun，模拟切页签、改设置
    for _ in range(3):
        instance.run()
        assert list(instance.session_state["decisions"].values()) == ["应用"] * 3

    # 先生成标记版（会新增下载按钮），再生成已应用版
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    assert list(instance.session_state["decisions"].values()) == ["应用"] * 3
    next(b for b in instance.button if b.label == "生成已应用版").click().run()
    assert not instance.exception, [e.value for e in instance.exception]
    messages = [block.value for block in instance.success]
    assert any("直接插入 3 人" in text for text in messages), messages


def test_unticked_rows_stay_undecided():
    """只渲染表格不勾选，不能把所有人变成"取消"（否则标记版会空掉）。"""
    instance = AppTest.from_file(APP, default_timeout=200).run()
    assert instance.session_state["decisions"] == {}
    instance.run()
    assert instance.session_state["decisions"] == {}
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    messages = [block.value for block in instance.success]
    assert any("新增 279 人" in text for text in messages), messages


def test_cancel_column_removes_person_from_both_exports():
    instance = AppTest.from_file(APP, default_timeout=200).run()
    _tick(instance, "main_editor_新入职员工_0", [0, 1], column="取消")
    assert list(instance.session_state["decisions"].values()) == ["取消"] * 2
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    messages = [block.value for block in instance.success]
    assert any("新增 277 人" in text for text in messages), messages


def test_preview_counts_match_generated_counts():
    """预览"将新增 N 人"必须等于生成结果里的 N，未指定车间的人不能算进预览。"""
    instance = AppTest.from_file(APP, default_timeout=200).run()
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **279** 人" in markdown, markdown[-500:]
    captions = " ".join(block.value for block in instance.caption)
    assert "另有 23 人未指定车间" in captions
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    assert any("新增 279 人" in block.value for block in instance.success)

    _tick(instance, "main_editor_新入职员工_0", [0], column="取消")
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **278** 人" in markdown
