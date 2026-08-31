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


def test_pending_delete_tab_offers_a_per_person_action_choice(app):
    """待定·核算有清单无：给出自动判定，但每个人的动作都能手工改。"""
    text = " ".join(block.value for block in app.info)
    assert "动作列可以逐人改" in text
    assert "保留在一线人员" in text and "移动到副主任表格" in text
    values = {metric.label: metric.value for metric in app.metric}
    assert values["保留在一线人员"] == "1"
    assert values["移动到副主任表格"] == "15"
    labels = [button.label for button in app.button]
    assert "全部改为「保留在一线人员」" in labels
    assert "全部改为「移动到副主任表格」" in labels
    assert "恢复自动判定" in labels


def test_pending_delete_action_override_reaches_the_export(app):
    """把全部人改成"保留在一线人员"后，导出里就不该再有"移到副主任表"。"""
    next(b for b in app.button if b.label == "全部改为「保留在一线人员」").click().run()
    assert not app.exception, [e.value for e in app.exception]
    values = {metric.label: metric.value for metric in app.metric}
    assert values["保留在一线人员"] == "16"
    assert values["移动到副主任表格"] == "0"
    markdown = " ".join(block.value for block in app.markdown)
    assert "按清单更新 **16** 人" in markdown
    assert "移到副主任表" not in markdown.split("将新增")[1].split("将新增")[0]

    next(b for b in app.button if b.label == "恢复自动判定").click().run()
    values = {metric.label: metric.value for metric in app.metric}
    assert values["移动到副主任表格"] == "15"


def test_action_edit_updates_the_metrics_in_the_same_run():
    """指标渲染在表格上方，改完动作必须立刻重跑一次，否则数字要等下次交互才变。"""
    instance = AppTest.from_file(APP, default_timeout=250).run()
    key = "main_pending_editor_待定需填入人员（核算有·清单无）_0"
    keys = [name for name in instance.session_state.filtered_state if "pending_editor" in name]
    assert keys == [key], keys
    instance.session_state[key] = {
        "edited_rows": {0: {"动作": "保留在一线人员"}},
        "added_rows": [],
        "deleted_rows": [],
    }
    instance.run()
    assert not instance.exception, [e.value for e in instance.exception]
    values = {metric.label: metric.value for metric in instance.metric}
    assert values["保留在一线人员"] == "2"
    assert values["移动到副主任表格"] == "14"
    assert len(instance.session_state["action_override"]) == 1


def test_checkbox_edit_updates_the_metrics_in_the_same_run():
    instance = AppTest.from_file(APP, default_timeout=250).run()
    _tick(instance, "main_editor_新入职员工_0", [0, 1, 2])
    metrics = [m for m in instance.metric if m.label == "已应用"]
    assert metrics[0].value == "3"


def test_choice_labels_spell_out_the_action(app):
    """按钮/勾选框不能只写"应用/取消"，要写清楚是新增还是删除。"""
    labels = [button.label for button in app.button]
    assert "全部新增到一线人员" in labels
    assert "全部不新增" in labels
    assert "全部从一线人员删除" in labels
    assert "全部保留在一线人员" in labels
    assert "全部应用" not in labels
    assert "全部取消" not in labels


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
FEATURE_COMBINED = "③ 一键生成全部"


def _walk(node, out=None):
    out = out if out is not None else []
    out.append(node)
    children = getattr(node, "children", None)
    for child in children.values() if isinstance(children, dict) else (children or []):
        _walk(child, out)
    return out


def _stray_elements(instance):
    """找出被 Streamlit magic 意外渲染出来的原始对象。

    回归：`st.info(x) if cond else st.warning(x)` 是个裸表达式，magic 会给它套一层
    `st.write`，把 DeltaGenerator 的 repr 当正文打到页面上（用户看到的就是"一坨代码"）。
    magic 能识别并跳过直接的 `st.xxx(...)` 调用，但识别不了条件表达式。
    """
    bad = []
    for node in _walk(instance.main) + _walk(instance.sidebar):
        if type(node).__name__ == "UnknownElement":
            bad.append(repr(node)[:120])
        for attr in ("value", "body"):
            text = getattr(node, attr, None)
            if isinstance(text, str) and ("DeltaGenerator" in text or "_provided_cursor" in text):
                bad.append(f"{type(node).__name__}: {text[:120]}")
    return bad


def test_no_raw_objects_leak_into_the_page(app):
    assert _stray_elements(app) == []


def test_no_raw_objects_leak_into_the_supervisor_page():
    instance = AppTest.from_file(APP, default_timeout=250).run()
    instance.radio(key="feature").set_value(FEATURE_SUPERVISOR).run()
    assert not instance.exception, [e.value for e in instance.exception]
    assert _stray_elements(instance) == []


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


def test_supervisor_scope_is_fixed_to_strict(supervisor_app):
    """口径不再让用户选：一律排除功能一能放进一线车间的人。"""
    keys = [widget.label for widget in supervisor_app.radio]
    assert "「不在一线清单里」怎么算" not in keys
    with pytest.raises(KeyError):
        supervisor_app.radio(key="sup_scope")
    # 严格口径的目标是 150 人，本表现有 128 人
    info = " ".join(block.value for block in supervisor_app.info)
    assert "目标共 150 人" in info
    assert "本表现有 128 人" in info
    caption = " ".join(block.value for block in supervisor_app.caption)
    assert "功能一能放进一线的那批人不会重复出现在这里" in caption
    values = {metric.label: metric.value for metric in supervisor_app.metric}
    assert values["新入职员工"] == "20"


def test_supervisor_unmapped_warning_is_live(supervisor_app):
    """未映射车间的提示只出现一次，且由界面按当前映射实时计算。"""
    warnings = [block.value for block in supervisor_app.warning]
    live = [text for text in warnings if "没有对应车间" in text]
    assert len(live) == 1, warnings
    assert "还有 14 名待新增人员" in live[0]


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


def test_combined_feature_is_on_the_radio():
    instance = AppTest.from_file(APP, default_timeout=200).run()
    feature = instance.radio(key="feature")
    assert FEATURE_COMBINED in feature.options
    feature.set_value(FEATURE_COMBINED).run()
    assert not instance.exception, [e.value for e in instance.exception]
    assert _stray_elements(instance) == []
    labels = [button.label for button in instance.button]
    assert "生成全部已应用版" in labels
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **279** 人" in markdown
    assert "移到副主任表 **15** 人" in markdown
    captions = " ".join(block.value for block in instance.caption)
    assert "另有 15 人从一线移入" in captions
    info = " ".join(block.value for block in instance.info)
    assert "有 6 人同时出现在一线移出和副主任表待新增里" in info


def test_combined_approved_only_uses_existing_decisions():
    """没勾应用时，第二种范围只留下待定类的自动动作。"""
    instance = AppTest.from_file(APP, default_timeout=200).run()
    instance.radio(key="feature").set_value(FEATURE_COMBINED).run()
    instance.radio(key="combined_scope").set_value("只处理已经勾选「应用」的人").run()
    assert not instance.exception, [e.value for e in instance.exception]
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **0** 人" in markdown
    assert "移到副主任表 **15** 人" in markdown
    assert "按清单更新 **1** 人" in markdown


def test_combined_generate_offers_one_download():
    instance = AppTest.from_file(APP, default_timeout=300).run()
    instance.radio(key="feature").set_value(FEATURE_COMBINED).run()
    next(b for b in instance.button if b.label == "生成全部已应用版").click().run()
    assert not instance.exception, [e.value for e in instance.exception]
    success = " ".join(block.value for block in instance.success)
    assert "直接插入 279 人" in success
    assert "直接删除 41 人" in success
    assert "移到「副主任&工艺组长及其他」15 人" in success
    assert "副主任表直接插入 15 人" in success
    assert "副主任表直接删除 11 人" in success
    labels = [d.label for d in instance.get("download_button")]
    assert any("两表已应用版" in label for label in labels)


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
