"""用 Streamlit 的 AppTest 做端到端冒烟测试：不起浏览器、不截图，直接断言。"""

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
    assert "TJ4 安全质量奖人员对账" in app.title[0].value
    labels = {metric.label for metric in app.metric}
    assert {"两表一致", "新入职员工", "离职人员", "已在其他子表"} <= labels
    values = {metric.label: metric.value for metric in app.metric}
    assert values["两表一致"] == "616"
    assert values["新入职员工"] == "213"
    assert values["离职人员"] == "41"
    assert values["已在其他子表"] == "59"
    assert "参照日期 2026-07-31" in text


def test_unmapped_groups_are_summarised_not_dumped(app):
    warnings = [block.value for block in app.warning]
    assert any("22 名待新增人员" in text for text in warnings)
    # 长长的分组清单要收进折叠区，而不是塞在警告文字里
    assert all(len(text) < 120 for text in warnings), warnings


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
    # 22 名分组无法自动对应车间的人员应被跳过而不是硬塞进去
    assert "跳过 22 人" in success


def test_window_change_reruns_without_error(app):
    app.sidebar.number_input[0].set_value(6).run()
    assert not app.exception, [e.value for e in app.exception]
    values = {metric.label: metric.value for metric in app.metric}
    assert int(values["新入职员工"]) > 213
    app.sidebar.number_input[0].set_value(1).run()
    assert not app.exception


def test_duty_field_switch_reruns_without_error(app):
    app.sidebar.radio[0].set_value("岗位").run()
    assert not app.exception, [e.value for e in app.exception]
    app.sidebar.radio[0].set_value("职位").run()
    assert not app.exception


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
    key = "editor_新入职员工_0"
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
    assert any("新增 238 人" in text for text in messages), messages


def test_cancel_column_removes_person_from_both_exports():
    instance = AppTest.from_file(APP, default_timeout=200).run()
    _tick(instance, "editor_新入职员工_0", [0, 1], column="取消")
    assert list(instance.session_state["decisions"].values()) == ["取消"] * 2
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    messages = [block.value for block in instance.success]
    assert any("新增 236 人" in text for text in messages), messages


def test_preview_counts_match_generated_counts():
    """预览"将新增 N 人"必须等于生成结果里的 N，未指定车间的人不能算进预览。"""
    instance = AppTest.from_file(APP, default_timeout=200).run()
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **238** 人" in markdown, markdown[-500:]
    captions = " ".join(block.value for block in instance.caption)
    assert "另有 22 人未指定车间" in captions
    next(b for b in instance.button if b.label == "生成对照标记版").click().run()
    assert any("新增 238 人" in block.value for block in instance.success)

    _tick(instance, "editor_新入职员工_0", [0], column="取消")
    markdown = " ".join(block.value for block in instance.markdown)
    assert "将新增 **237** 人" in markdown
