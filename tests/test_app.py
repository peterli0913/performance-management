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
