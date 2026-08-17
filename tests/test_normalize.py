import datetime as dt

from tj4tools.normalize import (
    date_from_filename,
    fmt_date,
    is_blank,
    months_before,
    norm_eid,
    norm_name,
    parse_date,
)


def test_norm_name_strips_invisible_and_suffix():
    assert norm_name("曹\u200b睿\u200b晟") == "曹睿晟"
    assert norm_name("王鑫（ALS8215）") == "王鑫"
    assert norm_name("张志强(ALS15906)") == "张志强"
    assert norm_name("田志勇\xa0") == "田志勇"
    assert norm_name("梁光强 / 李晓刚") == "梁光强/李晓刚"
    assert norm_name(None) == ""


def test_norm_eid():
    assert norm_eid(" als12990 ") == "ALS12990"
    assert norm_eid("12345.0") == "12345"
    assert norm_eid("ＡＬＳ1") == "ALS1"


def test_is_blank_handles_placeholders():
    assert is_blank("--")
    assert is_blank("")
    assert is_blank(None)
    assert not is_blank("2026/07/29")


def test_parse_date_variants():
    assert parse_date(dt.datetime(2026, 7, 29)) == dt.date(2026, 7, 29)
    assert parse_date("2026/07/29") == dt.date(2026, 7, 29)
    assert parse_date("2026年7月9日") == dt.date(2026, 7, 9)
    assert parse_date("--") is None
    assert parse_date("放弃入职") is None
    assert parse_date(46204) == dt.date(2026, 7, 1)


def test_date_from_filename():
    assert date_from_filename("TJ4生产部&生产设备部人员清单，07-31-2026.xlsx") == dt.date(2026, 7, 31)
    assert date_from_filename("2026年07月份安全质量奖核算数据.xlsx") == dt.date(2026, 7, 31)
    assert date_from_filename("清单20260731.xlsx") == dt.date(2026, 7, 31)
    assert date_from_filename("无日期.xlsx") is None


def test_months_before_clamps_month_end():
    assert months_before(dt.date(2026, 7, 31)) == dt.date(2026, 6, 30)
    assert months_before(dt.date(2026, 3, 31)) == dt.date(2026, 2, 28)
    assert months_before(dt.date(2026, 1, 15)) == dt.date(2025, 12, 15)


def test_fmt_date_passthrough():
    assert fmt_date(None) == ""
    assert fmt_date("放弃入职") == "放弃入职"
