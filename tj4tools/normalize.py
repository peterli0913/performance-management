"""姓名 / 员工编号 / 日期 的规范化工具。

真实表格里的脏数据（零宽字符、不换行空格、姓名带编号后缀、"--" 当日期）
必须在比对之前统一处理，否则会产生大量假差异。
"""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata

# 零宽字符 / BOM / 词连接符，中文表格里出现频率很高
_INVISIBLE = dict.fromkeys(map(ord, "\u200b\u200c\u200d\ufeff\u2060\u180e"), None)

# 姓名尾部的编号后缀，例如 "王鑫（ALS8215）" / "张志强(ALS15906)"
_NAME_SUFFIX = re.compile(r"[（(][^（()）]*[)）]\s*$")

# 表示"空"的占位文本
_BLANK_TOKENS = {"", "-", "--", "---", "—", "——", "/", "N/A", "NA", "无", "nan", "NaN", "None", "null"}

_DATE_PATTERNS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y年%m月%d日",
    "%m/%d/%Y",
    "%m-%d-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y%m%d",
)


def clean_text(value) -> str:
    """去掉不可见字符与全角差异，返回 strip 后的文本。"""
    if value is None:
        return ""
    if isinstance(value, float) and value != value:  # NaN
        return ""
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.strftime("%Y-%m-%d")
    text = str(value)
    text = text.translate(_INVISIBLE)
    text = text.replace("\xa0", " ")
    text = unicodedata.normalize("NFKC", text)
    return re.sub(r"\s+", " ", text).strip()


def is_blank(value) -> bool:
    """判断是否为业务意义上的空值（含 "--" 这类占位符）。"""
    return clean_text(value) in _BLANK_TOKENS


def norm_name(value) -> str:
    """规范化姓名：去不可见字符、去掉尾部括号里的编号、去掉所有内部空格。"""
    text = clean_text(value)
    if not text:
        return ""
    stripped = _NAME_SUFFIX.sub("", text).strip()
    if stripped:
        text = stripped
    return text.replace(" ", "")


def norm_eid(value) -> str:
    """规范化员工编号：转大写、去空格、去掉浮点尾巴（Excel 数字编号常见）。"""
    text = clean_text(value)
    if not text:
        return ""
    text = text.replace(" ", "").upper()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def person_key(name, eid) -> tuple[str, str]:
    """比对主键：规范化后的 (姓名, 员工编号)。"""
    return norm_name(name), norm_eid(eid)


def key_label(key: tuple[str, str]) -> str:
    """主键的字符串形式，用作 Streamlit widget key 与决策字典的 key。"""
    return f"{key[0]}|{key[1]}"


def parse_date(value) -> _dt.date | None:
    """尽量把各种写法解析成 date；无法解析或表示空则返回 None。"""
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, (int, float)):
        # Excel 序列号（1900 日期系统）
        if isinstance(value, float) and value != value:
            return None
        if 1 <= float(value) <= 2958465:
            return excel_serial_to_date(float(value))
        return None
    text = clean_text(value)
    if text in _BLANK_TOKENS:
        return None
    text = text.split(" ")[0] if " " in text and ":" in text else text
    for pattern in _DATE_PATTERNS:
        try:
            return _dt.datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    match = re.search(r"(\d{4})\D{1,2}(\d{1,2})\D{1,2}(\d{1,2})", text)
    if match:
        try:
            return _dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def fmt_date(value) -> str:
    """展示用日期文本；解析不出来时原样返回清洗后的文本。"""
    parsed = parse_date(value)
    if parsed is not None:
        return parsed.isoformat()
    return clean_text(value)


_EXCEL_EPOCH = _dt.date(1899, 12, 30)


def excel_serial_to_date(serial: float) -> _dt.date:
    return _EXCEL_EPOCH + _dt.timedelta(days=int(serial))


def date_to_excel_serial(value: _dt.date) -> int:
    return (value - _EXCEL_EPOCH).days


def months_before(anchor: _dt.date, months: int = 1) -> _dt.date:
    """anchor 往前推 N 个自然月（用于"入职时间在文件名日期一个月以内"）。"""
    month = anchor.month - months
    year = anchor.year
    while month <= 0:
        month += 12
        year -= 1
    day = anchor.day
    while day > 0:
        try:
            return _dt.date(year, month, day)
        except ValueError:
            day -= 1
    raise ValueError("无法计算前 %d 个月" % months)


_FILENAME_DATES = (
    # 07-31-2026 / 07_31_2026 / 07.31.2026
    (re.compile(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})"), ("m", "d", "y")),
    # 2026-07-31
    (re.compile(r"(\d{4})[-_.](\d{1,2})[-_.](\d{1,2})"), ("y", "m", "d")),
    # 2026年07月31日
    (re.compile(r"(\d{4})年(\d{1,2})月(\d{1,2})日"), ("y", "m", "d")),
    # 20260731
    (re.compile(r"(?<!\d)(\d{4})(\d{2})(\d{2})(?!\d)"), ("y", "m", "d")),
    # 2026年07月（取月末）
    (re.compile(r"(\d{4})年(\d{1,2})月"), ("y", "m")),
)


def date_from_filename(filename: str) -> _dt.date | None:
    """从文件名里提取参照日期，例如 "...，07-31-2026.xlsx" -> 2026-07-31。

    只给出年月时取该月最后一天。
    """
    text = clean_text(filename)
    for pattern, order in _FILENAME_DATES:
        for match in pattern.finditer(text):
            parts = dict(zip(order, (int(g) for g in match.groups())))
            year, month = parts.get("y"), parts.get("m")
            if not year or not month or not (1 <= month <= 12):
                continue
            day = parts.get("d")
            if day is None:
                day = _month_end(year, month)
            try:
                return _dt.date(year, month, day)
            except ValueError:
                continue
    return None


def _month_end(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (_dt.date(year, month + 1, 1) - _dt.timedelta(days=1)).day
