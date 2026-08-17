"""在不破坏格式的前提下对 xlsx 做行级手术（插入 / 删除 / 单元格着色）。

为什么不用 openpyxl 往返：目标文件含 externalLinks（`[1]生产部!$C:$C` 公式依赖它）、
printerSettings、customProperty、批注 VML，openpyxl 保存时会全部丢弃。
这里直接改写 OOXML 的 XML part，未修改的 part 逐字节原样复制。

行号重映射使用三种语义，缺一不可：
  * ``map_single``       —— 单点引用、``<row r>``、``<c r>``
  * ``map_range_start``  —— 区间起点
  * ``map_range_end``    —— 区间终点；插入锚点处的终点需要顺延，否则
                            合并区 / 区间公式 / 打印区域会把新增行排除在外
"""

from __future__ import annotations

import bisect
import copy
import datetime as _dt
import io
import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field

from .normalize import date_to_excel_serial

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

EXCEL_MAX_ROW = 1048576

FILL_GREEN = "FF92D050"
FILL_RED = "FFFF0000"

_XMLNS_RE = re.compile(r'xmlns(?::(?P<prefix>[\w.-]+))?="(?P<uri>[^"]+)"')

# 工作表/外部引用前缀，例如  Sheet1!  '4号楼'!  [1]生产部!
_SHEET_PREFIX = (
    r"(?:'(?:[^']|'')*'|\[\d+\][^\s!,()+\-*/&=<>:'\"]*|[A-Za-z0-9_.\u4e00-\u9fff]+)!"
)
_CELL = r"\$?[A-Z]{1,3}\$?\d{1,7}"
_REF_RE = re.compile(
    rf"(?P<prefix>{_SHEET_PREFIX})?(?P<c1>{_CELL})(?::(?P<c2>{_CELL}))?"
)
_CELL_PARTS_RE = re.compile(r"^(?P<cd>\$?)(?P<col>[A-Z]{1,3})(?P<rd>\$?)(?P<row>\d{1,7})$")
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
_IDENT_BEFORE = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_$.")
_COORD_RE = re.compile(r"^([A-Z]{1,3})(\d+)$")


def col_to_index(col: str) -> int:
    """列字母转 1 基序号。"""
    result = 0
    for char in col:
        result = result * 26 + (ord(char) - 64)
    return result


def index_to_col(index: int) -> str:
    """1 基序号转列字母。"""
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def split_coord(coord: str) -> tuple[str, int]:
    match = _COORD_RE.match(coord)
    if not match:
        raise ValueError(f"非法单元格坐标: {coord!r}")
    return match.group(1), int(match.group(2))


# --------------------------------------------------------------------------- #
# 请求对象
# --------------------------------------------------------------------------- #


@dataclass
class NewRow:
    """一条待插入的新行。"""

    values: dict[str, object] = field(default_factory=dict)
    fills: dict[str, str] = field(default_factory=dict)


@dataclass
class InsertGroup:
    """在 ``anchor_row`` 之后插入若干新行。

    ``new_block=True`` 表示这是一个全新的分组块（如原表没有的"13号楼"车间），
    此时不会顺延锚点处的块内合并区/统计公式，而是为新块单独建立合并区。
    """

    anchor_row: int
    template_row: int
    rows: list[NewRow]
    new_block: bool = False
    block_col: str = "A"
    block_label: str | None = None


@dataclass
class Highlight:
    """给已有行的若干单元格加填充色。"""

    row: int
    cols: list[str]
    color: str


# --------------------------------------------------------------------------- #
# 行号重映射
# --------------------------------------------------------------------------- #


class RowMap:
    """把原始行号映射到编辑后的行号。"""

    def __init__(
        self,
        deletes: set[int],
        extend_counts: dict[int, int],
        block_counts: dict[int, int],
        span_threshold: int,
    ):
        self.deletes = set(deletes)
        self.extend_counts = dict(extend_counts)
        self.block_counts = dict(block_counts)
        self.span_threshold = span_threshold

        anchors = sorted(set(extend_counts) | set(block_counts))
        self._anchors = anchors
        totals = []
        running = 0
        for anchor in anchors:
            running += extend_counts.get(anchor, 0) + block_counts.get(anchor, 0)
            totals.append(running)
        self._anchor_cum = totals
        self._total_inserted = running

        self._deleted_sorted = sorted(self.deletes)

    def _inserted_before(self, row: int) -> int:
        # 锚点 a < row 的插入行都落在 row 之前
        idx = bisect.bisect_left(self._anchors, row)
        return self._anchor_cum[idx - 1] if idx > 0 else 0

    def _deleted_before(self, row: int) -> int:
        return bisect.bisect_left(self._deleted_sorted, row)

    def new_row(self, row: int) -> int:
        """行本身的新行号（不判断是否被删除）。"""
        value = row + self._inserted_before(row) - self._deleted_before(row)
        return min(value, EXCEL_MAX_ROW)

    def inserted_rows_at(self, anchor: int) -> tuple[list[int], list[int]]:
        """返回 (顺延块内区间的新行号, 新建块的新行号)。"""
        base = self.new_row(anchor)
        extend = self.extend_counts.get(anchor, 0)
        block = self.block_counts.get(anchor, 0)
        extend_rows = list(range(base + 1, base + extend + 1))
        block_rows = list(range(base + extend + 1, base + extend + block + 1))
        return extend_rows, block_rows

    def map_single(self, row: int) -> int:
        if row in self.deletes:
            return self.map_range_start(row)
        return self.new_row(row)

    def map_range_start(self, row: int) -> int:
        if row not in self.deletes:
            return self.new_row(row)
        # 起点被删：顺移到下一个存活行
        candidate = row + 1
        while candidate in self.deletes:
            candidate += 1
        return self.new_row(candidate)

    def map_range_end(self, row: int, span_start: int | None = None) -> int:
        if row in self.deletes:
            candidate = row - 1
            while candidate > 0 and candidate in self.deletes:
                candidate -= 1
            return self.new_row(candidate) if candidate > 0 else self.new_row(row)
        end = self.new_row(row)
        extend = self.extend_counts.get(row, 0)
        block = self.block_counts.get(row, 0)
        if extend or block:
            # 跨越整个数据区的区间（合计行、打印区域）连新建块一起顺延；
            # 块内区间（车间合并区、分块统计公式）只顺延本块新增行。
            whole_sheet = span_start is not None and span_start <= self.span_threshold
            end += extend + (block if whole_sheet else 0)
        return min(end, EXCEL_MAX_ROW)

    @property
    def total_inserted(self) -> int:
        return self._total_inserted

    @property
    def total_deleted(self) -> int:
        return len(self.deletes)


# --------------------------------------------------------------------------- #
# 公式处理
# --------------------------------------------------------------------------- #


def _iter_refs(formula: str):
    """扫描公式中的单元格引用，产出 (match, prefix, c1, c2)。已跳过字符串字面量。"""
    masked = list(formula)
    for literal in _STRING_LITERAL_RE.finditer(formula):
        for i in range(*literal.span()):
            masked[i] = "\x00"
    masked_text = "".join(masked)
    for match in _REF_RE.finditer(masked_text):
        start, end = match.span()
        if "\x00" in masked_text[start:end]:
            continue
        if match.group("prefix") is None and start > 0:
            if masked_text[start - 1] in _IDENT_BEFORE:
                continue
        after = masked_text[end : end + 1]
        if after.isdigit() or after == "(":
            continue
        yield match


def _rewrite_refs(formula: str, rewrite) -> str:
    """用 ``rewrite(prefix, c1, c2) -> str | None`` 重写公式中的引用。"""
    out: list[str] = []
    cursor = 0
    for match in _iter_refs(formula):
        start, end = match.span()
        replacement = rewrite(match.group("prefix"), match.group("c1"), match.group("c2"))
        if replacement is None:
            continue
        out.append(formula[cursor:start])
        out.append(replacement)
        cursor = end
    out.append(formula[cursor:])
    return "".join(out)


def _prefix_is_self(prefix: str | None, sheet_name: str) -> bool:
    """判断引用前缀是否指向本表（None 视为本表）。"""
    if prefix is None:
        return True
    name = prefix[:-1]
    if name.startswith("["):
        return False
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return name == sheet_name


def remap_formula(formula: str, rowmap: RowMap, sheet_name: str) -> str:
    """把公式里指向本表的行号按 rowmap 重写。"""

    def rewrite(prefix, c1, c2):
        if not _prefix_is_self(prefix, sheet_name):
            return None
        p1 = _CELL_PARTS_RE.match(c1)
        if not p1:
            return None
        row1 = int(p1.group("row"))
        if c2 is None:
            new1 = rowmap.map_single(row1)
            return f"{prefix or ''}{p1.group('cd')}{p1.group('col')}{p1.group('rd')}{new1}"
        p2 = _CELL_PARTS_RE.match(c2)
        if not p2:
            return None
        row2 = int(p2.group("row"))
        new1 = rowmap.map_range_start(row1)
        new2 = rowmap.map_range_end(row2, span_start=row1)
        return (
            f"{prefix or ''}{p1.group('cd')}{p1.group('col')}{p1.group('rd')}{new1}"
            f":{p2.group('cd')}{p2.group('col')}{p2.group('rd')}{new2}"
        )

    return _rewrite_refs(formula, rewrite)


def translate_formula(formula: str, row_delta: int, col_delta: int = 0) -> str:
    """把公式按相对偏移平移（用于展开共享公式）。"""

    def rewrite(prefix, c1, c2):
        parts = []
        for coord in (c1, c2):
            if coord is None:
                continue
            match = _CELL_PARTS_RE.match(coord)
            if not match:
                return None
            col = match.group("col")
            row = int(match.group("row"))
            if not match.group("cd") and col_delta:
                col = index_to_col(max(1, col_to_index(col) + col_delta))
            if not match.group("rd") and row_delta:
                row = max(1, row + row_delta)
            parts.append(f"{match.group('cd')}{col}{match.group('rd')}{row}")
        return (prefix or "") + ":".join(parts)

    return _rewrite_refs(formula, rewrite)


def remap_sqref(sqref: str, rowmap: RowMap) -> str:
    """重写 ``sqref`` / ``ref`` 里的行号（形如 "A1:B2 D5 D7:D9"）。"""
    tokens = []
    for token in sqref.split():
        if ":" in token:
            left, right = token.split(":", 1)
            pl, pr = _CELL_PARTS_RE.match(left), _CELL_PARTS_RE.match(right)
            if not pl or not pr:
                tokens.append(token)
                continue
            row1 = int(pl.group("row"))
            row2 = int(pr.group("row"))
            new1 = rowmap.map_range_start(row1)
            new2 = rowmap.map_range_end(row2, span_start=row1)
            if new2 < new1:
                new2 = new1
            tokens.append(
                f"{pl.group('cd')}{pl.group('col')}{pl.group('rd')}{new1}"
                f":{pr.group('cd')}{pr.group('col')}{pr.group('rd')}{new2}"
            )
        else:
            part = _CELL_PARTS_RE.match(token)
            if not part:
                tokens.append(token)
                continue
            new = rowmap.map_single(int(part.group("row")))
            tokens.append(f"{part.group('cd')}{part.group('col')}{part.group('rd')}{new}")
    return " ".join(tokens)


# --------------------------------------------------------------------------- #
# XML part 读写
# --------------------------------------------------------------------------- #


def _parse_part(data: bytes) -> tuple[ET.Element, str, str]:
    """解析 XML part，返回 (root, 原始声明, 原始根标签)。同时注册命名空间前缀。

    命名空间要在**整篇文档**里扫（有些前缀声明在子元素上，如 ``x15ac``），
    这样 ElementTree 序列化时才会沿用原始前缀。
    """
    text = data.decode("utf-8")
    match = re.search(r"<(?P<name>[A-Za-z_][\w.-]*)\b[^>]*>", text)
    if match is None:
        raise ValueError("无法定位根元素")
    root_tag = match.group(0)
    declaration = text[: match.start()]
    for ns in _XMLNS_RE.finditer(text):
        ET.register_namespace(ns.group("prefix") or "", ns.group("uri"))
    return ET.fromstring(text), declaration, root_tag


def _serialize_part(root: ET.Element, declaration: str, root_tag: str) -> bytes:
    """序列化，并把原始根标签里 ElementTree 没写出来的 xmlns 声明补回去。

    ``mc:Ignorable="x15"`` 这类属性会引用只在根标签声明过的前缀，
    整体替换根标签会丢掉 ET 新增的声明，只补差集才安全。
    """
    body = ET.tostring(root, encoding="unicode")
    match = re.match(r"<[A-Za-z_][\w.-]*\b[^>]*?(?P<selfclose>/?)>", body)
    if match is None:
        raise ValueError("序列化结果异常")
    new_tag = match.group(0)
    present = {ns.group(0) for ns in _XMLNS_RE.finditer(new_tag)}
    missing = [ns.group(0) for ns in _XMLNS_RE.finditer(root_tag) if ns.group(0) not in present]
    if missing:
        closing = "/>" if match.group("selfclose") else ">"
        new_tag = new_tag[: -len(closing)].rstrip() + " " + " ".join(missing) + closing
    body = new_tag + body[match.end() :]
    return (declaration + body).encode("utf-8")


def _q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


# --------------------------------------------------------------------------- #
# 样式表
# --------------------------------------------------------------------------- #


class StyleTable:
    """按需向 styles.xml 追加填充色与 cellXfs 条目。"""

    def __init__(self, data: bytes):
        self.root, self._decl, self._root_tag = _parse_part(data)
        self._fills = self.root.find(_q("fills"))
        self._cell_xfs = self.root.find(_q("cellXfs"))
        if self._fills is None or self._cell_xfs is None:
            raise ValueError("styles.xml 缺少 fills/cellXfs")
        self._xfs = list(self._cell_xfs)
        self._fill_cache: dict[str, int] = {}
        self._xf_cache: dict[tuple[int, int], int] = {}
        self.dirty = False
        for index, fill in enumerate(self._fills):
            pattern = fill.find(_q("patternFill"))
            if pattern is None or pattern.get("patternType") != "solid":
                continue
            fg = pattern.find(_q("fgColor"))
            rgb = fg.get("rgb") if fg is not None else None
            if rgb and rgb.upper() not in self._fill_cache:
                self._fill_cache[rgb.upper()] = index

    def fill_id(self, argb: str) -> int:
        argb = argb.upper()
        if argb in self._fill_cache:
            return self._fill_cache[argb]
        fill = ET.SubElement(self._fills, _q("fill"))
        pattern = ET.SubElement(fill, _q("patternFill"))
        pattern.set("patternType", "solid")
        ET.SubElement(pattern, _q("fgColor")).set("rgb", argb)
        ET.SubElement(pattern, _q("bgColor")).set("indexed", "64")
        index = len(self._fills) - 1
        self._fills.set("count", str(len(self._fills)))
        self._fill_cache[argb] = index
        self.dirty = True
        return index

    def styled(self, base_xf: int, argb: str) -> int:
        """返回"基于 base_xf、填充色为 argb"的 cellXfs 索引。"""
        fill = self.fill_id(argb)
        key = (base_xf, fill)
        if key in self._xf_cache:
            return self._xf_cache[key]
        if 0 <= base_xf < len(self._xfs):
            new_xf = copy.deepcopy(self._xfs[base_xf])
        else:
            new_xf = ET.Element(_q("xf"))
            new_xf.set("numFmtId", "0")
            new_xf.set("fontId", "0")
            new_xf.set("borderId", "0")
            new_xf.set("xfId", "0")
        new_xf.set("fillId", str(fill))
        new_xf.set("applyFill", "1")
        self._cell_xfs.append(new_xf)
        self._xfs.append(new_xf)
        index = len(self._xfs) - 1
        self._cell_xfs.set("count", str(index + 1))
        self._xf_cache[key] = index
        self.dirty = True
        return index

    def to_bytes(self) -> bytes:
        return _serialize_part(self.root, self._decl, self._root_tag)


# --------------------------------------------------------------------------- #
# 主编辑器
# --------------------------------------------------------------------------- #


class XlsxEditor:
    """加载 xlsx 全部 part，对指定工作表做行手术后重新打包。"""

    def __init__(self, data: bytes):
        self._order: list[str] = []
        self._parts: dict[str, bytes] = {}
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                self._order.append(info.filename)
                self._parts[info.filename] = archive.read(info.filename)
        self._sheet_paths = self._resolve_sheet_paths()
        self._styles: StyleTable | None = None
        self._used: dict[tuple, int] = {}
        self.warnings: list[str] = []

    # -- 元数据 ------------------------------------------------------------ #

    def _resolve_sheet_paths(self) -> dict[str, str]:
        wb_root, _, _ = _parse_part(self._parts["xl/workbook.xml"])
        rels_root, _, _ = _parse_part(self._parts["xl/_rels/workbook.xml.rels"])
        targets = {}
        for rel in rels_root:
            target = rel.get("Target", "")
            if target.startswith("/"):
                path = target.lstrip("/")
            elif target.startswith("../"):
                path = target[3:]
            else:
                path = "xl/" + target
            targets[rel.get("Id")] = path
        mapping = {}
        sheets = wb_root.find(_q("sheets"))
        for sheet in [] if sheets is None else list(sheets):
            rel_id = sheet.get(f"{{{REL_NS}}}id")
            if rel_id in targets:
                mapping[sheet.get("name")] = targets[rel_id]
        return mapping

    @property
    def sheet_names(self) -> list[str]:
        return list(self._sheet_paths)

    @property
    def styles(self) -> StyleTable:
        if self._styles is None:
            self._styles = StyleTable(self._parts["xl/styles.xml"])
        return self._styles

    # -- 行手术 ------------------------------------------------------------ #

    def edit_rows(
        self,
        sheet_name: str,
        *,
        deletes=(),
        inserts=(),
        highlights=(),
        first_data_row: int = 1,
    ) -> None:
        if sheet_name not in self._sheet_paths:
            raise KeyError(f"工作表不存在: {sheet_name}")
        path = self._sheet_paths[sheet_name]
        root, decl, root_tag = _parse_part(self._parts[path])
        sheet_data = root.find(_q("sheetData"))
        if sheet_data is None:
            raise ValueError("工作表缺少 sheetData")

        rows = {int(el.get("r")): el for el in sheet_data.findall(_q("row"))}
        delete_set = {row for row in deletes if row in rows}
        groups = sorted(inserts, key=lambda g: (g.anchor_row, g.new_block))

        extend_counts: dict[int, int] = {}
        block_counts: dict[int, int] = {}
        for group in groups:
            if not group.rows:
                continue
            bucket = block_counts if group.new_block else extend_counts
            bucket[group.anchor_row] = bucket.get(group.anchor_row, 0) + len(group.rows)

        rowmap = RowMap(delete_set, extend_counts, block_counts, span_threshold=first_data_row)

        shared = self._collect_shared_formulas(sheet_data)
        self._expand_shared_formulas(sheet_data, shared)

        # 模板行必须在重编号之前快照，否则复制到的是已经改过行号的公式
        templates = {}
        for group in groups:
            if group.rows and group.template_row not in templates:
                source = rows.get(group.template_row)
                if source is None:
                    raise ValueError(f"模板行不存在: {group.template_row}")
                templates[group.template_row] = copy.deepcopy(source)

        # 逐行重编号 / 重写公式 / 着色
        highlight_map: dict[int, list[Highlight]] = {}
        for item in highlights:
            highlight_map.setdefault(item.row, []).append(item)

        for old_row, element in rows.items():
            if old_row in delete_set:
                sheet_data.remove(element)
                continue
            new_row = rowmap.new_row(old_row)
            self._renumber_row(element, new_row, rowmap, sheet_name)
            for item in highlight_map.get(old_row, ()):
                self._apply_highlight(element, new_row, item)

        # 插入新行
        new_merges: list[str] = []
        for group in groups:
            if not group.rows:
                continue
            template = templates[group.template_row]
            extend_rows, block_rows = rowmap.inserted_rows_at(group.anchor_row)
            targets = block_rows if group.new_block else extend_rows
            # 同一锚点上可能有多个 group，按顺序取用未被占用的行号
            used = self._used.setdefault((sheet_name, group.anchor_row, group.new_block), 0)
            slots = targets[used : used + len(group.rows)]
            self._used[(sheet_name, group.anchor_row, group.new_block)] = used + len(group.rows)
            created: list[ET.Element] = []
            for spec, new_row in zip(group.rows, slots):
                element = self._build_row(
                    template, group.template_row, new_row, rowmap, sheet_name, spec
                )
                sheet_data.append(element)
                created.append(element)
            if group.new_block and created:
                if group.block_label is not None:
                    cell = self._ensure_cell(
                        created[0], group.block_col, slots[0], template
                    )
                    self._set_cell_value(cell, group.block_label)
                col = group.block_col
                new_merges.append(f"{col}{slots[0]}:{col}{slots[-1]}")

        # 按行号排序（插入的新行是 append 上去的）
        ordered = sorted(sheet_data.findall(_q("row")), key=lambda el: int(el.get("r")))
        for element in list(sheet_data):
            sheet_data.remove(element)
        for element in ordered:
            sheet_data.append(element)

        last_row = int(ordered[-1].get("r")) if ordered else 1
        self._remap_sheet_metadata(root, rowmap, new_merges, last_row)
        self._parts[path] = _serialize_part(root, decl, root_tag)
        self._remap_defined_names(sheet_name, rowmap)
        self._drop_calc_chain()
        self._force_full_calc()

    # -- 共享公式 ---------------------------------------------------------- #

    def _collect_shared_formulas(self, sheet_data: ET.Element) -> dict[str, tuple[str, int, int]]:
        """收集 si -> (公式文本, 主单元格列序号, 主单元格行号)。"""
        masters = {}
        for row in sheet_data.findall(_q("row")):
            for cell in row.findall(_q("c")):
                formula = cell.find(_q("f"))
                if formula is None or formula.get("t") != "shared":
                    continue
                si = formula.get("si")
                if si is None or not (formula.text or "").strip():
                    continue
                col, row_index = split_coord(cell.get("r"))
                masters[si] = (formula.text, col_to_index(col), row_index)
        return masters

    def _expand_shared_formulas(self, sheet_data: ET.Element, masters: dict) -> None:
        """把共享公式全部展开成普通公式，避免行号变化后 si/ref 失效。"""
        for row in sheet_data.findall(_q("row")):
            for cell in row.findall(_q("c")):
                formula = cell.find(_q("f"))
                if formula is None or formula.get("t") != "shared":
                    continue
                si = formula.get("si")
                master = masters.get(si)
                if master is None:
                    continue
                text, master_col, master_row = master
                col, row_index = split_coord(cell.get("r"))
                formula.text = translate_formula(
                    text, row_index - master_row, col_to_index(col) - master_col
                )
                for attr in ("t", "si", "ref"):
                    formula.attrib.pop(attr, None)

    # -- 行/单元格构造 ------------------------------------------------------ #

    def _renumber_row(
        self, element: ET.Element, new_row: int, rowmap: RowMap, sheet_name: str
    ) -> None:
        element.set("r", str(new_row))
        for cell in element.findall(_q("c")):
            col, _ = split_coord(cell.get("r"))
            cell.set("r", f"{col}{new_row}")
            formula = cell.find(_q("f"))
            if formula is not None and formula.text:
                formula.text = remap_formula(formula.text, rowmap, sheet_name)

    def _build_row(
        self,
        template: ET.Element,
        template_row: int,
        new_row: int,
        rowmap: RowMap,
        sheet_name: str,
        spec: NewRow,
    ) -> ET.Element:
        element = copy.deepcopy(template)
        element.set("r", str(new_row))
        # 模板行的同行引用应指向新行，其余引用照常映射
        local = _LocalRowMap(rowmap, template_row, new_row)
        for cell in element.findall(_q("c")):
            col, _ = split_coord(cell.get("r"))
            cell.set("r", f"{col}{new_row}")
            formula = cell.find(_q("f"))
            if formula is not None and formula.text:
                formula.text = remap_formula(formula.text, local, sheet_name)
                self._clear_cached_value(cell)
            else:
                self._clear_cell(cell)
        for col, value in spec.values.items():
            cell = self._ensure_cell(element, col, new_row, template)
            self._set_cell_value(cell, value)
        for col, argb in spec.fills.items():
            cell = self._ensure_cell(element, col, new_row, template)
            base = int(cell.get("s") or 0)
            cell.set("s", str(self.styles.styled(base, argb)))
        return element

    @staticmethod
    def _clear_cached_value(cell: ET.Element) -> None:
        for tag in ("v", "is"):
            child = cell.find(_q(tag))
            if child is not None:
                cell.remove(child)
        cell.attrib.pop("t", None)

    @staticmethod
    def _clear_cell(cell: ET.Element) -> None:
        for tag in ("f", "v", "is"):
            child = cell.find(_q(tag))
            if child is not None:
                cell.remove(child)
        cell.attrib.pop("t", None)

    @staticmethod
    def _find_cell(row: ET.Element, col: str) -> ET.Element | None:
        for cell in row.findall(_q("c")):
            if split_coord(cell.get("r"))[0] == col:
                return cell
        return None

    def _ensure_cell(
        self, row: ET.Element, col: str, row_index: int, template: ET.Element
    ) -> ET.Element:
        existing = self._find_cell(row, col)
        if existing is not None:
            return existing
        cell = ET.Element(_q("c"))
        cell.set("r", f"{col}{row_index}")
        source = self._find_cell(template, col)
        if source is not None and source.get("s"):
            cell.set("s", source.get("s"))
        target = col_to_index(col)
        position = len(row)
        for index, other in enumerate(row.findall(_q("c"))):
            if col_to_index(split_coord(other.get("r"))[0]) > target:
                position = index
                break
        row.insert(position, cell)
        return cell

    @staticmethod
    def _set_cell_value(cell: ET.Element, value) -> None:
        XlsxEditor._clear_cell(cell)
        if value is None or value == "":
            return
        if isinstance(value, bool):
            cell.set("t", "b")
            ET.SubElement(cell, _q("v")).text = "1" if value else "0"
            return
        if isinstance(value, _dt.datetime):
            value = value.date()
        if isinstance(value, _dt.date):
            ET.SubElement(cell, _q("v")).text = str(date_to_excel_serial(value))
            return
        if isinstance(value, (int, float)):
            ET.SubElement(cell, _q("v")).text = repr(value) if isinstance(value, float) else str(value)
            return
        cell.set("t", "inlineStr")
        is_el = ET.SubElement(cell, _q("is"))
        text_el = ET.SubElement(is_el, _q("t"))
        text = str(value)
        if text != text.strip():
            text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text_el.text = text

    def _apply_highlight(self, row: ET.Element, row_index: int, item: Highlight) -> None:
        for col in item.cols:
            cell = self._find_cell(row, col)
            if cell is None:
                cell = ET.Element(_q("c"))
                cell.set("r", f"{col}{row_index}")
                row.append(cell)
            base = int(cell.get("s") or 0)
            cell.set("s", str(self.styles.styled(base, item.color)))

    # -- 工作表其余元素 ---------------------------------------------------- #

    def _remap_sheet_metadata(
        self, root: ET.Element, rowmap: RowMap, new_merges: list[str], last_row: int
    ) -> None:
        dimension = root.find(_q("dimension"))
        if dimension is not None and dimension.get("ref"):
            ref = dimension.get("ref")
            if ":" in ref:
                left, right = ref.split(":", 1)
                col_right, _ = split_coord(right)
                dimension.set("ref", f"{left}:{col_right}{last_row}")

        merge = root.find(_q("mergeCells"))
        if merge is not None:
            for cell in merge.findall(_q("mergeCell")):
                cell.set("ref", remap_sqref(cell.get("ref"), rowmap))
            for ref in new_merges:
                ET.SubElement(merge, _q("mergeCell")).set("ref", ref)
            merge.set("count", str(len(merge.findall(_q("mergeCell")))))
        elif new_merges:
            merge = ET.Element(_q("mergeCells"))
            for ref in new_merges:
                ET.SubElement(merge, _q("mergeCell")).set("ref", ref)
            merge.set("count", str(len(new_merges)))
            self._insert_after(root, merge, ("sheetData",))

        auto_filter = root.find(_q("autoFilter"))
        if auto_filter is not None and auto_filter.get("ref"):
            auto_filter.set("ref", remap_sqref(auto_filter.get("ref"), rowmap))

        for element in root.findall(_q("conditionalFormatting")):
            if element.get("sqref"):
                element.set("sqref", remap_sqref(element.get("sqref"), rowmap))

        validations = root.find(_q("dataValidations"))
        if validations is not None:
            for element in validations.findall(_q("dataValidation")):
                if element.get("sqref"):
                    element.set("sqref", remap_sqref(element.get("sqref"), rowmap))

        hyperlinks = root.find(_q("hyperlinks"))
        if hyperlinks is not None:
            for element in hyperlinks.findall(_q("hyperlink")):
                if element.get("ref"):
                    element.set("ref", remap_sqref(element.get("ref"), rowmap))

    @staticmethod
    def _insert_after(root: ET.Element, element: ET.Element, after_tags: tuple[str, ...]) -> None:
        children = list(root)
        position = len(children)
        for index, child in enumerate(children):
            if child.tag.split("}")[-1] in after_tags:
                position = index + 1
        root.insert(position, element)

    def _remap_defined_names(self, sheet_name: str, rowmap: RowMap) -> None:
        path = "xl/workbook.xml"
        root, decl, root_tag = _parse_part(self._parts[path])
        names = root.find(_q("definedNames"))
        if names is None:
            return
        changed = False
        for element in names.findall(_q("definedName")):
            if not element.text:
                continue
            updated = remap_formula(element.text, rowmap, sheet_name)
            if updated != element.text:
                element.text = updated
                changed = True
        if changed:
            self._parts[path] = _serialize_part(root, decl, root_tag)

    def _force_full_calc(self) -> None:
        path = "xl/workbook.xml"
        text = self._parts[path].decode("utf-8")
        if "fullCalcOnLoad" in text:
            return
        if "<calcPr" in text:
            text = re.sub(
                r"<calcPr\b([^/>]*?)\s*/>", r'<calcPr\1 fullCalcOnLoad="1"/>', text, count=1
            )
        else:
            text = text.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self._parts[path] = text.encode("utf-8")

    def _drop_calc_chain(self) -> None:
        """行数变化后旧 calcChain 会触发 Excel 修复提示，直接移除。"""
        target = "xl/calcChain.xml"
        if target not in self._parts:
            return
        self._parts.pop(target)
        self._order = [name for name in self._order if name != target]
        types = self._parts["[Content_Types].xml"].decode("utf-8")
        types = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^/]*/>', "", types)
        self._parts["[Content_Types].xml"] = types.encode("utf-8")
        rels_path = "xl/_rels/workbook.xml.rels"
        rels = self._parts[rels_path].decode("utf-8")
        rels = re.sub(r'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', "", rels)
        self._parts[rels_path] = rels.encode("utf-8")

    # -- 输出 -------------------------------------------------------------- #

    def to_bytes(self) -> bytes:
        if self._styles is not None and self._styles.dirty:
            self._parts["xl/styles.xml"] = self._styles.to_bytes()
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in self._order:
                archive.writestr(name, self._parts[name])
        return buffer.getvalue()


class _LocalRowMap:
    """把模板行的同行引用改指向新行，其余引用沿用原 rowmap。"""

    def __init__(self, base: RowMap, template_row: int, new_row: int):
        self._base = base
        self._template_row = template_row
        self._new_row = new_row

    def map_single(self, row: int) -> int:
        if row == self._template_row:
            return self._new_row
        return self._base.map_single(row)

    def map_range_start(self, row: int) -> int:
        if row == self._template_row:
            return self._new_row
        return self._base.map_range_start(row)

    def map_range_end(self, row: int, span_start: int | None = None) -> int:
        if row == self._template_row and span_start == self._template_row:
            return self._new_row
        return self._base.map_range_end(row, span_start=span_start)
