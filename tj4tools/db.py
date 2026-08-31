"""把入库结果写成一个 SQLite 文件。

选 SQLite 的理由：标准库自带、单文件可下载、支持 SQL 交叉查询，
业务方拿到 .db 也能用任意工具打开。

结构：
  ``_files``      本次入库的文件清单
  ``_sheets``     每张表的元数据（原文件、原 sheet 名、表名、行列数）
  ``_documents``  PDF/Word 的纯文本
  ``t_<n>``       每张二维表一张物理表，列名取清洗后的表头
"""

from __future__ import annotations

import datetime as _dt
import re
import sqlite3
from dataclasses import dataclass

from .ingest import IngestResult, SheetTable
from .normalize import clean_text

_IDENT_BAD = re.compile(r"[^0-9A-Za-z_\u4e00-\u9fff]+")


def safe_identifier(name: str, fallback: str = "col") -> str:
    cleaned = _IDENT_BAD.sub("_", clean_text(name)).strip("_")
    if not cleaned or cleaned[0].isdigit():
        cleaned = f"{fallback}_{cleaned}" if cleaned else fallback
    return cleaned[:60]


def _quote(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def _cell_to_sql(value):
    if value is None:
        return None
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()[:10] if isinstance(value, _dt.date) and not isinstance(value, _dt.datetime) else value.isoformat(sep=" ")
    if isinstance(value, (int, float, str, bytes)):
        return value
    return str(value)


@dataclass
class TableInfo:
    table_name: str
    file_name: str
    sheet_name: str
    n_rows: int
    n_cols: int
    columns: list[str]


class Workspace:
    """内存 SQLite 工作区，可 dump 成字节供下载。"""

    def __init__(self):
        self.conn = sqlite3.connect(":memory:", check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.tables: list[TableInfo] = []
        self._init_schema()

    def _init_schema(self) -> None:
        self.conn.executescript(
            """
            CREATE TABLE _files (
                file_name TEXT, sha256 TEXT, size INTEGER, kind TEXT, n_tables INTEGER
            );
            CREATE TABLE _sheets (
                table_name TEXT, file_name TEXT, sheet_name TEXT,
                header_row INTEGER, n_rows INTEGER, n_cols INTEGER, columns TEXT
            );
            CREATE TABLE _documents (file_name TEXT, kind TEXT, text TEXT);
            CREATE TABLE _problems (message TEXT);
            """
        )
        self.conn.commit()

    # -- 写入 -------------------------------------------------------------- #

    def load(self, result: IngestResult) -> None:
        for entry in result.files:
            self.conn.execute(
                "INSERT INTO _files VALUES (?,?,?,?,?)",
                (entry["file_name"], entry["sha256"], entry["size"], entry["kind"], entry["n_tables"]),
            )
        for document in result.documents:
            self.conn.execute(
                "INSERT INTO _documents VALUES (?,?,?)",
                (document.file_name, document.kind, document.text),
            )
        for message in result.problems:
            self.conn.execute("INSERT INTO _problems VALUES (?)", (message,))
        used: set[str] = set()
        for index, table in enumerate(result.tables, start=1):
            self.add_table(table, index, used)
        self.conn.commit()

    def add_table(self, table: SheetTable, index: int, used: set[str]) -> TableInfo:
        base = safe_identifier(f"t{index}_{table.sheet_name}", f"t{index}")
        name = base
        suffix = 2
        while name in used:
            name = f"{base}_{suffix}"
            suffix += 1
        used.add(name)

        columns: list[str] = []
        seen: dict[str, int] = {}
        for position, raw in enumerate(table.header):
            column = safe_identifier(raw, f"c{position + 1}")
            if column in seen:
                seen[column] += 1
                column = f"{column}_{seen[column]}"
            else:
                seen[column] = 1
            columns.append(column)
        if not columns:
            columns = ["c1"]

        ddl = ", ".join(f"{_quote(column)} TEXT" for column in columns)
        self.conn.execute(f"CREATE TABLE {_quote(name)} (_row INTEGER, {ddl})")
        placeholders = ",".join(["?"] * (len(columns) + 1))
        payload = []
        for offset, row in enumerate(table.rows):
            values = list(row[: len(columns)]) + [None] * max(0, len(columns) - len(row))
            payload.append([table.header_row_index + 1 + offset] + [_cell_to_sql(v) for v in values])
        if payload:
            self.conn.executemany(f"INSERT INTO {_quote(name)} VALUES ({placeholders})", payload)
        info = TableInfo(
            table_name=name,
            file_name=table.file_name,
            sheet_name=table.sheet_name,
            n_rows=len(table.rows),
            n_cols=len(columns),
            columns=columns,
        )
        self.tables.append(info)
        self.conn.execute(
            "INSERT INTO _sheets VALUES (?,?,?,?,?,?,?)",
            (
                name,
                table.file_name,
                table.sheet_name,
                table.header_row_index,
                info.n_rows,
                info.n_cols,
                "|".join(columns),
            ),
        )
        return info

    # -- 读取 -------------------------------------------------------------- #

    def query(self, sql: str) -> tuple[list[str], list[tuple]]:
        cursor = self.conn.execute(sql)
        headers = [d[0] for d in cursor.description] if cursor.description else []
        return headers, cursor.fetchall()

    def to_bytes(self) -> bytes:
        buffer = bytearray()
        for line in self.conn.iterdump():
            buffer.extend((line + "\n").encode("utf-8"))
        return bytes(buffer)

    def to_sqlite_file(self, path: str) -> None:
        target = sqlite3.connect(path)
        with target:
            self.conn.backup(target)
        target.close()

    def to_sqlite_bytes(self) -> bytes:
        """序列化为真正的 .db 字节（优先用 serialize，回退到临时文件）。"""
        serialize = getattr(self.conn, "serialize", None)
        if serialize is not None:
            return serialize()
        import os
        import tempfile

        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "workspace.db")
            self.to_sqlite_file(path)
            with open(path, "rb") as handle:
                return handle.read()
