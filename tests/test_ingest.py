import io
import zipfile

from tj4tools.db import Workspace, safe_identifier
from tj4tools.ingest import ingest_files, pick_header_row


def test_ingest_two_sample_workbooks(roster_bytes, bonus_bytes):
    result = ingest_files(
        [
            ("TJ4生产部&生产设备部人员清单，07-31-2026.xlsx", roster_bytes),
            ("2026年07月份安全质量奖核算数据.xlsx", bonus_bytes),
        ]
    )
    assert not result.problems
    assert len(result.files) == 2
    labels = {table.sheet_name for table in result.tables}
    assert {"生产部", "离职人员&调出人员", "人员变动说明", "一线人员", "副主任&工艺组长及其他"} <= labels
    production = next(t for t in result.tables if t.sheet_name == "生产部")
    assert production.header_row_index == 1
    assert production.n_rows == 1151
    # 「人员变动说明」第 1 行是合并标题，表头必须被探测到第 2 行
    changes = next(t for t in result.tables if t.sheet_name == "人员变动说明")
    assert changes.header_row_index == 2
    assert "姓名" in changes.header


def test_frontline_workshop_is_filled_down_on_ingest(bonus_bytes):
    result = ingest_files([("bonus.xlsx", bonus_bytes)])
    frontline = next(t for t in result.tables if t.sheet_name == "一线人员")
    # 表头是第 1 行（第 2 行全是日期，不能被当成表头）
    assert frontline.header_row_index == 1
    assert frontline.header[:5] == ["车间", "职务", "姓名", "员工编号", "入职日期"]
    people = [row for row in frontline.rows if row[2]]
    assert len(people) == 673
    assert all(row[0] for row in people), "合并的车间列必须向下填充"


def test_zip_is_expanded_and_deduplicated(roster_bytes, bonus_bytes):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("子目录/清单，07-31-2026.xlsx", roster_bytes)
        archive.writestr("子目录/核算.xlsx", bonus_bytes)
        archive.writestr("../坏路径.xlsx", b"x")
    result = ingest_files([("上传.zip", buffer.getvalue())])
    names = [entry["file_name"] for entry in result.files]
    assert any("清单" in name for name in names)
    assert any("核算" in name for name in names)
    assert any("路径非法" in problem or "解析失败" in problem for problem in result.problems)


def test_duplicate_upload_is_skipped(bonus_bytes):
    result = ingest_files([("a.xlsx", bonus_bytes), ("b.xlsx", bonus_bytes)])
    assert len(result.files) == 1
    assert any("重复文件" in problem for problem in result.problems)


def test_unsupported_type_is_reported():
    result = ingest_files([("说明.md", b"# hi")])
    # 仍然登记在清单里，只是没有表；静默丢弃才是事故
    assert [entry["n_tables"] for entry in result.files] == [0]
    assert any("暂不支持" in problem for problem in result.problems)


def test_broken_file_does_not_abort_batch(bonus_bytes):
    result = ingest_files([("坏的.xlsx", b"not an xlsx"), ("好的.xlsx", bonus_bytes)])
    assert any("解析失败" in problem for problem in result.problems)
    assert len(result.files) == 2
    assert next(e for e in result.files if e["file_name"] == "坏的.xlsx")["n_tables"] == 0
    assert next(e for e in result.files if e["file_name"] == "好的.xlsx")["n_tables"] == 8


def test_csv_with_gbk_encoding():
    data = "姓名,员工编号\n张三,ALS1\n".encode("gb18030")
    result = ingest_files([("名单.csv", data)])
    assert not result.problems
    table = result.tables[0]
    assert table.header == ["姓名", "员工编号"]
    assert table.rows[0][0] == "张三"


def test_pick_header_row_skips_title_row():
    grid = [["合并标题", None, None], ["序号", "姓名", "员工号"], [1, "张三", "A1"]]
    assert pick_header_row(grid) == 1


def test_safe_identifier():
    assert safe_identifier("离职人员&调出人员") == "离职人员_调出人员"
    assert safe_identifier("公司\n编制") == "公司_编制"
    assert safe_identifier("") == "col"
    assert safe_identifier("2025年").startswith("col")


def test_workspace_roundtrip(roster_bytes, bonus_bytes):
    result = ingest_files([("清单.xlsx", roster_bytes), ("核算.xlsx", bonus_bytes)])
    workspace = Workspace()
    workspace.load(result)
    assert len(workspace.tables) == len(result.tables)
    headers, rows = workspace.query("SELECT file_name, sheet_name, n_rows FROM _sheets")
    assert headers == ["file_name", "sheet_name", "n_rows"]
    assert len(rows) == len(result.tables)
    production = next(t for t in workspace.tables if t.sheet_name == "生产部")
    _, count = workspace.query(f'SELECT COUNT(*) FROM "{production.table_name}"')
    assert count[0][0] == 1151
    # 跨表 SQL 查询要能跑通
    frontline = next(t for t in workspace.tables if t.sheet_name == "一线人员")
    _, joined = workspace.query(
        f'SELECT COUNT(*) FROM "{production.table_name}" p '
        f'JOIN "{frontline.table_name}" f ON p.姓名 = f.姓名 AND p.员工编号 = f.员工编号'
    )
    assert joined[0][0] > 500
    blob = workspace.to_sqlite_bytes()
    assert blob[:15] == b"SQLite format 3"
