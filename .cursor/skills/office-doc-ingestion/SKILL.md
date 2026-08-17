---
name: office-doc-ingestion
description: Use when an app must ingest a mixed pile of user files (xlsx/xls/csv/docx/pdf, folders, zips) into one queryable store. Covers safe zip extraction, multi-header sheets, merged cells, and choosing the store.
---

# 混合文档批量入库

## 存储选型

| 场景 | 选择 |
| --- | --- |
| 需要 SQL 交叉查询、要给业务方下载"一个数据库文件" | **SQLite**（标准库，单文件，零部署） |
| 只在一次会话内计算，不需要持久化 | dict of DataFrame |
| 列结构完全不可预知、要全文检索 | SQLite + FTS5 或长表 `(file, sheet, row, col, value)` |

推荐结构：每个 sheet 一张表（列名清洗+去重）+ `_files` / `_sheets` 元数据表 + 文本类文件进 `_documents`。
既能 `st.dataframe` 直接展示，也能开个 SQL 输入框给高级用户。

## zip 解压必须防目录穿越

```python
for m in zf.infolist():
    if m.is_dir():
        continue
    p = os.path.normpath(m.filename)
    if p.startswith(("/", "..")) or os.path.isabs(p):
        continue          # 跳过恶意路径
```

同时限制单文件解压体积和总条目数，防 zip bomb。zip 里的中文文件名常是 GBK 编码，
`m.flag_bits & 0x800 == 0` 时用 `m.filename.encode("cp437").decode("gbk", "replace")` 还原。

## Excel 读取的三个必做处理

1. **表头行不一定是第 1 行**：向下扫描前若干行，取"包含预期关键列名最多"的那一行当表头，
   而不是硬编码 `header=0` 或 `skiprows=1`。
2. **纵向合并单元格**：`openpyxl` 读到的合并区只有左上角有值，其余是 `None`。
   分组列必须**向下填充**（forward fill），否则分组全丢。
   需要写回时还要单独保留"块边界"（合并区的 min_row/max_row）。
3. **公式 vs 值**：`data_only=True` 只有 Excel 缓存过才有值；要改写文件必须 `data_only=False`。
   需要两者时就加载两次，别指望一次拿全。

`.xls`（老格式）需要 `xlrd<2` 或先转换；`pandas` 新版已不支持 xlrd 读 xlsx。

## PDF / Word

- Word 表格：`python-docx` 的 `doc.tables[i].rows[j].cells[k].text`；合并单元格会重复出现同一个 cell 对象，按 `id()` 去重。
- PDF 文字型：`pdfplumber.extract_table()` 够用；扫描件必须先 OCR，不要假装能解析。
- 两者都要把"提取失败"作为一等结果记录下来（文件名 + 原因），而不是静默跳过——
  用户上传 30 个文件时，静默跳过 = 数据缺失事故。

## 幂等与追踪

每个文件算 sha256 存进 `_files`，重复上传直接跳过。入库结果一定要有一张
"本次入库清单"（文件名 / 类型 / 表数 / 行数 / 状态），这是用户建立信任的唯一途径。
