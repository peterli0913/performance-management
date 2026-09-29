---
name: xlsx-format-preserving-edit
description: Use when inserting/deleting/highlighting rows in an existing .xlsx that must keep its exact formatting, formulas, filters, conditional formatting and other sheets. Covers why openpyxl round-trips lose data and how to edit the OOXML parts directly.
---

# 保格式编辑 xlsx（行插入/删除/着色）

## 何时用

- 需求里出现"格式全部一样""保留公式和筛选""保留其他子表"。
- 目标文件是业务方长期维护的模板（含条件格式、外部链接、打印设置、批注）。

## 铁律

```
先用 zipfile 列出 xlsx 内部 part 清单，再决定用 openpyxl 还是改 XML。
```

## 第一步：体检（必做，成本极低）

```python
import zipfile
with zipfile.ZipFile(path) as z:
    for n in z.namelist():
        print(n, z.getinfo(n).file_size)
```

出现下列任一 part，**不要**用 openpyxl 读写往返：

| part | openpyxl 往返后果 |
| --- | --- |
| `xl/externalLinks/*` | 外部链接被丢弃，`[1]Sheet!$C:$C` 类公式全部失效 |
| `xl/printerSettings/*` | 打印设置丢失 |
| `xl/customProperty*.bin` | `customProperties` 丢失，Excel 可能提示修复 |
| `xl/drawings/vmlDrawing*.vml` | 批注框/图形丢失 |
| `xl/media/*`, `xl/charts/*` | 图片和图表丢失 |
| `xl/pivotCache/*` | 数据透视丢失 |

只有当清单里仅有 `workbook.xml / worksheets / styles.xml / sharedStrings.xml / theme` 时，openpyxl 往返才是安全的。

## 第二步：即使用 openpyxl，行插入也是坑

`ws.insert_rows()` **不会**：调整公式引用、调整合并单元格、调整条件格式 sqref、调整 autoFilter/Print_Area、复制样式和行高。
所以无论走哪条路，都得自己写行号重映射逻辑——那就干脆直接改 XML。

## 第三步：XML 级行手术的正确做法

核心是**三张行号映射表**，而不是一个简单 delta：

```
map_single[old]      -> 单点引用、行 r 属性
map_range_start[old] -> 区间起点
map_range_end[old]   -> 区间终点；若 old 是插入锚点，则 = new(old) + 插入行数
```

`map_range_end` 是关键：车间块末尾追加人员时，`A3:A73` 合并区、`COUNTIF($L$3:$AP$73,"a")`
区间公式、`Print_Area $A$1:$F$675` 都必须把终点顺延，否则新增行落在统计范围之外。
被删除的行做钳制：区间起点取下一个存活行，区间终点取上一个存活行。

必须一起重映射的位置（漏一个 Excel 就报修复）：

1. `<row r>` 和每个 `<c r>`
2. `<f>` 公式里的引用（跳过带 `!` 的他表引用和 `[n]` 外链引用）
3. `<mergeCells>`
4. `<conditionalFormatting sqref>`、`<dataValidation sqref>`、`<hyperlink ref>`
5. `<autoFilter ref>`、`<dimension ref>`
6. `workbook.xml` 的 `<definedNames>`（`_xlnm._FilterDatabase`、`_xlnm.Print_Area`）
7. 同表的 `xl/comments*.xml` 的 `ref`（本表有批注时）

## 公式引用正则的两个陷阱

```python
REF = re.compile(
    r"(?P<sheet>'[^']*'!|\[\d+\][^\s!,()+\-*/&=<>:]*!|[A-Za-z0-9_.\u4e00-\u9fff]+!)?"
    r"(?<![A-Za-z0-9_$.])(?P<col>\$?[A-Z]{1,3})(?P<row>\$?\d{1,7})(?![\d(])"
)
```

- `(?![\d(])`：否则 `LOG10(` 会被当成单元格 `LOG10`。
- `(?<![A-Za-z0-9_$.])`：否则函数名尾部会被误匹配。
- 先处理 `A1:B2` 区间再处理单点，才能区分 start/end 语义。
- 处理前先把字符串字面量 `"..."` 挖出来占位，避免改到文本。

## 新增行怎么来

深拷贝一个模板行（通常是同块最后一行）的 `<row>` 元素，然后：

- 改 `r` 和所有 `<c r>`；
- 公式用"临时映射"重写：把 `map_single[模板行] = 新行号` 后跑一遍 remap，
  这样同行引用变新行、跨块区间引用仍然正确；
- 业务列写值，其余列清空 `<v>/<f>` 但**保留 `s`（样式索引）**，行高来自模板行的 `ht/customHeight`。

## 写值不要动 sharedStrings

用内联字符串，零副作用：

```xml
<c r="C74" s="123" t="inlineStr"><is><t>张三</t></is></c>
```

日期写序列号（`(date - 1899-12-30).days`）并保留模板样式，格式自然正确。

## 填充色

往 `styles.xml` 的 `<fills>` 追加 solid fill（追加不会改变已有索引，`dxfs` 不受影响），
再按 `(原 xf 索引, fillId)` 克隆 `<cellXfs>` 条目并缓存，避免样式表膨胀。

## 收尾三件事

1. 删除 `xl/calcChain.xml`，同时删掉 `[Content_Types].xml` 的 Override 和 `workbook.xml.rels` 的 Relationship——行数变化后旧 calcChain 会触发"修复"提示。
2. `workbook.xml` 的 `<calcPr>` 加 `fullCalcOnLoad="1"`，打开即重算。
3. 重写 zip 时**逐条复制未修改的 part 原始字节**，只替换改过的，其他一律不碰。

## 验证（便宜且有效）

```python
wb = openpyxl.load_workbook(out, data_only=False)   # 能加载 = XML 基本合法
```

再断言：行数、目标行的值、`F{n}` 公式里的行号等于 `n`、合并区终点已顺延、
`zip.namelist()` 与原文件的差集只有 `calcChain.xml`。
