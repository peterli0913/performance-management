---
name: streamlit-data-app-patterns
description: Use when building or reviewing a Streamlit app that uploads files, reviews hundreds of rows, or produces downloadable Excel/PDF output. Covers caching, per-row approve/reject UX at scale, and state that survives reruns.
---

# Streamlit 数据处理 App 模式

## 何时用

上传文件 → 计算 → 人工逐条复核 → 导出文件，这类内部工具。

## 状态与缓存

- **重计算全部走 `@st.cache_data`**，key 用文件字节的 sha256，不要用文件名：

```python
@st.cache_data(show_spinner=False)
def build_workspace(payload: tuple[tuple[str, bytes], ...]):
    ...
```

- **人工决策存 `st.session_state`**，且用稳定业务主键（如 `姓名|员工编号`）做 key，
  绝不能用行号——排序或筛选一变行号就错位。
- 上传控件返回的 `UploadedFile` 每次 rerun 都是新对象，读一次 `.getvalue()` 存成 bytes 再传给缓存函数。

## 几百行的逐条审批 UX

`st.button` 一行两个按钮 × 300 行 = 600 个 widget，会明显卡。正确做法是三层：

1. **批量层**：`全部应用 / 全部取消 / 重置` 按钮 + 按分类/车间的批量操作。
2. **表格层**：`st.data_editor` 只放一个 `CheckboxColumn`（或三态 `SelectboxColumn`），
   其余列 `disabled=True`。一次 rerun 提交所有改动，性能与行数几乎无关。
3. **逐条层**：分页（每页 20-25 行）后再渲染 per-row 按钮，满足"每人一个按钮"的需求
   而不让 widget 总数爆炸。

`data_editor` 回写要用返回值 diff，而不是原地改 DataFrame：

```python
edited = st.data_editor(view, key=f"editor_{page}", hide_index=True,
                        column_config={"应用": st.column_config.CheckboxColumn()},
                        disabled=[c for c in view.columns if c != "应用"])
for key, val in zip(view["_key"], edited["应用"]):
    st.session_state.decisions[key] = "应用" if val else "取消"
```

## 下载按钮

`st.download_button` 会在**每次 rerun 时**执行它的 `data=` 参数。生成 Excel 很贵，所以：

- 先用一个普通 `st.button("生成")` 把 bytes 算出来放进 `session_state`，
- 再用 `st.download_button(data=st.session_state.blob)` 提供下载。

否则每次点任何控件都会重算一遍导出文件。

## 大表展示

- `st.dataframe(df, use_container_width=True, hide_index=True)` + 显式 `column_config` 控制列宽。
- 超过 ~5000 行先聚合或分页，不要指望浏览器。
- 日期统一先格式化成 `YYYY-MM-DD` 字符串再展示，避免时区/NaT 噪音。

## 结构

```
app.py                # 只做 UI 编排，不放业务逻辑
<pkg>/                # 纯函数业务层，可被 pytest 直接测
  ingest.py  db.py  <domain>.py  xlsx_surgery.py
tests/                # 不依赖 Streamlit
```

UI 层零业务逻辑是能便宜验证的前提：所有断言都在 pytest 里跑，不需要起浏览器。

## 部署到 Streamlit Community Cloud

- `requirements.txt` 固定主版本；`openpyxl`、`pandas` 必列。
- 默认 200MB 上传上限，超了在 `.streamlit/config.toml` 调 `maxUploadSize`。
- 不要往仓库目录写临时文件，用 `tempfile` 或内存 `BytesIO`；容器文件系统是易失的。
