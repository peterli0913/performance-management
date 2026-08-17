"""TJ4 安全质量奖人员对账工具（Streamlit）。

流程：上传文件/文件夹压缩包 → 统一入库 → 四类差异复核 → 生成新的核算表。
业务逻辑全部在 ``tj4tools`` 包里，本文件只做界面编排。
"""

from __future__ import annotations

import datetime as dt
import os

import pandas as pd
import streamlit as st

from tj4tools.bonus_export import build_workbook
from tj4tools.db import Workspace
from tj4tools.ingest import ingest_files
from tj4tools.normalize import date_from_filename
from tj4tools.roster import (
    ADD_CATEGORIES,
    CATEGORIES,
    CATEGORY_LEFT,
    CATEGORY_NEW,
    CATEGORY_PENDING_ADD,
    CATEGORY_PENDING_DEL,
    build_workshop_mapping,
    parse_bonus,
    parse_roster,
    reconcile,
)

NEW_BLOCK_PREFIX = "＋新建车间块："
UNSET = "（未指定）"
UNDECIDED, APPLY, CANCEL = "待定", "应用", "取消"
PAGE_SIZE = 20

st.set_page_config(page_title="TJ4 安全质量奖人员对账", page_icon="📊", layout="wide")


# --------------------------------------------------------------------------- #
# 缓存层：重计算按文件内容做 key，避免每次交互都重跑
# --------------------------------------------------------------------------- #


@st.cache_data(show_spinner="正在读取文件…")
def cached_ingest(payload: tuple[tuple[str, bytes], ...]):
    return ingest_files(list(payload))


@st.cache_resource(show_spinner="正在建立综合数据库…")
def cached_workspace(payload: tuple[tuple[str, bytes], ...]) -> Workspace:
    # SQLite 连接不能被 pickle，所以用 cache_resource 而不是 cache_data
    workspace = Workspace()
    workspace.load(cached_ingest(payload))
    return workspace


@st.cache_data(show_spinner="正在解析人员清单与核算数据…")
def cached_analyze(
    roster_name: str,
    roster_bytes: bytes,
    bonus_name: str,
    bonus_bytes: bytes,
    duty_field: str,
    ref_date: dt.date | None,
    window_months: int,
    exclude_others: bool,
):
    roster = parse_roster(roster_bytes, roster_name, duty_field=duty_field)
    bonus = parse_bonus(bonus_bytes, bonus_name)
    mapping = build_workshop_mapping(roster, bonus)
    result = reconcile(
        roster,
        bonus,
        ref_date=ref_date,
        window_months=window_months,
        exclude_in_others=exclude_others,
        mapping=mapping,
    )
    return roster, bonus, result


@st.cache_data(show_spinner="正在导出数据库…")
def cached_db_bytes(payload: tuple[tuple[str, bytes], ...]) -> bytes:
    return cached_workspace(payload).to_sqlite_bytes()


# --------------------------------------------------------------------------- #
# 会话状态
# --------------------------------------------------------------------------- #


def init_state() -> None:
    st.session_state.setdefault("decisions", {})
    st.session_state.setdefault("workshop_override", {})
    st.session_state.setdefault("mapping_override", {})


def decision_of(label: str) -> str:
    return st.session_state["decisions"].get(label, UNDECIDED)


def set_decision(label: str, value: str) -> None:
    st.session_state["decisions"][label] = value


def bump(category: str) -> None:
    """改变 data_editor 的 key，让批量操作后的表格重新初始化。"""
    key = f"ver_{category}"
    st.session_state[key] = st.session_state.get(key, 0) + 1


# --------------------------------------------------------------------------- #
# 上传
# --------------------------------------------------------------------------- #


def collect_uploads() -> tuple[tuple[str, bytes], ...]:
    st.markdown(
        "上传 **人员清单** 与 **安全质量奖核算数据**（可多选）。"
        "整个文件夹请先打成 `.zip` 上传，压缩包会被自动递归展开。"
    )
    uploaded = st.file_uploader(
        "选择文件（xlsx / xls / csv / docx / pdf / zip，可多选）",
        type=["xlsx", "xlsm", "xls", "csv", "tsv", "txt", "docx", "pdf", "zip"],
        accept_multiple_files=True,
    )
    items = [(handle.name, handle.getvalue()) for handle in uploaded or []]
    samples = repo_samples()
    if samples and st.checkbox(f"同时载入仓库内的 {len(samples)} 个示例文件", value=not items):
        items.extend(samples)
    return tuple(items)


def repo_samples() -> list[tuple[str, bytes]]:
    here = os.path.dirname(os.path.abspath(__file__))
    out = []
    for name in sorted(os.listdir(here)):
        if not name.lower().endswith(".xlsx") or name.startswith("~$"):
            continue
        with open(os.path.join(here, name), "rb") as handle:
            out.append((name, handle.read()))
    return out


def guess_roles(result) -> tuple[str | None, str | None]:
    """猜哪份是人员清单、哪份是核算数据：先看子表名，再看文件名。"""
    sheets_by_file: dict[str, set[str]] = {}
    for table in result.tables:
        sheets_by_file.setdefault(table.file_name, set()).add(table.sheet_name)
    roster = bonus = None
    for name, sheets in sheets_by_file.items():
        if roster is None and ("生产部" in sheets or "人员清单" in name):
            roster = name
        if bonus is None and ("一线人员" in sheets or "安全质量奖" in name):
            bonus = name
    return roster, bonus


# --------------------------------------------------------------------------- #
# 车间映射
# --------------------------------------------------------------------------- #


def to_workshop(value: str) -> str:
    """把界面上的选项值还原成真正的车间名。"""
    if not value or value == UNSET:
        return ""
    return value[len(NEW_BLOCK_PREFIX) :] if value.startswith(NEW_BLOCK_PREFIX) else value


def to_option(value: str, options: list[str]) -> str:
    """把车间名转成下拉里合法的选项值。"""
    if not value:
        return UNSET
    return value if value in options else NEW_BLOCK_PREFIX + value


def build_options(bonus, groups) -> list[str]:
    """下拉选项 = 未指定 + 原有车间 + 「新建车间块」候选。"""
    return (
        [UNSET]
        + list(bonus.workshops)
        + [NEW_BLOCK_PREFIX + name for name in sorted({g for g in groups if g})]
    )


def effective_workshop(item, mapping) -> str:
    person = st.session_state["workshop_override"]
    group = st.session_state["mapping_override"]
    if item.label in person:
        return person[item.label]
    if item.group in group:
        return group[item.group]
    guess = mapping.get(item.group)
    return guess.workshop if guess else ""


def describe_guess(guess) -> str:
    if guess is None or not guess.workshop:
        return "需人工指定"
    if guess.source == "经验":
        return f"{guess.workshop}（两表已匹配 {guess.support} 人 · 置信度{guess.confidence}）"
    return f"{guess.workshop}（按命名规则推断 · 置信度{guess.confidence}）"


def render_mapping_editor(analysis, bonus, items, options) -> None:
    counts: dict[str, int] = {}
    for item in items:
        counts[item.group] = counts.get(item.group, 0) + 1
    if not counts:
        return
    overrides = st.session_state["mapping_override"]
    rows = []
    for group, count in sorted(counts.items(), key=lambda kv: -kv[1]):
        guess = analysis.mapping.get(group)
        current = overrides.get(group, guess.workshop if guess else "")
        rows.append(
            {
                "目前分组": group,
                "待新增": count,
                "自动建议": describe_guess(guess),
                "最终车间": to_option(current, options),
            }
        )
    frame = pd.DataFrame(rows)
    unknown = int((frame["最终车间"] == UNSET).sum())
    title = f"车间映射（{len(frame)} 个分组，"
    title += f"{unknown} 个待人工指定）" if unknown else "已全部对应）"
    with st.expander(title, expanded=bool(unknown)):
        st.caption(
            "「自动建议」由两表已匹配人员反推得出。置信度低或显示「需人工指定」的行请在右侧下拉里选择；"
            "选「＋新建车间块」会在一线人员子表最下方新建该分组。"
        )
        edited = st.data_editor(
            frame,
            hide_index=True,
            width="stretch",
            key="mapping_editor",
            column_config={
                "目前分组": st.column_config.TextColumn(width="medium"),
                "待新增": st.column_config.NumberColumn(width="small"),
                "自动建议": st.column_config.TextColumn(width="large"),
                "最终车间": st.column_config.SelectboxColumn(
                    options=options, required=True, width="medium"
                ),
            },
            disabled=["目前分组", "待新增", "自动建议"],
        )
        for group, before, after in zip(frame["目前分组"], frame["最终车间"], edited["最终车间"]):
            if after != before:
                overrides[group] = to_workshop(after)


# --------------------------------------------------------------------------- #
# 差异复核
# --------------------------------------------------------------------------- #


def render_category(category, items, bonus, mapping, options) -> None:
    if not items:
        st.success(f"没有「{category}」人员")
        return

    decisions = st.session_state["decisions"]
    applied = sum(1 for item in items if decisions.get(item.label) == APPLY)
    cancelled = sum(1 for item in items if decisions.get(item.label) == CANCEL)
    metrics = st.columns([1, 1, 1, 1, 3])
    metrics[0].metric("总数", len(items))
    metrics[1].metric("已应用", applied)
    metrics[2].metric("已取消", cancelled)
    metrics[3].metric("待定", len(items) - applied - cancelled)

    bulk = st.columns([1, 1, 1, 5])
    if bulk[0].button("全部应用", key=f"all_apply_{category}"):
        for item in items:
            decisions[item.label] = APPLY
        bump(category)
        st.rerun()
    if bulk[1].button("全部取消", key=f"all_cancel_{category}"):
        for item in items:
            decisions[item.label] = CANCEL
        bump(category)
        st.rerun()
    if bulk[2].button("清空决策", key=f"all_reset_{category}"):
        for item in items:
            decisions.pop(item.label, None)
        bump(category)
        st.rerun()

    mode = st.radio(
        "复核方式",
        ["表格批量", "逐条按钮"],
        horizontal=True,
        key=f"mode_{category}",
        help="人数多时用表格批量勾选更快；需要逐个确认时切到逐条按钮。",
    )
    editable = category in ADD_CATEGORIES
    if mode == "表格批量":
        render_table(category, items, mapping, options, editable)
    else:
        render_rows(category, items, mapping, editable)


def row_payload(item, mapping, options=None) -> dict:
    payload = item.as_dict()
    workshop = effective_workshop(item, mapping)
    payload["车间"] = to_option(workshop, options) if options else workshop
    return payload


def render_table(category, items, mapping, options, editable) -> None:
    """表格批量复核。

    「应用」和「取消」是两个独立复选框，并且无条件按编辑结果回写。
    这样表格状态与决策状态形成稳定不动点：既不会因为"没勾选"就把待定误判成取消，
    也不依赖"和上一帧比较"这种脆弱逻辑。
    """
    decisions = st.session_state["decisions"]
    overrides = st.session_state["workshop_override"]
    order = ["应用", "取消", "姓名", "员工编号", "职务", "车间", "目前分组", "入职时间",
             "离职时间", "离职/调出备注", "判定依据", "提示"]
    frame = pd.DataFrame(
        [
            {
                "应用": decisions.get(item.label) == APPLY,
                "取消": decisions.get(item.label) == CANCEL,
                **{k: v for k, v in row_payload(item, mapping, options).items() if k in order},
            }
            for item in items
        ]
    )[order]
    config = {
        "应用": st.column_config.CheckboxColumn("应用", width="small", help="纳入「已应用版」导出"),
        "取消": st.column_config.CheckboxColumn("取消", width="small", help="从两个导出里都剔除"),
        "姓名": st.column_config.TextColumn(width="small"),
        "员工编号": st.column_config.TextColumn(width="small"),
        "职务": st.column_config.TextColumn(width="small"),
        "判定依据": st.column_config.TextColumn(width="large"),
    }
    disabled = [column for column in order if column not in ("应用", "取消")]
    if editable:
        config["车间"] = st.column_config.SelectboxColumn(
            options=options, required=True, width="medium"
        )
        disabled.remove("车间")
    version = st.session_state.get(f"ver_{category}", 0)
    edited = st.data_editor(
        frame,
        hide_index=True,
        width="stretch",
        height=min(640, 90 + 35 * len(frame)),
        key=f"editor_{category}_{version}",
        column_config=config,
        disabled=disabled,
    )
    for index, item in enumerate(items):
        if bool(edited["应用"].iloc[index]):
            decisions[item.label] = APPLY
        elif bool(edited["取消"].iloc[index]):
            decisions[item.label] = CANCEL
        else:
            decisions.pop(item.label, None)
        if editable:
            now = edited["车间"].iloc[index]
            if now != frame["车间"].iloc[index]:
                overrides[item.label] = to_workshop(now)
    st.caption(
        "「应用」= 纳入「已应用版」导出；「取消」= 从两个导出里都剔除；"
        "两个都不勾 = 待定，只出现在「对照标记版」里。两个都勾时按「应用」处理。"
    )


def render_rows(category, items, mapping, editable) -> None:
    pages = max(1, (len(items) + PAGE_SIZE - 1) // PAGE_SIZE)
    page = 1
    if pages > 1:
        page = int(
            st.number_input(
                f"翻到第几页　（每页 {PAGE_SIZE} 人，共 {pages} 页 / {len(items)} 人）",
                min_value=1,
                max_value=pages,
                value=1,
                key=f"page_{category}",
            )
        )
        first = (page - 1) * PAGE_SIZE + 1
        st.caption(f"当前显示第 {first} ~ {min(page * PAGE_SIZE, len(items))} 人")
    badge = {APPLY: "✅ 已应用", CANCEL: "🚫 已取消", UNDECIDED: "⏳ 待定"}
    for item in items[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]:
        payload = row_payload(item, mapping)
        cols = st.columns([2.2, 1.6, 0.9, 1.1, 1.1, 0.8, 0.8, 1.0])
        cols[0].markdown(f"**{item.name}**　`{item.eid}`")
        cols[1].write(payload["车间"] or "待指定")
        cols[2].write(item.duty or "—")
        cols[3].write(payload["入职时间"] or "—")
        cols[4].write(payload["离职时间"] or "—")
        if cols[5].button("应用修改", key=f"ok_{category}_{item.label}"):
            set_decision(item.label, APPLY)
            st.rerun()
        if cols[6].button("取消", key=f"no_{category}_{item.label}"):
            set_decision(item.label, CANCEL)
            st.rerun()
        cols[7].write(badge[decision_of(item.label)])
        detail = payload["判定依据"]
        if payload["离职/调出备注"]:
            detail += f"｜备注：{payload['离职/调出备注']}"
        if payload["提示"]:
            detail += f"｜⚠️ {payload['提示']}"
        st.caption(f"{item.group or '—'}　{detail}")
        st.divider()


# --------------------------------------------------------------------------- #
# 导出
# --------------------------------------------------------------------------- #


def output_name(bonus_name: str, suffix: str) -> str:
    base = os.path.splitext(os.path.basename(bonus_name))[0]
    stamp = dt.datetime.now().strftime("%m%d-%H%M")
    return f"{base}（{suffix}{stamp}）.xlsx"


def render_export(bonus_bytes, bonus, analysis, include_pending) -> None:
    decisions = st.session_state["decisions"]

    def resolved(item):
        item.workshop = to_workshop(effective_workshop(item, analysis.mapping))
        return item

    selectable = [
        item
        for item in analysis.items
        if item.category in (CATEGORY_NEW, CATEGORY_LEFT)
        or include_pending.get(item.category, True)
    ]
    kept = [resolved(i) for i in selectable if decisions.get(i.label) != CANCEL]
    approved = [resolved(i) for i in selectable if decisions.get(i.label) == APPLY]

    left, right = st.columns(2)
    with left:
        st.subheader("① 对照标记版")
        st.caption(
            "按人员清单整理后的全量对照表：新增人员插到对应车间**最下方**、"
            "姓名与员工编号填**绿色**；需删除的人员保留原行、姓名与员工编号填**红色**。"
            "被点「取消」的人不纳入。"
        )
        adds = [i for i in kept if i.action == "add"]
        removes = [i for i in kept if i.action == "remove"]
        preview_counts(adds, removes, "标记删除")
        if st.button("生成对照标记版", type="primary", width="stretch"):
            generate("mark", bonus_bytes, bonus, adds, removes)
        offer_download("mark", bonus.file_name, "对照标记版")

    with right:
        st.subheader("② 已应用版")
        st.caption(
            "只处理被点「应用修改」的人员：删除类**直接删行**、"
            "新增类**直接插入对应位置**，不做任何着色。"
        )
        adds = [i for i in approved if i.action == "add"]
        removes = [i for i in approved if i.action == "remove"]
        preview_counts(adds, removes, "直接删除")
        if st.button("生成已应用版", type="primary", width="stretch"):
            generate("apply", bonus_bytes, bonus, adds, removes)
        offer_download("apply", bonus.file_name, "已应用版")

    st.info(
        "两个导出都保留原文件的全部子表、字体、行高列宽、公式、筛选和条件格式，"
        "只改动「一线人员」子表的人员行。打开后 Excel 会自动重算公式。"
    )


def preview_counts(adds, removes, remove_label: str) -> None:
    """预览数字必须和生成结果一致——未指定车间的人生成时会被跳过，这里就不能算进去。"""
    ready = [item for item in adds if item.workshop]
    pending = len(adds) - len(ready)
    st.write(f"将新增 **{len(ready)}** 人，{remove_label} **{len(removes)}** 人")
    if pending:
        st.caption(f"另有 {pending} 人未指定车间，生成时会被跳过（在上方「车间映射」里指定即可纳入）")


def generate(mode, bonus_bytes, bonus, adds, removes) -> None:
    if not adds and not removes:
        st.warning("没有需要处理的人员")
        return
    try:
        with st.spinner("正在生成 Excel…"):
            data, summary = build_workbook(bonus_bytes, bonus, adds, removes, mode=mode)
    except Exception as exc:  # noqa: BLE001 - 生成失败要给出可读原因
        st.error(f"生成失败：{type(exc).__name__}: {exc}")
        return
    st.session_state[f"blob_{mode}"] = data
    st.session_state[f"summary_{mode}"] = summary


def offer_download(mode, bonus_name, suffix) -> None:
    summary = st.session_state.get(f"summary_{mode}")
    if summary is None:
        return
    st.success(summary.text())
    if summary.per_workshop:
        st.caption("按车间：" + "、".join(f"{k} {v}人" for k, v in summary.per_workshop.items()))
    for warning in summary.warnings:
        st.warning(warning)
    if summary.skipped:
        with st.expander(f"已跳过 {len(summary.skipped)} 人"):
            for text in summary.skipped:
                st.write("•", text)
    st.download_button(
        f"下载{suffix}",
        data=st.session_state[f"blob_{mode}"],
        file_name=output_name(bonus_name, suffix),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        key=f"dl_{mode}",
    )


# --------------------------------------------------------------------------- #
# 综合数据库
# --------------------------------------------------------------------------- #


def render_database(payload, result) -> None:
    workspace = cached_workspace(payload)
    st.subheader("入库清单")
    if result.files:
        st.dataframe(
            pd.DataFrame(result.files).rename(
                columns={
                    "file_name": "文件",
                    "sha256": "指纹",
                    "size": "字节",
                    "kind": "类型",
                    "n_tables": "表数",
                }
            ),
            hide_index=True,
            width="stretch",
        )
    if result.problems:
        with st.expander(f"入库提示 / 问题（{len(result.problems)} 条）"):
            for problem in result.problems:
                st.write("•", problem)

    st.subheader("综合数据库")
    st.caption(
        "所有文件的所有子表都合并进了一个 SQLite 数据库：每个子表一张表，"
        "另有 `_files` / `_sheets` / `_documents` / `_problems` 元数据表，可直接写 SQL 交叉查询。"
    )
    sheets = pd.DataFrame(
        [
            {
                "表名": info.table_name,
                "来源文件": info.file_name,
                "子表": info.sheet_name,
                "行数": info.n_rows,
                "列数": info.n_cols,
            }
            for info in workspace.tables
        ]
    )
    if sheets.empty:
        st.warning("没有解析出任何二维表")
        return
    st.dataframe(sheets, hide_index=True, width="stretch")

    labels = dict(zip(sheets["表名"], sheets["子表"]))
    choice = st.selectbox("预览某张表", sheets["表名"], format_func=lambda n: f"{n}（{labels[n]}）")
    limit = st.slider("预览行数", 10, 500, 50, step=10)
    headers, rows = workspace.query(f'SELECT * FROM "{choice}" LIMIT {limit}')
    st.dataframe(pd.DataFrame(rows, columns=headers), hide_index=True, width="stretch")

    with st.expander("自定义 SQL 查询"):
        sql = st.text_area("SQL", value="SELECT * FROM _sheets", height=90)
        if st.button("执行查询"):
            try:
                headers, rows = workspace.query(sql)
                st.dataframe(
                    pd.DataFrame(rows, columns=headers), hide_index=True, width="stretch"
                )
            except Exception as exc:  # noqa: BLE001 - SQL 错误直接回显
                st.error(f"{type(exc).__name__}: {exc}")

    if st.button("准备数据库下载"):
        st.session_state["db_blob"] = cached_db_bytes(payload)
    if "db_blob" in st.session_state:
        st.download_button(
            "下载综合数据库（.db）",
            data=st.session_state["db_blob"],
            file_name="综合数据库.db",
            mime="application/octet-stream",
        )


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #


def render_sidebar(names, guess_roster, guess_bonus):
    with st.sidebar:
        st.header("设置")
        roster_name = st.selectbox(
            "人员清单文件", names,
            index=names.index(guess_roster) if guess_roster in names else 0,
        )
        bonus_name = st.selectbox(
            "安全质量奖核算数据文件", names,
            index=names.index(guess_bonus) if guess_bonus in names else min(1, len(names) - 1),
        )
        duty_field = st.radio(
            "「职务」取自清单的哪一列",
            ["职位", "岗位"],
            help="实习生的「岗位」记作助理工程师而「职位」记作实习生，"
                 "取「职位」可以把实习生排除在外，通常更贴近核算表口径。",
        )
        detected = date_from_filename(roster_name)
        ref_date = st.date_input(
            "参照日期（默认取自清单文件名）",
            value=detected or dt.date.today(),
            format="YYYY-MM-DD",
        )
        if detected is None:
            st.caption("⚠️ 清单文件名里没识别出日期，请确认上面的参照日期")
        window_months = int(st.number_input("新入职判定窗口（月）", 1, 12, 1))
        exclude_others = st.checkbox(
            "排除已在「副主任&工艺组长及其他」子表的人",
            value=True,
            help="这些人已经在核算文件的另一个子表里，重复添加会导致一人两处。",
        )
        st.divider()
        st.caption("待定人员是否纳入导出")
        include_pending = {
            CATEGORY_PENDING_ADD: st.checkbox("纳入「待定·清单有核算无」", value=True),
            CATEGORY_PENDING_DEL: st.checkbox("纳入「待定·核算有清单无」", value=True),
        }
        st.divider()
        if st.button("重置全部人工决策"):
            for key in ("decisions", "workshop_override", "mapping_override"):
                st.session_state[key] = {}
            st.rerun()
    return roster_name, bonus_name, duty_field, ref_date, window_months, exclude_others, include_pending


def main() -> None:
    init_state()
    st.title("📊 TJ4 安全质量奖人员对账")
    st.caption(
        "比对人员清单「生产部」里职务为助工/操作工/工程师/班长的人员与核算数据「一线人员」子表，"
        "分出新入职、离职、两类待定，并生成保留原格式与公式的新核算表。"
    )

    payload = collect_uploads()
    if not payload:
        st.info("请先上传文件。")
        return

    result = cached_ingest(payload)
    names = [entry["file_name"] for entry in result.files]
    if not names:
        st.error("没有成功读取到任何文件，请查看下方问题清单。")
        for problem in result.problems:
            st.write("•", problem)
        return

    guess_roster, guess_bonus = guess_roles(result)
    (
        roster_name,
        bonus_name,
        duty_field,
        ref_date,
        window_months,
        exclude_others,
        include_pending,
    ) = render_sidebar(names, guess_roster, guess_bonus)

    if roster_name == bonus_name:
        st.error("人员清单和核算数据不能是同一个文件，请在左侧重新选择。")
        return

    try:
        _, bonus, analysis = cached_analyze(
            roster_name,
            result.raw_files[roster_name],
            bonus_name,
            result.raw_files[bonus_name],
            duty_field,
            ref_date,
            window_months,
            exclude_others,
        )
    except Exception as exc:  # noqa: BLE001 - 解析失败要给出可读原因
        st.error(f"解析失败：{type(exc).__name__}: {exc}")
        return

    counts = analysis.counts
    metrics = st.columns(6)
    metrics[0].metric("两表一致", analysis.matched)
    metrics[1].metric("新入职员工", counts[CATEGORY_NEW])
    metrics[2].metric("待定·需新增", counts[CATEGORY_PENDING_ADD])
    metrics[3].metric("离职人员", counts[CATEGORY_LEFT])
    metrics[4].metric("待定·需核实", counts[CATEGORY_PENDING_DEL])
    metrics[5].metric("已在其他子表", analysis.excluded_in_others)
    st.caption(
        f"参照日期 {analysis.ref_date}，新入职窗口 {analysis.window_start} ~ {analysis.ref_date}；"
        f"清单目标职务 {analysis.matched + analysis.only_roster} 人，"
        f"核算一线人员 {analysis.matched + analysis.only_bonus} 人。"
    )
    for note in analysis.notes:
        st.warning(note)

    adds = [item for item in analysis.items if item.action == "add"]
    options = build_options(bonus, {item.group for item in adds})
    # 按当前生效的映射（含人工覆盖）实时统计，手工指定后这里的数字会立刻下降
    pending: dict[str, int] = {}
    for item in adds:
        if not to_workshop(effective_workshop(item, analysis.mapping)):
            pending[item.group] = pending.get(item.group, 0) + 1
    if pending:
        st.warning(
            f"还有 {sum(pending.values())} 名待新增人员的分组没有对应车间，"
            "请在下方「车间映射」里手工指定，否则导出时会被跳过。"
        )
        with st.expander(f"查看这 {len(pending)} 个待指定分组"):
            st.dataframe(
                pd.DataFrame(
                    sorted(pending.items(), key=lambda kv: (-kv[1], kv[0])),
                    columns=["目前分组", "待新增人数"],
                ),
                hide_index=True,
                width="stretch",
            )
    render_mapping_editor(analysis, bonus, adds, options)

    tabs = st.tabs(
        [
            f"新入职员工（{counts[CATEGORY_NEW]}）",
            f"待定·清单有核算无（{counts[CATEGORY_PENDING_ADD]}）",
            f"离职人员（{counts[CATEGORY_LEFT]}）",
            f"待定·核算有清单无（{counts[CATEGORY_PENDING_DEL]}）",
            "生成核算表",
            "综合数据库",
        ]
    )
    for tab, category in zip(tabs, CATEGORIES):
        with tab:
            render_category(category, analysis.by_category(category), bonus, analysis.mapping, options)
    with tabs[4]:
        render_export(result.raw_files[bonus_name], bonus, analysis, include_pending)
    with tabs[5]:
        render_database(payload, result)


if __name__ == "__main__":
    main()
