"""TJ4 安全质量奖人员对账工具（Streamlit）。

流程：上传文件/文件夹压缩包 → 统一入库 → 四类差异复核 → 生成新的核算表。
业务逻辑全部在 ``tj4tools`` 包里，本文件只做界面编排。
"""

from __future__ import annotations

import datetime as dt
import os

import pandas as pd
import streamlit as st

from tj4tools.bonus_export import (
    build_combined_workbook,
    build_workbook,
    merge_others_inserts,
    split_by_action,
)
from tj4tools.db import Workspace
from tj4tools.ingest import ingest_files
from tj4tools.normalize import date_from_filename, fmt_date, months_before
from tj4tools.roster import (
    ADD_CATEGORIES,
    CATEGORIES,
    INTERN_MONTHS,
    CATEGORY_LEFT,
    CATEGORY_NEW,
    CATEGORY_PENDING_ADD,
    CATEGORY_PENDING_DEL,
    build_others_workshop_map,
    build_workshop_mapping,
    parse_bonus,
    parse_roster,
    placeable_keys,
    reconcile,
)
from tj4tools.supervisor import (
    SCOPE_STRICT,
    build_duty_map,
    describe_duty_map,
    reconcile_supervisors,
)
from tj4tools.supervisor_export import build_supervisor_workbook
from tj4tools.supervisor_export import split_by_action as sup_split_by_action

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
    new_hire_since: dt.date | None,
    intern_asof: dt.date | None,
    exclude_others: bool,
    include_interns: bool,
):
    roster = parse_roster(
        roster_bytes, roster_name, duty_field=duty_field, include_interns=include_interns
    )
    bonus = parse_bonus(bonus_bytes, bonus_name)
    mapping = build_workshop_mapping(roster, bonus)
    result = reconcile(
        roster,
        bonus,
        ref_date=ref_date,
        new_hire_since=new_hire_since,
        intern_asof=intern_asof,
        exclude_in_others=exclude_others,
        mapping=mapping,
    )
    return roster, bonus, result


@st.cache_data(show_spinner=False)
def _parsed_pair(roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns):
    roster = parse_roster(
        roster_bytes, roster_name, duty_field=duty_field, include_interns=include_interns
    )
    return roster, parse_bonus(bonus_bytes, bonus_name)


@st.cache_data(show_spinner=False)
def cached_others_workshop_map(
    roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns
):
    roster, bonus = _parsed_pair(
        roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns
    )
    return build_others_workshop_map(roster, bonus)


def resolved_placeable(review: "Review", analysis) -> frozenset:
    """功能一能放进一线的人：已指定车间，或带括号只是等确认、但已有建议。"""
    keys = set(placeable_keys(analysis))
    for item in analysis.items:
        if item.action != "add":
            continue
        if to_workshop(effective_workshop(review, item, analysis.mapping)):
            keys.add(item.key)
    return frozenset(keys)


@st.cache_data(show_spinner=False)
def cached_supervisor_duty_map(
    roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns
):
    roster, bonus = _parsed_pair(
        roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns
    )
    return build_duty_map(roster, bonus)


@st.cache_data(show_spinner="正在对账「副主任&工艺组长及其他」…")
def cached_supervisor_analyze(
    roster_name,
    roster_bytes,
    bonus_name,
    bonus_bytes,
    duty_field,
    include_interns,
    ref_date,
    new_hire_since,
    intern_asof,
    scope,
    placeable,
):
    roster, bonus = _parsed_pair(
        roster_name, roster_bytes, bonus_name, bonus_bytes, duty_field, include_interns
    )
    return reconcile_supervisors(
        roster,
        bonus,
        ref_date=ref_date,
        new_hire_since=new_hire_since,
        intern_asof=intern_asof,
        scope=scope,
        placeable_keys=set(placeable),
    )


@st.cache_data(show_spinner=False)
def count_interns(roster_bytes: bytes, roster_name: str) -> int:
    roster = parse_roster(roster_bytes, roster_name, duty_field="职位", include_interns=True)
    return sum(1 for person in roster.production.values() if person.is_intern)


@st.cache_data(show_spinner="正在导出数据库…")
def cached_db_bytes(payload: tuple[tuple[str, bytes], ...]) -> bytes:
    return cached_workspace(payload).to_sqlite_bytes()


# --------------------------------------------------------------------------- #
# 会话状态
# --------------------------------------------------------------------------- #


STATE_KEYS = ("decisions", "workshop_override", "mapping_override", "action_override")

ACTION_KEEP = "保留在一线人员"
ACTION_TO_OTHERS = "移动到副主任表格"
PENDING_DEL_ACTIONS = (ACTION_TO_OTHERS, ACTION_KEEP)


class Review:
    """一个功能的人工复核状态。

    两个功能（一线人员 / 副主任&工艺组长及其他）会出现同一个人，
    所以决策必须按功能分开存，否则在一个功能里点的"应用"会串到另一个功能。
    """

    def __init__(self, ns: str):
        self.ns = ns
        for name in STATE_KEYS:
            st.session_state.setdefault(self.key(name), {})

    def key(self, name: str) -> str:
        return f"{self.ns}__{name}" if self.ns else name

    @property
    def decisions(self) -> dict:
        return st.session_state[self.key("decisions")]

    @property
    def workshop_override(self) -> dict:
        return st.session_state[self.key("workshop_override")]

    @property
    def mapping_override(self) -> dict:
        return st.session_state[self.key("mapping_override")]

    @property
    def action_override(self) -> dict:
        return st.session_state[self.key("action_override")]

    def decision_of(self, label: str) -> str:
        return self.decisions.get(label, UNDECIDED)

    def set_decision(self, label: str, value: str) -> None:
        self.decisions[label] = value

    def reset(self) -> None:
        for name in STATE_KEYS:
            st.session_state[self.key(name)] = {}

    def bump(self, category: str) -> None:
        """改变 data_editor 的 key，让批量操作后的表格重新初始化。"""
        key = self.key(f"ver_{category}")
        st.session_state[key] = st.session_state.get(key, 0) + 1

    def version(self, category: str) -> int:
        return st.session_state.get(self.key(f"ver_{category}"), 0)

    def widget(self, *parts) -> str:
        return "_".join([self.ns or "main", *(str(p) for p in parts)])


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


def build_options(workshops, groups) -> list[str]:
    """下拉选项 = 未指定 + 原有车间 + 「新建车间块」候选。"""
    return (
        [UNSET]
        + list(workshops)
        + [NEW_BLOCK_PREFIX + name for name in sorted({g for g in groups if g})]
    )


def effective_workshop(review: Review, item, mapping) -> str:
    if item.label in review.workshop_override:
        return review.workshop_override[item.label]
    if item.group in review.mapping_override:
        return review.mapping_override[item.group]
    guess = mapping.get(item.group)
    if guess is not None:
        return guess.workshop if hasattr(guess, "workshop") else str(guess)
    return item.target_workshop or item.workshop


def describe_guess(guess) -> str:
    if guess is None:
        return "需人工指定"
    if isinstance(guess, str):
        return guess or "需人工指定"
    hint = guess.suggested or guess.workshop
    if guess.needs_manual:
        if not hint:
            return "需人工指定（分组名带括号）"
        extra = ""
        if guess.source == "经验" and guess.support:
            extra = f"；两表已匹配 {guess.support} 人 · 置信度{guess.confidence}"
        elif guess.source == "规则":
            extra = "；按命名规则推断"
        return f"{hint}（分组名带括号，请人工确认{extra}）"
    if not guess.workshop:
        return "需人工指定"
    if guess.source == "经验":
        return f"{guess.workshop}（两表已匹配 {guess.support} 人 · 置信度{guess.confidence}）"
    return f"{guess.workshop}（按命名规则推断 · 置信度{guess.confidence}）"


def render_mapping_editor(review: Review, mapping, items, options, hint: str) -> None:
    counts: dict[str, int] = {}
    for item in items:
        counts[item.group] = counts.get(item.group, 0) + 1
    if not counts:
        return
    overrides = review.mapping_override
    rows = []
    for group, count in sorted(counts.items(), key=lambda kv: -kv[1]):
        guess = mapping.get(group)
        rows.append(
            {
                "目前分组": group,
                "待新增": count,
                "自动建议": describe_guess(guess),
                "最终车间": to_option(effective_group_workshop(review, group, mapping), options),
            }
        )
    frame = pd.DataFrame(rows)
    unknown = int((frame["最终车间"] == UNSET).sum())
    title = f"车间映射（{len(frame)} 个分组，"
    title += f"{unknown} 个待人工指定）" if unknown else "已全部对应）"
    with st.expander(title, expanded=bool(unknown)):
        st.caption(hint)
        edited = st.data_editor(
            frame,
            hide_index=True,
            width="stretch",
            key=review.widget("mapping_editor"),
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


def effective_group_workshop(review: Review, group: str, mapping) -> str:
    if group in review.mapping_override:
        return review.mapping_override[group]
    guess = mapping.get(group)
    if guess is None:
        return ""
    return guess if isinstance(guess, str) else guess.workshop


# --------------------------------------------------------------------------- #
# 差异复核
# --------------------------------------------------------------------------- #


def choice_labels(category: str) -> tuple[str, str]:
    """勾选框/按钮的文案必须写清楚具体动作，不能只写"应用/取消"。

    由 render_category 的调用方按"这一类要做什么"传进来。
    """
    return CHOICE_LABELS.get(category, ("应用", "取消"))


# 每一类的动作文案：(勾上的动作, 不勾的动作)
LABELS_ADD_FRONTLINE = ("新增到一线人员", "不新增")
LABELS_REMOVE_FRONTLINE = ("从一线人员删除", "保留在一线人员")
LABELS_ADD_OTHERS = ("新增到副主任表", "不新增")
LABELS_REMOVE_OTHERS = ("从副主任表删除", "保留在副主任表")

CHOICE_LABELS: dict[str, tuple[str, str]] = {}


def set_choice_labels(mapping: dict[str, tuple[str, str]]) -> None:
    """切换功能时重设各分类的动作文案（两张子表的动作说法不一样）。"""
    CHOICE_LABELS.clear()
    CHOICE_LABELS.update(mapping)


def effective_pending_action(review: Review, item) -> str:
    """「待定·核算有清单无」的最终动作：人工选过就用人工的，否则用自动判定的。"""
    override = review.action_override.get(item.label)
    if override in PENDING_DEL_ACTIONS:
        return override
    return ACTION_TO_OTHERS if item.action == "move" else ACTION_KEEP


def render_pending_delete(review: Review, category, items) -> None:
    """「待定·核算有清单无」：自动给出动作，但每个人都可以手工改成另一个动作。"""
    if not items:
        st.success(f"没有「{category}」人员")
        return
    decided = [effective_pending_action(review, item) for item in items]
    metrics = st.columns([1, 1, 1, 4])
    metrics[0].metric("总数", len(items))
    metrics[1].metric("保留在一线人员", decided.count(ACTION_KEEP))
    metrics[2].metric("移动到副主任表格", decided.count(ACTION_TO_OTHERS))
    st.info(
        "这些人在核算表里有、在清单目标职务里没有，但都能在「生产部」里找到本人"
        "（多为职务变动或姓名录错），所以都按人员清单的信息处理。"
        "**动作列可以逐人改**：\n\n"
        f"- **{ACTION_KEEP}**：留在「一线人员」，并按清单改写姓名/员工编号/职务/入职日期\n"
        f"- **{ACTION_TO_OTHERS}**：从「一线人员」移除，写进「副主任&工艺组长及其他」对应车间\n\n"
        "默认按清单职务判：仍是助工/操作工/工程师/班长的保留，不是这四种的移走。"
    )
    bulk = st.columns([1.7, 1.7, 1.1, 4])
    if bulk[0].button(f"全部改为「{ACTION_KEEP}」", key=review.widget("all_keep", category)):
        for item in items:
            review.action_override[item.label] = ACTION_KEEP
        review.bump(category)
        st.rerun()
    if bulk[1].button(f"全部改为「{ACTION_TO_OTHERS}」", key=review.widget("all_move", category)):
        for item in items:
            review.action_override[item.label] = ACTION_TO_OTHERS
        review.bump(category)
        st.rerun()
    if bulk[2].button("恢复自动判定", key=review.widget("all_auto", category)):
        for item in items:
            review.action_override.pop(item.label, None)
        review.bump(category)
        st.rerun()

    order = ["动作", "姓名", "员工编号", "原一线车间", "核算职务", "清单职务", "目标车间",
             "修改内容", "入职时间", "实习生", "自动判定", "判定依据"]
    frame = pd.DataFrame(
        [
            {
                "动作": action,
                "姓名": item.name,
                "员工编号": item.eid,
                "原一线车间": item.workshop,
                "核算职务": item.duty,
                "清单职务": str(item.new_values.get("职务") or ""),
                "目标车间": item.target_workshop if action == ACTION_TO_OTHERS else "—",
                "修改内容": item.update_text or "（无字段变化）",
                "入职时间": fmt_date(item.hire_date),
                "实习生": item.intern_text,
                "自动判定": ACTION_TO_OTHERS if item.action == "move" else ACTION_KEEP,
                "判定依据": item.reason,
            }
            for item, action in zip(items, decided)
        ]
    )[order]
    edited = st.data_editor(
        frame,
        hide_index=True,
        width="stretch",
        height=min(640, 90 + 35 * len(frame)),
        key=review.widget("pending_editor", category, review.version(category)),
        column_config={
            "动作": st.column_config.SelectboxColumn(
                options=list(PENDING_DEL_ACTIONS), required=True, width="medium"
            ),
            "姓名": st.column_config.TextColumn(width="small"),
            "员工编号": st.column_config.TextColumn(width="small"),
            "判定依据": st.column_config.TextColumn(width="large"),
        },
        disabled=[column for column in order if column != "动作"],
    )
    changed = False
    for item, before, after in zip(items, frame["动作"], edited["动作"]):
        if after != before:
            review.action_override[item.label] = after
            changed = True
    st.caption("这一类不用勾「应用/取消」——动作列选什么就执行什么，两个导出都会照此处理。")
    if changed:
        # 指标渲染在表格上方，不重跑一次的话要等下次交互才更新
        st.rerun()

    moves = [item for item, action in zip(items, decided) if action == ACTION_TO_OTHERS]
    unmapped = [item for item in moves if not item.target_workshop]
    if unmapped:
        st.warning(
            "以下人员没能确定在副主任表里的车间，导出时会追加在该表最下方："
            + "、".join(f"{i.name}({i.eid})" for i in unmapped)
        )
    fallback = sorted({i.target_workshop for i in moves if i.target_workshop_source == "沿用一线车间"})
    if fallback:
        st.caption(
            "副主任表里原本没有这些车间，相关人员会追加在该表人员区最下方并写上车间名："
            + "、".join(fallback)
        )


def render_category(review: Review, category, items, mapping, options, editable) -> None:
    if not items:
        st.success(f"没有「{category}」人员")
        return

    decisions = review.decisions
    applied = sum(1 for item in items if decisions.get(item.label) == APPLY)
    cancelled = sum(1 for item in items if decisions.get(item.label) == CANCEL)
    metrics = st.columns([1, 1, 1, 1, 3])
    metrics[0].metric("总数", len(items))
    metrics[1].metric("已应用", applied)
    metrics[2].metric("已取消", cancelled)
    metrics[3].metric("待定", len(items) - applied - cancelled)

    apply_label, cancel_label = choice_labels(category)
    bulk = st.columns([1, 1, 1, 5])
    if bulk[0].button(f"全部{apply_label}", key=review.widget("all_apply", category)):
        for item in items:
            decisions[item.label] = APPLY
        review.bump(category)
        st.rerun()
    if bulk[1].button(f"全部{cancel_label}", key=review.widget("all_cancel", category)):
        for item in items:
            decisions[item.label] = CANCEL
        review.bump(category)
        st.rerun()
    if bulk[2].button("清空决策", key=review.widget("all_reset", category)):
        for item in items:
            decisions.pop(item.label, None)
        review.bump(category)
        st.rerun()

    mode = st.radio(
        "复核方式",
        ["表格批量", "逐条按钮"],
        horizontal=True,
        key=review.widget("mode", category),
        help="人数多时用表格批量勾选更快；需要逐个确认时切到逐条按钮。",
    )
    if mode == "表格批量":
        render_table(review, category, items, mapping, options, editable)
    else:
        render_rows(review, category, items, mapping, editable)


def row_payload(review: Review, item, mapping, options=None) -> dict:
    payload = item.as_dict()
    workshop = effective_workshop(review, item, mapping)
    payload["车间"] = to_option(workshop, options) if options else workshop
    return payload


def render_table(review: Review, category, items, mapping, options, editable) -> None:
    """表格批量复核。

    「应用」和「取消」是两个独立复选框，并且无条件按编辑结果回写。
    这样表格状态与决策状态形成稳定不动点：既不会因为"没勾选"就把待定误判成取消，
    也不依赖"和上一帧比较"这种脆弱逻辑。
    """
    decisions = review.decisions
    overrides = review.workshop_override
    apply_label, cancel_label = choice_labels(category)
    order = ["应用", "取消", "姓名", "员工编号", "职务", "车间", "目前分组", "入职时间",
             "离职时间", "清单数据差异", "实习生", "离职/调出备注", "判定依据", "提示"]
    frame = pd.DataFrame(
        [
            {
                "应用": decisions.get(item.label) == APPLY,
                "取消": decisions.get(item.label) == CANCEL,
                **{k: v for k, v in row_payload(review, item, mapping, options).items() if k in order},
            }
            for item in items
        ]
    )[order]
    if not frame["清单数据差异"].any():
        frame = frame.drop(columns=["清单数据差异"])
    if not frame["实习生"].any():
        frame = frame.drop(columns=["实习生"])
    order = list(frame.columns)
    config = {
        "应用": st.column_config.CheckboxColumn(
            apply_label, width="small", help="纳入「已应用版」导出"
        ),
        "取消": st.column_config.CheckboxColumn(
            cancel_label, width="small", help="从两个导出里都剔除，保持核算表原样"
        ),
        "姓名": st.column_config.TextColumn(width="small"),
        "员工编号": st.column_config.TextColumn(width="small"),
        "职务": st.column_config.TextColumn(width="small"),
        "实习生": st.column_config.TextColumn(width="small"),
        "清单数据差异": st.column_config.TextColumn(width="medium"),
        "判定依据": st.column_config.TextColumn(width="large"),
    }
    disabled = [column for column in order if column not in ("应用", "取消")]
    if editable:
        config["车间"] = st.column_config.SelectboxColumn(
            options=options, required=True, width="medium"
        )
        disabled.remove("车间")
    edited = st.data_editor(
        frame,
        hide_index=True,
        width="stretch",
        height=min(640, 90 + 35 * len(frame)),
        key=review.widget("editor", category, review.version(category)),
        column_config=config,
        disabled=disabled,
    )
    changed = False
    for index, item in enumerate(items):
        before = decisions.get(item.label)
        if bool(edited["应用"].iloc[index]):
            decisions[item.label] = APPLY
        elif bool(edited["取消"].iloc[index]):
            decisions[item.label] = CANCEL
        else:
            decisions.pop(item.label, None)
        changed = changed or decisions.get(item.label) != before
        if editable:
            now = edited["车间"].iloc[index]
            if now != frame["车间"].iloc[index]:
                overrides[item.label] = to_workshop(now)
                changed = True
    st.caption(
        f"「{apply_label}」= 纳入「已应用版」导出；「{cancel_label}」= 从两个导出里都剔除；"
        f"两个都不勾 = 待定，只出现在「对照标记版」里。两个都勾时按「{apply_label}」处理。"
    )
    if changed:
        # 上方的统计指标要立刻跟着变，否则得等下一次交互才刷新
        st.rerun()


def render_rows(review: Review, category, items, mapping, editable) -> None:
    pages = max(1, (len(items) + PAGE_SIZE - 1) // PAGE_SIZE)
    page = 1
    if pages > 1:
        page = int(
            st.number_input(
                f"翻到第几页　（每页 {PAGE_SIZE} 人，共 {pages} 页 / {len(items)} 人）",
                min_value=1,
                max_value=pages,
                value=1,
                key=review.widget("page", category),
            )
        )
        first = (page - 1) * PAGE_SIZE + 1
        st.caption(f"当前显示第 {first} ~ {min(page * PAGE_SIZE, len(items))} 人")
    apply_label, cancel_label = choice_labels(category)
    badge = {APPLY: f"✅ {apply_label}", CANCEL: f"🚫 {cancel_label}", UNDECIDED: "⏳ 待定"}
    for item in items[(page - 1) * PAGE_SIZE : page * PAGE_SIZE]:
        payload = row_payload(review, item, mapping)
        cols = st.columns([2.2, 1.6, 0.9, 1.1, 1.1, 1.0, 1.0, 1.1])
        name = f"**{item.name}**　`{item.eid}`"
        cols[0].markdown(name + ("　🟡实习生" if item.is_intern else ""))
        cols[1].write(payload["车间"] or "待指定")
        cols[2].write(item.duty or "—")
        cols[3].write(payload["入职时间"] or "—")
        cols[4].write(payload["离职时间"] or "—")
        if cols[5].button(apply_label, key=review.widget("ok", category, item.label)):
            review.set_decision(item.label, APPLY)
            st.rerun()
        if cols[6].button(cancel_label, key=review.widget("no", category, item.label)):
            review.set_decision(item.label, CANCEL)
            st.rerun()
        cols[7].write(badge[review.decision_of(item.label)])
        detail = payload["判定依据"]
        if payload["清单数据差异"]:
            detail = f"将改为：{payload['清单数据差异']}｜{detail}"
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


def collect_frontline_items(review: Review, analysis, include_pending, *, approved_only: bool):
    """按当前复核结果收集一线人员导出清单。

    ``approved_only=True`` 只要勾了「应用」的人（已应用版）；
    ``False`` 则排除勾了「取消」的人（对照标记版 / 一键全部应用）。
    「待定·核算有清单无」始终按动作列执行，不看应用/取消。
    """

    def resolved(item):
        item.workshop = to_workshop(effective_workshop(review, item, analysis.mapping))
        return item

    selectable = [
        item
        for item in analysis.items
        if item.category in (CATEGORY_NEW, CATEGORY_LEFT)
        or include_pending.get(item.category, True)
    ]
    auto = []
    for item in selectable:
        if item.category != CATEGORY_PENDING_DEL:
            continue
        action = effective_pending_action(review, item)
        item.action = "move" if action == ACTION_TO_OTHERS else "update"
        auto.append(resolved(item))
    reviewed = [i for i in selectable if i.category != CATEGORY_PENDING_DEL]
    chosen = [
        resolved(item)
        for item in reviewed
        if (review.decisions.get(item.label) == APPLY)
        or (not approved_only and review.decisions.get(item.label) != CANCEL)
    ]
    return split_by_action(auto + chosen)


def render_export(review: Review, bonus_bytes, bonus, analysis, include_pending) -> None:
    kept = collect_frontline_items(review, analysis, include_pending, approved_only=False)
    approved = collect_frontline_items(review, analysis, include_pending, approved_only=True)

    left, right = st.columns(2)
    with left:
        st.subheader("① 对照标记版")
        st.caption(
            "按人员清单整理后的全量对照表：新增人员插到对应车间**该职务**的最下方、"
            "姓名与员工编号填**绿色**；需删除的人员保留原行、填**红色**；"
            "按清单更新的人员就地改值、改动的单元格填**橙色**。被点「取消」的人不纳入。"
        )
        adds, removes, updates, moves = kept
        preview_counts(adds, removes, updates, moves, "标记删除")
        if st.button("生成对照标记版", type="primary", width="stretch",
                     key=review.widget("gen_mark")):
            generate(review, "mark", bonus_bytes, bonus, adds, removes, updates, moves)
        offer_download(review, "mark", bonus.file_name, "对照标记版")

    with right:
        st.subheader("② 已应用版")
        st.caption(
            "只处理被点「应用」的人员：删除类**直接删行**、新增类**直接插入对应位置**"
            "且内容用**红色字体**、更新类就地改成清单里的值。"
        )
        adds, removes, updates, moves = approved
        preview_counts(adds, removes, updates, moves, "直接删除")
        if st.button("生成已应用版", type="primary", width="stretch",
                     key=review.widget("gen_apply")):
            generate(review, "apply", bonus_bytes, bonus, adds, removes, updates, moves)
        offer_download(review, "apply", bonus.file_name, "已应用版")

    st.info(
        "两个导出都保留原文件的全部子表、字体、行高列宽、公式、筛选和条件格式；"
        "改动只发生在「一线人员」和「副主任&工艺组长及其他」两张子表的人员行。"
        "实习生的职务列填**黄色**底纹。打开后 Excel 会自动重算公式。"
    )


def preview_counts(adds, removes, updates, moves, remove_label: str) -> None:
    """预览数字必须和生成结果一致——未指定车间的人生成时会被跳过，这里就不能算进去。"""
    ready = [item for item in adds if item.workshop]
    pending = len(adds) - len(ready)
    text = f"将新增 **{len(ready)}** 人，{remove_label} **{len(removes)}** 人"
    if moves:
        text += f"，移到副主任表 **{len(moves)}** 人"
    if updates:
        text += f"，按清单更新 **{len(updates)}** 人"
    st.write(text)
    if pending:
        st.caption(f"另有 {pending} 人未指定车间，生成时会被跳过（在上方「车间映射」里指定即可纳入）")


def generate(review: Review, mode, bonus_bytes, bonus, adds, removes, updates=(), moves=()) -> None:
    if not adds and not removes and not updates and not moves:
        st.warning("没有需要处理的人员")
        return
    try:
        with st.spinner("正在生成 Excel…"):
            data, summary = build_workbook(
                bonus_bytes, bonus, adds, removes, list(updates), list(moves), mode=mode
            )
    except Exception as exc:  # noqa: BLE001 - 生成失败要给出可读原因
        st.error(f"生成失败：{type(exc).__name__}: {exc}")
        return
    st.session_state[review.key(f"blob_{mode}")] = data
    st.session_state[review.key(f"summary_{mode}")] = summary


def offer_download(review: Review, mode, bonus_name, suffix) -> None:
    summary = st.session_state.get(review.key(f"summary_{mode}"))
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
        data=st.session_state[review.key(f"blob_{mode}")],
        file_name=output_name(bonus_name, suffix),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
        key=review.widget("dl", mode),
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


def render_sidebar(names, guess_roster, guess_bonus, intern_total):
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
            help="实习生的「职位」记作实习生、「岗位」记作实际岗位，"
                 "取「职位」会天然把实习生排除在外。",
        )
        include_interns = st.checkbox(
            f"纳入实习生（{intern_total} 人，导出时职务列标黄底）",
            value=True,
            help="「职位」写实习生的人按「岗位」列的实际岗位归入职务分组；"
                 "备注里写着「校招实习生」的人本来就在名单里，不受这个开关影响。",
        )
        detected = date_from_filename(roster_name)
        ref_date = st.date_input(
            "参照日期（默认取自清单文件名）",
            value=detected or dt.date.today(),
            format="YYYY-MM-DD",
        )
        if detected is None:
            st.caption("⚠️ 清单文件名里没识别出日期，请确认上面的参照日期")
        default_since = months_before(ref_date, 1)
        new_hire_since = st.date_input(
            "新入职判定窗口日期",
            value=default_since,
            format="YYYY-MM-DD",
            help="窗口的起点；入职时间要落在这一天和参照日期之间才算新入职。",
        )
        st.caption(f"即：入职时间在 {new_hire_since} ~ {ref_date} 之间的算新入职员工")
        intern_asof = st.date_input(
            "实习生判断日期",
            value=ref_date,
            format="YYYY-MM-DD",
            help="按这一天计算实习生的入职时长，分成入职超过 3 个月和不到 3 个月。",
        )
        st.caption(
            f"即：入职不晚于 {months_before(intern_asof, INTERN_MONTHS)} 的实习生算"
            f"「入职超过 {INTERN_MONTHS} 个月」"
        )
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
            Review("").reset()
            Review("sup").reset()
            Review("all").reset()
            st.rerun()
    return {
        "roster_name": roster_name,
        "bonus_name": bonus_name,
        "duty_field": duty_field,
        "include_interns": include_interns,
        "ref_date": ref_date,
        "new_hire_since": new_hire_since,
        "intern_asof": intern_asof,
        "exclude_others": exclude_others,
        "include_pending": include_pending,
    }


FEATURE_FRONTLINE = "① 一线人员"
FEATURE_SUPERVISOR = "② 副主任&工艺组长及其他"
FEATURE_COMBINED = "③ 一键生成全部"


def main() -> None:
    st.title("📊 TJ4 安全质量奖核算表生成")

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
    intern_total = 0
    if guess_roster in result.raw_files:
        try:
            intern_total = count_interns(result.raw_files[guess_roster], guess_roster)
        except Exception:  # noqa: BLE001 - 只是侧边栏的提示数字，失败不影响主流程
            intern_total = 0
    options_ = render_sidebar(names, guess_roster, guess_bonus, intern_total)
    roster_name = options_["roster_name"]
    bonus_name = options_["bonus_name"]

    if roster_name == bonus_name:
        st.error("人员清单和核算数据不能是同一个文件，请在左侧重新选择。")
        return

    try:
        roster, bonus, analysis = cached_analyze(
            roster_name,
            result.raw_files[roster_name],
            bonus_name,
            result.raw_files[bonus_name],
            options_["duty_field"],
            options_["ref_date"],
            options_["new_hire_since"],
            options_["intern_asof"],
            options_["exclude_others"],
            options_["include_interns"],
        )
    except Exception as exc:  # noqa: BLE001 - 解析失败要给出可读原因
        st.error(f"解析失败：{type(exc).__name__}: {exc}")
        return

    feature = st.radio(
        "要处理哪张子表",
        [FEATURE_FRONTLINE, FEATURE_SUPERVISOR, FEATURE_COMBINED],
        horizontal=True,
        key="feature",
        help="①② 的复核状态各自独立。③ 把两侧改动写进同一份核算表。",
    )
    if feature == FEATURE_SUPERVISOR:
        render_supervisor_feature(payload, result, roster, bonus, analysis, options_)
    elif feature == FEATURE_COMBINED:
        render_combined_feature(payload, result, roster, bonus, analysis, options_)
    else:
        render_frontline_feature(payload, result, bonus, analysis, options_)


def render_frontline_feature(payload, result, bonus, analysis, options_) -> None:
    review = Review("")
    bonus_name = options_["bonus_name"]
    include_pending = options_["include_pending"]
    set_choice_labels(
        {
            CATEGORY_NEW: LABELS_ADD_FRONTLINE,
            CATEGORY_PENDING_ADD: LABELS_ADD_FRONTLINE,
            CATEGORY_LEFT: LABELS_REMOVE_FRONTLINE,
        }
    )
    st.caption(
        "比对人员清单「生产部」里职务为助工/操作工/工程师/班长的人员与核算数据「一线人员」子表，"
        "分出新入职、离职、两类待定，并生成保留原格式与公式的新核算表。"
    )
    counts = analysis.counts
    metrics = st.columns(6)
    metrics[0].metric("两表一致", analysis.matched)
    metrics[1].metric("新入职员工", counts[CATEGORY_NEW])
    metrics[2].metric("待定·需新增", counts[CATEGORY_PENDING_ADD])
    metrics[3].metric("离职人员", counts[CATEGORY_LEFT])
    metrics[4].metric("待定·需核实", counts[CATEGORY_PENDING_DEL])
    metrics[5].metric("已在其他子表", analysis.excluded_in_others)
    st.caption(
        f"参照日期 {analysis.ref_date}，入职时间在 {analysis.new_hire_since} ~ {analysis.ref_date} "
        f"之间的算新入职；清单目标职务 {analysis.matched + analysis.only_roster} 人，"
        f"核算一线人员 {analysis.matched + analysis.only_bonus} 人。"
    )
    interns = analysis.intern_counts
    if interns:
        st.caption(
            f"实习生 {sum(interns.values())} 人（按 {analysis.intern_asof} 判断）："
            + "、".join(f"{name} {count} 人" for name, count in interns.items())
        )
    for note in analysis.notes:
        st.warning(note)

    adds = [item for item in analysis.items if item.action == "add"]
    options = build_options(bonus.workshops, {item.group for item in adds})
    render_unmapped_warning(review, adds, analysis.mapping)
    render_mapping_editor(
        review,
        analysis.mapping,
        adds,
        options,
        "「自动建议」由两表已匹配人员反推得出。置信度低、带括号或显示「需人工指定」的行请在右侧下拉里选择；"
        "目前分组带括号的（委培、分区等）一律不自动落车间，必须手选。"
        "选「＋新建车间块」会在一线人员子表最下方新建该分组。",
    )

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
            if category == CATEGORY_PENDING_DEL:
                render_pending_delete(review, category, analysis.by_category(category))
            else:
                render_category(
                    review,
                    category,
                    analysis.by_category(category),
                    analysis.mapping,
                    options,
                    category in ADD_CATEGORIES,
                )
    with tabs[4]:
        render_export(review, result.raw_files[bonus_name], bonus, analysis, include_pending)
    with tabs[5]:
        render_database(payload, result)


def render_supervisor_feature(payload, result, roster, bonus, frontline_analysis, options_) -> None:
    review = Review("sup")
    bonus_name = options_["bonus_name"]
    set_choice_labels(
        {
            CATEGORY_NEW: LABELS_ADD_OTHERS,
            CATEGORY_PENDING_ADD: LABELS_ADD_OTHERS,
            CATEGORY_LEFT: LABELS_REMOVE_OTHERS,
            CATEGORY_PENDING_DEL: LABELS_REMOVE_OTHERS,
        }
    )
    st.caption(
        "从人员清单生成「副主任&工艺组长及其他」子表：取**管理类职务**（车间副主任/工艺组长/经理），"
        "外加**不在「一线人员」子表、且功能一也放不进一线车间的**助工/工程师/操作工/班长——"
        "功能一能放进一线的那批人不会重复出现在这里。"
    )
    if bonus.others_layout is None:
        st.error("核算文件里找不到「副主任&工艺组长及其他」子表。")
        return

    # 功能一能放进一线车间的人不该再进这张表，否则同一个人会出现在两张子表里。
    # 带括号的分组还没手选车间，但只要有建议就不算"放不进一线"。
    placeable = resolved_placeable(Review(""), frontline_analysis)

    try:
        analysis = cached_supervisor_analyze(
            options_["roster_name"],
            result.raw_files[options_["roster_name"]],
            bonus_name,
            result.raw_files[bonus_name],
            options_["duty_field"],
            options_["include_interns"],
            options_["ref_date"],
            options_["new_hire_since"],
            options_["intern_asof"],
            SCOPE_STRICT,
            placeable,
        )
    except Exception as exc:  # noqa: BLE001 - 解析失败要给出可读原因
        st.error(f"对账失败：{type(exc).__name__}: {exc}")
        return

    counts = analysis.counts
    metrics = st.columns(5)
    metrics[0].metric("已在本表", analysis.matched)
    metrics[1].metric("新入职员工", counts[CATEGORY_NEW])
    metrics[2].metric("待定·需新增", counts[CATEGORY_PENDING_ADD])
    metrics[3].metric("离职人员", counts[CATEGORY_LEFT])
    metrics[4].metric("待定·需核实", counts[CATEGORY_PENDING_DEL])
    # 注意：这里必须写成 if/else 语句。写成条件表达式的话它是个裸表达式，
    # Streamlit 的 magic 会给它套一层 st.write，把 DeltaGenerator 的 repr 打到页面上。
    for note in analysis.notes:
        if note.startswith("取人口径"):
            st.info(note)
        else:
            st.warning(note)

    layout = bonus.others_layout
    duty_map = cached_supervisor_duty_map(
        options_["roster_name"],
        result.raw_files[options_["roster_name"]],
        bonus_name,
        result.raw_files[bonus_name],
        options_["duty_field"],
        options_["include_interns"],
    )
    with st.expander(f"职务写法映射（{len(duty_map)} 条，由两表已匹配人员反推）"):
        st.caption(
            "清单写「车间副主任」，本表写「副主任」；清单写「助理工程师」，本表也写「助理工程师」。"
            "映射由已匹配的人反推众数得出。"
        )
        st.dataframe(
            pd.DataFrame(describe_duty_map(duty_map)), hide_index=True, width="stretch"
        )

    adds = [item for item in analysis.items if item.action == "add"]
    workshop_mapping = cached_others_workshop_map(
        options_["roster_name"],
        result.raw_files[options_["roster_name"]],
        bonus_name,
        result.raw_files[bonus_name],
        options_["duty_field"],
        options_["include_interns"],
    )
    options = build_options(layout.workshops, {item.group for item in adds})
    render_unmapped_warning(review, adds, workshop_mapping)
    render_mapping_editor(
        review,
        workshop_mapping,
        adds,
        options,
        "把清单的「目前分组」对应到本表的车间。本表的车间叫法和一线人员不同"
        "（如「11号楼车间D级区域」）。目前分组带括号的一律不自动落车间，必须手选；"
        "选「＋新建车间块」会在本表人员区最下方新建。",
    )

    tabs = st.tabs(
        [
            f"新入职员工（{counts[CATEGORY_NEW]}）",
            f"待定·清单有本表无（{counts[CATEGORY_PENDING_ADD]}）",
            f"离职人员（{counts[CATEGORY_LEFT]}）",
            f"待定·本表有清单无（{counts[CATEGORY_PENDING_DEL]}）",
            "生成核算表",
            "综合数据库",
        ]
    )
    for tab, category in zip(tabs, CATEGORIES):
        with tab:
            render_category(
                review,
                category,
                analysis.by_category(category),
                workshop_mapping,
                options,
                category in ADD_CATEGORIES,
            )
    with tabs[4]:
        render_supervisor_export(review, result.raw_files[bonus_name], bonus, analysis)
    with tabs[5]:
        render_database(payload, result)


def collect_supervisor_items(review: Review, analysis, *, approved_only: bool):
    """按当前复核结果收集副主任表导出清单，语义与 ``collect_frontline_items`` 对齐。"""
    mapping = {
        item.group: item.target_workshop for item in analysis.items if item.action == "add"
    }

    def resolved(item):
        if item.action == "add":
            item.target_workshop = to_workshop(effective_workshop(review, item, mapping))
        return item

    items = [
        resolved(item)
        for item in analysis.items
        if (review.decisions.get(item.label) == APPLY)
        or (not approved_only and review.decisions.get(item.label) != CANCEL)
    ]
    return sup_split_by_action(items)


def render_supervisor_export(review: Review, bonus_bytes, bonus, analysis) -> None:
    kept = collect_supervisor_items(review, analysis, approved_only=False)
    approved = collect_supervisor_items(review, analysis, approved_only=True)

    left, right = st.columns(2)
    with left:
        st.subheader("① 对照标记版")
        st.caption(
            "新增人员插到本表对应车间**该职务**的最下方、姓名与员工编号填**绿色**；"
            "需删除的人员保留原行、填**红色**。被点「取消」的人不纳入。"
        )
        adds, removes = kept
        supervisor_preview(adds, removes, "标记删除")
        if st.button("生成对照标记版", type="primary", width="stretch",
                     key=review.widget("gen_mark")):
            supervisor_generate(review, "mark", bonus_bytes, bonus, adds, removes)
        offer_download(review, "mark", bonus.file_name, "副主任表·对照标记版")

    with right:
        st.subheader("② 已应用版")
        st.caption("只处理被点「应用」的人：删除类**直接删行**、新增类**直接插入**且内容用**红色字体**。")
        adds, removes = approved
        supervisor_preview(adds, removes, "直接删除")
        if st.button("生成已应用版", type="primary", width="stretch",
                     key=review.widget("gen_apply")):
            supervisor_generate(review, "apply", bonus_bytes, bonus, adds, removes)
        offer_download(review, "apply", bonus.file_name, "副主任表·已应用版")

    st.info(
        "只改动「副主任&工艺组长及其他」子表的人员行，「一线人员」和其余子表一字不动；"
        "格式、公式、筛选、条件格式全部保留，实习生的职务列填**黄色**底纹。"
    )


def sup_split(items):
    return sup_split_by_action(items)


def supervisor_preview(adds, removes, remove_label: str) -> None:
    ready = [item for item in adds if item.target_workshop]
    pending = len(adds) - len(ready)
    st.write(f"将新增 **{len(ready)}** 人，{remove_label} **{len(removes)}** 人")
    if pending:
        st.caption(f"另有 {pending} 人未指定车间，生成时会被跳过（在上方「车间映射」里指定即可纳入）")


def supervisor_generate(review: Review, mode, bonus_bytes, bonus, adds, removes) -> None:
    if not adds and not removes:
        st.warning("没有需要处理的人员")
        return
    try:
        with st.spinner("正在生成 Excel…"):
            data, summary = build_supervisor_workbook(bonus_bytes, bonus, adds, removes, mode=mode)
    except Exception as exc:  # noqa: BLE001 - 生成失败要给出可读原因
        st.error(f"生成失败：{type(exc).__name__}: {exc}")
        return
    st.session_state[review.key(f"blob_{mode}")] = data
    st.session_state[review.key(f"summary_{mode}")] = summary


def render_combined_feature(payload, result, roster, bonus, frontline_analysis, options_) -> None:
    """把两个功能里的已应用改动写进同一份核算表。"""
    frontline_review = Review("")
    supervisor_review = Review("sup")
    bundle = Review("all")
    bonus_name = options_["bonus_name"]
    include_pending = options_["include_pending"]

    st.caption(
        "一次生成一份核算表，同时改「一线人员」和「副主任&工艺组长及其他」。"
        "默认把两侧**还没点取消**的人都按「应用」处理（直接删行 / 插入，新增红字，实习生黄底）；"
        "已经在 ①② 里点过「取消」的人仍然排除。"
        "同一个人既要从一线移出、又是副主任表待新增时，副主任表只插一行"
        "（一线侧仍是「移到副主任表」时，即使②点了不新增，这个人还是会写进旧表，职务沿用清单写法）。"
        "未指定车间的人请回 ①② 的「车间映射」里指定，否则生成时会被跳过。"
    )
    if bonus.others_layout is None:
        st.error("核算文件里找不到「副主任&工艺组长及其他」子表，无法一键生成两表。")
        return

    placeable = resolved_placeable(frontline_review, frontline_analysis)
    try:
        supervisor_analysis = cached_supervisor_analyze(
            options_["roster_name"],
            result.raw_files[options_["roster_name"]],
            bonus_name,
            result.raw_files[bonus_name],
            options_["duty_field"],
            options_["include_interns"],
            options_["ref_date"],
            options_["new_hire_since"],
            options_["intern_asof"],
            SCOPE_STRICT,
            placeable,
        )
    except Exception as exc:  # noqa: BLE001 - 解析失败要给出可读原因
        st.error(f"对账失败：{type(exc).__name__}: {exc}")
        return

    scope = st.radio(
        "纳入范围",
        ["两侧未取消的全部按应用处理", "只处理已经勾选「应用」的人"],
        horizontal=True,
        key="combined_scope",
        help="第一种就是「一键」：不用先去 ①② 逐条勾选。第二种沿用你在 ①② 里勾过的应用。",
    )
    approved_only = scope == "只处理已经勾选「应用」的人"

    fl_adds, fl_removes, fl_updates, fl_moves = collect_frontline_items(
        frontline_review, frontline_analysis, include_pending, approved_only=approved_only
    )
    sup_adds, sup_removes = collect_supervisor_items(
        supervisor_review, supervisor_analysis, approved_only=approved_only
    )
    move_keys = {item.key for item in fl_moves}
    unique_sup_adds = [
        item for item in merge_others_inserts(fl_moves, sup_adds) if item.key not in move_keys
    ]

    left, right = st.columns(2)
    with left:
        st.subheader("一线人员")
        preview_counts(fl_adds, fl_removes, fl_updates, fl_moves, "直接删除")
    with right:
        st.subheader("副主任&工艺组长及其他")
        supervisor_preview(unique_sup_adds, sup_removes, "直接删除")
        if fl_moves:
            st.caption(
                f"另有 {len(fl_moves)} 人从一线移入，与上表「移到副主任表」是同一批，不会重复插行。"
            )

    overlap = sum(1 for item in sup_adds if item.key in move_keys)
    if overlap:
        st.info(f"有 {overlap} 人同时出现在一线移出和副主任表待新增里，导出时合并为副主任表的一行。")

    if st.button("生成全部已应用版", type="primary", width="stretch", key="combined_gen"):
        generate_combined(
            bundle,
            result.raw_files[bonus_name],
            bonus,
            fl_adds,
            fl_removes,
            fl_updates,
            fl_moves,
            sup_adds,
            sup_removes,
        )
    offer_download(bundle, "apply", bonus.file_name, "两表已应用版")
    st.info(
        "只改「一线人员」和「副主任&工艺组长及其他」两张子表的人员行，"
        "其余子表、字体、行高列宽、公式、筛选和条件格式全部保留；打开后 Excel 会自动重算。"
    )


def generate_combined(
    review: Review,
    bonus_bytes,
    bonus,
    fl_adds,
    fl_removes,
    fl_updates,
    fl_moves,
    sup_adds,
    sup_removes,
) -> None:
    if not any((fl_adds, fl_removes, fl_updates, fl_moves, sup_adds, sup_removes)):
        st.warning("没有需要处理的人员")
        return
    try:
        with st.spinner("正在生成 Excel…"):
            data, summary = build_combined_workbook(
                bonus_bytes,
                bonus,
                fl_adds,
                fl_removes,
                fl_updates,
                fl_moves,
                sup_adds,
                sup_removes,
                mode="apply",
            )
    except Exception as exc:  # noqa: BLE001 - 生成失败要给出可读原因
        st.error(f"生成失败：{type(exc).__name__}: {exc}")
        return
    st.session_state[review.key("blob_apply")] = data
    st.session_state[review.key("summary_apply")] = summary


def render_unmapped_warning(review: Review, adds, mapping) -> None:
    """按当前生效的映射（含人工覆盖）实时统计，手工指定后这里的数字会立刻下降。"""
    pending: dict[str, int] = {}
    for item in adds:
        if not to_workshop(effective_workshop(review, item, mapping)):
            pending[item.group] = pending.get(item.group, 0) + 1
    if not pending:
        return
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


if __name__ == "__main__":
    main()
