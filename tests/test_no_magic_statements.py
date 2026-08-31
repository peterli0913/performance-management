"""静态检查：app.py 里不能有会被 Streamlit magic 误渲染的裸表达式。

Streamlit 默认开着 magic，会把裸表达式语句包一层 `transparent_write()` 打到页面上。
按 `streamlit/runtime/scriptrunner/magic.py` 的实现，只有这几类会被跳过：
  * `ast.Call`（任何函数调用，所以正常的 `st.xxx(...)` 语句没事）
  * 文档字符串
  * `yield` / `yield from` / `await`

所以 `st.info(x) if cond else st.warning(x)` 这种条件表达式会被包起来，
把 `DeltaGenerator` 的 repr 当正文渲染出来——用户看到的就是"一坨代码"。
这类问题肉眼审查很容易漏，用 AST 静态兜住。
"""

import ast
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP = os.path.join(ROOT, "app.py")

_DOCSTRING_PARENTS = (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)


def _docstring_nodes(tree: ast.AST) -> set[int]:
    """收集所有文档字符串节点的 id，它们不会被 magic 包装。"""
    found = set()
    for node in ast.walk(tree):
        if not isinstance(node, _DOCSTRING_PARENTS):
            continue
        body = getattr(node, "body", None)
        if not body:
            continue
        first = body[0]
        if isinstance(first, ast.Expr) and isinstance(first.value, ast.Constant):
            if isinstance(first.value.value, str):
                found.add(id(first))
    return found


def _offenders(path: str) -> list[tuple[int, str]]:
    tree = ast.parse(open(path, encoding="utf-8").read())
    docstrings = _docstring_nodes(tree)
    skip = (ast.Call, ast.Yield, ast.YieldFrom, ast.Await)
    out = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Expr) or id(node) in docstrings:
            continue
        if isinstance(node.value, skip):
            continue
        out.append((node.lineno, type(node.value).__name__))
    return out


def test_no_bare_expressions_that_magic_would_render():
    offenders = _offenders(APP)
    assert offenders == [], (
        "下面这些裸表达式会被 Streamlit magic 包成 transparent_write() 打到页面上，"
        "请改写成 if/else 之类的语句：\n"
        + "\n".join(f"  app.py:{line} ({kind})" for line, kind in offenders)
    )


def test_the_check_actually_catches_the_pattern(tmp_path):
    """守住这个守卫本身：改坏了要能报出来。"""
    sample = tmp_path / "bad_app.py"
    sample.write_text(
        "import streamlit as st\n"
        "for note in []:\n"
        "    st.info(note) if note else st.warning(note)\n",
        encoding="utf-8",
    )
    assert _offenders(str(sample)) == [(3, "IfExp")]
