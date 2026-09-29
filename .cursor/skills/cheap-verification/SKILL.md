---
name: cheap-verification
description: Use when verifying data-processing or document-generating code without burning tokens and time. Covers writing assertions instead of printing dumps, golden-sample tests on real files, and never pasting large outputs into context.
---

# 低成本验证

## 铁律

```
让程序判断对错并只输出结论，不要把数据倒进上下文自己读。
```

一次 `print(df.head(50))` 可能是几千 token，而且下一轮还得重读。
换成 `assert` + 一行 `OK/FAIL` 摘要，成本降两个数量级，且形成回归保护。

## 具体做法

1. **探查脚本一次写全**：把所有想看的结构（sheet 名、表头、distinct 值、计数、样例 3 条）
   写进一个脚本跑一次，而不是来回十次交互。看完就删掉探查脚本。
2. **输出只打摘要**：`print(len(x), Counter(...).most_common(5))`，不要 `print(x)`。
   样例最多 3 条，字段截断到 20 字符。
3. **命令输出重定向到文件**，只在真正需要时读，且优先用 `Grep` 定位而不是整篇读。
4. **黄金样本测试**：拿真实文件当 fixture，断言已知个体的分类结果和守恒关系。
   比自造 mock 数据更能抓真 bug，且不需要维护 fixture。
5. **生成类功能的验证 = 重新加载 + 断言**：生成 xlsx 后用 openpyxl 载回来断言值/公式/样式/合并区，
   不要靠肉眼看截图。截图只用于最后确认 UI 布局。
6. **UI 只验证"能起来"**：`streamlit run` 后检查端口和日志无 traceback 即可；
   逻辑正确性全部由 pytest 保证。真要看界面就一次截图，不要反复截。

## 反面清单

- 反复 `cat` 整个文件 —— 用 `Read` 带 offset/limit 或 `Grep`。
- 把大 JSON/DataFrame 打到 stdout 再人工比对。
- 每改一行就跑一遍全量流程 —— 先写测试，改完一次性跑。
- 用截图代替断言。
