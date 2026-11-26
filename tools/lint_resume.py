# -*- coding: utf-8 -*-
"""简历文案与排版自检：把「规范」变成可执行的检查项。

用法：
    python lint_resume.py 简历.docx

检查依据（公开的简历写作规范）：
- 避免第一人称与口语化表达，全文用动宾结构（超级简历）
- 中文简历避免叹号、问号、省略号；中英文标点不混用（超级简历）
- 中英文之间留空格（中英混排排版惯例）
- 每条描述尽量含量化结果（STAR + 量化）
- 段落长度可控，避免一条 bullet 塞三件事
- 页边距、字号、行距在常规区间（A4 页边距 2cm、正文 10–10.5pt、行距 1.15–1.3）
"""
from __future__ import annotations

import pathlib
import re
import sys

from docx import Document
from docx.shared import Cm

# 口语 / 主观评价 / 演讲句式
COLLOQUIAL = [
    "拍脑袋", "挡在", "打到", "搞定", "弄", "跑通", "卡住", "顺手", "一堆",
    "很", "非常", "比较", "大概", "差不多", "其实", "所以", "然后", "而且",
    "但是", "而不是", "关注的不是", "代价是", "结论是", "我觉得", "可以看到",
    "值得注意的是", "众所周知", "极大地", "有效地", "显著地",
]
FIRST_PERSON = ["我们", "本人", "我", "咱们"]
BAD_PUNCT = ["！", "？", "…", "!!", "??"]
# 中英文之间缺空格：中文紧跟英文/数字，或英文紧跟中文
CJK = r"\u4e00-\u9fff"
MISSING_SPACE_1 = re.compile(rf"[{CJK}][A-Za-z]")
MISSING_SPACE_2 = re.compile(rf"[A-Za-z][{CJK}]")
# 允许例外：英文缩写与中文常见搭配（如 A/B 测试）先不排除，人工复核
PERIOD_END = re.compile(r"[。；;]\s*$")


def iter_paragraphs(doc):
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            yield p, text


def check_layout(doc) -> list[str]:
    issues: list[str] = []
    sec = doc.sections[0]
    margins = {
        "上": sec.top_margin.cm,
        "下": sec.bottom_margin.cm,
        "左": sec.left_margin.cm,
        "右": sec.right_margin.cm,
    }
    for side, value in margins.items():
        if value < 1.8:
            issues.append(f"[排版] {side}边距 {value:.1f}cm，低于常规 2cm（A4 简历建议 1.8–2.5cm）")

    sizes = {
        round(run.font.size.pt, 1)
        for p, _ in iter_paragraphs(doc)
        for run in p.runs
        if run.font.size is not None
    }
    body = sorted(s for s in sizes if 8 <= s <= 12)
    if body and min(body) < 10:
        issues.append(f"[排版] 正文字号最小 {min(body)}pt，低于常规 10–10.5pt（打印与投屏偏小）")

    spacings = {
        round(p.paragraph_format.line_spacing, 2)
        for p, _ in iter_paragraphs(doc)
        if isinstance(p.paragraph_format.line_spacing, float)
    }
    if spacings and max(spacings) < 1.15:
        issues.append(f"[排版] 行距 {max(spacings)}，低于常规 1.15–1.3（中文正文建议固定值 20–24pt）")
    return issues


def check_text(doc) -> list[str]:
    issues: list[str] = []
    body_lines = [t for _, t in iter_paragraphs(doc)]

    for line in body_lines:
        for word in FIRST_PERSON:
            if word in line:
                issues.append(f"[用语] 出现第一人称「{word}」：{line[:36]}…")
        for word in COLLOQUIAL:
            if word in line:
                issues.append(f"[用语] 口语/评价性表达「{word}」：{line[:36]}…")
        for mark in BAD_PUNCT:
            if mark in line:
                issues.append(f"[标点] 出现「{mark}」：{line[:36]}…")
        if MISSING_SPACE_1.search(line):
            sample = MISSING_SPACE_1.search(line).group(0)
            issues.append(f"[排版] 中英之间缺空格「{sample}」：{line[:36]}…")
        if MISSING_SPACE_2.search(line):
            sample = MISSING_SPACE_2.search(line).group(0)
            issues.append(f"[排版] 中英之间缺空格「{sample}」：{line[:36]}…")
        if PERIOD_END.search(line):
            issues.append(f"[标点] 句末带标点（中文简历惯例不加）：{line[:36]}…")
        if len(line) > 95 and line.startswith("•"):
            issues.append(f"[长度] 单条 {len(line)} 字，超过 95 字易丢重点：{line[:30]}…")
    return issues


def check_quantification(doc) -> list[str]:
    """bullet 是否含量化结果（数字 / 百分比 / 时长）。"""
    issues: list[str] = []
    for _, text in iter_paragraphs(doc):
        if not text.startswith("•"):
            continue
        if not re.search(r"\d", text):
            issues.append(f"[量化] 该条无数字：{text[:40]}…")
    return issues


def check_duplicates(doc) -> list[str]:
    """同一个具体数字被反复提及（简历里重复等于浪费版面）。"""
    issues: list[str] = []
    text = "\n".join(t for _, t in iter_paragraphs(doc))
    for token in ("58 条", "226 条", "5pp", "48/48", "22/33", "1,639", "748 条"):
        count = text.count(token)
        if count > 1:
            issues.append(f"[重复] 「{token}」出现 {count} 次")
    return issues


def main() -> int:
    path = sys.argv[1]
    doc = Document(path)
    issues = check_layout(doc) + check_text(doc) + check_quantification(doc) + check_duplicates(doc)

    lines = [f"== {path} =="]
    if not issues:
        lines.append("未发现问题")
    else:
        lines.extend(f" - {item}" for item in issues)
        lines.append(f"\n共 {len(issues)} 项")
    report = "\n".join(lines)
    # 控制台可能是 GBK，统一写一份 UTF-8 报告，便于查看
    out = pathlib.Path(path).with_suffix(".lint.txt")
    out.write_text(report, encoding="utf-8")
    print("报告已写出:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
