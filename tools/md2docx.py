# -*- coding: utf-8 -*-
"""把 interview-prep.md 转成 docx（标题/列表/表格/引用/代码块 + 中文字体）。

只为这份文档服务，不追求通用性；但处理了 Markdown 里实际出现的语法。
"""
from __future__ import annotations

import re
import sys

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "微软雅黑"
MONO = "Consolas"
ACCENT = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x44, 0x44, 0x44)
BODY = 10.5


def style_run(run, size=BODY, bold=False, color=GREY, mono=False):
    name = MONO if mono else FONT
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def add_rich_text(paragraph, text, size=BODY, base_bold=False, color=GREY, mono=False):
    """处理 **粗体** 与 `代码` 内联标记。"""
    tokens = re.split(r"(\*\*.+?\*\*|`[^`]+`)", text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**"):
            style_run(paragraph.add_run(token[2:-2]), size=size, bold=True, color=RGBColor(0, 0, 0), mono=mono)
        elif token.startswith("`") and token.endswith("`"):
            style_run(paragraph.add_run(token[1:-1]), size=size - 0.5, bold=base_bold, color=GREY, mono=True)
        else:
            style_run(paragraph.add_run(token), size=size, bold=base_bold, color=color, mono=mono)


def setup(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    sec.left_margin = sec.right_margin = Cm(1.8)
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(BODY)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def add_table(doc, rows):
    header, body = rows[0], rows[1:]
    table = doc.add_table(rows=1, cols=len(header))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, cell in enumerate(header):
        paragraph = table.rows[0].cells[i].paragraphs[0]
        add_rich_text(paragraph, cell, size=9.0, base_bold=True, color=ACCENT)
    for row in body:
        cells = table.add_row().cells
        for i, cell in enumerate(row[: len(header)]):
            add_rich_text(cells[i].paragraphs[0], cell, size=9.0)
    doc.add_paragraph()


def main() -> int:
    src, dst = sys.argv[1], sys.argv[2]
    lines = open(src, encoding="utf-8").read().splitlines()

    doc = Document()
    setup(doc)

    index = 0
    while index < len(lines):
        line = lines[index].rstrip()
        stripped = line.strip()

        # 表格
        if stripped.startswith("|") and index + 1 < len(lines) and set(lines[index + 1].strip()) <= set("|-: "):
            block = []
            while index < len(lines) and lines[index].strip().startswith("|"):
                raw = lines[index].strip()
                if not set(raw) <= set("|-: "):
                    cells = [c.strip() for c in raw.strip("|").split("|")]
                    block.append(cells)
                index += 1
            if block:
                add_table(doc, block)
            continue

        # 代码块
        if stripped.startswith("```"):
            index += 1
            code: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith("```"):
                code.append(lines[index])
                index += 1
            index += 1
            for code_line in code:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.left_indent = Cm(0.4)
                style_run(p.add_run(code_line if code_line else " "), size=9.0, mono=True)
            doc.add_paragraph()
            continue

        if not stripped:
            index += 1
            continue

        if stripped == "---":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            style_run(p.add_run(" "), size=4)
            index += 1
            continue

        # 标题
        if stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            text = stripped[level:].strip()
            size = {1: 20, 2: 14, 3: 12, 4: 11}.get(level, 11)
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(10 if level <= 2 else 6)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            if level == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                style_run(p.add_run(text), size=size, bold=True, color=ACCENT)
            else:
                style_run(p.add_run(text), size=size, bold=True, color=ACCENT)
                if level == 2:
                    pPr = p._p.get_or_add_pPr()
                    pbdr = OxmlElement("w:pBdr")
                    bottom = OxmlElement("w:bottom")
                    bottom.set(qn("w:val"), "single")
                    bottom.set(qn("w:sz"), "6")
                    bottom.set(qn("w:space"), "2")
                    bottom.set(qn("w:color"), "1F3A5F")
                    pbdr.append(bottom)
                    pPr.append(pbdr)
            index += 1
            continue

        # 引用
        if stripped.startswith(">"):
            text = stripped.lstrip(">").strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            add_rich_text(p, text, size=BODY, base_bold=True, color=RGBColor(0x1F, 0x3A, 0x5F))
            index += 1
            continue

        # 列表
        if stripped.startswith("- ") or re.match(r"^\d+\. ", stripped):
            text = stripped[2:] if stripped.startswith("- ") else re.sub(r"^\d+\. ", "", stripped)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.first_line_indent = Cm(-0.35)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            add_rich_text(p, "•  " + text)
            index += 1
            continue

        # 普通段落
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        add_rich_text(p, stripped)
        index += 1

    doc.save(dst)
    print("saved:", dst)
    return 0


if __name__ == "__main__":
    sys.exit(main())
