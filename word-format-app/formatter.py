from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

TITLE_FONT = "方正小标宋_GBK"
BODY_FONT = "方正仿宋_GBK"
LEVEL1_FONT = "方正黑体_GBK"
LEVEL2_FONT = "方正楷体_GBK"
WESTERN_FONT = "Times New Roman"
PAGE_FONT = "宋体"

TITLE_SIZE = Pt(22)
BODY_SIZE = Pt(16)
PAGE_SIZE = Pt(14)
LINE_SPACING = Pt(29.5)

DOC_NO_RE = re.compile(r".+〔\d{4}〕\d+号$")
LEVEL1_RE = re.compile(r"^[一二三四五六七八九十]+、")
LEVEL2_RE = re.compile(r"^（[一二三四五六七八九十]+）")
LEVEL3_RE = re.compile(r"^\d+\.")
LEVEL4_RE = re.compile(r"^（\d+）")
ARABIC_ITEM_RE = re.compile(r"^\s*(\d+)[.．、]?\s*")
DATE_RE = re.compile(r"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$")
ATTACHMENT_RE = re.compile(r"^附件[：:]")
LATIN_RE = re.compile(r"[A-Za-z0-9]")


@dataclass
class ParagraphReport:
    index: int
    kind: str
    text: str
    font_east_asia: str
    font_ascii: str
    font_size_pt: float
    alignment: str
    first_line_indent_pt: float
    left_indent_pt: float
    line_spacing_pt: float
    actions: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class FormatSummary:
    paragraphs_updated: int = 0
    tables_updated: int = 0
    footer_updated: bool = False
    warnings: list[str] = field(default_factory=list)
    paragraph_reports: list[ParagraphReport] = field(default_factory=list)

    def to_dict(self) -> dict:
        payload = asdict(self)
        payload["warning_count"] = len(self.warnings)
        payload["paragraph_count"] = len(self.paragraph_reports)
        payload["kind_counts"] = _count_kinds(self.paragraph_reports)
        return payload


def format_document(input_path: Path, output_path: Path) -> FormatSummary:
    doc = Document(str(input_path))
    summary = FormatSummary()

    _set_document_layout(doc)
    _normalize_paragraph_structure(doc)
    _normalize_numbered_sequences(doc)
    _format_paragraphs(doc, summary)
    _format_tables(doc, summary)
    summary.footer_updated = _set_page_footer(doc)

    if summary.footer_updated:
        summary.warnings.append("页码已统一设置为页脚右侧；如需区分奇偶页外侧，还需在 Word 中进一步处理。")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return summary


def _count_kinds(paragraph_reports: list[ParagraphReport]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in paragraph_reports:
        counts[item.kind] = counts.get(item.kind, 0) + 1
    return counts


def _set_document_layout(doc: DocumentObject) -> None:
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(37)
        section.bottom_margin = Mm(32)
        section.left_margin = Mm(28)
        section.right_margin = Mm(26)
        section.header_distance = Mm(15)
        section.footer_distance = Mm(15)
        section.start_type = WD_SECTION.CONTINUOUS


def _normalize_paragraph_structure(doc: DocumentObject) -> None:
    for paragraph in list(doc.paragraphs):
        text = paragraph.text
        if not text or "\n" not in text:
            continue

        parts = [part.strip() for part in text.splitlines() if part.strip()]
        if len(parts) <= 1:
            continue

        paragraph.text = parts[0]
        anchor = paragraph
        for part in parts[1:]:
            anchor = _insert_paragraph_after(anchor, part)


def _insert_paragraph_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para


def _normalize_numbered_sequences(doc: DocumentObject) -> None:
    counter = None
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        if LEVEL1_RE.match(text):
            counter = None
            continue

        if _has_word_numbering(paragraph) or ARABIC_ITEM_RE.match(text):
            counter = 1 if counter is None else counter + 1
            paragraph.text = _replace_leading_number(text, counter)
            _remove_word_numbering(paragraph)
        else:
            counter = None


def _has_word_numbering(paragraph: Paragraph) -> bool:
    p_pr = paragraph._p.pPr
    return bool(p_pr is not None and p_pr.numPr is not None)


def _remove_word_numbering(paragraph: Paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    num_pr = p_pr.numPr
    if num_pr is not None:
        p_pr.remove(num_pr)


def _replace_leading_number(text: str, number: int) -> str:
    stripped = ARABIC_ITEM_RE.sub("", text, count=1).lstrip()
    return f"{number}.{stripped}"


def _format_paragraphs(doc: DocumentObject, summary: FormatSummary) -> None:
    title_indices = _detect_title_indices(doc.paragraphs)

    for idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:
            _normalize_paragraph_spacing(paragraph)
            continue

        kind = _classify_paragraph(text, idx, title_indices)
        actions: list[str] = []
        warnings: list[str] = []

        if kind == "title":
            _normalize_title_text(paragraph)
            _apply_paragraph_style(
                paragraph,
                TITLE_FONT,
                TITLE_SIZE,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
            )
            actions.append("按标题设置为方正小标宋_GBK二号居中，自动优化断行")
        elif kind == "doc_number":
            _apply_paragraph_style(
                paragraph,
                BODY_FONT,
                BODY_SIZE,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                space_before_pt=LINE_SPACING.pt * 2,
            )
            actions.append("按发文字号设置为方正仿宋_GBK三号居中，下空2行")
            if "第" in text or "001" in text:
                warnings.append("发文字号可能包含“第”字或虚位序号，建议人工复核。")
        elif kind == "level1":
            _apply_paragraph_style(
                paragraph,
                LEVEL1_FONT,
                BODY_SIZE,
                bold=False,
                first_line_chars=2,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )
            actions.append("按一级标题设置为方正黑体_GBK三号，首行缩进2字")
        elif kind == "level2":
            _apply_paragraph_style(
                paragraph,
                LEVEL2_FONT,
                BODY_SIZE,
                bold=False,
                first_line_chars=2,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )
            actions.append("按二级标题设置为方正楷体_GBK三号，首行缩进2字")
        elif kind in {"level3", "level4", "attachment"}:
            _apply_paragraph_style(
                paragraph,
                BODY_FONT,
                BODY_SIZE,
                bold=False,
                first_line_chars=2,
                left_chars=2 if kind == "attachment" else None,
                space_before_pt=LINE_SPACING.pt if kind == "attachment" else 0,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )
            if kind == "attachment":
                actions.append("按附件格式设置为左空2字、首行缩进2字，正文下空1行")
                if text.endswith(("。", "；", "，")):
                    warnings.append("附件名称后建议不加标点符号。")
            else:
                actions.append("按正文层级设置为方正仿宋_GBK三号，首行缩进2字")
        elif kind == "salutation":
            _apply_paragraph_style(
                paragraph,
                BODY_FONT,
                BODY_SIZE,
                bold=False,
                first_line_chars=2,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )
            actions.append("按称呼段设置为方正仿宋_GBK三号，首行缩进2字")
        elif kind == "date":
            _apply_paragraph_style(
                paragraph,
                BODY_FONT,
                BODY_SIZE,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                first_line_chars=2,
            )
            actions.append("按成文日期设置为方正仿宋_GBK三号右对齐，首行缩进2字")
            if any(ch in text for ch in "〇零壹贰叁肆伍陆柒捌玖拾"):
                warnings.append("成文日期疑似使用大写汉字，建议改为阿拉伯数字。")
        else:
            _apply_paragraph_style(
                paragraph,
                BODY_FONT,
                BODY_SIZE,
                bold=False,
                first_line_chars=2,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
            )
            actions.append("按正文设置为方正仿宋_GBK三号，首行缩进2字")

        _apply_western_font_normalization(paragraph, actions)
        _collect_sequence_warnings(kind, text, warnings)
        if len(text) > 120 and kind in {"title", "level1", "level2"}:
            warnings.append("该段较长但被识别为标题类，建议人工复核。")

        report = _build_paragraph_report(paragraph, idx + 1, kind, text, actions, warnings)
        summary.paragraph_reports.append(report)
        summary.warnings.extend(f"第{report.index}段：{warning}" for warning in warnings)
        summary.paragraphs_updated += 1


def _classify_paragraph(text: str, index: int, title_indices: set[int]) -> str:
    if DOC_NO_RE.match(text):
        return "doc_number"
    if index in title_indices:
        return "title"
    if LEVEL1_RE.match(text):
        return "level1"
    if LEVEL2_RE.match(text):
        return "level2"
    if LEVEL3_RE.match(text):
        return "level3"
    if LEVEL4_RE.match(text):
        return "level4"
    if ATTACHMENT_RE.match(text):
        return "attachment"
    if text.endswith(("：", ":")):
        return "salutation"
    if DATE_RE.match(text):
        return "date"
    return "body"


def _detect_title_indices(paragraphs: list[Paragraph]) -> set[int]:
    non_empty = [(idx, p.text.strip()) for idx, p in enumerate(paragraphs) if p.text.strip()]
    if not non_empty:
        return set()

    doc_no_index = next((idx for idx, text in non_empty if DOC_NO_RE.match(text)), None)
    salutation_index = next((idx for idx, text in non_empty if text.endswith(("：", ":"))), None)
    first_level1_index = next((idx for idx, text in non_empty if LEVEL1_RE.match(text)), None)

    start = 0
    if doc_no_index is not None:
        start = next((pos for pos, (idx, _) in enumerate(non_empty) if idx == doc_no_index), 0) + 1

    end = None
    if salutation_index is not None:
        end = next((pos for pos, (idx, _) in enumerate(non_empty) if idx == salutation_index), None)
    elif first_level1_index is not None:
        end = next((pos for pos, (idx, _) in enumerate(non_empty) if idx == first_level1_index), None)

    candidates = non_empty[start:end]
    if not candidates and end is not None:
        candidates = non_empty[:end]

    title_indices = set()
    for idx, text in candidates[:2]:
        if len(text) <= 40 and not re.search(r"[。；，：:]", text):
            title_indices.add(idx)

    if not title_indices and non_empty:
        first_idx, first_text = non_empty[0]
        if len(first_text) <= 40 and first_level1_index not in {None, first_idx}:
            title_indices.add(first_idx)

    return title_indices


def _apply_paragraph_style(
    paragraph: Paragraph,
    east_asia_font: str,
    font_size,
    *,
    bold: bool,
    alignment: WD_ALIGN_PARAGRAPH | None = None,
    first_line_chars: int | None = None,
    left_chars: int | None = None,
    space_before_pt: float = 0,
    space_after_pt: float = 0,
) -> None:
    _normalize_paragraph_spacing(paragraph)
    if alignment is not None:
        paragraph.alignment = alignment
    _clear_paragraph_border(paragraph)

    paragraph_format = paragraph.paragraph_format
    paragraph_format.line_spacing = LINE_SPACING
    paragraph_format.space_before = Pt(space_before_pt)
    paragraph_format.space_after = Pt(space_after_pt)
    paragraph_format.left_indent = Pt(BODY_SIZE.pt * left_chars) if left_chars is not None else Pt(0)
    paragraph_format.first_line_indent = Pt(BODY_SIZE.pt * first_line_chars) if first_line_chars is not None else Pt(0)

    if not paragraph.runs:
        run = paragraph.add_run("")
        _set_run_fonts(run, east_asia_font, font_size, bold)
        return

    for run in paragraph.runs:
        _set_run_fonts(run, east_asia_font, font_size, bold)


def _set_run_fonts(run: Run, east_asia_font: str, font_size, bold: bool) -> None:
    run.font.size = font_size
    run.font.bold = bold
    run.font.name = WESTERN_FONT
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.highlight_color = None
    run.font.underline = False

    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)

    r_fonts.set(qn("w:ascii"), WESTERN_FONT)
    r_fonts.set(qn("w:hAnsi"), WESTERN_FONT)
    r_fonts.set(qn("w:cs"), WESTERN_FONT)
    r_fonts.set(qn("w:eastAsia"), east_asia_font)
    color = r_pr.color
    if color is None:
        color = OxmlElement("w:color")
        r_pr.append(color)
    color.set(qn("w:val"), "000000")


def _normalize_title_text(paragraph: Paragraph) -> None:
    text = "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text
    text = text.strip()
    if not text:
        return

    for run in list(paragraph.runs):
        run._element.getparent().remove(run._element)

    lines = _split_title_lines(text)
    for idx, line in enumerate(lines):
        run = paragraph.add_run(line)
        if idx < len(lines) - 1:
            run.add_break()


def _split_title_lines(text: str) -> list[str]:
    if len(text) <= 22:
        return [text]

    if "（" in text and text.endswith("）"):
        main, suffix = text.split("（", 1)
        main = main.strip()
        suffix = "（" + suffix.strip()
        if len(main) >= 10:
            return [main, suffix]

    midpoint = len(text) // 2
    candidates = []
    markers = ["汇总", "情况", "意见", "建议", "报告", "方案", "通知", "年度", "会议", "组织生活会"]
    for marker in markers:
        pos = text.find(marker)
        if pos != -1:
            break_pos = pos + len(marker)
            if 6 <= break_pos <= len(text) - 6:
                candidates.append(break_pos)

    if candidates:
        break_pos = min(candidates, key=lambda pos: abs(pos - midpoint))
        return [text[:break_pos], text[break_pos:]]

    break_pos = midpoint
    return [text[:break_pos], text[break_pos:]]


def _clear_paragraph_border(paragraph: Paragraph) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is not None:
        p_pr.remove(p_bdr)


def _normalize_paragraph_spacing(paragraph: Paragraph) -> None:
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    paragraph_format.line_spacing = LINE_SPACING
    if paragraph.alignment is None:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _apply_western_font_normalization(paragraph: Paragraph, actions: list[str]) -> None:
    if LATIN_RE.search(paragraph.text):
        actions.append("数字和西文字母统一设置为 Times New Roman")


def _collect_sequence_warnings(kind: str, text: str, warnings: list[str]) -> None:
    if kind == "level1" and not LEVEL1_RE.match(text):
        warnings.append("一级标题序号格式可能不符合“一、”规范。")
    if kind == "level2" and not LEVEL2_RE.match(text):
        warnings.append("二级标题序号格式可能不符合“（一）”规范。")
    if kind == "level3" and not LEVEL3_RE.match(text):
        warnings.append("三级标题序号格式可能不符合“1.”规范。")
    if kind == "level4" and not LEVEL4_RE.match(text):
        warnings.append("四级标题序号格式可能不符合“（1）”规范。")
    if kind == "body" and re.match(r"^[（(]?[一二三四五六七八九十\d]+[）).、]", text):
        warnings.append("该段看起来像有层次序号，但未被识别到对应级别，建议人工复核。")


def _build_paragraph_report(
    paragraph: Paragraph,
    index: int,
    kind: str,
    text: str,
    actions: list[str],
    warnings: list[str],
) -> ParagraphReport:
    run = next((item for item in paragraph.runs if item.text), paragraph.runs[0] if paragraph.runs else None)
    east_font = ""
    ascii_font = ""
    font_size_pt = 0.0
    if run is not None:
        r_pr = run._element.get_or_add_rPr()
        r_fonts = r_pr.rFonts
        if r_fonts is not None:
            east_font = r_fonts.get(qn("w:eastAsia")) or ""
            ascii_font = r_fonts.get(qn("w:ascii")) or ""
        if run.font.size is not None:
            font_size_pt = float(run.font.size.pt)

    pf = paragraph.paragraph_format
    line_spacing_pt = float(pf.line_spacing.pt) if hasattr(pf.line_spacing, "pt") else float(LINE_SPACING.pt)
    return ParagraphReport(
        index=index,
        kind=kind,
        text=text,
        font_east_asia=east_font,
        font_ascii=ascii_font,
        font_size_pt=font_size_pt,
        alignment=_alignment_name(paragraph.alignment),
        first_line_indent_pt=float(pf.first_line_indent.pt) if pf.first_line_indent is not None else 0.0,
        left_indent_pt=float(pf.left_indent.pt) if pf.left_indent is not None else 0.0,
        line_spacing_pt=line_spacing_pt,
        actions=actions,
        warnings=warnings,
    )


def _alignment_name(alignment: WD_ALIGN_PARAGRAPH | None) -> str:
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    return mapping.get(alignment, "unknown")


def _format_tables(doc: DocumentObject, summary: FormatSummary) -> None:
    for table in doc.tables:
        _format_table(table)
        summary.tables_updated += 1
    if doc.tables:
        summary.warnings.append("表格内容已统一基础字体，但未对表格边框、列宽、跨页策略做专项优化。")


def _format_table(table: Table) -> None:
    for row in table.rows:
        for cell in row.cells:
            _format_cell(cell)


def _format_cell(cell: _Cell) -> None:
    for paragraph in cell.paragraphs:
        text = paragraph.text.strip()
        kind = "level1" if LEVEL1_RE.match(text) else "body"
        if kind == "level1":
            _apply_paragraph_style(paragraph, LEVEL1_FONT, BODY_SIZE, bold=False, first_line_chars=2, alignment=WD_ALIGN_PARAGRAPH.LEFT)
        else:
            _apply_paragraph_style(paragraph, BODY_FONT, BODY_SIZE, bold=False, first_line_chars=2, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    for table in cell.tables:
        _format_table(table)


def _set_page_footer(doc: DocumentObject) -> bool:
    updated = False
    for section in doc.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.clear()
        run_left = paragraph.add_run("- ")
        _set_page_fonts(run_left)
        _append_page_field(paragraph)
        run_right = paragraph.add_run(" -")
        _set_page_fonts(run_right)
        updated = True
    return updated


def _set_page_fonts(run: Run) -> None:
    run.font.size = PAGE_SIZE
    run.font.name = PAGE_FONT
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:ascii"), WESTERN_FONT)
    r_fonts.set(qn("w:hAnsi"), WESTERN_FONT)
    r_fonts.set(qn("w:cs"), WESTERN_FONT)
    r_fonts.set(qn("w:eastAsia"), PAGE_FONT)


def _append_page_field(paragraph: Paragraph) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "

    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    run = paragraph.add_run()
    _set_page_fonts(run)
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)
