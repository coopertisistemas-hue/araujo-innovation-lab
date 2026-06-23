"""
Gerador do documento institucional:
Framework Conceitual para Modernização do Curso Técnico em Marketing
Colégio Araújo — Araujo Innovation Lab
v1.0 — Junho de 2026
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[2]
OUT_DIR = ROOT / "output" / "doc"
OUTPUT_DOCX = OUT_DIR / "marketing-course-modernization-framework-v1.docx"


# ── Paleta institucional aprovada ─────────────────────────────────────────────
C = {
    "blue":       "1F3A45",
    "blue_soft":  "EAF0EE",
    "blue_soft2": "F1F5F3",
    "gold":       "C7A446",
    "gold_soft":  "F5EFDB",
    "ink":        "1F2C31",
    "body":       "34434A",
    "muted":      "6E7A80",
    "line":       "D7DEDA",
    "bg":         "F7F7F5",
    "panel":      "FFFFFF",
    "green_soft": "EEF4F0",
    "white":      "FFFFFF",
}


# ── Helpers de formatação ──────────────────────────────────────────────────────

def run_style(run, font_name: str, size: int, color: str,
              bold: bool = False, italic: bool = False) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.bold = bold
    run.font.italic = italic


def cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def cell_borders(cell, color: str = "D7DEDA", size: str = "8") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("left", "top", "right", "bottom"):
        el = tc_borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tc_borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:color"), color)


def no_borders(cell) -> None:
    cell_borders(cell, color="FFFFFF", size="0")


def para_bottom_border(paragraph, color: str = "D7DEDA", size: str = "6") -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.first_child_found_in("w:pBdr")
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def add_page_field(paragraph) -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    fld.append(run)
    paragraph._p.append(fld)


def add_numpages_field(paragraph) -> None:
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "NUMPAGES")
    run = OxmlElement("w:r")
    fld.append(run)
    paragraph._p.append(fld)


# ── Estilos do documento ───────────────────────────────────────────────────────

def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string(C["body"])
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.22

    title_style = doc.styles["Title"]
    title_style.font.name = "Georgia"
    title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Georgia")
    title_style.font.size = Pt(22)
    title_style.font.color.rgb = RGBColor.from_string(C["ink"])
    title_style.font.bold = True

    for style_name, size in [("Heading 1", 15), ("Heading 2", 12), ("Heading 3", 11)]:
        s = doc.styles[style_name]
        s.font.name = "Georgia" if style_name == "Heading 1" else "Arial"
        s._element.rPr.rFonts.set(qn("w:eastAsia"), s.font.name)
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = RGBColor.from_string(C["blue"])
        s.paragraph_format.space_before = Pt(18 if style_name == "Heading 1" else 12)
        s.paragraph_format.space_after = Pt(6)

    if "Caption" in doc.styles:
        cap = doc.styles["Caption"]
        cap.font.name = "Arial"
        cap._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
        cap.font.size = Pt(9)
        cap.font.color.rgb = RGBColor.from_string(C["muted"])


def configure_section(section, header_text: str, footer_right: str,
                      with_hf: bool = True) -> None:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.2)

    if not with_hf:
        return

    header = section.header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.text = header_text
    for run in p.runs:
        run_style(run, "Arial", 9, C["muted"])
    para_bottom_border(p)

    footer = section.footer
    table = footer.add_table(rows=1, cols=3, width=Inches(6.0))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cells = table.rows[0].cells

    # Esquerda: classificação
    p_left = cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_left = p_left.add_run("Circulação Interna Restrita  |  v1.0")
    run_style(r_left, "Arial", 8, C["muted"])
    no_borders(cells[0])

    # Centro: número da página
    p_center = cells[1].paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_pg = p_center.add_run("Página ")
    run_style(r_pg, "Arial", 8, C["muted"])
    add_page_field(p_center)
    r_of = p_center.add_run(" de ")
    run_style(r_of, "Arial", 8, C["muted"])
    add_numpages_field(p_center)
    no_borders(cells[1])

    # Direita: identificação institucional
    p_right = cells[2].paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_right = p_right.add_run(footer_right)
    run_style(r_right, "Arial", 8, C["muted"])
    no_borders(cells[2])


# ── Primitivos de conteúdo ────────────────────────────────────────────────────

def add_body(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    run_style(r, "Arial", 11, C["body"])


def add_body_bold_intro(doc: Document, bold_part: str, rest: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r1 = p.add_run(bold_part)
    run_style(r1, "Arial", 11, C["body"], bold=True)
    r2 = p.add_run(rest)
    run_style(r2, "Arial", 11, C["body"])


def add_bullet(doc: Document, text: str, indent: float = 0.6) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(f"•  {text}")
    run_style(r, "Arial", 11, C["body"])


def add_sub_bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.2)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"–  {text}")
    run_style(r, "Arial", 10, C["body"])


def add_section_heading(doc: Document, number: str, title: str) -> None:
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(22)
    p.clear()
    r_num = p.add_run(f"{number}.  ")
    run_style(r_num, "Georgia", 15, C["gold"], bold=True)
    r_title = p.add_run(title)
    run_style(r_title, "Georgia", 15, C["blue"], bold=True)
    para_bottom_border(p, color=C["line"], size="4")


def add_subsection(doc: Document, title: str) -> None:
    p = doc.add_paragraph(style="Heading 2")
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(5)
    p.clear()
    r = p.add_run(title)
    run_style(r, "Arial", 12, C["blue"], bold=True)


def add_highlight_block(doc: Document, label: str, text: str,
                        fill: str = "F1F5F3", border_color: str = "D7DEDA") -> None:
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell_shading(cell, fill)
    cell_borders(cell, color=border_color, size="12")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    if label:
        r1 = p.add_run(f"{label}:  ")
        run_style(r1, "Arial", 10, C["blue"], bold=True)
    r2 = p.add_run(text)
    run_style(r2, "Arial", 10, C["body"])
    doc.add_paragraph()


def add_note_block(doc: Document, text: str) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell_shading(cell, C["gold_soft"])
    cell_borders(cell, color=C["gold"], size="10")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r1 = p.add_run("Nota:  ")
    run_style(r1, "Arial", 10, C["blue"], bold=True)
    r2 = p.add_run(text)
    run_style(r2, "Arial", 10, C["body"])
    doc.add_paragraph()


def add_spacer(doc: Document, pts: int = 8) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(pts)
    p.paragraph_format.space_after = Pt(0)


# ── Tabela de metadados ───────────────────────────────────────────────────────

def add_metadata_table(doc: Document) -> None:
    rows_data = [
        ("Nome do Documento",       "Framework Conceitual para Modernização do Curso Técnico em Marketing"),
        ("Versão",                   "1.0"),
        ("Data",                     "Junho de 2026"),
        ("Status",                   "Documento de alinhamento institucional — nenhuma decisão curricular foi tomada"),
        ("Elaborado por",            "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários Primários",  "Direção Escolar · Orientação Educacional · Coordenação Pedagógica · Corpo Docente"),
        ("Destinatários Secundários","Futuros grupos de trabalho curricular"),
        ("Instituição",              "EEB Araújo Figueiredo — Araujo Innovation Lab — Urubici, SC"),
    ]
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, value) in enumerate(rows_data):
        row = table.rows[i]
        row.cells[0].width = Inches(2.3)
        row.cells[1].width = Inches(4.0)
        cell_shading(row.cells[0], C["blue_soft"])
        p_l = row.cells[0].paragraphs[0]
        r_l = p_l.add_run(label)
        run_style(r_l, "Arial", 10, C["blue"], bold=True)
        p_v = row.cells[1].paragraphs[0]
        r_v = p_v.add_run(value)
        run_style(r_v, "Arial", 10, C["body"])
        for cell in row.cells:
            cell_borders(cell, color=C["line"], size="6")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    doc.add_paragraph()


# ── Tabela genérica com cabeçalho azul ────────────────────────────────────────

def add_data_table(doc: Document, headers: list[str],
                   rows_data: list[list[str]],
                   col_widths: list[float] | None = None) -> None:
    col_count = len(headers)
    table = doc.add_table(rows=1 + len(rows_data), cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    header_row = table.rows[0]
    for j, h in enumerate(headers):
        cell_shading(header_row.cells[j], C["blue"])
        p = header_row.cells[j].paragraphs[0]
        r = p.add_run(h)
        run_style(r, "Arial", 10, C["white"], bold=True)
        cell_borders(header_row.cells[j], color=C["blue"], size="6")
        if col_widths:
            header_row.cells[j].width = Inches(col_widths[j])

    for i, row_vals in enumerate(rows_data):
        row = table.rows[i + 1]
        fill = C["blue_soft2"] if i % 2 == 0 else C["panel"]
        for j, val in enumerate(row_vals):
            is_first = j == 0
            cell_shading(row.cells[j], C["blue_soft"] if is_first else fill)
            p = row.cells[j].paragraphs[0]
            r = p.add_run(val)
            run_style(r, "Arial", 10, C["blue"] if is_first else C["body"],
                      bold=is_first)
            cell_borders(row.cells[j], color=C["line"], size="6")
            row.cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if col_widths:
                row.cells[j].width = Inches(col_widths[j])

    doc.add_paragraph()


# ── Capa ───────────────────────────────────────────────────────────────────────

def add_cover(doc: Document) -> None:
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    configure_section(section, "", "", with_hf=False)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.2)

    # Label institucional
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(36)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("ARAUJO INNOVATION LAB  |  EEB ARAÚJO FIGUEIREDO")
    run_style(r, "Arial", 9, C["blue"], bold=True)

    # Linha dourada
    bar = doc.add_paragraph()
    bar.paragraph_format.space_before = Pt(2)
    bar.paragraph_format.space_after = Pt(10)
    para_bottom_border(bar, color=C["gold"], size="18")

    # Classificação do documento
    tag_p = doc.add_paragraph()
    tag_p.paragraph_format.space_after = Pt(4)
    tag_run = tag_p.add_run("DOCUMENTO INSTITUCIONAL — ALINHAMENTO ESTRATÉGICO E PEDAGÓGICO")
    run_style(tag_run, "Arial", 8, C["muted"], bold=True)

    # Título principal
    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_before = Pt(28)
    title.paragraph_format.space_after = Pt(14)
    title.add_run("Framework Conceitual para Modernização do Curso Técnico em Marketing")

    # Subtítulo descritivo
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(36)
    r_sub = sub.add_run(
        "Documento preparatório para alinhamento estratégico e pedagógico entre "
        "Direção, Orientação Educacional, Coordenação Pedagógica e Corpo Docente."
    )
    run_style(r_sub, "Arial", 12, C["muted"])

    # Bloco de identificação do documento (tabela)
    id_table = doc.add_table(rows=5, cols=2)
    id_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    id_rows = [
        ("Versão",      "1.0 — Primeira Emissão Oficial"),
        ("Data",        "Junho de 2026"),
        ("Status",      "Documento de alinhamento — nenhuma decisão curricular tomada"),
        ("Elaborado por","Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários","Direção · Orientação Educacional · Coordenação · Corpo Docente"),
    ]
    for i, (label, value) in enumerate(id_rows):
        row = id_table.rows[i]
        row.cells[0].width = Inches(2.1)
        row.cells[1].width = Inches(4.2)
        cell_shading(row.cells[0], C["blue_soft"])
        p_label = row.cells[0].paragraphs[0]
        r_label = p_label.add_run(label)
        run_style(r_label, "Arial", 10, C["blue"], bold=True)
        p_value = row.cells[1].paragraphs[0]
        r_value = p_value.add_run(value)
        run_style(r_value, "Arial", 10, C["body"])
        for cell in row.cells:
            cell_borders(cell, color=C["line"], size="6")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Espaço e assinatura institucional
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(70)
    para_bottom_border(spacer, color=C["line"], size="4")

    inst = doc.add_paragraph()
    inst.paragraph_format.space_before = Pt(12)
    inst.paragraph_format.space_after = Pt(4)
    r_inst = inst.add_run("EEB Araújo Figueiredo")
    run_style(r_inst, "Georgia", 13, C["ink"], bold=True)

    city = doc.add_paragraph()
    city.paragraph_format.space_after = Pt(2)
    r_city = city.add_run("Urubici — Santa Catarina")
    run_style(r_city, "Arial", 10, C["body"])

    circ = doc.add_paragraph()
    r_circ = circ.add_run("Circulação Interna Restrita  ·  v1.0  ·  Junho de 2026")
    run_style(r_circ, "Arial", 9, C["muted"])


# ── Bloco de encerramento institucional ──────────────────────────────────────

def add_closing_block(doc: Document) -> None:
    doc.add_paragraph()
    end_table = doc.add_table(rows=1, cols=1)
    end_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    end_cell = end_table.rows[0].cells[0]
    cell_shading(end_cell, C["blue"])
    cell_borders(end_cell, color=C["blue"], size="6")

    end_p = end_cell.paragraphs[0]
    end_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p.paragraph_format.space_before = Pt(10)
    end_p.paragraph_format.space_after = Pt(4)
    r_e1 = end_p.add_run("Este documento foi elaborado como instrumento de alinhamento institucional.")
    run_style(r_e1, "Arial", 10, C["white"])

    end_p2 = end_cell.add_paragraph()
    end_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p2.paragraph_format.space_before = Pt(2)
    end_p2.paragraph_format.space_after = Pt(4)
    r_e2 = end_p2.add_run(
        "Nenhuma decisão curricular está sendo tomada ou proposta neste documento."
    )
    run_style(r_e2, "Arial", 10, C["gold"], bold=True)

    end_p3 = end_cell.add_paragraph()
    end_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p3.paragraph_format.space_before = Pt(4)
    end_p3.paragraph_format.space_after = Pt(4)
    r_e3 = end_p3.add_run(
        "O passo seguinte pertence à equipe pedagógica — construído juntos, no tempo que for necessário."
    )
    run_style(r_e3, "Arial", 10, C["white"], italic=True)

    end_p4 = end_cell.add_paragraph()
    end_p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p4.paragraph_format.space_before = Pt(8)
    end_p4.paragraph_format.space_after = Pt(4)
    r_e4 = end_p4.add_run("Elaborado por: ")
    run_style(r_e4, "Arial", 9, C["gold"], bold=True)
    r_e5 = end_p4.add_run("Profª Letícia Meurer · Professor Orientador José Alexandre")
    run_style(r_e5, "Arial", 9, C["white"])

    end_p5 = end_cell.add_paragraph()
    end_p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p5.paragraph_format.space_before = Pt(2)
    end_p5.paragraph_format.space_after = Pt(10)
    r_e6 = end_p5.add_run("Araujo Innovation Lab — EEB Araújo Figueiredo — Urubici, SC  ·  Junho de 2026")
    run_style(r_e6, "Arial", 8, "D7DEDA", italic=True)


# ── Bloco de governança documental ────────────────────────────────────────────

def add_governance_block(doc: Document) -> None:
    doc.add_page_break()

    # Título da seção de governança
    p_gov = doc.add_paragraph()
    p_gov.paragraph_format.space_before = Pt(14)
    p_gov.paragraph_format.space_after = Pt(10)
    r_gov = p_gov.add_run("GOVERNANÇA DOCUMENTAL")
    run_style(r_gov, "Georgia", 14, C["blue"], bold=True)
    para_bottom_border(p_gov, color=C["line"], size="4")

    rows_data = [
        ("Documento-fonte institucional",    "marketing-course-modernization-framework-v1.md"),
        ("Documento derivado",               "marketing-course-modernization-framework-v1.docx"),
        ("Padrão aplicável — documentos",    "document-template-standard.md"),
        ("Padrão aplicável — identidade",    "branding-standard.md"),
        ("Versão",                           "v1.0"),
        ("Instituição",                      "Colégio Araújo — EEB Araújo Figueiredo"),
        ("Local",                            "Urubici – SC"),
        ("Status",                           "Documento Institucional Publicado"),
    ]
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, value) in enumerate(rows_data):
        row = table.rows[i]
        row.cells[0].width = Inches(2.5)
        row.cells[1].width = Inches(3.8)
        cell_shading(row.cells[0], C["blue_soft"])
        p_l = row.cells[0].paragraphs[0]
        r_l = p_l.add_run(label)
        run_style(r_l, "Arial", 10, C["blue"], bold=True)
        p_v = row.cells[1].paragraphs[0]
        r_v = p_v.add_run(value)
        run_style(r_v, "Arial", 10, C["body"])
        for cell in row.cells:
            cell_borders(cell, color=C["line"], size="6")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    doc.add_paragraph()

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(8)
    sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sig = sig.add_run("Araujo Innovation Lab — EEB Araújo Figueiredo — Urubici, SC")
    run_style(r_sig, "Arial", 9, C["muted"], italic=True)


# ── Construção do documento ────────────────────────────────────────────────────

def build_document() -> None:
    doc = Document()
    configure_styles(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # CAPA
    # ═══════════════════════════════════════════════════════════════════════════
    add_cover(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO DE CORPO — nova seção com cabeçalho e rodapé
    # ═══════════════════════════════════════════════════════════════════════════
    body_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(
        body_section,
        header_text="Araujo Innovation Lab  |  Framework Conceitual para Modernização do Curso Técnico em Marketing",
        footer_right="EEB Araújo Figueiredo  |  Junho 2026",
        with_hf=True,
    )

    # ── Identificação do Documento ────────────────────────────────────────────
    p_id = doc.add_paragraph()
    p_id.paragraph_format.space_before = Pt(8)
    p_id.paragraph_format.space_after = Pt(6)
    r_id = p_id.add_run("IDENTIFICAÇÃO DO DOCUMENTO")
    run_style(r_id, "Georgia", 13, C["blue"], bold=True)
    para_bottom_border(p_id, color=C["line"], size="4")

    add_metadata_table(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 1 — APRESENTAÇÃO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "1", "Apresentação")

    add_body(doc,
        "O presente documento foi elaborado com o propósito de oferecer à comunidade educacional do "
        "Colégio Araújo um conjunto de reflexões estratégicas, pedagógicas e institucionais que "
        "contextualizam a importância de iniciar, de forma organizada e colaborativa, uma discussão "
        "sobre o futuro do Curso Técnico em Marketing."
    )
    add_body(doc,
        "Não se trata de uma proposta de mudança. Não há, neste documento, qualquer prescrição "
        "curricular, disciplina nova, alteração de carga horária ou plano de implementação. O que se "
        "apresenta aqui é um convite ao diálogo — fundamentado em evidências sobre as transformações "
        "em curso na sociedade, na economia e na profissão de marketing — para que a equipe educacional "
        "possa, juntos, avaliar em que medida esses movimentos são relevantes para a formação dos "
        "estudantes desta instituição."
    )
    add_body(doc,
        "A iniciativa de produzir este framework nasce do reconhecimento do trabalho já realizado. "
        "O Curso Técnico em Marketing da EEB Araújo Figueiredo tem uma história de dedicação, "
        "competência docente e compromisso com os estudantes. Qualquer reflexão sobre o futuro "
        "precisa partir desse reconhecimento — e não de qualquer forma de crítica ao que existe."
    )
    add_body(doc,
        "O momento presente, no entanto, é singular. O mercado de trabalho, a tecnologia, o perfil "
        "dos estudantes e as metodologias disponíveis para a educação profissional estão se "
        "transformando com uma velocidade sem precedente. Ignorar essas transformações seria uma "
        "escolha tão arriscada quanto adotá-las precipitadamente, sem o devido processo institucional "
        "de análise e planejamento."
    )
    add_body(doc,
        "Este documento pretende apoiar exatamente esse processo: fornecer um referencial conceitual "
        "que permita à equipe educacional discutir, com qualidade e profundidade, quais elementos do "
        "cenário atual merecem atenção, quais princípios devem orientar decisões futuras e qual seria "
        "um caminho razoável e respeitoso para uma eventual fase de planejamento colaborativo."
    )
    add_highlight_block(doc, "Compromisso deste documento",
        "Ao concluir a leitura deste material, espera-se que gestores, orientadores e professores "
        "tenham clareza sobre por que esta conversa importa, sem que nenhuma decisão tenha sido "
        "tomada por eles ou para eles. A decisão pertence à instituição — construída coletivamente, "
        "no tempo que for necessário.",
        fill=C["blue_soft2"], border_color=C["blue"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 2 — CENÁRIO ATUAL DA FORMAÇÃO PROFISSIONAL
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "2", "Cenário Atual da Formação Profissional")

    add_body(doc,
        "A área de marketing atravessa uma das transformações mais profundas de sua história recente. "
        "O que era um conjunto relativamente estável de práticas e saberes — publicidade, pesquisa de "
        "mercado, gestão de produto — tornou-se um campo em contínua reconfiguração, atravessado por "
        "forças econômicas, tecnológicas e culturais que se retroalimentam e se aceleram mutuamente."
    )

    add_subsection(doc, "A Economia Digital como Novo Contexto de Trabalho")
    add_body(doc,
        "A digitalização da economia não é mais uma tendência futura: é a realidade operacional de "
        "negócios de todos os tamanhos, em todos os setores, inclusive em cidades de pequeno e médio "
        "porte como Urubici. Pequenos comércios, prestadores de serviços, pousadas e empreendedores "
        "rurais já operam, em graus variados, em ambientes digitais — gerenciando redes sociais, "
        "respondendo a avaliações online, recebendo pedidos por aplicativos e construindo reputação "
        "em plataformas globais."
    )
    add_body(doc,
        "Esse deslocamento do ponto de contato entre negócios e consumidores para o ambiente digital "
        "não eliminou a necessidade de competências tradicionais de marketing — análise de público, "
        "posicionamento de marca, comunicação persuasiva. Pelo contrário, ele as tornou mais "
        "exigentes, porque agora precisam operar simultaneamente em múltiplos canais, com velocidade "
        "maior e com métricas de resultado em tempo real."
    )

    add_subsection(doc, "O Comportamento do Consumidor em Transformação")
    add_body(doc,
        "O consumidor contemporâneo chegou ao mercado dotado de recursos sem precedentes: acesso "
        "imediato à informação, comparação instantânea entre fornecedores, capacidade de amplificar "
        "experiências positivas ou negativas por meio das redes sociais, e um nível de exigência "
        "construído por anos de exposição a interfaces digitais de alta qualidade."
    )
    add_body(doc,
        "Esse novo consumidor não apenas compra — ele avalia, compartilha, influencia e, cada vez "
        "mais, cria. A linha entre consumidor e produtor de conteúdo tornou-se fluida. Para o "
        "profissional de marketing, isso significa que as competências necessárias incluem não apenas "
        "produzir mensagens, mas interpretar comportamentos, construir relacionamentos e gerenciar "
        "a percepção da marca ao longo de uma jornada complexa e multicanal."
    )

    add_subsection(doc, "A Comunicação como Campo Expandido")
    add_body(doc,
        "O campo da comunicação profissional expandiu-se enormemente. Copywriting, design visual, "
        "produção audiovisual, storytelling, comunicação por áudio e criação de experiências "
        "interativas — competências que antes eram domínio de especialistas separados — tornaram-se "
        "progressivamente parte do repertório esperado de um profissional de marketing. Ferramentas "
        "acessíveis democratizaram a produção; ao mesmo tempo, elevaram o nível de qualidade "
        "percebida pelo mercado."
    )

    add_subsection(doc, "Empreendedorismo como Competência Transversal")
    add_body(doc,
        "O empreendedorismo deixou de ser um traço de personalidade para tornar-se uma competência "
        "educável. Em um mercado de trabalho cada vez mais marcado por relações não-lineares — "
        "projetos, contratos, parcerias, trabalho autônomo — o profissional que compreende a lógica "
        "do negócio, que sabe criar valor, propor soluções e comunicar resultados tem vantagem "
        "significativa, independentemente de trabalhar em uma grande empresa, em um pequeno negócio "
        "local ou para si mesmo."
    )
    add_body(doc,
        "Para o Curso Técnico em Marketing, esse cenário representa uma oportunidade de posicionar "
        "a formação técnica não apenas como preparação para o emprego formal, mas como desenvolvimento "
        "de uma mentalidade produtiva e autônoma que serve ao estudante em qualquer trajetória "
        "profissional."
    )

    add_subsection(doc, "Inteligência Artificial como Variável Estrutural")
    add_body(doc,
        "A inteligência artificial generativa incorporou-se ao cotidiano profissional de marketing "
        "com uma velocidade que surpreendeu especialistas. Ferramentas de geração de texto, imagem, "
        "vídeo e análise de dados estão acessíveis e são amplamente utilizadas por profissionais e "
        "pequenas empresas. Isso não elimina a necessidade de competência humana — ao contrário, "
        "eleva a importância do julgamento crítico, da curadoria de qualidade, da ética na "
        "comunicação e da compreensão estratégica que nenhuma ferramenta pode substituir."
    )
    add_highlight_block(doc, "Síntese do cenário",
        "Esses elementos — não isoladamente, mas em conjunto — configuram um cenário que merece "
        "atenção institucional. Não como urgência, mas como responsabilidade formativa.",
        fill=C["blue_soft2"], border_color=C["line"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 3 — DESAFIOS CONTEMPORÂNEOS DA EDUCAÇÃO TÉCNICA
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "3", "Desafios Contemporâneos da Educação Técnica")

    add_body(doc,
        "A educação profissional e tecnológica enfrenta, neste momento histórico, desafios que "
        "transcendem as fronteiras de qualquer disciplina ou curso específico. Compreender esses "
        "desafios é essencial para que as discussões sobre o futuro do Curso Técnico em Marketing "
        "se apoiem em uma visão realista e qualificada do contexto educacional."
    )

    add_subsection(doc, "A Aceleração do Ciclo de Relevância das Competências")
    add_body(doc,
        "Historicamente, a educação profissional técnica operava com a premissa de que um conjunto "
        "de competências ensinadas em sala permaneceria relevante por um longo período — talvez por "
        "toda a carreira do estudante. Esse pressuposto está sendo desafiado. O ciclo de vida das "
        "ferramentas digitais, das práticas de comunicação e das plataformas de marketing "
        "reduziu-se significativamente. Uma plataforma amplamente utilizada pode tornar-se obsoleta "
        "em poucos anos; uma competência técnica específica pode ser automatizada antes que o "
        "estudante complete sua formação."
    )
    add_body(doc,
        "Isso não significa que as instituições de ensino devam perseguir modismos tecnológicos a "
        "cada semestre. Significa, ao contrário, que a formação precisa desenvolver nos estudantes "
        "a capacidade de aprender continuamente — tornando o processo de atualização uma habilidade "
        "em si, e não apenas um pressuposto implícito."
    )

    add_subsection(doc, "O Perfil do Estudante em Transformação")
    add_body(doc,
        "Os estudantes que chegam ao Curso Técnico em Marketing trazem consigo experiências digitais "
        "ricas, mas frequentemente não estruturadas. Eles consomem conteúdo em múltiplas plataformas, "
        "interagem com criadores digitais, conhecem as interfaces de redes sociais com naturalidade "
        "e têm acesso a informações que antes eram restritas a especialistas. Ao mesmo tempo, muitos "
        "deles carecem de estrutura para transformar esse repertório informal em competência "
        "profissional reconhecida."
    )

    add_subsection(doc, "A Dimensão Prática como Exigência Crescente")
    add_body(doc,
        "O mercado de trabalho valoriza cada vez mais a capacidade de fazer, não apenas de saber. "
        "Portfólios, projetos concluídos, experiências práticas documentadas e evidências concretas "
        "de competência tendem a pesar mais do que certificados isolados. Para a escola técnica, "
        "isso representa uma oportunidade e um desafio: como criar, dentro dos limites do ambiente "
        "educacional, experiências de aprendizagem que sejam suficientemente próximas da realidade "
        "profissional para preparar os estudantes de forma efetiva?"
    )

    add_subsection(doc, "A Expectativa de Aprendizagem Ativa")
    add_body(doc,
        "A geração atual de estudantes cresceu em ambientes interativos. A relação passiva com o "
        "conhecimento — onde um especialista transmite e o estudante recebe — não ressoa com a mesma "
        "eficácia que experiências em que o estudante participa, cria, resolve e apresenta. Isso não "
        "invalida a aula expositiva ou a transmissão de conteúdo conceitual — que continuam "
        "fundamentais. Mas sugere que o equilíbrio entre instrução direta e aprendizagem ativa é "
        "um elemento pedagógico que merece reflexão."
    )

    add_subsection(doc, "A Tecnologia como Aliada da Prática Docente")
    add_body(doc,
        "Para os professores, a tecnologia representa simultaneamente uma demanda e uma oportunidade. "
        "A demanda é de atualização contínua em ferramentas e plataformas que evoluem rapidamente. "
        "A oportunidade é de ampliar o repertório metodológico, diversificar as estratégias de "
        "avaliação e acompanhar a trajetória dos estudantes com mais profundidade."
    )
    add_highlight_block(doc, "Perspectiva institucional",
        "Nenhum desses desafios tem solução simples ou imediata. Mas reconhecê-los coletivamente "
        "é o primeiro passo para construir respostas que façam sentido para a realidade específica "
        "desta instituição.",
        fill=C["green_soft"], border_color=C["line"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 4 — PERFIL DO ESTUDANTE DAS NOVAS GERAÇÕES
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "4", "Perfil do Estudante das Novas Gerações")

    add_body(doc,
        "Compreender quem são os estudantes que chegam ao Curso Técnico em Marketing é essencial "
        "para pensar com pertinência qualquer discussão sobre a formação que esta instituição deseja "
        "oferecer. Os jovens que compõem o corpo discente atual cresceram em um contexto radicalmente "
        "diferente daquele que formou as gerações anteriores — e essa diferença tem implicações "
        "educacionais concretas."
    )

    add_subsection(doc, "Nativos Digitais com Repertório Informal Rico")
    add_body(doc,
        "Os estudantes das novas gerações são, na expressão consagrada, nativos digitais. Mas mais "
        "do que apenas navegar em plataformas, eles as habitam. Redes sociais, aplicativos de vídeo "
        "curto, mensagens instantâneas e streams de conteúdo fazem parte do tecido de sua experiência "
        "cotidiana desde a infância. Muitos deles já criaram conteúdo, gerenciaram perfis pessoais, "
        "editaram vídeos e construíram narrativas digitais muito antes de ingressar no ensino técnico."
    )
    add_body(doc,
        "Esse repertório é um ponto de partida valioso para a educação profissional em marketing. "
        "A escola não precisa introduzir esses estudantes ao mundo digital — precisa ajudá-los a "
        "compreendê-lo com profundidade, ética e intenção estratégica."
    )

    add_subsection(doc, "Consumo de Conteúdo e Aprendizagem Informal")
    add_body(doc,
        "Esses estudantes aprendem frequentemente fora da sala de aula — assistindo a tutoriais, "
        "acompanhando criadores de conteúdo especializados, participando de comunidades online. "
        "A aprendizagem informal já é parte estrutural de sua experiência. Isso cria uma oportunidade: "
        "a escola pode conectar o que o estudante já aprende informalmente com estrutura conceitual, "
        "crítica e profissional que o distinguirá no mercado."
    )

    add_subsection(doc, "Preferências de Aprendizagem e Ambientes Colaborativos")
    add_body(doc,
        "Pesquisas sobre aprendizagem das novas gerações indicam consistentemente que esses estudantes "
        "respondem melhor a ambientes que oferecem: sentido prático claro para o que estão aprendendo, "
        "possibilidade de interação e colaboração, variação de formatos e estímulos, e reconhecimento "
        "de seus saberes anteriores. Não se trata de uma exigência de entretenimento, mas de uma "
        "disposição para o engajamento que a escola pode aproveitar pedagogicamente."
    )
    add_body(doc,
        "Projetos coletivos, desafios reais ou realistas, apresentações para públicos além da sala "
        "de aula e a construção de algo concreto e visível ao longo do curso são elementos que tendem "
        "a estimular o comprometimento desses estudantes com sua própria formação."
    )

    add_subsection(doc, "Interação com Tecnologia: Familiaridade e Necessidade de Mediação")
    add_body(doc,
        "A familiaridade dos estudantes com tecnologia não significa, necessariamente, que eles a "
        "utilizam com eficiência profissional. Saber operar uma rede social como usuário é muito "
        "diferente de saber gerenciá-la como profissional. Saber usar ferramentas de inteligência "
        "artificial para fins pessoais é muito diferente de compreender seus princípios de "
        "funcionamento, suas limitações e suas implicações éticas."
    )

    add_subsection(doc, "Uma Geração com Potencial para o Marketing")
    add_body(doc,
        "Existe uma convergência natural entre o perfil das novas gerações e as competências que a "
        "área de marketing demanda. Criatividade, capacidade de comunicação visual e textual, "
        "sensibilidade para tendências e comportamentos, habilidade de produzir conteúdo e de "
        "construir narrativas — essas são características que muitos desses jovens já cultivam, "
        "mesmo sem perceber. A escola técnica tem a oportunidade única de dar nome, estrutura e "
        "profissionalismo a esse potencial."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 5 — TENDÊNCIAS DA EDUCAÇÃO PROFISSIONAL E TECNOLÓGICA
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "5", "Tendências da Educação Profissional e Tecnológica")

    add_body(doc,
        "A educação profissional brasileira vive um momento de reflexão e renovação metodológica. "
        "Em diferentes contextos e instituições, estão sendo exploradas e consolidadas abordagens "
        "pedagógicas que buscam aproximar a formação técnica da realidade do trabalho, sem abrir "
        "mão do rigor conceitual e da profundidade formativa que caracterizam a boa educação."
    )
    add_body(doc,
        "Conhecer essas tendências não significa adotá-las automaticamente. Significa ter um "
        "referencial informado para discutir, com qualidade, quais delas são pertinentes para o "
        "contexto específico do Curso Técnico em Marketing da EEB Araújo Figueiredo."
    )

    add_subsection(doc, "Aprendizagem Baseada em Projetos")
    add_body(doc,
        "A aprendizagem baseada em projetos organiza a experiência educacional em torno de desafios "
        "com começo, meio e fim — um problema para resolver, um produto para criar, uma solução para "
        "apresentar. Em vez de estudar conceitos em abstrato, o estudante os mobiliza para alcançar "
        "um resultado concreto. Essa abordagem não substitui a aula expositiva ou o ensino "
        "estruturado; ela oferece um contexto de aplicação que dá sentido ao conteúdo aprendido."
    )

    add_subsection(doc, "Aprendizagem Baseada em Desafios")
    add_body(doc,
        "Próxima da aprendizagem por projetos, a aprendizagem baseada em desafios parte de problemas "
        "reais — da comunidade, de empresas locais, de situações que os próprios estudantes "
        "identificam — e organiza o processo formativo em torno da busca por soluções. Ela estimula "
        "o pensamento crítico, a pesquisa, a colaboração e a comunicação de resultados."
    )
    add_body(doc,
        "Para uma cidade como Urubici, com um ecossistema econômico rico e diversificado, os desafios "
        "reais disponíveis para esse tipo de aprendizagem são numerosos: como um pequeno negócio pode "
        "ampliar sua presença digital? Como uma pousada pode melhorar sua reputação online? Como uma "
        "feira gastronômica pode se comunicar melhor com visitantes? Essas perguntas têm resposta "
        "e podem ser trabalhadas pedagogicamente."
    )

    add_subsection(doc, "Abordagens Interdisciplinares")
    add_body(doc,
        "A interdisciplinaridade reconhece que as fronteiras entre disciplinas são, muitas vezes, "
        "mais administrativas do que intelectuais. O marketing dialoga profundamente com língua "
        "portuguesa, com arte e design, com matemática e estatística, com sociologia, com tecnologia, "
        "com ética. Uma formação que reconhece e cultiva essas conexões prepara o estudante para "
        "a complexidade do mundo real, onde problemas raramente pertencem a uma única área do "
        "conhecimento."
    )

    add_subsection(doc, "Resolução de Problemas do Mundo Real")
    add_body(doc,
        "Quando o estudante percebe que o que aprende tem aplicação imediata e relevância concreta, "
        "o engajamento tende a ser significativamente maior. Problemas do mundo real — especialmente "
        "quando envolvem a comunidade local, pessoas reais e consequências observáveis — criam um "
        "ambiente de aprendizagem que nenhuma simulação perfeita consegue reproduzir completamente."
    )

    add_subsection(doc, "Portfólio como Instrumento de Desenvolvimento")
    add_body(doc,
        "O portfólio é mais do que uma pasta de trabalhos. É um registro da trajetória do estudante "
        "— evidência da evolução, da diversidade de competências desenvolvidas, da capacidade de "
        "reflexão sobre o próprio aprendizado. Em áreas criativas e comunicacionais como o marketing, "
        "o portfólio tem valor especial: ele demonstra, de forma concreta e visual, o que o estudante "
        "é capaz de fazer."
    )

    add_subsection(doc, "Experiências Práticas como Componente Formativo")
    add_body(doc,
        "Visitas técnicas a empresas, observações de campo, contato com profissionais da área, "
        "imersões em contextos reais de trabalho — essas experiências têm valor educacional que vai "
        "além do que qualquer aula pode oferecer. Elas humanizam a formação, contextualizam o "
        "aprendizado e preparam o estudante para as dimensões interpessoais e situacionais do trabalho "
        "que nenhuma teoria consegue antecipar completamente."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 6 — INTELIGÊNCIA ARTIFICIAL COMO FERRAMENTA EDUCACIONAL
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "6", "Inteligência Artificial como Ferramenta Educacional")

    add_body(doc,
        "A inteligência artificial entrou definitivamente no universo educacional — e na prática "
        "profissional de marketing — com uma presença que não pode ser ignorada nem adotada de "
        "forma acrítica. Este é talvez o tema que mais exige equilíbrio na reflexão institucional: "
        "entre a realidade do uso já disseminado e a responsabilidade pedagógica de mediar esse uso "
        "com critério, ética e discernimento."
    )

    add_subsection(doc, "A Centralidade Permanente do Educador")
    add_body(doc,
        "O ponto de partida desta reflexão deve ser inequívoco: nenhuma ferramenta de inteligência "
        "artificial substitui o educador. A mediação pedagógica, a relação de confiança entre "
        "professor e estudante, a capacidade de perceber dificuldades individuais, de adaptar a "
        "abordagem, de transmitir valores e de construir o ambiente de sala de aula são dimensões "
        "humanas, insubstituíveis e essenciais à educação de qualidade."
    )
    add_highlight_block(doc, "Princípio fundamental",
        "A inteligência artificial pode ampliar certas capacidades, mas o faz como instrumento "
        "sob orientação humana, não como agente autônomo de formação. O educador permanece central "
        "em todo processo de aprendizagem.",
        fill=C["blue_soft2"], border_color=C["blue"])

    add_subsection(doc, "IA como Apoio ao Processo de Aprendizagem")
    add_body(doc,
        "Quando utilizada de forma orientada e pedagogicamente intencionada, a inteligência artificial "
        "pode enriquecer experiências de aprendizagem de diversas maneiras. Ela pode ajudar estudantes "
        "a explorar perspectivas múltiplas sobre um tema, gerar variações de um texto para análise "
        "comparativa, simular situações de atendimento ao cliente, ou oferecer feedback instantâneo "
        "em tarefas específicas."
    )

    add_subsection(doc, "Uso Responsável e Pensamento Crítico")
    add_body(doc,
        "O uso responsável da inteligência artificial é uma competência que precisa ser desenvolvida, "
        "não presumida. Estudantes precisam aprender a avaliar criticamente os resultados gerados "
        "por ferramentas de IA, reconhecendo limitações, imprecisões e vieses possíveis. Também "
        "precisam distinguir entre o que é genuinamente seu — seu pensamento, sua criatividade, "
        "seu julgamento — e o que foi gerado automaticamente."
    )
    add_bullet(doc, "Avaliar criticamente os resultados gerados por ferramentas de IA;")
    add_bullet(doc, "Distinguir entre pensamento próprio e conteúdo gerado automaticamente;")
    add_bullet(doc, "Citar e referenciar adequadamente quando material gerado com IA é entregue formalmente;")
    add_bullet(doc, "Compreender que a qualidade de uma entrega com IA depende do repertório e julgamento humano que a orienta.")

    add_subsection(doc, "Dimensões Éticas do Uso de IA")
    add_body(doc,
        "A ética do uso de inteligência artificial em comunicação e marketing merece atenção "
        "específica. O profissional de marketing que utiliza IA para produzir conteúdo tem "
        "responsabilidade sobre o que é publicado — não a ferramenta. Isso inclui verificar a "
        "veracidade de informações, evitar a reprodução de estereótipos, não criar conteúdo "
        "enganoso e respeitar os direitos de privacidade e de imagem."
    )

    add_subsection(doc, "Preparação Profissional para um Mercado que Já Usa IA")
    add_body(doc,
        "O mercado de trabalho em marketing já utiliza inteligência artificial em diversas frentes: "
        "geração de conteúdo, segmentação de públicos, análise de dados, automação de comunicação, "
        "criação de imagens e personalização de campanhas. O estudante que conclui o curso técnico "
        "sem qualquer familiaridade com esse universo chega ao mercado em desvantagem que não é "
        "sua culpa, mas que pode ser evitada."
    )
    add_highlight_block(doc, "Posicionamento institucional",
        "Inserir o tema da inteligência artificial na formação técnica em marketing não é uma "
        "concessão à modernidade tecnológica — é uma responsabilidade formativa com os estudantes "
        "que a escola acolhe e com o futuro profissional que eles merecem.",
        fill=C["green_soft"], border_color=C["line"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 7 — CONEXÃO COM A REALIDADE LOCAL DE URUBICI
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "7", "Conexão com a Realidade Local de Urubici")

    add_body(doc,
        "Um dos maiores patrimônios pedagógicos disponíveis para o Curso Técnico em Marketing da "
        "EEB Araújo Figueiredo é, paradoxalmente, o território em que a escola está inserida. "
        "Urubici — com sua paisagem única, seu turismo em crescimento, sua gastronomia reconhecida, "
        "seu artesanato, seu agronegócio e sua comunidade empreendedora — oferece um campo de "
        "aprendizagem contextualizada de valor inestimável."
    )
    add_body(doc,
        "A educação em marketing que parte do que é local, concreto e próximo tem uma qualidade "
        "que nenhum caso importado de São Paulo ou Nova York consegue reproduzir: ela é real para "
        "o estudante. O negócio que ele vai analisar, a marca que vai ajudar a comunicar, o "
        "consumidor que vai aprender a compreender — tudo isso faz parte do mundo em que ele vive."
    )

    add_subsection(doc, "Turismo e Hospitalidade como Campo de Aprendizagem")
    add_body(doc,
        "Urubici é um destino turístico de relevância crescente na Serra Catarinense. Pousadas, "
        "hotéis, acomodações de imersão na natureza e experiências de turismo rural compõem um "
        "setor que depende, cada vez mais, de comunicação digital eficaz para atrair visitantes "
        "de outras cidades e estados. A gestão da reputação online, a produção de conteúdo visual "
        "que capture a beleza do local e a construção de presença em plataformas de reserva são "
        "desafios reais que qualquer estudante de marketing poderia observar, analisar e — em "
        "momentos apropriados de sua formação — apoiar."
    )

    add_subsection(doc, "Gastronomia e Culinária Regional")
    add_body(doc,
        "A gastronomia de Urubici — marcada por produtos da Serra, pela culinária típica de "
        "descendentes de imigrantes europeus e por restaurantes que valorizam ingredientes locais "
        "— representa um campo rico para comunicação e marketing. Como uma cozinha local conta a "
        "história de suas escolhas? Como um restaurante constrói identidade de marca que vá além "
        "do cardápio? Como a fotografia de alimentos e a narrativa gastronômica podem ser trabalhadas "
        "como competências de marketing? Essas são questões com resposta pedagógica concreta."
    )

    add_subsection(doc, "Comércio Local e Pequenos Negócios")
    add_body(doc,
        "O comércio local de Urubici — lojas, farmácias, prestadores de serviços, profissionais "
        "autônomos — enfrenta os mesmos desafios que qualquer pequena empresa em um contexto de "
        "transformação digital: como ser encontrado por clientes online, como construir credibilidade, "
        "como responder a avaliações, como comunicar promoções e eventos. Esses desafios são reais, "
        "próximos e suficientemente concretos para servirem como contexto de aprendizagem."
    )

    add_subsection(doc, "Agronegócio e Produtos Regionais")
    add_body(doc,
        "A Serra Catarinense tem uma tradição agrícola significativa: maçã, pera, alho, mel, produtos "
        "coloniais e outros itens produzidos na região chegam a mercados locais, estaduais e "
        "nacionais. A comunicação dos valores desses produtos — origem, método de produção, qualidade, "
        "conexão com o território — é um desafio de marketing que o estudante formado nesta região "
        "está em posição privilegiada para compreender e trabalhar."
    )

    add_subsection(doc, "Eventos, Festivais e Manifestações Culturais")
    add_body(doc,
        "A vida cultural e comunitária de Urubici — com seus eventos sazonais, festivais, feiras "
        "e manifestações culturais — oferece oportunidades de aprendizagem sobre comunicação de "
        "eventos, produção de conteúdo ao vivo, cobertura digital e mobilização de comunidades. "
        "Esses são contextos de alta relevância pedagógica que, quando trabalhados intencionalmente, "
        "conectam a formação técnica com a identidade local."
    )

    add_subsection(doc, "Empreendedorismo Regional Emergente")
    add_body(doc,
        "Urubici é também uma cidade de empreendedores. Novos negócios surgem, iniciativas de "
        "turismo de base comunitária ganham força, cooperativas e associações se organizam. Esse "
        "ecossistema de empreendedorismo regional é um campo fértil para que estudantes de marketing "
        "aprendam a identificar oportunidades, a comunicar propostas de valor e a apoiar o "
        "desenvolvimento de marcas nascentes."
    )
    add_highlight_block(doc, "Valor pedagógico do território",
        "A escola que conecta sua formação ao território não apenas enriquece o currículo — ela "
        "também constrói seu papel social na comunidade e prepara profissionais que conhecem e "
        "valorizam o lugar onde vivem.",
        fill=C["gold_soft"], border_color=C["gold"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 8 — VISÃO PROPOSTA PARA O FUTURO DO CURSO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "8", "Visão Proposta para o Futuro do Curso")

    add_body(doc,
        "Uma reflexão institucional sobre o futuro do Curso Técnico em Marketing precisa ser orientada "
        "por uma visão compartilhada das competências que se deseja que os estudantes desenvolvam ao "
        "longo de sua formação. Essa visão não é um currículo — é um horizonte. É a imagem do tipo "
        "de profissional e de cidadão que a escola aspira a contribuir para formar."
    )
    add_body(doc,
        "A proposta a seguir não é definitiva. É um ponto de partida para a conversa — um conjunto "
        "de competências amplas que refletem tanto as demandas contemporâneas da área de marketing "
        "quanto os valores educativos que toda escola de qualidade deve cultivar."
    )

    add_subsection(doc, "Comunicação como Competência Fundamental")
    add_body(doc,
        "O profissional de marketing do futuro precisa comunicar-se com clareza, precisão e adequação "
        "a diferentes públicos e contextos. Isso inclui comunicação escrita — em textos de divulgação, "
        "relatórios, e-mails, roteiros — mas também comunicação oral, visual e digital. A capacidade "
        "de adaptar a mensagem ao canal, ao receptor e ao objetivo é um dos diferenciais mais "
        "valorizados pelo mercado."
    )

    add_subsection(doc, "Criatividade com Propósito")
    add_body(doc,
        "Criatividade no marketing não é apenas originalidade estética — é a capacidade de resolver "
        "problemas com soluções inesperadas, de ver conexões onde outros não veem, de propor "
        "abordagens que combinem relevância, pertinência e impacto. O estudante que desenvolve "
        "criatividade com propósito — criatividade a serviço de objetivos claros — está preparado "
        "para uma das competências mais difíceis de ser automatizada."
    )

    add_subsection(doc, "Resolução de Problemas e Pensamento Estratégico")
    add_body(doc,
        "A capacidade de identificar um problema, compreender seu contexto, mapear possibilidades "
        "e propor soluções informadas é uma das competências mais universais e duradouras que a "
        "educação pode desenvolver. No marketing, ela aparece continuamente: como aumentar o "
        "engajamento? Como recuperar uma reputação afetada? Como comunicar um produto para um "
        "público que ainda não o conhece?"
    )

    add_subsection(doc, "Trabalho em Equipe e Colaboração")
    add_body(doc,
        "O marketing é, por natureza, uma área colaborativa. Campanhas, projetos e estratégias "
        "raramente são obra de uma única pessoa. A capacidade de trabalhar em equipe — de contribuir, "
        "de ouvir, de mediar conflitos, de reconhecer habilidades alheias e de construir coletivamente "
        "— é uma competência que a escola pode cultivar intencionalmente em todas as etapas da "
        "formação."
    )

    add_subsection(doc, "Letramento Digital e Fluência Tecnológica")
    add_body(doc,
        "Compreender como as ferramentas digitais funcionam — não apenas operá-las, mas entender "
        "sua lógica, suas possibilidades e seus limites — é uma competência que o mercado espera "
        "de qualquer profissional de marketing. Isso inclui plataformas de redes sociais, ferramentas "
        "de análise de dados, ambientes de produção de conteúdo e, cada vez mais, ferramentas de "
        "inteligência artificial."
    )

    add_subsection(doc, "Uso Ético da Tecnologia")
    add_body(doc,
        "A dimensão ética do uso da tecnologia é indissociável da competência técnica. Profissionais "
        "de marketing têm poder de influência sobre comportamentos, percepções e decisões de consumo. "
        "Esse poder exige responsabilidade — para com os consumidores, para com a verdade, para com "
        "a privacidade e para com os valores que a comunicação veicula. A formação que negligencia "
        "essa dimensão forma técnicos, não profissionais."
    )

    add_subsection(doc, "Mentalidade Empreendedora")
    add_body(doc,
        "A mentalidade empreendedora — a disposição para identificar oportunidades, assumir "
        "responsabilidade pelo próprio aprendizado e pela própria trajetória, propor soluções e "
        "não apenas seguir instruções — é uma competência transversal de alto valor. Ela serve ao "
        "estudante que empreende, ao que se emprega, e ao cidadão que participa ativamente da "
        "comunidade onde vive."
    )
    add_highlight_block(doc, "Nota de escopo",
        "Essas competências não precisam de uma nova disciplina para existir. Elas podem ser "
        "cultivadas em qualquer ambiente de aprendizagem bem projetado — o que, por si só, é um "
        "convite para a reflexão sobre as práticas pedagógicas atuais.",
        fill=C["blue_soft2"], border_color=C["line"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 9 — PRINCÍPIOS PEDAGÓGICOS NORTEADORES
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "9", "Princípios Pedagógicos Norteadores")

    add_body(doc,
        "Os princípios pedagógicos que seguem não são regras a serem impostas. São valores educativos "
        "que podem orientar as discussões e decisões da equipe docente ao pensar sobre o futuro do "
        "Curso Técnico em Marketing. São propostas de bússola, não de mapa."
    )

    principles = [
        ("Aprender Fazendo",
         "O conhecimento que é apenas ouvido ou lido tende a permanecer superficial. O conhecimento "
         "que é praticado, aplicado, testado e revisado consolida-se de forma duradoura. Este princípio "
         "— consagrado na pedagogia e reafirmado por décadas de pesquisa educacional — sugere que a "
         "formação técnica tem todo a ganhar quando equilibra instrução e experiência, teoria e prática, "
         "conceito e aplicação."),
        ("Aprendizagem Significativa",
         "O estudante aprende melhor quando percebe sentido no que estuda — quando compreende para que "
         "serve, como se conecta com o que já sabe e como pode ser utilizado em contextos que importam "
         "para ele. A aprendizagem significativa não exige que todo conteúdo seja imediatamente prático; "
         "exige que haja conexão inteligível entre o que é ensinado e uma estrutura de compreensão que "
         "faça sentido para o aprendiz."),
        ("Colaboração como Método e Valor",
         "A colaboração é, ao mesmo tempo, uma estratégia pedagógica eficaz e um valor que a educação "
         "deve cultivar. Aprender com os colegas, construir em grupo, dar e receber feedback, compartilhar "
         "o mérito das conquistas e a responsabilidade pelos erros — essas são experiências formativas "
         "que preparam o estudante tanto para o mercado de trabalho quanto para a vida em comunidade."),
        ("Inovação com Responsabilidade",
         "A abertura para o novo — para novas ideias, novas abordagens, novas ferramentas — é uma "
         "disposição que a educação técnica contemporânea deve cultivar. Mas a inovação sem "
         "responsabilidade pode ser imprudente, e a responsabilidade sem abertura para o novo pode ser "
         "estagnante. O equilíbrio entre experimentar com critério e avançar com cuidado é um princípio "
         "pedagógico que honra tanto a criatividade quanto a integridade."),
        ("Ética como Princípio Estruturante",
         "A ética não é uma disciplina isolada na grade curricular — é um princípio que deve atravessar "
         "todas as experiências de aprendizagem. Na comunicação, a ética diz respeito à veracidade das "
         "mensagens, ao respeito pelo consumidor, à responsabilidade sobre os efeitos do que se publica. "
         "No uso de tecnologia, diz respeito à privacidade, ao consentimento e à transparência."),
        ("Conexão com a Comunidade",
         "A escola que reconhece o território como recurso pedagógico amplia seu próprio valor educativo "
         "e social. Conectar a formação com os desafios, as oportunidades e as pessoas da comunidade "
         "local não apenas enriquece o aprendizado — fortalece o vínculo entre a escola e o lugar que "
         "ela serve."),
        ("Protagonismo Estudantil Estruturado",
         "O estudante aprende mais quando é agente de sua própria formação — quando propõe, decide, "
         "executa e avalia. Mas o protagonismo precisa de estrutura para florescer: objetivos claros, "
         "suporte docente consistente, limites bem definidos e um ambiente seguro para experimentar. "
         "O protagonismo sem estrutura é abandono; a estrutura sem protagonismo é passividade."),
        ("Continuidade e Consistência",
         "Competências sólidas não se desenvolvem em um semestre. Elas se constroem ao longo de "
         "experiências repetidas, reflexão acumulada e aplicação progressivamente mais complexa. "
         "Um currículo que valoriza a continuidade — que conecta o que foi aprendido com o que está "
         "sendo aprendido e com o que virá — prepara estudantes com mais profundidade e durabilidade."),
    ]

    for principle_name, principle_text in principles:
        add_body_bold_intro(doc, f"{principle_name}:  ", principle_text)
        add_spacer(doc, 4)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 10 — BENEFÍCIOS ESPERADOS
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "10", "Benefícios Esperados")

    add_body(doc,
        "Uma discussão institucional qualificada sobre o futuro do Curso Técnico em Marketing, "
        "conduzida de forma colaborativa e responsável, tem o potencial de gerar benefícios concretos "
        "para diferentes grupos que compõem a comunidade educacional. É importante nomear esses "
        "benefícios não como promessas, mas como horizontes que fundamentam o esforço coletivo de "
        "diagnóstico e planejamento."
    )

    add_subsection(doc, "10.1 Para os Estudantes")
    add_body(doc,
        "Os estudantes são, em última instância, os principais beneficiários de qualquer processo "
        "de melhoria da qualidade formativa. Uma formação técnica em marketing que se mantém "
        "atualizada, contextualizada e pedagogicamente consistente oferece a esses jovens:"
    )
    add_bullet(doc, "Preparação mais efetiva para o mercado de trabalho, com competências reconhecidas e evidências concretas de aprendizagem;")
    add_bullet(doc, "Maior sentido de propósito durante a formação, ao perceberem conexão entre o que aprendem e sua aplicação real no contexto local e nacional;")
    add_bullet(doc, "Desenvolvimento de um portfólio profissional que pode acompanhá-los em processos seletivos, cursos superiores ou em iniciativas empreendedoras;")
    add_bullet(doc, "Familiaridade ética e crítica com as ferramentas digitais que definirão o mercado de trabalho da próxima década;")
    add_bullet(doc, "Vínculos com o território — compreensão e valorização das oportunidades econômicas e culturais de Urubici e região;")
    add_bullet(doc, "Habilidades transversais — comunicação, colaboração, resolução de problemas, autogestão — que têm valor em qualquer área de atuação.")

    add_subsection(doc, "10.2 Para os Professores")
    add_body(doc,
        "O processo de reflexão e planejamento colaborativo que este documento propõe não é uma "
        "sobrecarga para os professores — é uma oportunidade de reconhecimento e de desenvolvimento "
        "profissional. Entre os benefícios esperados para o corpo docente:"
    )
    add_bullet(doc, "Reconhecimento formal do trabalho já realizado, por meio de um processo de diagnóstico que valoriza as práticas existentes antes de propor qualquer mudança;")
    add_bullet(doc, "Acesso a referenciais atualizados sobre metodologias, ferramentas e tendências da educação profissional, de forma contextualizada e prática;")
    add_bullet(doc, "Espaço para diálogo interdisciplinar que frequentemente não existe no cotidiano escolar, mas que enriquece a prática pedagógica de todos;")
    add_bullet(doc, "Participação ativa nas decisões que afetarão o curso, com voz genuína no processo de planejamento;")
    add_bullet(doc, "Fortalecimento do senso de comunidade profissional entre os docentes que compartilham o compromisso com a formação dos estudantes.")

    add_subsection(doc, "10.3 Para a Instituição")
    add_body(doc,
        "Para a EEB Araújo Figueiredo como instituição, os benefícios de um processo bem conduzido "
        "de reflexão e eventual evolução do Curso Técnico em Marketing são de natureza tanto "
        "pedagógica quanto reputacional e estratégica:"
    )
    add_bullet(doc, "Fortalecimento da qualidade formativa como marca institucional reconhecida pela comunidade, pelas famílias e pelos estudantes;")
    add_bullet(doc, "Posicionamento da escola como referência regional em educação técnica aplicada e contextualizada;")
    add_bullet(doc, "Alinhamento com as expectativas do mercado de trabalho local, tornando os egressos do curso ativos valorizados pela economia regional;")
    add_bullet(doc, "Construção de memória institucional — registros, processos e saberes que protegem a continuidade do curso mesmo diante de mudanças de equipe;")
    add_bullet(doc, "Fortalecimento das parcerias com a comunidade empresarial e civil de Urubici e região.")

    add_subsection(doc, "10.4 Para a Comunidade Local")
    add_body(doc,
        "A escola pública que forma bem não beneficia apenas seus estudantes — ela contribui para "
        "o desenvolvimento do território em que está inserida. Os benefícios esperados para a "
        "comunidade de Urubici incluem:"
    )
    add_bullet(doc, "Profissionais de marketing formados localmente com capacidade de apoiar o desenvolvimento econômico da região — no turismo, no comércio, no agronegócio, nos serviços;")
    add_bullet(doc, "Jovens que permanecem ou retornam à região por encontrarem oportunidades profissionais conectadas à sua formação;")
    add_bullet(doc, "Pequenos negócios e empreendedores locais que podem contar, progressivamente, com o apoio de estudantes e egressos qualificados para suas necessidades de comunicação digital;")
    add_bullet(doc, "Valorização da identidade cultural e econômica local por meio de projetos educativos que tomam Urubici como objeto de aprendizagem e como território de aplicação;")
    add_bullet(doc, "Fortalecimento do papel social da escola como instituição que contribui ativamente para o desenvolvimento humano e econômico da comunidade que a mantém e que ela serve.")

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 11 — PRÓXIMOS PASSOS SUGERIDOS
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "11", "Próximos Passos Sugeridos")

    add_body(doc,
        "Este documento conclui com uma proposição simples: antes de qualquer decisão sobre o futuro "
        "do Curso Técnico em Marketing, é necessário conhecer melhor o presente. Isso significa "
        "conduzir um processo cuidadoso, respeitoso e participativo de diagnóstico — que escute os "
        "professores, compreenda os estudantes, mapeie o que já existe de valor e identifique, com "
        "clareza, os pontos que merecem atenção."
    )
    add_note_block(doc,
        "Os próximos passos que se propõem não são um plano de implementação. São um processo de "
        "escuta e de construção coletiva de conhecimento, sem o qual qualquer decisão curricular "
        "seria prematura."
    )

    add_subsection(doc, "Processo Diagnóstico com o Corpo Docente")
    add_body(doc,
        "O primeiro e mais importante passo é ouvir os professores. Cada professor que atua no curso "
        "carrega percepções, experiências e conhecimentos que nenhum documento externo pode captar. "
        "Um processo de escuta individual — realizado com respeito, sem agenda oculta e com genuíno "
        "interesse no que cada docente tem a dizer — é a base indispensável para qualquer reflexão "
        "institucional bem fundamentada."
    )
    add_body(doc,
        "Esse processo poderia incluir conversas individuais estruturadas, espaços de compartilhamento "
        "coletivo e instrumentos simples de coleta de percepções sobre o curso, os estudantes e as "
        "possibilidades que os próprios professores identificam. O objetivo não é avaliar práticas "
        "— é compreender realidades."
    )

    add_subsection(doc, "Processo Diagnóstico com os Estudantes")
    add_body(doc,
        "Os estudantes são atores centrais da vida do curso e têm perspectivas insubstituíveis sobre "
        "sua experiência de aprendizagem. Ouvir o que percebem como pontos fortes da formação, o que "
        "sentem falta e o que acreditam que poderia ser diferente é um recurso pedagógico de grande "
        "valor — e uma demonstração de respeito pela sua condição de sujeitos ativos da educação "
        "que recebem."
    )
    add_body(doc,
        "Instrumentos simples de escuta — questionários anônimos, rodas de conversa, formulários "
        "de avaliação — podem oferecer uma visão qualificada da experiência discente, sem criar "
        "expectativas de mudança imediata."
    )

    add_subsection(doc, "Análise e Síntese das Informações Coletadas")
    add_body(doc,
        "Os dados coletados no diagnóstico — percepções docentes, perspectivas discentes, análise "
        "do contexto atual — precisam ser sistematizados e analisados com cuidado. Essa análise não "
        "deve ser realizada por uma única pessoa, mas por um grupo representativo da equipe "
        "pedagógica, que traga diferentes perspectivas para a leitura do material."
    )
    add_body(doc,
        "O produto dessa etapa seria um documento de diagnóstico — uma fotografia honesta do curso "
        "em seu estado atual, com seus pontos fortes, seus desafios e suas oportunidades — que "
        "servirá de base para qualquer conversa futura sobre planejamento."
    )

    add_subsection(doc, "Discussão Curricular em Grupo de Trabalho Colaborativo")
    add_body(doc,
        "Somente após a conclusão do diagnóstico seria oportuno iniciar conversas mais específicas "
        "sobre competências, conteúdos, metodologias e estrutura do curso. Essa discussão deve ser "
        "conduzida em um grupo de trabalho colaborativo — com representação da gestão, da coordenação "
        "pedagógica, dos professores e, quando apropriado, dos próprios estudantes."
    )
    add_body(doc,
        "Não há pressa nessa etapa. A qualidade do processo é mais importante do que a velocidade "
        "das conclusões. Decisões curriculares tomadas com base em diagnóstico sólido e construção "
        "coletiva têm muito mais chance de ser bem implementadas e sustentadas ao longo do tempo."
    )

    add_subsection(doc, "Planejamento Colaborativo como Próxima Etapa")
    add_body(doc,
        "O encaminhamento natural de todo o processo descrito acima é a elaboração de um plano "
        "pedagógico colaborativo — construído pela equipe que conhece o curso, os estudantes e a "
        "comunidade — que traduza as reflexões deste framework em propostas concretas, viáveis e "
        "alinhadas com a identidade e os valores da EEB Araújo Figueiredo."
    )
    add_body(doc,
        "Esse plano será, em seu momento, um documento diferente deste. Enquanto este framework "
        "estabelece fundamentos e convida ao diálogo, o plano futuro — quando existir — será "
        "construído a partir das vozes da equipe, dos dados do diagnóstico e do compromisso "
        "coletivo com a qualidade da formação que os estudantes merecem."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # BLOCO DE ENCERRAMENTO
    # ═══════════════════════════════════════════════════════════════════════════
    add_closing_block(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # GOVERNANÇA DOCUMENTAL
    # ═══════════════════════════════════════════════════════════════════════════
    add_governance_block(doc)

    # ── Salvar ────────────────────────────────────────────────────────────────
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f"✓ Documento gerado: {OUTPUT_DOCX}")


if __name__ == "__main__":
    build_document()
