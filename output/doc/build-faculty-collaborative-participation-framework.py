"""
Gerador do documento institucional:
Framework de Participação Docente e Construção Colaborativa
para o Fortalecimento do Curso Técnico em Marketing
EEB Araújo Figueiredo — Araujo Innovation Lab
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
OUTPUT_DOCX = OUT_DIR / "faculty-collaborative-participation-framework-v1.docx"


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

    p_left = cells[0].paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r_left = p_left.add_run("Circulação Interna Restrita  |  v1.0")
    run_style(r_left, "Arial", 8, C["muted"])
    no_borders(cells[0])

    p_center = cells[1].paragraphs[0]
    p_center.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_pg = p_center.add_run("Página ")
    run_style(r_pg, "Arial", 8, C["muted"])
    add_page_field(p_center)
    r_of = p_center.add_run(" de ")
    run_style(r_of, "Arial", 8, C["muted"])
    add_numpages_field(p_center)
    no_borders(cells[1])

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


def add_subsection3(doc: Document, title: str) -> None:
    p = doc.add_paragraph(style="Heading 3")
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.clear()
    r = p.add_run(title)
    run_style(r, "Arial", 11, C["blue"], bold=True)


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
    r1 = p.add_run("Nota institucional:  ")
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
        ("Nome do Documento",
         "Framework de Participação Docente e Construção Colaborativa para o Fortalecimento do Curso Técnico em Marketing"),
        ("Versão",                    "v1.0"),
        ("Data",                      "Junho de 2026"),
        ("Status",                    "Documento preparatório para participação docente institucional"),
        ("Elaborado por",             "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários Primários",   "Corpo Docente · Orientação Educacional · Coordenação Pedagógica · Direção Escolar"),
        ("Destinatários Secundários", "Futuros grupos de trabalho institucionais"),
        ("Instituição",               "EEB Araújo Figueiredo — Araujo Innovation Lab — Urubici, SC"),
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
    tag_run = tag_p.add_run("DOCUMENTO ESTRATÉGICO INSTITUCIONAL — ALINHAMENTO PEDAGÓGICO INTERNO")
    run_style(tag_run, "Arial", 8, C["muted"], bold=True)

    # Título principal
    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_before = Pt(28)
    title.paragraph_format.space_after = Pt(14)
    title.add_run(
        "Framework de Participação Docente e Construção Colaborativa "
        "para o Fortalecimento do Curso Técnico em Marketing"
    )

    # Subtítulo descritivo
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(36)
    r_sub = sub.add_run(
        "Documento preparatório para participação docente institucional, voltado ao "
        "alinhamento pedagógico interno entre Corpo Docente, Orientação Educacional, "
        "Coordenação Pedagógica e Direção Escolar."
    )
    run_style(r_sub, "Arial", 12, C["muted"])

    # Bloco de identificação do documento (tabela)
    id_table = doc.add_table(rows=5, cols=2)
    id_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    id_rows = [
        ("Versão",        "v1.0 — Primeira Emissão Oficial"),
        ("Data",          "Junho de 2026"),
        ("Status",        "Documento preparatório — nenhuma decisão curricular tomada"),
        ("Elaborado por", "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários", "Corpo Docente · Orientação · Coordenação · Direção"),
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
        "Nenhuma decisão curricular foi tomada ou está sendo proposta neste documento."
    )
    run_style(r_e2, "Arial", 10, C["gold"], bold=True)

    end_p3 = end_cell.add_paragraph()
    end_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p3.paragraph_format.space_before = Pt(4)
    end_p3.paragraph_format.space_after = Pt(4)
    r_e3 = end_p3.add_run(
        "A construção do futuro do Curso Técnico em Marketing pertence, coletivamente, a todos os seus educadores."
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

    p_gov = doc.add_paragraph()
    p_gov.paragraph_format.space_before = Pt(14)
    p_gov.paragraph_format.space_after = Pt(10)
    r_gov = p_gov.add_run("GOVERNANÇA DOCUMENTAL")
    run_style(r_gov, "Georgia", 14, C["blue"], bold=True)
    para_bottom_border(p_gov, color=C["line"], size="4")

    rows_data = [
        ("Documento-fonte institucional",  "faculty-collaborative-participation-framework-v1.md"),
        ("Documento derivado obrigatório", "faculty-collaborative-participation-framework-v1.docx"),
        ("Padrão aplicável — documentos",  "document-template-standard.md"),
        ("Padrão aplicável — identidade",  "branding-standard.md"),
        ("Classificação",                  "Documento Estratégico Institucional"),
        ("Uso",                            "Alinhamento Pedagógico Interno"),
        ("Versão",                         "v1.0"),
        ("Instituição",                    "EEB Araújo Figueiredo — Araujo Innovation Lab"),
        ("Local",                          "Urubici – SC"),
        ("Status",                         "Aprovado para conversão DOCX institucional"),
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
        header_text="Araujo Innovation Lab  |  Framework de Participação Docente — Curso Técnico em Marketing",
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
        "O presente documento tem como propósito instituir, de forma transparente e fundamentada, "
        "um processo de participação docente voltado à construção coletiva de reflexões sobre o "
        "Curso Técnico em Marketing da EEB Araújo Figueiredo. Trata-se de uma iniciativa que nasce "
        "do reconhecimento — genuíno e irrestrito — do percurso já trilhado por todos aqueles que, "
        "ao longo dos anos, dedicaram seu conhecimento, sua experiência e seu comprometimento "
        "profissional à formação dos estudantes desta instituição."
    )
    add_body(doc,
        "O Curso Técnico em Marketing da EEB Araújo Figueiredo tem uma história. Essa história é "
        "feita de aulas preparadas com cuidado, de relações pedagógicas construídas no dia a dia, "
        "de estudantes formados e encaminhados ao mercado de trabalho, e de práticas educativas "
        "consolidadas por profissionais que conhecem profundamente o que ensinam e para quem ensinam. "
        "Reconhecer esse percurso não é uma formalidade protocolar — é a base sobre a qual qualquer "
        "reflexão institucional precisa, obrigatoriamente, se assentar."
    )
    add_body(doc,
        "Este documento não é um instrumento de avaliação. Não há aqui, em nenhuma seção, qualquer "
        "intenção de mensurar desempenho, identificar lacunas individuais ou propor correções de "
        "trajetória. A natureza deste processo é estritamente colaborativa e reflexiva: trata-se de "
        "uma escuta qualificada, estruturada institucionalmente, que busca reunir o conhecimento que "
        "os professores possuem — e que apenas eles possuem — sobre a realidade do curso, a "
        "experiência dos estudantes e as possibilidades que o futuro pode reservar."
    )
    add_body(doc,
        "O convite que este documento formaliza é, em essência, um convite ao diálogo. Um diálogo "
        "entre pares, entre profissionais que compartilham o compromisso com a qualidade da formação "
        "oferecida aos jovens desta comunidade. Um diálogo sem agendas ocultas, sem conclusões "
        "previamente definidas e sem qualquer pressuposto de que o que existe precisa ser transformado."
    )
    add_highlight_block(doc, "Compromisso deste documento",
        "O futuro do curso será construído coletivamente. Este documento é o primeiro passo formal "
        "desse processo. Nenhuma decisão curricular está sendo tomada ou proposta aqui.",
        fill=C["blue_soft2"], border_color=C["blue"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 2 — CONTEXTO DA INICIATIVA
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "2", "Contexto da Iniciativa")

    add_body(doc,
        "A educação profissional técnica existe em diálogo permanente com o mundo ao qual os "
        "estudantes serão, ao término da formação, lançados. Esse mundo, nas últimas décadas, passou "
        "por transformações de uma amplitude e de uma velocidade que poucos períodos históricos "
        "conheceram. Compreender essas transformações — sem alarmismo e sem ingenuidade — é parte "
        "da responsabilidade formativa de toda instituição de ensino que leva a sério o compromisso "
        "com seus estudantes."
    )

    add_subsection(doc, "A Comunicação em Transformação Contínua")
    add_body(doc,
        "A área da comunicação, que constitui o núcleo identitário da formação em marketing, passou "
        "por uma reconfiguração profunda. O surgimento e a consolidação das plataformas digitais "
        "criaram novos espaços de interação entre marcas, organizações e pessoas. A produção de "
        "mensagens, antes circunscrita a profissionais com acesso a estruturas de comunicação, "
        "tornou-se uma prática disseminada. A circulação de conteúdo ganhou velocidade e abrangência "
        "sem precedentes. Novos formatos de comunicação — audiovisual, interativo, imersivo — "
        "passaram a coexistir e, por vezes, a competir com os formatos tradicionais."
    )
    add_body(doc,
        "Esse movimento não representa uma ruptura com os fundamentos da comunicação estratégica "
        "— a compreensão do público, a clareza da mensagem, a consistência da identidade e a ética "
        "no relacionamento com o consumidor continuam sendo as bases sobre as quais qualquer prática "
        "profissional de marketing se sustenta. O que se alterou foi o ecossistema em que essas "
        "práticas se realizam, com novos canais, novas ferramentas e novas expectativas por parte "
        "daqueles que recebem e que produzem conteúdo."
    )

    add_subsection(doc, "A Evolução Tecnológica como Variável Estrutural")
    add_body(doc,
        "A tecnologia digital tornou-se, nas últimas décadas, uma variável estrutural da vida "
        "econômica, social e cultural. Pequenas empresas, negócios locais, profissionais autônomos "
        "e organizações de todos os portes operam, em graus variados, em ambientes digitais — "
        "gerenciando sua presença online, construindo reputação em plataformas, comunicando seus "
        "produtos e serviços por meios que antes eram inacessíveis a operações de menor escala."
    )
    add_body(doc,
        "Mais recentemente, a inteligência artificial generativa incorporou-se à prática profissional "
        "com uma presença que não pode ser desconsiderada. Ferramentas capazes de apoiar a produção "
        "de textos, imagens, análises e estratégias passaram a fazer parte do cotidiano de trabalho "
        "de agências, equipes de marketing e empreendedores individuais. Essa realidade não elimina "
        "— e em nenhuma medida diminui — a necessidade de julgamento humano, de criatividade, de "
        "ética e de perspectiva estratégica."
    )

    add_subsection(doc, "A Transformação do Mercado de Trabalho")
    add_body(doc,
        "O mercado de trabalho para profissionais de marketing passou por uma diversificação "
        "significativa. A demanda por profissionais capazes de operar em ambientes digitais, de "
        "produzir e gerir conteúdo, de analisar dados de comportamento de consumidores e de articular "
        "comunicação em múltiplos canais cresceu de forma consistente. Ao mesmo tempo, o "
        "empreendedorismo — a capacidade de identificar oportunidades, construir propostas de valor "
        "e comunicá-las com eficácia — tornou-se uma competência cada vez mais valorizada, "
        "independentemente do vínculo de trabalho formal ou autônomo do profissional."
    )
    add_body(doc,
        "Para uma cidade como Urubici — com seu ecossistema de turismo, gastronomia, comércio local, "
        "agronegócio e cultura regional — essa demanda é concreta e próxima. Há, no território em "
        "que a escola está inserida, um contexto econômico vivo que pode servir tanto como campo de "
        "aprendizagem quanto como destino profissional para os egressos do curso."
    )

    add_subsection(doc, "A Transformação do Perfil dos Estudantes")
    add_body(doc,
        "Os jovens que chegam ao Curso Técnico em Marketing trazem consigo experiências digitais "
        "ricas e um repertório de consumo cultural e comunicacional que distingue sua geração de "
        "todas as anteriores. Eles interagem com conteúdo de forma intensa, habitam plataformas "
        "digitais com naturalidade e, em muitos casos, já produziram alguma forma de conteúdo antes "
        "de ingressar no ensino técnico. Ao mesmo tempo, frequentemente carecem de estrutura para "
        "transformar esse repertório informal em competência profissional reconhecida e aplicável."
    )

    add_subsection(doc, "A Centralidade do Conhecimento Docente")
    add_body(doc,
        "Nenhum conjunto de transformações externas pode ser compreendido e traduzido "
        "pedagogicamente sem a mediação dos professores. São eles que conhecem como os estudantes "
        "chegam ao curso, quais são suas dificuldades e suas potencialidades, o que funciona em "
        "sala de aula e o que não funciona, quais conexões entre conteúdo e realidade ressoam com "
        "mais profundidade. É esse conhecimento acumulado — singular, insubstituível e de alto valor "
        "institucional — que este processo busca, acima de tudo, ouvir e valorizar."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 3 — O PAPEL DOS PROFESSORES
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "3", "O Papel dos Professores na Construção do Futuro do Curso")

    add_body(doc,
        "Em qualquer processo de reflexão institucional sobre educação, o ponto de partida mais "
        "sólido não é um conjunto de dados externos, nem uma tendência de mercado, nem um referencial "
        "teórico — por mais valioso que sejam esses elementos. O ponto de partida mais sólido é a "
        "experiência acumulada dos educadores que constroem, dia a dia, a realidade da formação que "
        "a instituição oferece."
    )
    add_body(doc,
        "Os professores do Curso Técnico em Marketing da EEB Araújo Figueiredo são portadores de "
        "um conhecimento que vai muito além do domínio dos conteúdos de suas disciplinas — embora "
        "esse domínio seja, em si mesmo, de alto valor. Eles conhecem os estudantes que chegam ao "
        "curso: suas motivações, seus contextos familiares, suas dificuldades iniciais, suas "
        "transformações ao longo da formação. Conhecem o que funciona pedagogicamente neste contexto "
        "específico — nesta escola, nesta cidade, com estes estudantes. Conhecem as limitações e as "
        "potencialidades do ambiente institucional em que trabalham."
    )

    add_subsection(doc, "Experiência Docente como Patrimônio Institucional")
    add_body(doc,
        "A trajetória profissional de cada professor é um patrimônio que pertence não apenas ao "
        "indivíduo, mas à instituição e, em última análise, aos estudantes que se beneficiam dela. "
        "Anos de prática pedagógica acumulam percepções que não se transferem por documentos — elas "
        "se constroem na relação com os estudantes, no enfrentamento das dificuldades do cotidiano "
        "escolar, nas adaptações que cada educador realiza silenciosamente para que o aprendizado "
        "aconteça. Reconhecer esse patrimônio é o ponto de partida ético de qualquer iniciativa "
        "institucional que se proponha a ser verdadeiramente colaborativa."
    )
    add_body(doc,
        "Professores que vivenciaram diferentes momentos da história do curso conhecem sua evolução. "
        "Sabem o que foi tentado e o que funcionou, o que foi descontinuado e por quê, quais eram "
        "os desafios de dez anos atrás e quais são os de hoje. Esse conhecimento histórico é "
        "inestimável para qualquer reflexão sobre o futuro — porque o futuro, quando bem construído, "
        "é sempre uma continuidade inteligente do passado, não uma negação dele."
    )

    add_subsection(doc, "Conhecimento Pedagógico Situado")
    add_body(doc,
        "O conhecimento pedagógico dos professores deste curso é situado — construído a partir de "
        "uma realidade concreta, com estudantes reais, em um contexto específico. Esse tipo de "
        "conhecimento tem uma qualidade que os saberes genéricos não possuem: ele foi testado na "
        "prática. As estratégias pedagógicas que os professores desenvolveram, as formas de apresentar "
        "conceitos que ressoam com os estudantes, as sequências didáticas que facilitam a compreensão "
        "— tudo isso representa uma inteligência pedagógica desenvolvida ao longo do tempo, cujo "
        "valor é difícil de superestimar."
    )

    add_subsection(doc, "A Articulação entre Experiência e Renovação")
    add_body(doc,
        "A força de qualquer processo educativo reside na articulação entre o que foi consolidado "
        "e o que está em transformação. A experiência docente acumulada não se opõe à renovação — "
        "ela é o que torna a renovação possível sem que o essencial se perca. São os professores "
        "experientes que sabem distinguir o que é modismo passageiro do que representa uma mudança "
        "estrutural relevante. São eles que conseguem avaliar, com critério, quais inovações "
        "pedagógicas têm potencial real para enriquecer a formação dos estudantes e quais "
        "representam apenas novidade sem substância."
    )

    add_subsection(doc, "Professores como Co-autores do Futuro")
    add_body(doc,
        "A distinção entre ser objeto de um processo e ser co-autor de um processo é fundamental "
        "para compreender a natureza desta iniciativa. Os professores não estão sendo convidados a "
        "validar decisões já tomadas, nem a receber prescrições elaboradas por outros. Estão sendo "
        "convidados a construir, a partir de suas perspectivas e experiências, as reflexões que "
        "servirão de base para as discussões futuras."
    )
    add_highlight_block(doc, "Posicionamento institucional",
        "Essa co-autoria não é apenas uma declaração de intenção — ela é a condição de possibilidade "
        "de qualquer resultado que tenha valor e sustentabilidade. Um processo educativo que não "
        "incorpora o conhecimento dos seus protagonistas docentes é, desde o início, um processo "
        "incompleto.",
        fill=C["blue_soft2"], border_color=C["blue"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 4 — OBJETIVOS DO PROCESSO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "4", "Objetivos do Processo de Participação Docente")

    add_body(doc,
        "O processo de participação docente que este documento formaliza tem objetivos "
        "cuidadosamente delimitados. Essa delimitação não é uma restrição — é uma garantia de "
        "que o processo permanecerá fiel ao seu propósito essencial: ouvir, compreender e reunir "
        "o conhecimento que os professores possuem sobre o curso e sobre os estudantes."
    )

    add_subsection(doc, "Escutar com Profundidade e Respeito")
    add_body(doc,
        "O primeiro e mais fundamental objetivo deste processo é escutar. Não uma escuta formal e "
        "superficial, cumprida por protocolo, mas uma escuta genuína, atenta e respeitosa às "
        "experiências, percepções e perspectivas de cada professor. A qualidade desta escuta é o "
        "critério mais importante pelo qual o processo será avaliado — não pelo número de instrumentos "
        "aplicados, mas pela profundidade e pela autenticidade das contribuições que conseguir reunir."
    )

    add_subsection(doc, "Compreender a Realidade Atual do Curso")
    add_body(doc,
        "O processo busca compreender, com fidelidade e nuance, a realidade atual do Curso Técnico "
        "em Marketing — tal como ela é vivida e percebida pelos professores que nela atuam. Essa "
        "compreensão abrange tanto os aspectos que funcionam bem e que merecem ser preservados e "
        "valorizados quanto os aspectos que os próprios professores identificam como oportunidades "
        "de crescimento e desenvolvimento."
    )
    add_body(doc,
        "Essa compreensão não é um diagnóstico de problemas. É uma fotografia qualificada da "
        "realidade atual, construída a partir do ponto de observação mais privilegiado que existe: "
        "a perspectiva de quem vive essa realidade cotidianamente."
    )

    add_subsection(doc, "Reunir Perspectivas Diversas e Complementares")
    add_body(doc,
        "O Curso Técnico em Marketing envolve professores com diferentes formações, trajetórias e "
        "experiências profissionais. Essa diversidade é uma riqueza — não uma fragmentação. "
        "Perspectivas diferentes sobre o mesmo fenômeno educacional enriquecem a compreensão "
        "coletiva, revelam dimensões que uma perspectiva única não alcançaria e constroem uma "
        "visão mais completa e realista da realidade do curso."
    )

    add_subsection(doc, "Identificar Oportunidades a Partir do Olhar Docente")
    add_body(doc,
        "Os professores são, em sua prática cotidiana, observadores privilegiados das oportunidades "
        "que o curso ainda não explorou plenamente. Eles percebem onde os estudantes demonstram "
        "maior engajamento, quais temas despertam curiosidade, quais conexões entre conteúdo e "
        "realidade ressoam com mais força. Reunir essas percepções — de forma organizada e "
        "institucionalizada — é um dos objetivos mais valiosos deste processo."
    )

    add_subsection(doc, "Apoiar Reflexões Futuras com Fundamento Sólido")
    add_body(doc,
        "O processo de participação docente não se propõe a tomar decisões — propõe-se a criar as "
        "condições para que decisões futuras, quando necessárias, sejam tomadas com base em "
        "informações qualificadas, perspectivas diversas e construção coletiva. As contribuições "
        "dos professores tornar-se-ão parte do substrato sobre o qual qualquer reflexão institucional "
        "futura será desenvolvida."
    )
    add_body(doc,
        "Esse objetivo confere ao processo um valor que transcende o momento atual. As contribuições "
        "reunidas agora terão pertinência para além do processo imediato — elas constituirão uma "
        "memória institucional viva sobre como os educadores desta escola percebem a formação que "
        "oferecem e as possibilidades que vislumbram para o futuro."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 5 — PRINCÍPIOS NORTEADORES
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "5", "Princípios Norteadores do Processo")

    add_body(doc,
        "Os princípios a seguir constituem os fundamentos éticos e institucionais sobre os quais "
        "este processo se organiza. Eles não são apenas diretrizes operacionais — são compromissos "
        "que a instituição assume com os professores ao convidá-los a participar. Cada princípio "
        "deve ser compreendido, aplicado e verificado ao longo de todas as etapas do processo."
    )

    add_subsection3(doc, "Respeito à Trajetória Profissional")
    add_body(doc,
        "O princípio do respeito à trajetória profissional parte do reconhecimento de que cada "
        "professor carrega consigo uma história de formação, desenvolvimento e prática que merece "
        "consideração incondicional. A trajetória profissional não é apenas um conjunto de "
        "credenciais — é a expressão de escolhas, dedicação, superação de desafios e construção "
        "de um modo singular de exercer a docência. Esse percurso, independentemente de seu "
        "conteúdo específico, é digno de reconhecimento institucional."
    )
    add_body(doc,
        "Este princípio implica que nenhuma contribuição seja recebida como mais ou menos válida "
        "em função do tempo de serviço, da titulação ou de qualquer outro marcador hierárquico. "
        "Respeitar a trajetória profissional significa também não criar situações em que o professor "
        "sinta que sua prática está sendo questionada ou que suas escolhas pedagógicas estão sob "
        "escrutínio. O processo de participação docente não é uma avaliação de desempenho — é uma "
        "oportunidade de contribuição que pressupõe, como ponto de partida, a competência e a "
        "boa-fé de cada participante."
    )

    add_subsection3(doc, "Escuta Qualificada")
    add_body(doc,
        "A escuta qualificada distingue-se da escuta ordinária pelo grau de atenção, cuidado e "
        "intencionalidade com que é realizada. Escutar de forma qualificada significa criar condições "
        "para que o professor possa expressar suas perspectivas com liberdade, sem receio de "
        "julgamento e com a confiança de que o que compartilha será recebido com seriedade."
    )
    add_body(doc,
        "A escuta qualificada também implica a capacidade de receber perspectivas que não se alinham "
        "com expectativas previamente formadas. Contribuições que apontam para direções inesperadas, "
        "que expressam ceticismo fundamentado ou que identificam desafios que outros não perceberam "
        "têm valor igual — ou maior — do que aquelas que confirmam o que já se sabia. O processo "
        "ganha profundidade e credibilidade exatamente quando consegue acolher essa diversidade de "
        "perspectivas com igual respeito."
    )

    add_subsection3(doc, "Construção Colaborativa")
    add_body(doc,
        "A colaboração genuína pressupõe que todos os participantes de um processo são, de fato, "
        "construtores — e não apenas consultados ou informados. Este processo é colaborativo em seu "
        "sentido mais pleno: os professores não estão sendo chamados a validar um caminho já "
        "decidido, mas a contribuir para a construção de um caminho que ainda não existe."
    )
    add_body(doc,
        "Esse princípio tem implicações práticas importantes. Significa que as contribuições dos "
        "professores não serão utilizadas apenas para confirmar percepções já formadas, mas para "
        "ampliar, corrigir e enriquecer a compreensão institucional sobre o curso. A mesma "
        "disposição que valoriza a colaboração como estratégia de aprendizagem dos estudantes deve "
        "orientar as práticas de desenvolvimento institucional da equipe docente."
    )

    add_subsection3(doc, "Transparência")
    add_body(doc,
        "O princípio da transparência exige que este processo seja, em todas as suas dimensões, "
        "compreensível para os participantes. Os professores têm o direito de saber com clareza: "
        "por que este processo está sendo conduzido, o que será feito com as contribuições que "
        "oferecerem, quem terá acesso às informações coletadas, como os resultados serão utilizados "
        "e quais são os limites do uso das informações compartilhadas."
    )
    add_body(doc,
        "A transparência não é apenas uma questão ética — é uma condição de confiança. Um processo "
        "que não explica suas intenções e seus encaminhamentos não inspira a confiança necessária "
        "para que os participantes contribuam de forma autêntica e aprofundada. Este documento é, "
        "em si mesmo, uma expressão do compromisso com a transparência."
    )

    add_subsection3(doc, "Foco na Formação dos Estudantes")
    add_body(doc,
        "Em todos os momentos e em todas as etapas, a centralidade dos estudantes deve ser o "
        "horizonte que orienta as reflexões. Os professores são convidados a participar não em "
        "nome de seus interesses individuais, nem em nome de uma agenda institucional abstrata, "
        "mas em nome dos jovens que chegam ao Curso Técnico em Marketing com expectativas reais, "
        "trajetórias específicas e um futuro profissional que a escola tem a responsabilidade de "
        "preparar."
    )
    add_body(doc,
        "O foco nos estudantes é, também, o terreno de convergência mais sólido entre perspectivas "
        "docentes diversas. Professores com diferentes formações, experiências e visões pedagógicas "
        "compartilham, invariavelmente, o compromisso com a formação dos estudantes que lhes são "
        "confiados. É esse compromisso compartilhado que torna possível a construção colaborativa."
    )

    add_subsection3(doc, "Valorização da Diversidade de Perspectivas")
    add_body(doc,
        "O corpo docente de um curso multidisciplinar como o Técnico em Marketing é, por natureza, "
        "diverso. Essa diversidade — de formações acadêmicas, de experiências profissionais, de "
        "abordagens pedagógicas, de tempos de serviço, de formas de perceber os estudantes e o "
        "mundo — é um recurso intelectual de alto valor que este processo se propõe a reconhecer "
        "e a aproveitar."
    )
    add_body(doc,
        "A valorização da diversidade de perspectivas exige que nenhuma visão seja tratada como a "
        "perspectiva correta ou como o referencial a partir do qual as demais serão julgadas. Cada "
        "professor percebe o curso de um ângulo que lhe é próprio; a composição dessas diferentes "
        "percepções é que produz uma compreensão mais completa e mais fiel da realidade do que "
        "qualquer ponto de vista singular poderia alcançar."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 6 — ESTRUTURA DO PROCESSO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "6", "Estrutura do Processo de Participação")

    add_body(doc,
        "O processo de participação docente organiza-se em quatro etapas sequenciais, cada uma "
        "com propósito e dinâmica próprios. Essa estrutura foi desenhada para garantir que as "
        "contribuições individuais dos professores sejam cuidadosamente coletadas, organizadas, "
        "refletidas coletivamente e transformadas em subsídios institucionais de qualidade."
    )

    add_subsection(doc, "Etapa 1 — Participação por Instrumento Institucional")
    add_body(doc,
        "A primeira etapa do processo consiste na participação individual de cada professor por "
        "meio de um instrumento institucional estruturado. Esse instrumento tem como propósito "
        "criar um espaço protegido e organizado para que cada professor possa expressar suas "
        "percepções, suas perspectivas e suas reflexões sobre o curso de acordo com seu próprio "
        "ritmo e sua própria forma de elaboração."
    )
    add_body(doc,
        "A participação por instrumento institucional garante que todos os professores tenham a "
        "mesma oportunidade de contribuir — independentemente de sua disponibilidade para participar "
        "de encontros coletivos, de sua facilidade com a expressão oral em grupo ou de qualquer "
        "outra variável que possa criar assimetrias de participação. A confidencialidade das "
        "contribuições individuais é um compromisso institucional desta etapa."
    )

    add_subsection(doc, "Etapa 2 — Consolidação das Contribuições")
    add_body(doc,
        "A segunda etapa consiste na organização e consolidação das contribuições coletadas na "
        "etapa anterior. Esse trabalho, realizado pela equipe pedagógica responsável pelo processo, "
        "tem como objetivo transformar o conjunto de perspectivas individuais em um material "
        "organizado que preserve a diversidade de vozes sem perder a coerência analítica."
    )
    add_body(doc,
        "A consolidação não é uma interpretação — é uma organização cuidadosa das contribuições, "
        "que busca identificar convergências, preservar dissonâncias relevantes e construir uma "
        "síntese fiel ao que os professores expressaram. Perspectivas minoritárias ou inesperadas "
        "não são descartadas em favor da maioria — elas são registradas com igual atenção, porque "
        "frequentemente revelam dimensões da realidade que a convergência majoritária não alcança."
    )

    add_subsection(doc, "Etapa 3 — Reflexão Coletiva")
    add_body(doc,
        "A terceira etapa organiza-se em torno de espaços de reflexão coletiva — encontros "
        "estruturados nos quais os professores, a partir da síntese elaborada na etapa anterior, "
        "podem aprofundar as discussões, complementar as perspectivas individuais com o diálogo "
        "entre pares e construir, coletivamente, uma compreensão mais rica e mais fundamentada "
        "sobre os temas em reflexão."
    )
    add_body(doc,
        "A reflexão coletiva não tem como objetivo chegar a conclusões definitivas ou a posições "
        "de consenso sobre questões que exigem maturidade e tempo. Seu propósito é criar "
        "oportunidades para que o conhecimento individual ganhe dimensão coletiva — que as "
        "percepções de cada professor sejam enriquecidas pelo contato com as perspectivas dos "
        "colegas, e que essa troca produza uma inteligência coletiva superior à soma das partes."
    )

    add_subsection(doc, "Etapa 4 — Subsídios para Futuras Discussões")
    add_body(doc,
        "A quarta etapa consiste na sistematização das reflexões coletivas em subsídios "
        "institucionais que servirão de base para discussões futuras sobre o fortalecimento do "
        "curso. Esses subsídios são documentos de apoio — não decisões, não propostas e não "
        "planos de ação. São a materialização do conhecimento coletivo produzido ao longo do "
        "processo, organizado de forma que possa ser utilizado com qualidade nas etapas seguintes "
        "da vida institucional do curso."
    )
    add_body(doc,
        "Os subsídios elaborados nesta etapa serão compartilhados com todos os professores "
        "participantes, garantindo que cada um tenha acesso ao produto final do processo em que "
        "contribuiu. Essa devolução é parte do compromisso de transparência que orienta o processo "
        "e representa o reconhecimento formal da contribuição de cada educador."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 7 — UTILIZAÇÃO DAS CONTRIBUIÇÕES
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "7", "Utilização das Contribuições")

    add_body(doc,
        "A clareza sobre o uso das contribuições coletadas ao longo do processo é um elemento "
        "essencial de respeito e de confiança. Os professores que participam têm o direito de "
        "saber, com precisão, como suas perspectivas serão utilizadas — e como não serão utilizadas. "
        "Esta seção estabelece, de forma explícita e inequívoca, os dois lados desse compromisso."
    )

    add_subsection(doc, "Como as contribuições serão utilizadas")
    add_body(doc,
        "As contribuições dos professores serão utilizadas exclusivamente como subsídios para "
        "reflexões institucionais futuras sobre o Curso Técnico em Marketing. Elas alimentarão "
        "um processo de construção coletiva de conhecimento sobre a realidade do curso, sobre o "
        "perfil dos estudantes, sobre as práticas pedagógicas que os professores consideram "
        "relevantes e sobre as possibilidades que eles identificam para o futuro da formação."
    )
    add_body(doc,
        "As contribuições serão sistematizadas e organizadas de forma a preservar a diversidade "
        "de perspectivas e a profundidade das reflexões individuais. Sínteses elaboradas a partir "
        "dessas contribuições serão compartilhadas com os professores participantes antes de "
        "qualquer encaminhamento institucional, garantindo que o material seja fiel ao que foi "
        "expresso. As perspectivas reunidas poderão informar, em etapas futuras e apropriadas, "
        "discussões sobre possíveis fortalecimentos da proposta formativa do curso."
    )

    add_subsection(doc, "Como as contribuições não serão utilizadas")
    add_body(doc,
        "As contribuições dos professores não serão utilizadas, em nenhuma hipótese, para fins "
        "avaliativos. Nenhuma contribuição individual será interpretada como evidência de "
        "desempenho, de qualificação ou de adequação ao cargo. O processo de participação docente "
        "não é, em nenhum de seus aspectos, um instrumento de avaliação de professores."
    )
    add_bullet(doc, "Não para a construção de rankings, comparações ou classificações de desempenho individual ou coletivo;")
    add_bullet(doc, "Não para justificar decisões disciplinares, administrativas ou de gestão de pessoal;")
    add_bullet(doc, "Não para tomar decisões curriculares ou pedagógicas de forma unilateral, sem o devido processo de construção coletiva;")
    add_bullet(doc, "Não para medição de performance ou auditoria de práticas pedagógicas;")
    add_bullet(doc, "Não para compartilhamento com instâncias externas à instituição sem o consentimento explícito dos participantes.")

    add_note_block(doc,
        "A participação dos professores neste processo não implica sua concordância com qualquer "
        "encaminhamento específico — ela é, exclusivamente, uma contribuição a um processo de "
        "reflexão coletiva."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 8 — TEMÁTICAS DA ESCUTA DOCENTE
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "8", "Temáticas da Escuta Docente")

    add_body(doc,
        "As temáticas a seguir representam os campos de reflexão que o processo de participação "
        "docente explorará. Cada uma delas foi selecionada por sua relevância para uma compreensão "
        "abrangente e fundamentada do Curso Técnico em Marketing e por sua capacidade de reunir "
        "contribuições significativas dos professores. As temáticas são apresentadas como áreas de "
        "reflexão — não como perguntas fechadas, não como categorias de avaliação e não como pautas "
        "com respostas esperadas."
    )

    themes = [
        ("Identidade e Proposta Formativa do Curso",
         "A identidade de um curso técnico é construída ao longo do tempo, a partir das escolhas "
         "pedagógicas de seus professores, das características dos estudantes que nele ingressam "
         "e das relações que o curso estabelece com o contexto social e econômico em que está "
         "inserido. Refletir sobre a identidade atual do curso — o que ele representa, o que o "
         "distingue e quais valores formativos ele expressa — é uma das temáticas centrais deste "
         "processo."),
        ("Perfil dos Estudantes",
         "Os professores são os observadores mais qualificados do perfil dos estudantes que chegam "
         "ao curso e de como esse perfil se transforma ao longo da formação. Suas percepções sobre "
         "as características, os interesses, as dificuldades e as potencialidades dos estudantes "
         "são um recurso informacional de alto valor para qualquer reflexão sobre a formação "
         "oferecida."),
        ("Práticas Pedagógicas e Experiências de Aprendizagem",
         "As estratégias pedagógicas que cada professor desenvolve e as experiências de aprendizagem "
         "que organiza para seus estudantes são expressões do seu conhecimento profissional em ação. "
         "As reflexões sobre o que funciona bem, o que inspira os estudantes e o que parece gerar "
         "engajamento e aprendizagem significativa constituem uma das contribuições mais valiosas "
         "que os professores podem oferecer a este processo."),
        ("Mercado de Trabalho e Inserção Profissional",
         "A percepção dos professores sobre como o mercado de trabalho regional e nacional valoriza "
         "a formação técnica em marketing — e sobre como os egressos do curso têm se inserido "
         "profissionalmente — é uma dimensão importante da reflexão sobre a qualidade e a "
         "pertinência da formação oferecida."),
        ("Tecnologia e Ferramentas Digitais no Contexto da Formação",
         "A relação entre tecnologia e formação em marketing atravessa múltiplas dimensões da "
         "experiência docente — da familiaridade dos estudantes com ferramentas digitais ao "
         "repertório metodológico disponível para os professores, passando pelas expectativas do "
         "mercado de trabalho sobre as competências técnicas dos egressos."),
        ("Inteligência Artificial: Presença, Possibilidades e Limites",
         "A inteligência artificial tornou-se uma realidade presente na prática profissional de "
         "marketing e, progressivamente, também nas experiências de aprendizagem dos estudantes. "
         "Como os professores percebem essa presença, quais possibilidades e quais limites "
         "identificam e como compreendem o papel da escola na mediação desse tema são questões "
         "que merecem espaço privilegiado na escuta institucional."),
        ("Conexões com o Território e a Comunidade Local",
         "Urubici oferece um contexto econômico, cultural e social rico para a formação em "
         "marketing. As percepções dos professores sobre as conexões entre o curso e o território "
         "— as oportunidades que o contexto local oferece para a aprendizagem e para a inserção "
         "profissional dos egressos — representam uma temática de alto valor para reflexões sobre "
         "o futuro da formação."),
        ("Oportunidades Identificadas pelos Professores",
         "As oportunidades que os próprios professores identificam para o fortalecimento do curso "
         "— a partir de sua experiência prática, de suas observações sobre os estudantes e de seu "
         "conhecimento do contexto educacional — constituem talvez a temática mais diretamente "
         "orientada à construção de reflexões futuras. É aqui que o conhecimento acumulado dos "
         "professores se traduz mais claramente em contribuição prospectiva."),
    ]

    for theme_name, theme_text in themes:
        add_body_bold_intro(doc, f"{theme_name}:  ", theme_text)
        add_spacer(doc, 4)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 9 — BENEFÍCIOS ESPERADOS
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "9", "Benefícios Esperados")

    add_body(doc,
        "Uma iniciativa institucional de participação docente que seja conduzida com respeito, "
        "seriedade e compromisso genuíno com a colaboração tem o potencial de gerar benefícios "
        "concretos e duradouros para todos os que fazem parte da comunidade educacional. Os "
        "benefícios descritos a seguir são expectativas fundamentadas — não promessas, não "
        "garantias, mas horizontes possíveis quando um processo é conduzido com integridade e "
        "cuidado."
    )

    add_subsection(doc, "9.1 Para os Estudantes")
    add_body(doc,
        "Os estudantes são, em última instância, os principais beneficiários de qualquer iniciativa "
        "que fortaleça a qualidade da formação oferecida pelo Curso Técnico em Marketing. Um "
        "processo de escuta docente que produz reflexões qualificadas contribui para uma formação "
        "mais contextualizada e pertinente — que reconhece quem são os estudantes de hoje e o que "
        "eles precisarão saber e saber fazer para atuar com competência em um mercado de trabalho "
        "em transformação."
    )
    add_bullet(doc, "Formação mais contextualizada e pertinente ao mercado de trabalho atual e à realidade local;")
    add_bullet(doc, "Maior engajamento nas experiências de aprendizagem, percebendo conexão entre o que estudam e sua aplicação real;")
    add_bullet(doc, "Desenvolvimento de competências valorizadas pelo mercado regional e nacional, documentadas em portfólio profissional;")
    add_bullet(doc, "Familiaridade ética e crítica com as ferramentas digitais que definirão o mercado da próxima década;")
    add_bullet(doc, "Vínculo mais profundo com o território de Urubici e suas oportunidades econômicas e culturais.")

    add_subsection(doc, "9.2 Para os Professores")
    add_body(doc,
        "Os professores que participam deste processo têm a oportunidade de vivenciar, de forma "
        "concreta, o reconhecimento institucional de seu conhecimento e de sua trajetória. Ser "
        "chamado a contribuir — não como executor de uma decisão já tomada, mas como construtor "
        "de reflexões que ainda estão por ser elaboradas — é uma forma de reconhecimento "
        "profissional que vai além das formas tradicionais de valorização docente."
    )
    add_bullet(doc, "Reconhecimento formal do conhecimento e da trajetória profissional como patrimônio institucional;")
    add_bullet(doc, "Oportunidade de diálogo qualificado com colegas sobre temas que raramente têm espaço no cotidiano escolar;")
    add_bullet(doc, "Participação ativa na construção das reflexões que poderão orientar decisões futuras sobre o curso;")
    add_bullet(doc, "Fortalecimento do senso de comunidade profissional entre os docentes do curso;")
    add_bullet(doc, "Registro formal da contribuição individual à memória institucional da escola.")

    add_subsection(doc, "9.3 Para a Instituição")
    add_body(doc,
        "Para a EEB Araújo Figueiredo, este processo representa uma oportunidade de construir, "
        "de forma organizada e documentada, um corpus de conhecimento sobre o seu Curso Técnico "
        "em Marketing que apenas os professores possuem. Esse conhecimento, sistematizado e "
        "preservado, constitui um patrimônio institucional — uma memória viva da inteligência "
        "pedagógica que a escola desenvolveu ao longo do tempo."
    )
    add_bullet(doc, "Construção de memória institucional qualificada sobre o curso, preservada em registros formais;")
    add_bullet(doc, "Fortalecimento da cultura institucional de colaboração e valorização do conhecimento docente;")
    add_bullet(doc, "Fortalecimento da confiança entre docentes e gestão — base de qualquer cultura organizacional saudável;")
    add_bullet(doc, "Posicionamento da escola como referência regional em processos pedagógicos participativos e responsáveis.")

    add_subsection(doc, "9.4 Para a Comunidade")
    add_body(doc,
        "A comunidade de Urubici e região se beneficia, de forma indireta mas significativa, de "
        "um processo que fortalece a qualidade da formação técnica em marketing oferecida pela "
        "EEB Araújo Figueiredo. Profissionais formados com qualidade, pertinência e conexão ao "
        "contexto local têm maior probabilidade de permanecer ou retornar à região, contribuindo "
        "para o desenvolvimento econômico e cultural do território."
    )
    add_bullet(doc, "Profissionais de marketing formados localmente com capacidade de apoiar o desenvolvimento econômico regional;")
    add_bullet(doc, "Jovens que permanecem ou retornam à região por encontrarem oportunidades profissionais conectadas à sua formação;")
    add_bullet(doc, "Pequenos negócios e empreendedores locais com acesso progressivo a profissionais qualificados para comunicação digital;")
    add_bullet(doc, "Valorização da identidade cultural e econômica local por meio de projetos educativos que tomam Urubici como território de aprendizagem;")
    add_bullet(doc, "Fortalecimento do papel social da escola como instituição que contribui ativamente para o desenvolvimento humano e econômico da comunidade.")

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 10 — PRÓXIMOS PASSOS INSTITUCIONAIS
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "10", "Próximos Passos Institucionais")

    add_body(doc,
        "O processo de participação docente que este documento formaliza organiza-se em uma "
        "sequência de etapas institucionais que serão conduzidas com cuidado e transparência. "
        "Essa sequência não constitui um plano de implementação com prazos e responsabilidades "
        "fixadas — constitui, antes, uma trajetória de construção coletiva cujo ritmo será "
        "determinado pela qualidade do processo e não pela urgência de resultados."
    )

    add_subsection(doc, "Participação Docente")
    add_body(doc,
        "A etapa inicial consiste no envolvimento de cada professor por meio do instrumento "
        "institucional descrito anteriormente. Cada professor será formalmente convidado a "
        "participar, com informação clara sobre os objetivos do processo, sobre a natureza da "
        "participação esperada e sobre os compromissos que a instituição assume em relação ao "
        "uso das contribuições. A participação é uma forma de contribuição ao coletivo — "
        "valorizada e esperada, mas exercida com liberdade e autonomia por cada professor."
    )

    add_subsection(doc, "Consolidação e Organização das Contribuições")
    add_body(doc,
        "As contribuições coletadas serão organizadas com cuidado pela equipe pedagógica "
        "responsável pelo processo. Esse trabalho de consolidação buscará preservar a riqueza "
        "e a diversidade das perspectivas individuais, identificando convergências e registrando "
        "com igual atenção as perspectivas que se distinguem da visão majoritária. A síntese "
        "elaborada será compartilhada com os professores como parte do compromisso de "
        "transparência que orienta o processo."
    )

    add_subsection(doc, "Reflexão Coletiva em Espaços Institucionais")
    add_body(doc,
        "A partir da consolidação das contribuições individuais, serão organizados espaços de "
        "reflexão coletiva nos quais os professores poderão aprofundar as discussões, complementar "
        "as perspectivas individuais e construir uma compreensão compartilhada sobre os temas em "
        "reflexão. Esses espaços serão organizados de forma a garantir a participação equitativa "
        "de todos e a qualidade do diálogo entre pares."
    )

    add_subsection(doc, "Construção de Subsídios para Discussões Futuras")
    add_body(doc,
        "O resultado das etapas anteriores será sistematizado em subsídios institucionais que "
        "apoiarão as discussões futuras sobre o fortalecimento do curso. Esses subsídios não são "
        "decisões — são contribuições qualificadas ao processo institucional de reflexão, "
        "construídas coletivamente e fundamentadas no conhecimento dos professores que vivem o "
        "curso cotidianamente."
    )
    add_body(doc,
        "As decisões sobre o futuro do curso — quando forem necessárias e quando o processo de "
        "reflexão coletiva tiver produzido as condições para que sejam tomadas com responsabilidade "
        "— serão construídas nessas instâncias coletivas, com a participação informada e ativa "
        "dos professores. Nenhuma decisão relevante sobre a formação oferecida pelo Curso Técnico "
        "em Marketing será tomada sem esse processo."
    )

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 11 — MENSAGEM INSTITUCIONAL DE ENCERRAMENTO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "11", "Mensagem Institucional de Encerramento")

    add_body(doc,
        "O Curso Técnico em Marketing da EEB Araújo Figueiredo é mais do que uma oferta de "
        "formação profissional — é uma expressão do compromisso desta instituição com o futuro "
        "dos jovens que vivem e crescem neste território. Por detrás de cada disciplina, de cada "
        "aula e de cada estudante que conclui a formação, há o trabalho silencioso, persistente "
        "e muitas vezes invisível de professores que escolheram a educação como vocação e que a "
        "exercem com seriedade e dedicação."
    )
    add_body(doc,
        "Este documento nasce do reconhecimento desse trabalho. Não como formalidade — mas como "
        "convicção. O futuro de qualquer curso de qualidade não se constrói pela substituição de "
        "sua história, mas pelo seu aprofundamento. A experiência acumulada pelos professores desta "
        "instituição não é um obstáculo à evolução — é o solo sobre o qual a evolução se torna "
        "possível. É o conhecimento sobre os estudantes que chegam, sobre o que os faz aprender, "
        "sobre o que os prepara para o mundo, que torna qualquer reflexão sobre o futuro relevante "
        "e responsável."
    )
    add_body(doc,
        "Convidamos cada professor a participar deste processo não como cumprimento de uma "
        "obrigação institucional, mas como exercício de um papel que é genuinamente seu: o de "
        "co-autor do futuro da formação que esta escola oferece. Sua perspectiva importa. Sua "
        "experiência é insubstituível. Sua contribuição fará diferença."
    )
    add_body(doc,
        "Os estudantes que chegam ao Curso Técnico em Marketing chegam com sonhos, com expectativas "
        "e com a confiança de que a escola — e seus professores — os ajudarão a encontrar seu "
        "caminho. Essa confiança é um compromisso que a instituição assume coletivamente. E é a "
        "lembrança desse compromisso que torna este processo não apenas importante, mas necessário."
    )
    add_highlight_block(doc, "",
        "Construiremos juntos. No tempo que for necessário. "
        "Com o cuidado que a formação dos jovens desta comunidade merece.",
        fill=C["gold_soft"], border_color=C["gold"])

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
    print(f"Documento gerado: {OUTPUT_DOCX}")


if __name__ == "__main__":
    build_document()
