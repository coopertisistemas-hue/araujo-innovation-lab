"""
Gerador do documento institucional:
Instrumento Institucional de Participação Docente
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
OUTPUT_DOCX = OUT_DIR / "teacher-participation-instrument-v1.docx"


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
    "resp":       "F4F7F5",
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


def add_spacer(doc: Document, pts: int = 8) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(pts)
    p.paragraph_format.space_after = Pt(0)


# ── Primitivos específicos para o instrumento ─────────────────────────────────

def add_block_intro(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(text)
    run_style(r, "Arial", 10, C["muted"], italic=True)


def add_question(doc: Document, number: str, text: str,
                 optional: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(6)
    r_num = p.add_run(f"{number}.  ")
    run_style(r_num, "Arial", 11, C["gold"], bold=True)
    r_text = p.add_run(text)
    run_style(r_text, "Arial", 11, C["ink"], bold=True)
    if optional:
        r_opt = p.add_run("  (opcional)")
        run_style(r_opt, "Arial", 10, C["muted"], italic=True)
    para_bottom_border(p, color=C["line"], size="3")


def add_mcq_option(doc: Document, text: str, last: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3) if not last else Pt(12)
    r = p.add_run("○   " + text)   # ○ WHITE CIRCLE
    run_style(r, "Arial", 11, C["body"])


def add_checkbox_option(doc: Document, text: str, last: bool = False) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3) if not last else Pt(12)
    r = p.add_run("☐   " + text)   # ☐ BALLOT BOX
    run_style(r, "Arial", 11, C["body"])


def add_other_mcq(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("○   Outro:  _____________________________________________")
    run_style(r, "Arial", 11, C["body"])


def add_other_checkbox(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("☐   Outro:  _____________________________________________")
    run_style(r, "Arial", 11, C["body"])


def add_short_text_field(doc: Document) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("_______________________________________________________________")
    run_style(r, "Arial", 11, C["body"])


def add_response_area(doc: Document, lines: int = 5) -> None:
    """Shaded response box with ruled lines for open-ended answers."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = table.rows[0].cells[0]
    cell_shading(cell, C["resp"])
    cell_borders(cell, color=C["line"], size="8")

    # Label
    p_label = cell.paragraphs[0]
    p_label.paragraph_format.space_before = Pt(6)
    p_label.paragraph_format.space_after = Pt(6)
    r_label = p_label.add_run("Sua resposta:")
    run_style(r_label, "Arial", 9, C["muted"], italic=True)

    # Response lines
    for i in range(lines):
        p_line = cell.add_paragraph()
        p_line.paragraph_format.space_before = Pt(0)
        p_line.paragraph_format.space_after = Pt(6)
        r_line = p_line.add_run("")
        run_style(r_line, "Arial", 11, C["body"])
        para_bottom_border(p_line, color=C["line"], size="4")

    # Bottom padding
    p_pad = cell.add_paragraph()
    p_pad.paragraph_format.space_before = Pt(4)
    p_pad.paragraph_format.space_after = Pt(4)

    doc.add_paragraph()


# ── Tabela de identificação ───────────────────────────────────────────────────

def add_metadata_table(doc: Document) -> None:
    rows_data = [
        ("Nome do Documento",
         "Instrumento Institucional de Participação Docente para o Fortalecimento "
         "do Curso Técnico em Marketing"),
        ("Versão",         "v1.0"),
        ("Data",           "Junho de 2026"),
        ("Status",         "Documento preparatório para participação docente institucional"),
        ("Elaborado por",  "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários",  "Corpo Docente do Curso Técnico em Marketing"),
        ("Classificação",  "Documento Estratégico Institucional"),
        ("Instituição",    "EEB Araújo Figueiredo — Araujo Innovation Lab — Urubici, SC"),
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


# ── Bloco de créditos institucionais ─────────────────────────────────────────

def add_credits_block(doc: Document) -> None:
    p_heading = doc.add_paragraph()
    p_heading.paragraph_format.space_before = Pt(16)
    p_heading.paragraph_format.space_after = Pt(8)
    r_h = p_heading.add_run("EQUIPE INSTITUCIONAL")
    run_style(r_h, "Georgia", 11, C["blue"], bold=True)
    para_bottom_border(p_heading, color=C["line"], size="4")

    credits_rows = [
        ("Coordenação Técnica",
         "Profª Letícia Meurer",
         "Docente do Curso Técnico em Marketing"),
        ("",
         "José Alexandre",
         "Professor Orientador — Araujo Innovation Lab"),
        ("Apoio Pedagógico",
         "Suzane",
         "Orientadora Pedagógica"),
        ("Apoio Institucional",
         "Aline Liz",
         "Secretária Escolar"),
        ("Instituição",
         "EEB Araújo Figueiredo",
         "Urubici – SC"),
    ]

    table = doc.add_table(rows=len(credits_rows), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, (role, name, desc) in enumerate(credits_rows):
        row = table.rows[i]
        row.cells[0].width = Inches(2.1)
        row.cells[1].width = Inches(4.2)

        fill_role = C["blue_soft"] if role else C["blue_soft2"]
        cell_shading(row.cells[0], fill_role)

        p_role = row.cells[0].paragraphs[0]
        p_role.paragraph_format.space_before = Pt(5)
        p_role.paragraph_format.space_after = Pt(5)
        if role:
            r_role = p_role.add_run(role)
            run_style(r_role, "Arial", 9, C["blue"], bold=True)

        p_name = row.cells[1].paragraphs[0]
        p_name.paragraph_format.space_before = Pt(5)
        p_name.paragraph_format.space_after = Pt(5)
        r_name = p_name.add_run(name)
        run_style(r_name, "Arial", 10, C["ink"], bold=True)
        r_desc = p_name.add_run(f"\n{desc}")
        run_style(r_desc, "Arial", 9, C["body"])

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
    tag_run = tag_p.add_run("DOCUMENTO ESTRATÉGICO INSTITUCIONAL — PARTICIPAÇÃO DOCENTE")
    run_style(tag_run, "Arial", 8, C["muted"], bold=True)

    # Título principal
    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_before = Pt(28)
    title.paragraph_format.space_after = Pt(14)
    title.add_run(
        "Instrumento Institucional de Participação Docente "
        "para o Fortalecimento do Curso Técnico em Marketing"
    )

    # Subtítulo
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(36)
    r_sub = sub.add_run(
        "Instrumento de escuta qualificada para coleta de perspectivas, experiências e "
        "reflexões do corpo docente sobre o Curso Técnico em Marketing da "
        "EEB Araújo Figueiredo."
    )
    run_style(r_sub, "Arial", 12, C["muted"])

    # Bloco de identificação na capa
    id_table = doc.add_table(rows=5, cols=2)
    id_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    id_rows = [
        ("Versão",        "v1.0 — Primeira Emissão Oficial"),
        ("Data",          "Junho de 2026"),
        ("Classificação", "Documento Estratégico Institucional"),
        ("Status",        "Documento preparatório para participação docente"),
        ("Destinatários", "Corpo Docente — EEB Araújo Figueiredo"),
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

    # Assinatura institucional
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(55)
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


# ── Bloco de encerramento ─────────────────────────────────────────────────────

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
    r1 = end_p.add_run("Este instrumento não é uma ferramenta de avaliação docente.")
    run_style(r1, "Arial", 10, C["white"])

    end_p2 = end_cell.add_paragraph()
    end_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p2.paragraph_format.space_before = Pt(2)
    end_p2.paragraph_format.space_after = Pt(4)
    r2 = end_p2.add_run(
        "É um espaço de contribuição profissional para a construção coletiva do futuro do curso."
    )
    run_style(r2, "Arial", 10, C["gold"], bold=True)

    end_p3 = end_cell.add_paragraph()
    end_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p3.paragraph_format.space_before = Pt(4)
    end_p3.paragraph_format.space_after = Pt(4)
    r3 = end_p3.add_run(
        "As respostas individuais serão tratadas com absoluta confidencialidade "
        "e utilizadas exclusivamente para os fins institucionais descritos neste documento."
    )
    run_style(r3, "Arial", 10, C["white"], italic=True)

    end_p4 = end_cell.add_paragraph()
    end_p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p4.paragraph_format.space_before = Pt(8)
    end_p4.paragraph_format.space_after = Pt(10)
    r4 = end_p4.add_run("Araujo Innovation Lab — EEB Araújo Figueiredo — Urubici, SC  ·  Junho de 2026")
    run_style(r4, "Arial", 8, "D7DEDA", italic=True)


# ── Bloco de governança ───────────────────────────────────────────────────────

def add_governance_block(doc: Document) -> None:
    doc.add_page_break()

    p_gov = doc.add_paragraph()
    p_gov.paragraph_format.space_before = Pt(14)
    p_gov.paragraph_format.space_after = Pt(10)
    r_gov = p_gov.add_run("GOVERNANÇA DOCUMENTAL")
    run_style(r_gov, "Georgia", 14, C["blue"], bold=True)
    para_bottom_border(p_gov, color=C["line"], size="4")

    rows_data = [
        ("Documento-fonte institucional",  "teacher-participation-instrument-v1.md"),
        ("Documento derivado",             "teacher-participation-instrument-v1.docx"),
        ("Padrão — documentos",            "document-template-standard.md"),
        ("Padrão — identidade",            "branding-standard.md"),
        ("Classificação",                  "Documento Estratégico Institucional"),
        ("Uso",                            "Participação Docente Institucional"),
        ("Versão",                         "v1.0"),
        ("Instituição",                    "EEB Araújo Figueiredo — Araujo Innovation Lab"),
        ("Local",                          "Urubici – SC"),
        ("Status",                         "Documento Institucional Publicado"),
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
        header_text="Araujo Innovation Lab  |  Instrumento de Participação Docente — Curso Técnico em Marketing",
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

    # ── Créditos Institucionais ───────────────────────────────────────────────
    add_credits_block(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 1 — APRESENTAÇÃO INSTITUCIONAL
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "1", "Apresentação Institucional")

    add_body(doc,
        "Este instrumento foi elaborado com o propósito de convidar cada professor do Curso "
        "Técnico em Marketing a contribuir com suas percepções, experiências e reflexões em "
        "apoio a um processo institucional de escuta qualificada. Trata-se de uma iniciativa "
        "fundamentada no reconhecimento de que o conhecimento acumulado pelos educadores que "
        "constroem este curso é insubstituível — e que qualquer reflexão sobre o futuro da "
        "formação oferecida precisa, necessariamente, partir das vozes de quem o vive."
    )
    add_body(doc,
        "O processo que este instrumento inaugura não é avaliativo. Não há, em nenhuma das "
        "perguntas que seguem, qualquer intenção de mensurar desempenho, comparar práticas "
        "ou identificar deficiências. O que se busca é precisamente o oposto: reunir, com "
        "cuidado e respeito, as perspectivas que apenas quem está dentro do curso pode oferecer "
        "— sobre seus pontos fortes, sobre os estudantes que formam, sobre as experiências que "
        "funcionam, sobre as oportunidades que o futuro pode trazer."
    )
    add_body(doc,
        "As contribuições que cada professor oferecer integrarão um conjunto de subsídios que "
        "servirão de base para reflexões coletivas futuras sobre o fortalecimento do curso. O "
        "estudante é o centro desse processo. A qualidade da formação que ele recebe é a razão "
        "pela qual este instrumento existe."
    )
    add_highlight_block(doc, "Compromisso institucional",
        "Este instrumento não é uma ferramenta de avaliação docente. É um espaço de contribuição "
        "profissional.",
        fill=C["blue_soft2"], border_color=C["blue"])

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 2 — ORIENTAÇÕES PARA PREENCHIMENTO
    # ═══════════════════════════════════════════════════════════════════════════
    add_section_heading(doc, "2", "Orientações para Preenchimento")

    add_body(doc, "Tempo estimado: 15 a 20 minutos.")
    add_spacer(doc, 4)
    add_body(doc,
        "Sobre as respostas: Não há respostas certas ou erradas. O que importa é a autenticidade "
        "da sua perspectiva. Respostas breves e diretas são tão valiosas quanto respostas mais "
        "extensas — o que conta é que reflitam genuinamente sua experiência e seu olhar sobre "
        "o curso."
    )
    add_body(doc,
        "Sobre a confidencialidade: As contribuições individuais serão tratadas com discrição. "
        "O nome é opcional. As informações compartilhadas serão utilizadas exclusivamente para "
        "os fins descritos na Apresentação Institucional — nunca para fins avaliativos, "
        "disciplinares ou de gestão de pessoal."
    )
    add_body(doc,
        "Sobre o processo: Ao responder este instrumento, você está contribuindo para a "
        "construção de uma compreensão coletiva sobre o Curso Técnico em Marketing. Sua "
        "perspectiva será incorporada a uma síntese institucional que será compartilhada com "
        "todos os participantes ao final do processo."
    )
    add_spacer(doc, 6)

    p_instr = doc.add_paragraph()
    p_instr.paragraph_format.space_before = Pt(4)
    p_instr.paragraph_format.space_after = Pt(4)
    r_instr = p_instr.add_run("Instruções de preenchimento:")
    run_style(r_instr, "Arial", 11, C["blue"], bold=True)

    add_bullet(doc, "Nas questões de escolha única, marque apenas uma opção: ○")
    add_bullet(doc, "Nas questões de múltipla escolha, marque todas as opções que considerar pertinentes: ☐")
    add_bullet(doc, "Nas questões abertas, escreva livremente no espaço indicado.")
    add_bullet(doc, "Questões marcadas como (opcional) podem ser deixadas em branco.")

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 3 — BLOCO A — PERFIL PROFISSIONAL
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "3", "Bloco A — Perfil Profissional")
    add_block_intro(doc,
        "Este bloco tem como propósito contextualizar as contribuições oferecidas. "
        "Todas as informações são tratadas com confidencialidade."
    )

    # A.1 — Nome (opcional)
    add_question(doc, "A.1", "Nome", optional=True)
    add_short_text_field(doc)

    # A.2 — Formação acadêmica
    add_question(doc, "A.2",
        "Qual é o seu maior nível de formação acadêmica concluído?")
    add_mcq_option(doc, "Graduação")
    add_mcq_option(doc, "Especialização / Pós-graduação lato sensu")
    add_mcq_option(doc, "Mestrado")
    add_mcq_option(doc, "Doutorado")
    add_other_mcq(doc)

    # A.3 — Área e disciplinas
    add_question(doc, "A.3",
        "Qual é a sua área principal de formação acadêmica? Quais disciplinas você "
        "ministra ou já ministrou no Curso Técnico em Marketing?")
    add_response_area(doc, lines=3)

    # A.4 — Tempo de docência
    add_question(doc, "A.4",
        "Há quantos anos você atua como docente na educação básica ou técnica?")
    add_mcq_option(doc, "Menos de 3 anos")
    add_mcq_option(doc, "De 3 a 7 anos")
    add_mcq_option(doc, "De 8 a 15 anos")
    add_mcq_option(doc, "Mais de 15 anos", last=True)

    # A.5 — Tempo no curso
    add_question(doc, "A.5",
        "Há quanto tempo você atua especificamente no Curso Técnico em Marketing "
        "desta instituição?")
    add_mcq_option(doc, "Menos de 1 ano")
    add_mcq_option(doc, "De 1 a 3 anos")
    add_mcq_option(doc, "De 4 a 7 anos")
    add_mcq_option(doc, "Mais de 7 anos", last=True)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 4 — BLOCO B — IDENTIDADE E POTENCIALIDADES DO CURSO
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "4", "Bloco B — Identidade e Potencialidades do Curso")
    add_block_intro(doc,
        "Este bloco convida você a compartilhar sua percepção sobre o que torna este curso "
        "significativo e valioso. Suas respostas ajudarão a preservar e valorizar o que já "
        "funciona bem."
    )

    # B.1
    add_question(doc, "B.1",
        "Em sua percepção, quais são os pontos mais fortes do Curso Técnico em Marketing "
        "da EEB Araújo Figueiredo? O que você considera que este curso faz especialmente bem?")
    add_response_area(doc, lines=5)

    # B.2
    add_question(doc, "B.2",
        "Que aspectos da proposta atual do curso você considera mais importantes preservar "
        "e valorizar em qualquer processo de reflexão sobre o seu futuro?")
    add_response_area(doc, lines=5)

    # B.3
    add_question(doc, "B.3",
        "O que você acredita que distingue este curso e o torna especial para os estudantes "
        "que o frequentam? Que características tornam esta experiência de formação única?")
    add_response_area(doc, lines=5)

    # B.4
    add_question(doc, "B.4",
        "Das características a seguir, quais você considera mais presentes e valiosas no "
        "Curso Técnico em Marketing desta instituição? "
        "(selecione todas que considerar pertinentes)")
    add_checkbox_option(doc, "Qualidade e comprometimento do corpo docente")
    add_checkbox_option(doc, "Conexão com a realidade local de Urubici e região")
    add_checkbox_option(doc, "Relevância dos conteúdos para o mercado de trabalho")
    add_checkbox_option(doc, "Abertura para a criatividade e a expressão dos estudantes")
    add_checkbox_option(doc, "Ambiente de aprendizagem acolhedor e respeitoso")
    add_checkbox_option(doc, "Diversidade disciplinar e caráter multidisciplinar do curso")
    add_checkbox_option(doc, "Comprometimento dos estudantes com sua própria formação")
    add_checkbox_option(doc, "Vínculo entre a escola e a comunidade local")
    add_checkbox_option(doc, "Formação de caráter prático e aplicado")
    add_other_checkbox(doc)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 5 — BLOCO C — PERFIL DOS ESTUDANTES
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "5", "Bloco C — Perfil dos Estudantes")
    add_block_intro(doc,
        "Você é um dos observadores mais qualificados da realidade dos estudantes deste curso. "
        "Este bloco busca reunir suas percepções sobre quem são esses jovens e o que os "
        "caracteriza."
    )

    # C.1
    add_question(doc, "C.1",
        "Como você descreveria o perfil dos estudantes que chegam ao Curso Técnico em "
        "Marketing? Quais são suas características, interesses e motivações mais frequentes?")
    add_response_area(doc, lines=5)

    # C.2
    add_question(doc, "C.2",
        "Quais potencialidades, talentos e pontos fortes você identifica com mais frequência "
        "nos estudantes deste curso? O que eles trazem de valioso para a experiência de "
        "aprendizagem?")
    add_response_area(doc, lines=5)

    # C.3
    add_question(doc, "C.3",
        "Quais são os principais desafios que os estudantes enfrentam ao longo de sua "
        "formação no curso? Como você percebe essas dificuldades no contexto do aprendizado?")
    add_response_area(doc, lines=5)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 6 — BLOCO D — EXPERIÊNCIAS DE APRENDIZAGEM E PRÁTICAS PEDAGÓGICAS
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "6",
        "Bloco D — Experiências de Aprendizagem e Práticas Pedagógicas")
    add_block_intro(doc,
        "Este bloco valoriza o seu conhecimento pedagógico em ação. As reflexões que você "
        "compartilhar aqui representam uma das contribuições mais ricas que este processo "
        "pode reunir."
    )

    # D.1
    add_question(doc, "D.1",
        "Que tipos de atividades, experiências ou situações de aprendizagem você percebe "
        "que geram mais engajamento, curiosidade e motivação nos estudantes deste curso?")
    add_response_area(doc, lines=5)

    # D.2
    add_question(doc, "D.2",
        "Descreva uma experiência pedagógica — uma aula, um projeto, uma atividade ou um "
        "momento — que você considera particularmente significativa para o aprendizado dos "
        "estudantes. O que tornou essa experiência especial?")
    add_response_area(doc, lines=6)

    # D.3
    add_question(doc, "D.3",
        "Quais abordagens, estratégias ou práticas pedagógicas você considera mais eficazes "
        "no contexto específico do Curso Técnico em Marketing? O que, em sua experiência, "
        "faz diferença no aprendizado desses estudantes?")
    add_response_area(doc, lines=5)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 7 — BLOCO E — FORMAÇÃO PROFISSIONAL E MERCADO DE TRABALHO
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "7", "Bloco E — Formação Profissional e Mercado de Trabalho")
    add_block_intro(doc,
        "Este bloco busca compreender como você percebe a relação entre a formação oferecida "
        "pelo curso e as demandas do mundo profissional ao qual os estudantes serão "
        "encaminhados."
    )

    # E.1
    add_question(doc, "E.1",
        "Quais competências você considera mais essenciais para a formação do profissional "
        "de marketing nos dias atuais? Que capacidades um egresso deste curso deveria "
        "demonstrar ao concluir sua formação?")
    add_response_area(doc, lines=5)

    # E.2
    add_question(doc, "E.2",
        "Considerando as demandas atuais do mercado de trabalho para profissionais de "
        "marketing, quais das seguintes competências você considera prioritárias para o "
        "egresso deste curso? (selecione todas que considerar relevantes)")
    add_checkbox_option(doc, "Comunicação escrita e produção de conteúdo")
    add_checkbox_option(doc, "Gestão de presença digital e redes sociais")
    add_checkbox_option(doc, "Análise de dados e interpretação de métricas")
    add_checkbox_option(doc, "Criatividade, design e produção visual")
    add_checkbox_option(doc, "Empreendedorismo e mentalidade de negócios")
    add_checkbox_option(doc, "Fotografia e produção audiovisual")
    add_checkbox_option(doc, "Atendimento, relacionamento e gestão de clientes")
    add_checkbox_option(doc, "Marketing de conteúdo, narrativa e storytelling")
    add_checkbox_option(doc, "Planejamento estratégico de comunicação")
    add_checkbox_option(doc, "Uso ético e crítico de tecnologias, incluindo inteligência artificial")
    add_other_checkbox(doc)

    # E.3
    add_question(doc, "E.3",
        "Que perfil profissional você gostaria que os egressos deste curso apresentassem "
        "ao concluir sua formação? Como você descreveria o profissional que este curso "
        "aspira a formar?")
    add_response_area(doc, lines=5)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 8 — BLOCO F — TECNOLOGIA E INTELIGÊNCIA ARTIFICIAL
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "8", "Bloco F — Tecnologia e Inteligência Artificial")
    add_block_intro(doc,
        "Este bloco convida você a refletir sobre as transformações tecnológicas que "
        "atravessam tanto a prática profissional de marketing quanto a experiência "
        "educacional. Não há resposta esperada — o que importa é a sua perspectiva."
    )

    # F.1
    add_question(doc, "F.1",
        "Qual é o seu grau de familiaridade com ferramentas de inteligência artificial "
        "generativa — como assistentes de texto, geradores de imagem, ferramentas de "
        "análise automatizada, entre outras?")
    add_mcq_option(doc, "Ainda não utilizei ferramentas desse tipo")
    add_mcq_option(doc,
        "Já explorei de forma inicial, mas não as utilizo com regularidade")
    add_mcq_option(doc,
        "Utilizo com alguma regularidade, principalmente para uso pessoal ou profissional")
    add_mcq_option(doc,
        "Utilizo com frequência e já incorporei essas ferramentas à minha prática")
    add_mcq_option(doc,
        "Utilizo com frequência e tenho explorado seu potencial no contexto educacional",
        last=True)

    # F.2
    add_question(doc, "F.2",
        "Quais oportunidades e quais desafios você identifica para a formação em marketing "
        "diante do crescimento e da disseminação das ferramentas de inteligência artificial? "
        "Como você percebe esse cenário para os estudantes que você forma?")
    add_response_area(doc, lines=5)

    # F.3
    add_question(doc, "F.3",
        "Que considerações éticas você acredita que devem orientar o uso de inteligência "
        "artificial na comunicação, no marketing e na educação? Que responsabilidades você "
        "identifica nesse campo?")
    add_response_area(doc, lines=5)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 9 — BLOCO G — CONEXÃO COM URUBICI E A COMUNIDADE
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "9", "Bloco G — Conexão com Urubici e a Comunidade")
    add_block_intro(doc,
        "Urubici é o território em que este curso existe e onde os estudantes vivem. "
        "Este bloco convida você a refletir sobre as conexões entre a formação oferecida "
        "e o contexto local."
    )

    # G.1
    add_question(doc, "G.1",
        "Como você percebe as conexões entre o Curso Técnico em Marketing e o contexto "
        "econômico, cultural e comunitário de Urubici e região — com seu turismo, sua "
        "gastronomia, seu comércio, seu agronegócio e seu empreendedorismo?")
    add_response_area(doc, lines=5)

    # G.2
    add_question(doc, "G.2",
        "Que oportunidades oferecidas pelo território de Urubici você acredita que "
        "poderiam ser melhor exploradas como campo de aprendizagem para os estudantes do "
        "curso? Há experiências locais, negócios, eventos ou setores que você identifica "
        "como especialmente relevantes para a formação em marketing?")
    add_response_area(doc, lines=5)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 10 — BLOCO H — PROJETOS, IDEIAS E OPORTUNIDADES DE COLABORAÇÃO
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "10",
        "Bloco H — Projetos, Ideias e Oportunidades de Colaboração")
    add_block_intro(doc,
        "Este é o bloco mais prospectivo do instrumento. Aqui você é convidado a "
        "compartilhar suas ideias, seus interesses e as iniciativas que gostaria de ver "
        "ganhar vida. Suas respostas neste bloco têm potencial direto de inspirar os "
        "próximos passos institucionais."
    )

    # H.1
    add_question(doc, "H.1",
        "Há algum projeto, iniciativa ou experiência de aprendizagem que você gostaria "
        "de desenvolver com os estudantes do curso e que ainda não teve oportunidade de "
        "realizar? Como você imagina essa iniciativa?")
    add_response_area(doc, lines=5)

    # H.2
    add_question(doc, "H.2",
        "Que conhecimentos, experiências profissionais ou práticas pedagógicas que você "
        "acumulou ao longo de sua trajetória você gostaria de compartilhar de forma mais "
        "ampla com os estudantes deste curso?")
    add_response_area(doc, lines=5)

    # H.3
    add_question(doc, "H.3",
        "Que oportunidades de colaboração e diálogo entre as diferentes disciplinas e "
        "áreas do curso você identifica? Há conexões entre campos do conhecimento que "
        "poderiam ser mais exploradas em benefício da formação dos estudantes?")
    add_response_area(doc, lines=5)

    # H.4
    add_question(doc, "H.4",
        "Que tipos de parcerias com a comunidade local, com empresas, com organizações "
        "culturais ou com outras instituições educacionais você acredita que poderiam "
        "enriquecer a formação dos estudantes e fortalecer o curso?")
    add_response_area(doc, lines=5)

    # H.5
    add_question(doc, "H.5",
        "Há algo mais que você gostaria de compartilhar sobre o Curso Técnico em Marketing "
        "— uma ideia, uma percepção, uma reflexão ou uma sugestão — que não foi contemplado "
        "nas questões anteriores? Este espaço é inteiramente seu.")
    add_response_area(doc, lines=6)

    # ═══════════════════════════════════════════════════════════════════════════
    # SEÇÃO 11 — ENCERRAMENTO E AGRADECIMENTO
    # ═══════════════════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_section_heading(doc, "11", "Encerramento e Agradecimento")

    add_body(doc,
        "Ao chegar até aqui, você contribuiu com algo que nenhum documento, pesquisa ou "
        "referencial externo pode oferecer: o seu olhar. A perspectiva de quem conhece esses "
        "estudantes pelo nome, de quem percebe suas transformações ao longo da formação, de "
        "quem constrói o curso dia a dia com dedicação e profissionalismo."
    )
    add_body(doc,
        "Cada resposta que você compartilhou integra um processo de escuta que tem como único "
        "propósito fortalecer a qualidade da formação oferecida aos jovens desta comunidade. "
        "Suas percepções serão tratadas com o cuidado que merecem — organizadas, preservadas "
        "em sua diversidade e devolvidas ao coletivo docente em sínteses que serão "
        "compartilhadas com todos os participantes."
    )
    add_body(doc,
        "O futuro do Curso Técnico em Marketing da EEB Araújo Figueiredo será construído "
        "coletivamente — com o tempo necessário, com o respeito que a história deste curso "
        "exige e com a seriedade que os estudantes que nele confiam merecem. Você é parte "
        "essencial desse processo de construção."
    )
    add_highlight_block(doc, "",
        "Obrigado pela sua contribuição. Ela importa.",
        fill=C["gold_soft"], border_color=C["gold"])

    # ─────────────────────────────────────────────────────────────────────────
    add_closing_block(doc)
    add_governance_block(doc)

    # ── Salvar ────────────────────────────────────────────────────────────────
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f"Documento gerado: {OUTPUT_DOCX}")


if __name__ == "__main__":
    build_document()
