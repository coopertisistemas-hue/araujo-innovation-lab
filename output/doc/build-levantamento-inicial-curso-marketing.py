"""
Gerador do documento institucional:
Levantamento Inicial e Proposta de Organização do Curso Técnico em Marketing
Colégio Araújo — Araujo Innovation Lab
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[2]
OUT_DIR = ROOT / "output" / "doc"
OUTPUT_DOCX = OUT_DIR / "Colegio-Araujo-Levantamento-Inicial-Curso-Tecnico-Marketing-v0.1.docx"


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


def table_col_width(table, col_idx: int, width_inches: float) -> None:
    for row in table.rows:
        row.cells[col_idx].width = Inches(width_inches)


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
    contents = [
        "Documento de Trabalho | v0.1",
        "Página ",
        footer_right,
    ]
    aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT]
    for cell, text, align in zip(cells, contents, aligns):
        p2 = cell.paragraphs[0]
        p2.alignment = align
        run = p2.add_run(text)
        run_style(run, "Arial", 8, C["muted"])
        if text == "Página ":
            add_page_field(p2)
        no_borders(cell)


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
    run = p.add_run("ARAUJO INNOVATION LAB")
    run_style(run, "Arial", 9, C["blue"], bold=True)

    # Linha dourada
    bar = doc.add_paragraph()
    bar.paragraph_format.space_before = Pt(2)
    bar.paragraph_format.space_after = Pt(10)
    para_bottom_border(bar, color=C["gold"], size="16")

    # Subtítulo tipo
    tag_p = doc.add_paragraph()
    tag_p.paragraph_format.space_after = Pt(4)
    tag_run = tag_p.add_run("DOCUMENTO DE TRABALHO — VERSÃO INICIAL PARA ALINHAMENTO")
    run_style(tag_run, "Arial", 8, C["muted"], bold=True)

    # Título principal
    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_before = Pt(28)
    title.paragraph_format.space_after = Pt(12)
    title.add_run("Levantamento Inicial e Proposta de Organização do Curso Técnico em Marketing")

    # Subtítulo descritivo
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(32)
    r = sub.add_run("Documento preparatório para alinhamento interno entre gestão, orientação pedagógica e equipe docente.")
    run_style(r, "Arial", 12, C["muted"])

    # Bloco de identificação do documento (tabela)
    id_table = doc.add_table(rows=5, cols=2)
    id_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    id_rows = [
        ("Versão",        "0.1 — Versão Inicial"),
        ("Data",          "Junho de 2026"),
        ("Status",        "Em levantamento — nenhuma decisão finalizada"),
        ("Elaborado por", "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários", "Profª Letícia Meurer · Suzane · Aline Liz"),
    ]
    for i, (label, value) in enumerate(id_rows):
        row = id_table.rows[i]
        row.cells[0].width = Inches(2.0)
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
    spacer.paragraph_format.space_before = Pt(80)
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
    r_circ = circ.add_run("Circulação Interna Restrita")
    run_style(r_circ, "Arial", 9, C["muted"])


# ── Blocos de destaque ─────────────────────────────────────────────────────────

def add_highlight_block(doc: Document, label: str, text: str,
                        fill: str = "F1F5F3", border_color: str = "D7DEDA") -> None:
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell_shading(cell, fill)
    cell_borders(cell, color=border_color, size="12")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
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
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run("Nota importante:  ")
    run_style(r1, "Arial", 10, C["blue"], bold=True)
    r2 = p.add_run(text)
    run_style(r2, "Arial", 10, C["body"])
    doc.add_paragraph()


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


def add_sub_bullet(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.2)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"–  {text}")
    run_style(r, "Arial", 10, C["body"])


# ── Tabelas de conteúdo ────────────────────────────────────────────────────────

def add_section_heading(doc: Document, number: str, title: str) -> None:
    p = doc.add_paragraph(style="Heading 1")
    p.paragraph_format.space_before = Pt(22)
    p.clear()
    r_num = p.add_run(f"SEÇÃO {number} — ")
    run_style(r_num, "Georgia", 15, C["blue"], bold=True)
    r_title = p.add_run(title)
    run_style(r_title, "Georgia", 15, C["blue"], bold=True)
    para_bottom_border(p, color=C["line"], size="4")


def add_subsection(doc: Document, title: str) -> None:
    doc.add_paragraph(title, style="Heading 2")


def add_identification_table(doc: Document) -> None:
    rows_data = [
        ("Nome do Documento",
         "Levantamento Inicial e Proposta de Organização do Curso Técnico em Marketing"),
        ("Tipo",             "Documento de Trabalho — Versão Inicial para Alinhamento"),
        ("Versão",           "0.1"),
        ("Data",             "Junho de 2026"),
        ("Status",           "Em fase de levantamento e alinhamento inicial — nenhuma decisão finalizada"),
        ("Elaborado por",    "Profª Letícia Meurer · Professor Orientador José Alexandre"),
        ("Destinatários",    "Profª Letícia Meurer · Suzane (Orientadora Pedagógica) · Aline Liz (Secretária da Diretoria)"),
    ]
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for i, (label, value) in enumerate(rows_data):
        row = table.rows[i]
        row.cells[0].width = Inches(2.1)
        row.cells[1].width = Inches(4.1)
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


def add_participants_table(doc: Document, headers: list[str],
                           rows_data: list[list[str]]) -> None:
    col_count = len(headers)
    table = doc.add_table(rows=1 + len(rows_data), cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Cabeçalho
    header_row = table.rows[0]
    for j, h in enumerate(headers):
        cell_shading(header_row.cells[j], C["blue"])
        p = header_row.cells[j].paragraphs[0]
        r = p.add_run(h)
        run_style(r, "Arial", 10, C["white"], bold=True)
        cell_borders(header_row.cells[j], color=C["blue"], size="6")

    # Dados
    for i, row_data in enumerate(rows_data):
        row = table.rows[i + 1]
        fill = C["blue_soft2"] if i % 2 == 0 else C["panel"]
        for j, value in enumerate(row_data):
            cell_shading(row.cells[j], fill)
            p = row.cells[j].paragraphs[0]
            r = p.add_run(value)
            run_style(r, "Arial", 10, C["body"])
            cell_borders(row.cells[j], color=C["line"], size="6")
            row.cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph()


def add_next_steps_table(doc: Document) -> None:
    headers = ["Etapa", "Descrição", "Responsável Sugerido"]
    rows_data = [
        ("1. Validar o levantamento inicial",
         "Confirmar com os destinatários a precisão das informações sobre participantes e estrutura do curso",
         "Profª Letícia Meurer · Suzane · Aline Liz"),
        ("2. Identificar os participantes pendentes",
         "Localizar os contatos não identificados e completar a lista de participantes",
         "Aline Liz · Profª Letícia Meurer"),
        ("3. Confirmar as disciplinas vinculadas ao curso",
         "Mapear o conjunto completo de disciplinas da grade curricular e seus respectivos responsáveis",
         "Profª Letícia Meurer · Suzane"),
        ("4. Realizar reuniões individuais com os docentes",
         "Conduzir encontros exploratórios com cada professor identificado",
         "Prof. José Alexandre · Profª Letícia Meurer"),
        ("5. Consolidar o mapa das disciplinas",
         "Organizar as informações coletadas em um documento de referência para o planejamento",
         "Prof. José Alexandre"),
        ("6. Estruturar o planejamento do Curso",
         "Com base no levantamento consolidado, elaborar proposta de planejamento para apresentação institucional",
         "Profª Letícia Meurer · Prof. José Alexandre · Suzane"),
    ]
    table = doc.add_table(rows=1 + len(rows_data), cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    header_row = table.rows[0]
    for j, h in enumerate(headers):
        cell_shading(header_row.cells[j], C["blue"])
        p = header_row.cells[j].paragraphs[0]
        r = p.add_run(h)
        run_style(r, "Arial", 10, C["white"], bold=True)
        cell_borders(header_row.cells[j], color=C["blue"], size="6")

    for i, (etapa, descricao, responsavel) in enumerate(rows_data):
        row = table.rows[i + 1]
        fill = C["blue_soft2"] if i % 2 == 0 else C["panel"]
        values = [etapa, descricao, responsavel]
        col_fills = [C["blue_soft"], fill, fill]
        for j, val in enumerate(values):
            cell_shading(row.cells[j], col_fills[j])
            p = row.cells[j].paragraphs[0]
            r = p.add_run(val)
            is_etapa = j == 0
            run_style(r, "Arial", 10, C["blue"] if is_etapa else C["body"],
                      bold=is_etapa)
            cell_borders(row.cells[j], color=C["line"], size="6")
            row.cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph()


def add_contributions_table(doc: Document) -> None:
    rows_data = [
        "Observações sobre participantes identificados",
        "Observações sobre disciplinas vinculadas ao curso",
        "Sugestões para os encontros individuais com docentes",
        "Sugestões para uso do grupo de WhatsApp",
        "Outras recomendações ou informações relevantes",
        "Oportunidades identificadas para o curso",
    ]
    table = doc.add_table(rows=1 + len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Cabeçalho
    header_row = table.rows[0]
    for j, h in enumerate(["Campo", "Contribuição"]):
        cell_shading(header_row.cells[j], C["blue"])
        p = header_row.cells[j].paragraphs[0]
        r = p.add_run(h)
        run_style(r, "Arial", 10, C["white"], bold=True)
        cell_borders(header_row.cells[j], color=C["blue"], size="6")

    for i, campo in enumerate(rows_data):
        row = table.rows[i + 1]
        fill = C["blue_soft2"] if i % 2 == 0 else C["panel"]
        cell_shading(row.cells[0], C["blue_soft"])
        cell_shading(row.cells[1], fill)

        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(campo)
        run_style(r0, "Arial", 10, C["blue"], bold=False)

        # célula em branco para preenchimento
        p1 = row.cells[1].paragraphs[0]
        p1.paragraph_format.space_before = Pt(14)
        p1.paragraph_format.space_after = Pt(14)
        p1.add_run("")

        for cell in row.cells:
            cell_borders(cell, color=C["line"], size="6")
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    doc.add_paragraph()


# ── Construção do documento ────────────────────────────────────────────────────

def build_document() -> None:
    doc = Document()
    configure_styles(doc)

    # ── Capa ──────────────────────────────────────────────────────────────────
    add_cover(doc)

    # ── Seção de corpo ────────────────────────────────────────────────────────
    from docx.enum.section import WD_SECTION_START
    body_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(
        body_section,
        header_text="Araujo Innovation Lab  |  Levantamento Inicial — Curso Técnico em Marketing",
        footer_right="EEB Araújo Figueiredo  |  Junho 2026",
        with_hf=True,
    )

    # ── Seção 1 — Identificação ───────────────────────────────────────────────
    add_section_heading(doc, "1", "IDENTIFICAÇÃO DO DOCUMENTO")
    add_identification_table(doc)

    # ── Seção 2 — Objetivo ───────────────────────────────────────────────────
    add_section_heading(doc, "2", "OBJETIVO")
    add_body(doc,
        "Este documento tem como propósito apoiar o processo inicial de organização e levantamento "
        "do Curso Técnico em Marketing, reunindo as informações disponíveis sobre participantes, "
        "disciplinas e possibilidades de integração."
    )
    add_body(doc,
        "O material foi preparado como ponto de partida para conversas, não como conclusão. "
        "Ele reflete o estado atual do levantamento — incompleto por definição — e tem valor "
        "justamente na sua capacidade de facilitar o diálogo, reunir contribuições e orientar "
        "as próximas etapas."
    )
    add_body(doc, "Os destinatários deste documento são convidados a:")
    add_bullet(doc, "validar as informações apresentadas;")
    add_bullet(doc, "corrigir eventuais imprecisões;")
    add_bullet(doc, "sugerir complementações;")
    add_bullet(doc, "contribuir com observações relevantes para o planejamento do curso.")
    doc.add_paragraph()
    add_highlight_block(doc, "", "Nenhuma decisão foi tomada. Este é um documento vivo, aberto a ajustes.",
                        fill=C["gold_soft"], border_color=C["gold"])

    # ── Seção 3 — Contextualização ────────────────────────────────────────────
    add_section_heading(doc, "3", "CONTEXTUALIZAÇÃO")
    add_body(doc,
        "O Curso Técnico em Marketing possui uma natureza essencialmente multidisciplinar. "
        "Sua riqueza formativa não reside apenas nas disciplinas técnicas que lhe dão identidade, "
        "mas na articulação entre diferentes áreas do conhecimento que, juntas, ampliam o repertório "
        "dos estudantes e aprofundam a qualidade das experiências de aprendizagem."
    )
    add_body(doc,
        "Compreender quem são os professores que contribuem com o curso — e de que forma cada "
        "disciplina se conecta à formação em marketing — é um passo importante para fortalecer "
        "a aprendizagem dos estudantes e valorizar o trabalho docente já em curso."
    )
    add_body(doc,
        "Este levantamento não é uma avaliação. Não tem por objetivo medir desempenho ou redirecionar "
        "práticas. Seu propósito é, simplesmente, conhecer melhor o que existe, identificar "
        "oportunidades de integração e construir, de forma colaborativa, um planejamento mais "
        "consistente para o curso."
    )
    add_highlight_block(doc, "Princípio orientador",
                        "A participação de todos os envolvidos — gestores, orientadores, professores "
                        "e equipe administrativa — é fundamental para que este mapeamento seja preciso, "
                        "representativo e útil.",
                        fill=C["blue_soft2"], border_color=C["line"])

    # ── Seção 4 — Participantes ───────────────────────────────────────────────
    add_section_heading(doc, "4", "PARTICIPANTES ATUALMENTE IDENTIFICADOS")
    add_body(doc,
        "As informações a seguir representam o levantamento em curso, com base nas informações "
        "disponíveis até a data deste documento. Correções, complementações e eventuais ausências "
        "devem ser sinalizadas pelos destinatários."
    )
    doc.add_paragraph()

    add_subsection(doc, "4.1 Gestão e Apoio Institucional")
    add_participants_table(doc,
        headers=["Nome", "Função / Papel no Curso"],
        rows_data=[
            ["Prof. José Santana",    "Marketing — Docência Técnica"],
            ["Profª Letícia Meurer",  "Marketing — Docência Técnica"],
            ["Prof. Claudecir",       "Diretor"],
            ["Aline Liz",             "Secretária da Diretoria"],
            ["Suzane",                "Orientadora Pedagógica"],
        ]
    )

    add_subsection(doc, "4.2 Corpo Docente Identificado")
    add_participants_table(doc,
        headers=["Nome", "Disciplina"],
        rows_data=[
            ["Profª Amanda Cemin Elizeu",                "Química"],
            ["Profª Ana Carla de Souza",                 "Arte"],
            ["Profª Luciane Matias de Oliveira Krtiscka","Geografia"],
            ["Profª Jocélia Demétrio Vieira",            "Física"],
            ["Profª Juliana de Oliveira Ribeiro",        "Biologia"],
            ["Profª Rosana Maria Badalotti",             "Filosofia"],
            ["Profª Sheila da Cruz",                     "Língua Inglesa"],
        ]
    )

    add_subsection(doc, "4.3 Participantes a Identificar")
    add_body(doc,
        "Os contatos abaixo ainda não foram associados a nomes ou funções. "
        "Solicita-se apoio dos destinatários para identificação:"
    )
    add_participants_table(doc,
        headers=["Contato", "Situação"],
        rows_data=[
            ["49 9135-4656", "A identificar"],
            ["49 9824-137",  "Número incompleto — verificar dígito faltante"],
        ]
    )
    add_note_block(doc,
        "Qualquer informação adicional sobre participantes ainda não listados é bem-vinda. "
        "Este levantamento está em construção e pode estar incompleto."
    )

    # ── Seção 5 — Levantamento das Disciplinas ────────────────────────────────
    add_section_heading(doc, "5", "LEVANTAMENTO DAS DISCIPLINAS")
    add_body(doc,
        "A próxima etapa do processo consistirá no mapeamento detalhado das disciplinas vinculadas "
        "ao Curso Técnico em Marketing. Para cada disciplina identificada, o objetivo será compreender:"
    )
    add_bullet(doc, "o professor responsável pela disciplina;")
    add_bullet(doc, "os conteúdos desenvolvidos ao longo do ano letivo;")
    add_bullet(doc, "as competências trabalhadas com os estudantes;")
    add_bullet(doc, "as atividades práticas já realizadas ou em planejamento;")
    add_bullet(doc, "as oportunidades de integração com os objetivos do Curso Técnico em Marketing.")
    doc.add_paragraph()
    add_body(doc,
        "Este levantamento será construído em parceria com os próprios professores, por meio de "
        "conversas individuais e instrumentos simples de coleta de informações."
    )
    add_highlight_block(doc, "Foco",
                        "O propósito deste mapeamento é compreender e valorizar o que já acontece, "
                        "não avaliar práticas ou estabelecer comparações. Cada disciplina contribui, "
                        "à sua maneira, com a formação dos estudantes.",
                        fill=C["green_soft"], border_color=C["line"])

    # ── Seção 6 — Reuniões Individuais ────────────────────────────────────────
    add_section_heading(doc, "6", "PROPOSTA DE REUNIÕES INDIVIDUAIS COM OS DOCENTES")
    add_body(doc,
        "Como parte do processo de levantamento e alinhamento, propõe-se a realização de reuniões "
        "individuais com os professores vinculados ao curso. Esses encontros têm caráter "
        "exploratório e colaborativo."
    )
    doc.add_paragraph()
    add_subsection(doc, "Objetivos sugeridos para os encontros individuais")
    add_bullet(doc, "Compreender a realidade de cada disciplina — conteúdos, metodologias, dinâmicas de sala e desafios específicos;")
    add_bullet(doc, "Conhecer iniciativas já existentes que, mesmo sem articulação formal, dialogam com os objetivos do curso técnico;")
    add_bullet(doc, "Valorizar o trabalho já desenvolvido — reconhecer práticas, projetos e esforços que muitas vezes passam despercebidos institucionalmente;")
    add_bullet(doc, "Identificar oportunidades de integração que possam enriquecer tanto a disciplina do professor quanto a experiência do estudante no curso técnico;")
    add_bullet(doc, "Construir um planejamento colaborativo que respeite a autonomia pedagógica de cada docente e ao mesmo tempo fortaleça a coesão do curso.")
    doc.add_paragraph()
    add_subsection(doc, "Características dos encontros")
    add_body(doc,
        "Os encontros serão curtos, objetivos e conduzidos com escuta ativa. Não há agenda "
        "pré-definida além dos objetivos acima. O professor é o protagonista da conversa."
    )
    add_body(doc,
        "A agenda, os horários e os formatos (presencial ou por mensagem, conforme disponibilidade) "
        "serão combinados individualmente, respeitando a rotina de cada docente."
    )

    # ── Seção 7 — WhatsApp ────────────────────────────────────────────────────
    add_section_heading(doc, "7", "PROPOSTA PRELIMINAR PARA O GRUPO DE WHATSAPP")
    add_note_block(doc,
        "O conteúdo desta seção é uma proposta preliminar para discussão. Não representa regras, "
        "políticas ou obrigações. Toda e qualquer definição sobre o grupo dependerá do alinhamento "
        "com os participantes."
    )
    add_subsection(doc, "Contexto")
    add_body(doc,
        "O grupo de WhatsApp existente reúne professores e colaboradores vinculados ao Curso "
        "Técnico em Marketing. Antes de qualquer decisão sobre sua utilização, é relevante alinhar "
        "com os participantes qual seria o uso mais útil e menos invasivo desse espaço."
    )
    doc.add_paragraph()
    add_subsection(doc, "Possibilidades de uso — em discussão")
    add_bullet(doc, "Compartilhamento de projetos desenvolvidos pelas turmas do curso técnico;")
    add_bullet(doc, "Divulgação de atividades realizadas por alunos ou por disciplinas com potencial de integração;")
    add_bullet(doc, "Integração entre disciplinas — informes rápidos sobre conexões identificadas entre áreas;")
    add_bullet(doc, "Acompanhamento da evolução do curso — registros relevantes do dia a dia pedagógico;")
    add_bullet(doc, "Compartilhamento de boas práticas que possam inspirar outros docentes;")
    add_bullet(doc, "Fortalecimento da comunicação entre os participantes do curso, de forma leve e não invasiva.")
    doc.add_paragraph()
    add_highlight_block(doc, "",
                        "A frequência, o tipo de conteúdo e a dinâmica de participação são aspectos "
                        "que devem ser definidos coletivamente — e podem, inclusive, permanecer "
                        "exatamente como já estão, se essa for a avaliação dos participantes.",
                        fill=C["blue_soft2"], border_color=C["line"])

    # ── Seção 8 — Contribuições ───────────────────────────────────────────────
    doc.add_page_break()
    add_section_heading(doc, "8", "CONTRIBUIÇÕES E SUGESTÕES")
    add_body(doc,
        "Esta seção foi reservada para o registro de observações, sugestões, recomendações "
        "e oportunidades identificadas pelos destinatários deste documento."
    )
    add_body(doc,
        "As contribuições podem ser registradas diretamente neste documento, enviadas por mensagem "
        "ou compartilhadas verbalmente durante o alinhamento. Todas as contribuições serão "
        "incorporadas às versões seguintes do documento."
    )
    doc.add_paragraph()
    add_contributions_table(doc)

    # ── Seção 9 — Próximos Passos ─────────────────────────────────────────────
    add_section_heading(doc, "9", "PRÓXIMOS PASSOS")
    add_body(doc,
        "A sequência sugerida para os encaminhamentos deste processo é apresentada a seguir. "
        "As etapas são propostas de organização, não obrigações — podendo ser ajustadas conforme "
        "a realidade e disponibilidade dos envolvidos."
    )
    doc.add_paragraph()
    add_next_steps_table(doc)

    # ── Rodapé de encerramento ────────────────────────────────────────────────
    doc.add_paragraph()
    end_table = doc.add_table(rows=1, cols=1)
    end_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    end_cell = end_table.rows[0].cells[0]
    cell_shading(end_cell, C["blue"])
    cell_borders(end_cell, color=C["blue"], size="6")
    end_p = end_cell.paragraphs[0]
    end_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p.paragraph_format.space_before = Pt(8)
    end_p.paragraph_format.space_after = Pt(8)

    r_e1 = end_p.add_run("Elaborado por: ")
    run_style(r_e1, "Arial", 10, C["gold"], bold=True)
    r_e2 = end_p.add_run("Profª Letícia Meurer · Professor Orientador José Alexandre")
    run_style(r_e2, "Arial", 10, C["white"])

    end_p2 = end_cell.add_paragraph()
    end_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p2.paragraph_format.space_before = Pt(2)
    end_p2.paragraph_format.space_after = Pt(8)
    r_e3 = end_p2.add_run("Versão 0.1 — Junho de 2026 · Circulação Interna Restrita")
    run_style(r_e3, "Arial", 9, "D7DEDA")

    end_p3 = end_cell.add_paragraph()
    end_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p3.paragraph_format.space_after = Pt(8)
    r_e4 = end_p3.add_run("Araujo Innovation Lab — EEB Araújo Figueiredo — Urubici, SC")
    run_style(r_e4, "Arial", 8, "D7DEDA", italic=True)

    # ── Salvar ────────────────────────────────────────────────────────────────
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT_DOCX)
    print(f"Documento gerado: {OUTPUT_DOCX}")


if __name__ == "__main__":
    build_document()
