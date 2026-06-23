"""
Gerador do documento institucional:
Construindo Juntos o Futuro do Curso Técnico em Marketing
Guia para Conversa e Alinhamento Inicial
Colégio Araújo — Araujo Innovation Lab
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
OUTPUT_DOCX = OUT_DIR / "Colegio-Araujo-Construindo-Juntos-Futuro-Curso-Marketing-v0.1.docx"


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
    tag_run = tag_p.add_run("GUIA PARA CONVERSA E ALINHAMENTO INICIAL")
    run_style(tag_run, "Arial", 8, C["muted"], bold=True)

    # Título principal
    title = doc.add_paragraph(style="Title")
    title.paragraph_format.space_before = Pt(28)
    title.paragraph_format.space_after = Pt(12)
    title.add_run("Construindo Juntos o Futuro do Curso Técnico em Marketing")

    # Subtítulo descritivo
    sub = doc.add_paragraph()
    sub.paragraph_format.space_after = Pt(32)
    r = sub.add_run(
        "Documento de trabalho para estimular reflexão, diálogo e construção colaborativa "
        "sobre o futuro do Curso Técnico em Marketing."
    )
    run_style(r, "Arial", 12, C["muted"])

    # Bloco de identificação do documento (tabela)
    id_table = doc.add_table(rows=6, cols=2)
    id_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    id_rows = [
        ("Versão",          "0.1"),
        ("Data",            "Junho de 2026"),
        ("Status",          "Em Discussão — nenhuma decisão tomada"),
        ("Classificação",   "Documento de Trabalho Institucional"),
        ("Elaborado por",   "Prof. José Alexandre — Professor Orientador do Curso"),
        ("Participantes",   "Aline Liz · Prof. José Alexandre · Profª Letícia Meurer · Susane Custódio"),
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
    spacer.paragraph_format.space_before = Pt(60)
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


# ── Primitivos de texto ────────────────────────────────────────────────────────

def add_highlight_block(doc: Document, label: str, text: str,
                        fill: str = "F1F5F3", border_color: str = "D7DEDA") -> None:
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell_shading(cell, fill)
    cell_borders(cell, color=border_color, size="12")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
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
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    r1 = p.add_run("Nota importante:  ")
    run_style(r1, "Arial", 10, C["blue"], bold=True)
    r2 = p.add_run(text)
    run_style(r2, "Arial", 10, C["body"])
    doc.add_paragraph()


def add_question_block(doc: Document, text: str) -> None:
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.rows[0].cells[0]
    cell_shading(cell, C["green_soft"])
    cell_borders(cell, color=C["gold"], size="10")
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    r1 = p.add_run("Pergunta orientadora:  ")
    run_style(r1, "Arial", 10, C["blue"], bold=True)
    r2 = p.add_run(text)
    run_style(r2, "Arial", 10, C["body"], italic=True)
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


def add_discipline_block(doc: Document, discipline: str, items: list[tuple[str, str]]) -> None:
    """Painel de uma disciplina com rótulo de conceito e descrição curta."""
    doc.add_paragraph(discipline, style="Heading 3")
    for concept, description in items:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.35)
        p.paragraph_format.space_after = Pt(3)
        r_concept = p.add_run(f"•  {concept}")
        run_style(r_concept, "Arial", 11, C["blue"], bold=True)
        r_desc = p.add_run(f" — {description}")
        run_style(r_desc, "Arial", 11, C["body"])


# ── Construção do documento ────────────────────────────────────────────────────

def build_document() -> None:
    doc = Document()
    configure_styles(doc)

    # ── Capa ──────────────────────────────────────────────────────────────────
    add_cover(doc)

    # ── Seção de corpo ────────────────────────────────────────────────────────
    body_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(
        body_section,
        header_text=(
            "Araujo Innovation Lab  |  "
            "Construindo Juntos o Futuro do Curso Técnico em Marketing"
        ),
        footer_right="EEB Araújo Figueiredo  |  Junho 2026",
        with_hf=True,
    )

    # ── Seção 1 — Apresentação ────────────────────────────────────────────────
    add_section_heading(doc, "1", "APRESENTAÇÃO")
    add_body(doc,
        "Este documento foi criado para apoiar uma conversa colaborativa sobre "
        "oportunidades de fortalecimento do Curso Técnico em Marketing da EEB Araújo Figueiredo. "
        "Ele não representa um planejamento pronto, uma decisão tomada ou uma proposta a ser aprovada."
    )
    add_body(doc,
        "Sua única finalidade é estimular o diálogo. Cada seção propõe uma pergunta, uma reflexão "
        "ou um tema para que, juntos, possamos pensar sobre o futuro do curso e sobre o que gostaríamos "
        "de oferecer aos nossos estudantes até a conclusão da turma."
    )
    add_body(doc,
        "Todos os temas aqui apresentados estão abertos. Nada está decidido. "
        "O objetivo desta conversa é ouvir, refletir e construir conjuntamente — "
        "a partir das perspectivas, experiências e ideias de cada participante."
    )
    add_highlight_block(
        doc, "Princípio fundamental",
        "Este documento deve ajudar as pessoas a conversar sobre possibilidades futuras. "
        "Ele não deve definir o futuro.",
        fill=C["gold_soft"], border_color=C["gold"]
    )

    # ── Seção 2 — O Momento Atual do Curso ───────────────────────────────────
    add_section_heading(doc, "2", "O MOMENTO ATUAL DO CURSO")
    add_body(doc,
        "O Curso Técnico em Marketing está em andamento. Uma turma de estudantes caminha em direção "
        "à sua conclusão, prevista para 2027. Esse é um momento importante — e uma oportunidade "
        "significativa de fortalecer a experiência que esses jovens estão tendo."
    )
    add_body(doc,
        "Antes de conversar sobre possibilidades, é importante compreender melhor a realidade atual "
        "do curso: como os estudantes estão vivenciando essa formação, quais são os pontos fortes "
        "já presentes, e onde há espaço para enriquecer ainda mais essa trajetória."
    )
    add_body(doc,
        "O curso é conduzido por profissionais comprometidos e conta com o envolvimento de "
        "diferentes áreas da escola. Esse já é um patrimônio importante — e qualquer construção "
        "coletiva parte do que já existe e funciona."
    )
    add_highlight_block(
        doc, "Contexto",
        "Turma em andamento · Conclusão prevista em 2027 · "
        "Oportunidade de fortalecer a experiência dos estudantes nos próximos dois anos.",
        fill=C["blue_soft2"], border_color=C["line"]
    )

    # ── Seção 3 — Nosso Objetivo com Esta Conversa ───────────────────────────
    add_section_heading(doc, "3", "NOSSO OBJETIVO COM ESTA CONVERSA")
    add_body(doc,
        "A pergunta central que gostaríamos de explorar juntos é a seguinte:"
    )
    add_question_block(
        doc,
        "O que gostaríamos que nossos alunos vivenciassem até a conclusão do curso?"
    )
    add_body(doc,
        "Para ajudar a organizar essa reflexão, alguns temas podem orientar nossa conversa. "
        "Não são respostas — são pontos de partida para o diálogo."
    )
    doc.add_paragraph()
    add_bullet(doc, "Aprendizagem significativa — experiências que façam sentido para a vida dos estudantes;")
    add_bullet(doc, "Experiências práticas — contato com situações reais do universo do marketing;")
    add_bullet(doc, "Protagonismo — espaço para que os alunos sejam autores de suas próprias aprendizagens;")
    add_bullet(doc, "Criatividade — estímulo à geração de ideias, soluções e projetos autorais;")
    add_bullet(doc, "Trabalho em equipe — desenvolvimento da capacidade de colaborar e construir junto;")
    add_bullet(doc, "Comunicação — habilidade de expressar ideias com clareza e impacto;")
    add_bullet(doc, "Preparação para o mercado — aproximação com as demandas e realidades da profissão.")
    doc.add_paragraph()
    add_highlight_block(
        doc, "",
        "Esses temas são apenas sugestões. Cada participante pode trazer outros temas, "
        "perspectivas e prioridades — essa é exatamente a riqueza desta conversa.",
        fill=C["blue_soft2"], border_color=C["line"]
    )

    # ── Seção 4 — Como Cada Disciplina Pode Contribuir ───────────────────────
    add_section_heading(doc, "4", "COMO CADA DISCIPLINA PODE CONTRIBUIR")
    add_body(doc,
        "O objetivo desta seção não é transformar todas as disciplinas em Marketing. "
        "O propósito é compreender como cada área do conhecimento pode enriquecer "
        "a formação dos estudantes que estão construindo sua trajetória no Curso Técnico."
    )
    add_body(doc,
        "Os exemplos a seguir são convites para reflexão. Cada professor conhece melhor do que ninguém "
        "as possibilidades da sua própria disciplina — e essa é a conversa que gostaríamos de ter."
    )
    doc.add_paragraph()

    add_discipline_block(doc, "Arte", [
        ("Criatividade", "estímulo à geração de ideias e soluções inovadoras"),
        ("Identidade visual", "compreensão de elementos que representam marcas e projetos"),
        ("Comunicação visual", "utilização de recursos visuais para transmitir mensagens"),
        ("Branding", "construção da percepção e do posicionamento de marcas"),
    ])

    add_discipline_block(doc, "Geografia", [
        ("Mercado regional", "compreensão das características econômicas e sociais da região"),
        ("Turismo", "análise de oportunidades relacionadas ao desenvolvimento local"),
        ("Comportamento de consumo", "observação de hábitos e necessidades de diferentes públicos"),
        ("Desenvolvimento local", "entendimento do impacto das empresas e iniciativas na comunidade"),
    ])

    add_discipline_block(doc, "Filosofia", [
        ("Ética", "reflexão sobre responsabilidade e conduta profissional"),
        ("Pensamento crítico", "análise de informações, argumentos e decisões"),
        ("Responsabilidade social", "compreensão do papel das organizações na sociedade"),
        ("Reflexão sobre tecnologia", "discussão sobre os impactos das transformações digitais"),
    ])

    add_discipline_block(doc, "Língua Inglesa", [
        ("Ferramentas digitais", "contato com plataformas e conteúdos frequentemente disponíveis em inglês"),
        ("Marketing global", "compreensão de tendências e referências internacionais"),
        ("Tendências internacionais", "acesso a conteúdos e práticas adotadas em outros países"),
        ("Inteligência Artificial", "familiarização com recursos e tecnologias amplamente difundidos em inglês"),
    ])

    add_discipline_block(doc, "Biologia", [
        ("Sustentabilidade", "compreensão do equilíbrio entre desenvolvimento e preservação"),
        ("ESG", "reflexão sobre práticas ambientais, sociais e de governança"),
        ("Consumo consciente", "análise do impacto das escolhas de consumo"),
        ("Qualidade de vida", "relação entre produtos, serviços e bem-estar das pessoas"),
    ])

    add_discipline_block(doc, "Física", [
        ("Tecnologia", "compreensão dos princípios presentes em diversas soluções tecnológicas"),
        ("Inovação", "estímulo à busca de novas formas de resolver problemas"),
        ("Resolução de problemas", "desenvolvimento do raciocínio lógico e analítico"),
        ("Pensamento estruturado", "organização de informações para tomada de decisão"),
    ])

    add_discipline_block(doc, "Química", [
        ("Desenvolvimento de produtos", "compreensão de processos relacionados à criação de produtos"),
        ("Embalagens", "análise de materiais, funcionalidade e sustentabilidade"),
        ("Processos produtivos", "entendimento das etapas envolvidas na fabricação"),
        ("Sustentabilidade industrial", "reflexão sobre impactos ambientais e alternativas mais responsáveis"),
    ])

    doc.add_paragraph()
    add_note_block(doc,
        "Os exemplos apresentados possuem caráter exclusivamente ilustrativo e não representam "
        "definições curriculares ou direcionamentos pedagógicos. Seu objetivo é apenas apoiar "
        "a reflexão e estimular futuras conversas com os docentes."
    )

    # ── Seção 5 — O Mundo Está Mudando ───────────────────────────────────────
    doc.add_page_break()
    add_section_heading(doc, "5", "O MUNDO ESTÁ MUDANDO. E NOSSOS ALUNOS?")
    add_body(doc,
        "O Marketing é uma das áreas que mais tem se transformado nos últimos anos. "
        "O Marketing Digital, a Inteligência Artificial e as novas tecnologias estão "
        "reconfigurando as formas de comunicar, criar e se conectar com as pessoas."
    )
    add_body(doc,
        "Não se trata de algo distante ou assustador. Pelo contrário — essas transformações "
        "abrem oportunidades extraordinárias para profissionais que saibam se adaptar, "
        "que sejam curiosos, criativos e dispostos a aprender continuamente."
    )
    add_body(doc,
        "Nossos estudantes estão sendo formados justamente nesse momento de transição. "
        "A pergunta que queremos trazer para esta conversa não é sobre o que vai mudar — "
        "mas sobre o que podemos fazer, agora, para prepará-los bem para o que vem."
    )
    doc.add_paragraph()
    add_bullet(doc, "Marketing Digital — como aproximar os estudantes das ferramentas e práticas do marketing contemporâneo;")
    add_bullet(doc, "Inteligência Artificial — como desenvolver familiaridade crítica com recursos que já fazem parte do mercado;")
    add_bullet(doc, "Novas tecnologias — como estimular a curiosidade e a disposição para aprender o que ainda está surgindo;")
    add_bullet(doc, "Criatividade e adaptabilidade — como fortalecer competências que resistem às transformações tecnológicas;")
    add_bullet(doc, "Aprendizagem contínua — como cultivar o hábito e o prazer de continuar aprendendo ao longo da vida.")
    doc.add_paragraph()
    add_question_block(
        doc,
        "Como podemos ajudar nossos alunos a se tornarem profissionais adaptáveis, "
        "criativos, curiosos e preparados para aprender continuamente?"
    )

    # ── Seção 6 — O Que Podemos Fazer Além da Sala de Aula ──────────────────
    add_section_heading(doc, "6", "O QUE PODEMOS FAZER ALÉM DA SALA DE AULA?")
    add_body(doc,
        "A sala de aula é o coração da formação — mas não o único espaço de aprendizagem. "
        "Esta seção apresenta algumas possibilidades para reflexão. "
        "Não são propostas fechadas — são convites para pensar juntos sobre o que faz sentido "
        "para o nosso contexto e para os nossos estudantes."
    )
    doc.add_paragraph()

    add_subsection(doc, "Dentro do Colégio")
    add_bullet(doc, "Redes sociais — participação dos estudantes na comunicação institucional do colégio;")
    add_bullet(doc, "Produção de conteúdo — criação de materiais que representem o colégio e o curso;")
    add_bullet(doc, "Cobertura de eventos — registro e divulgação de atividades e momentos da escola;")
    add_bullet(doc, "Projetos institucionais — contribuição dos alunos para iniciativas do próprio colégio.")
    doc.add_paragraph()

    add_subsection(doc, "Fora do Colégio")
    add_bullet(doc, "Visitas técnicas — contato com empresas e profissionais do mercado regional;")
    add_bullet(doc, "Empresas locais — aproximação com o ecossistema de negócios de Urubici e região;")
    add_bullet(doc, "Eventos — participação em feiras, mostras e iniciativas da comunidade;")
    add_bullet(doc, "Projetos comunitários — contribuição dos estudantes para causas e iniciativas locais.")
    doc.add_paragraph()

    add_subsection(doc, "Vivências Profissionais")
    add_bullet(doc, "Portal Connect — acesso a experiências práticas e projetos reais por meio da plataforma;")
    add_bullet(doc, "Projetos reais — desenvolvimento de trabalhos com impacto além da sala de aula;")
    add_bullet(doc, "Ferramentas digitais — familiarização com plataformas utilizadas no mercado;")
    add_bullet(doc, "Marketing digital — práticas orientadas à comunicação, ao conteúdo e à presença online;")
    add_bullet(doc, "Inteligência Artificial — experimentação e uso crítico de recursos de IA aplicados ao marketing.")
    doc.add_paragraph()
    add_highlight_block(
        doc, "",
        "Todas as possibilidades aqui listadas são pontos de partida para o diálogo — "
        "não compromissos ou propostas a serem implementadas. O que faz sentido será "
        "construído coletivamente, no ritmo e na medida certa para o nosso contexto.",
        fill=C["blue_soft2"], border_color=C["line"]
    )

    # ── Seção 7 — Como Podemos Melhorar Nossa Comunicação ────────────────────
    add_section_heading(doc, "7", "COMO PODEMOS MELHORAR NOSSA COMUNICAÇÃO?")
    add_body(doc,
        "Um grupo que se comunica bem tem mais capacidade de construir junto. "
        "Esta seção convida à reflexão sobre como os participantes deste processo "
        "gostariam de se relacionar, compartilhar e se manter conectados."
    )
    doc.add_paragraph()
    add_bullet(doc, "Identidade visual do grupo — existe algo que nos representa como coletivo? Como gostaríamos de ser percebidos?")
    add_bullet(doc, "Nome do grupo — o grupo tem um nome que faça sentido e que represente o que construímos juntos?")
    add_bullet(doc, "Descrição institucional — como apresentamos este grupo para quem ainda não o conhece?")
    add_bullet(doc, "Compartilhamento de ações — que tipo de informações e atualizações faz sentido compartilhar entre os participantes?")
    add_bullet(doc, "Divulgação de resultados — como registramos e comunicamos o que o curso está produzindo?")
    add_bullet(doc, "Integração dos participantes — como fortalecemos os laços entre todos os envolvidos no curso?")
    doc.add_paragraph()
    add_highlight_block(
        doc, "",
        "Essas perguntas não têm respostas certas. São convites para que cada participante "
        "contribua com sua perspectiva — e para que, juntos, possamos construir "
        "formas de comunicação que sirvam a todos.",
        fill=C["green_soft"], border_color=C["line"]
    )

    # ── Seção 8 — Fundamentação Institucional ────────────────────────────────
    add_section_heading(doc, "8", "FUNDAMENTAÇÃO INSTITUCIONAL")
    add_body(doc,
        "A conversa que este documento propõe está alinhada às atribuições do Professor Orientador "
        "de Curso Técnico, conforme previsto na Instrução Normativa nº 128/2026. "
        "O trabalho de orientação inclui, entre outras responsabilidades, o acompanhamento "
        "das transformações da profissão, a atualização pedagógica e curricular, "
        "e o cuidado com a qualidade da experiência formativa dos estudantes."
    )
    add_body(doc,
        "Mais do que uma exigência normativa, essa fundamentação expressa um compromisso: "
        "o de garantir que o Curso Técnico em Marketing da EEB Araújo Figueiredo esteja "
        "sempre conectado às demandas do presente e preparado para o futuro."
    )
    add_body(doc,
        "A integração entre todos os participantes — professores, gestão, orientação pedagógica "
        "e equipe administrativa — é parte essencial desse compromisso. Este documento é, "
        "em si mesmo, uma expressão dessa integração."
    )
    add_highlight_block(
        doc, "Princípio",
        "A qualidade pedagógica de um curso não se constrói sozinha. "
        "Ela é resultado do cuidado coletivo, da escuta mútua e do compromisso compartilhado "
        "com a formação dos estudantes.",
        fill=C["blue_soft2"], border_color=C["line"]
    )

    # ── Seção 9 — Contribuições Esperadas ────────────────────────────────────
    doc.add_page_break()
    add_section_heading(doc, "9", "CONTRIBUIÇÕES ESPERADAS")
    add_body(doc,
        "Ao final desta conversa, esperamos ter avançado coletivamente em alguns pontos importantes. "
        "Não como resultado de decisões tomadas — mas como fruto do diálogo e da escuta de cada participante."
    )
    doc.add_paragraph()
    add_bullet(doc, "Compreender melhor a realidade atual do curso — seus pontos fortes, seus desafios e o que já existe de valioso;")
    add_bullet(doc, "Identificar oportunidades — caminhos que ainda não foram explorados e que podem enriquecer a formação dos estudantes;")
    add_bullet(doc, "Reunir ideias — diferentes perspectivas sobre o que poderia ser feito, testado ou fortalecido;")
    add_bullet(doc, "Fortalecer a integração — aproximar ainda mais todos os envolvidos no curso em torno de propósitos comuns;")
    add_bullet(doc, "Apoiar futuras etapas de planejamento — construir uma base de reflexões que possa orientar os próximos passos, quando o momento for oportuno.")
    doc.add_paragraph()
    add_highlight_block(
        doc, "",
        "O resultado mais importante desta conversa não é um documento — é o entendimento "
        "que construímos juntos. Tudo o que surgir aqui será cuidadosamente considerado "
        "nas etapas seguintes do processo.",
        fill=C["gold_soft"], border_color=C["gold"]
    )

    # ── Espaço para anotações ─────────────────────────────────────────────────
    doc.add_paragraph()
    add_subsection(doc, "Espaço para anotações e contribuições")
    add_body(doc,
        "Este espaço foi reservado para que os participantes registrem suas ideias, "
        "observações e contribuições durante ou após a conversa."
    )
    doc.add_paragraph()

    # Linhas em branco para anotações
    for _ in range(8):
        line_p = doc.add_paragraph()
        line_p.paragraph_format.space_after = Pt(16)
        para_bottom_border(line_p, color=C["line"], size="4")

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
    r_e2 = end_p.add_run("Prof. José Alexandre — Professor Orientador do Curso Técnico em Marketing")
    run_style(r_e2, "Arial", 10, C["white"])

    end_p2 = end_cell.add_paragraph()
    end_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_p2.paragraph_format.space_before = Pt(2)
    end_p2.paragraph_format.space_after = Pt(8)
    r_e3 = end_p2.add_run("Versão 0.1 — Junho de 2026 · Em Discussão · Circulação Interna Restrita")
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
