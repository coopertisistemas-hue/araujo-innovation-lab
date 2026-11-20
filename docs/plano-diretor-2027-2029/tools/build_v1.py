import re, sys, json, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC = sys.argv[1]; OUTDOCX = sys.argv[2]; CHARTS = sys.argv[3]; PAGEMAP = sys.argv[4] if len(sys.argv) > 4 else None
pagemap = json.load(open(PAGEMAP)) if PAGEMAP and os.path.exists(PAGEMAP) else {}

NAVY="1F2F4D"; GOLD="B8975A"; GREY="8C8279"; CREAM="F4F0E8"; INK="2B2B2B"; ZEBRA="FAF8F4"; RULE="D9CFC0"
HFONT="Georgia"; BFONT="Calibri"; CONTENT_W = 16.6  # cm

lines = open(SRC, encoding="utf-8").read().split("\n")

doc = Document()

# ---------- page setup ----------
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.2)
sec.top_margin = Cm(2.4); sec.bottom_margin = Cm(2.2)
sec.header_distance = Cm(1.1); sec.footer_distance = Cm(1.0)
sec.different_first_page_header_footer = True

def set_style_font(style, name, size, bold=None, italic=None, color=None):
    style.font.name = name
    style.font.size = Pt(size)
    if bold is not None: style.font.bold = bold
    if italic is not None: style.font.italic = italic
    if color: style.font.color.rgb = RGBColor.from_string(color)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ["w:asciiTheme","w:hAnsiTheme","w:eastAsiaTheme","w:cstheme"]:
        rf.attrib.pop(qn(a), None)
    for a in ["w:ascii","w:hAnsi","w:cs","w:eastAsia"]:
        rf.set(qn(a), name)

st = doc.styles
set_style_font(st["Normal"], BFONT, 10.5, color=INK)
st["Normal"].paragraph_format.space_after = Pt(6)
st["Normal"].paragraph_format.line_spacing = 1.2
set_style_font(st["Heading 1"], HFONT, 19, bold=False, color=NAVY)
st["Heading 1"].paragraph_format.space_before = Pt(0)
st["Heading 1"].paragraph_format.space_after = Pt(12)
st["Heading 1"].paragraph_format.keep_with_next = True
set_style_font(st["Heading 2"], HFONT, 13.5, bold=False, color=NAVY)
st["Heading 2"].paragraph_format.space_before = Pt(14)
st["Heading 2"].paragraph_format.space_after = Pt(6)
st["Heading 2"].paragraph_format.keep_with_next = True
set_style_font(st["Heading 3"], BFONT, 11, bold=True, color=NAVY)
st["Heading 3"].paragraph_format.space_before = Pt(8)
st["Heading 3"].paragraph_format.keep_with_next = True
set_style_font(st["List Bullet"], BFONT, 10.5, color=INK)

# ---------- helpers ----------
def shade(el_parent_pr, fill):
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
    el_parent_pr.append(shd)

def para_border(p, side, color, sz=12, space=4):
    pPr = p._p.get_or_add_pPr()
    bdr = pPr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr"); pPr.append(bdr)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(sz)); el.set(qn("w:space"), str(space)); el.set(qn("w:color"), color)
    bdr.append(el)

def add_hyperlink(p, url, text, size=None):
    part = p.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    h = OxmlElement("w:hyperlink"); h.set(qn("r:id"), rid)
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "1F4E79"); rpr.append(c)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rpr.append(u)
    if size:
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size*2))); rpr.append(sz)
    r.append(rpr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve"); r.append(t)
    h.append(r); p._p.append(h)

TOK = re.compile(r"(\*\*.+?\*\*|`[^`]+`|https?://[^\s)\]]+)")
def add_inline(p, text, size=None, color=None, bold=False, italic=False, font=None):
    for tok in TOK.split(text):
        if not tok: continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2]); r.bold = True
            _fmt(r, size, color, italic, font)
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); r.bold = bold
            _fmt(r, (size or 10.5) - 0.5, color, italic, "Consolas")
        elif tok.startswith("http"):
            m = re.match(r"^(.*?)([.,;:]*)$", tok); url, trail = m.group(1), m.group(2)
            add_hyperlink(p, url, url, size)
            if trail:
                r = p.add_run(trail); _fmt(r, size, color, italic, font)
        else:
            r = p.add_run(tok); r.bold = bold
            _fmt(r, size, color, italic, font)

def _fmt(r, size, color, italic, font):
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = RGBColor.from_string(color)
    if italic: r.italic = True
    if font:
        r.font.name = font

def para(text="", style=None, size=None, color=None, bold=False, italic=False, align=None, before=None, after=None, keep=False, font=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    add_inline(p, text, size=size, color=color, bold=bold, italic=italic, font=font)
    pf = p.paragraph_format
    if align: p.alignment = align
    if before is not None: pf.space_before = Pt(before)
    if after is not None: pf.space_after = Pt(after)
    if keep: pf.keep_with_next = True
    return p

def callout(text, fill=CREAM, bar=GOLD, size=10):
    p = doc.add_paragraph()
    add_inline(p, text, size=size, color=INK)
    p.paragraph_format.left_indent = Cm(0.35)
    p.paragraph_format.right_indent = Cm(0.2)
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(8)
    pPr = p._p.get_or_add_pPr()
    shade(pPr, fill)
    para_border(p, "left", bar, sz=24, space=8)
    return p

def add_field(p, instr, size=8.5, color=GREY):
    def fc(kind):
        r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
        c = OxmlElement("w:color"); c.set(qn("w:val"), color); rpr.append(c)
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size*2))); rpr.append(sz)
        r.append(rpr)
        f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), kind); r.append(f); return r
    p._p.append(fc("begin"))
    r = OxmlElement("w:r"); rpr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rpr.append(c)
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size*2))); rpr.append(sz)
    r.append(rpr)
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r.append(it)
    p._p.append(r)
    p._p.append(fc("separate"))
    r2 = OxmlElement("w:r"); rpr2 = OxmlElement("w:rPr")
    c2 = OxmlElement("w:color"); c2.set(qn("w:val"), color); rpr2.append(c2)
    sz2 = OxmlElement("w:sz"); sz2.set(qn("w:val"), str(int(size*2))); rpr2.append(sz2)
    r2.append(rpr2)
    t = OxmlElement("w:t"); t.text = "1"; r2.append(t); p._p.append(r2)
    p._p.append(fc("end"))

def page_break():
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)

def picture(path, width_cm, alt, keep=True):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4); p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = keep
    shape = p.add_run().add_picture(path, width=Cm(width_cm))
    shape._inline.docPr.set("descr", alt)
    shape._inline.docPr.set("title", alt[:60])
    return p

def caption(text, note=None):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(10)
    add_inline(p, text, size=9, color=NAVY, bold=True)
    if note:
        p2 = doc.add_paragraph(); p2.paragraph_format.space_after = Pt(10)
        add_inline(p2, note, size=8.5, color=GREY, italic=True)

def set_cell_shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); shade(tcPr, fill)

def set_cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    w = tcPr.find(qn("w:tcW"))
    if w is None:
        w = OxmlElement("w:tcW"); tcPr.insert(0, w)
    w.set(qn("w:w"), str(int(cm / 2.54 * 1440))); w.set(qn("w:type"), "dxa")

def table_props(tbl, widths_cm, header=True):
    t = tbl._tbl; tblPr = t.tblPr
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), str(int(sum(widths_cm)/2.54*1440))); tw.set(qn("w:type"), "dxa")
    old = tblPr.find(qn("w:tblW"))
    if old is not None: tblPr.remove(old)
    tblPr.append(tw)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    borders = OxmlElement("w:tblBorders")
    for side, val, col, sz in [("top","single",GREY,6),("left","nil",None,0),("bottom","single",GREY,6),("right","nil",None,0),("insideH","single",RULE,4),("insideV","nil",None,0)]:
        el = OxmlElement(f"w:{side}"); el.set(qn("w:val"), val)
        if val != "nil":
            el.set(qn("w:sz"), str(sz)); el.set(qn("w:space"), "0"); el.set(qn("w:color"), col)
        borders.append(el)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for side, v in [("top",50),("left",90),("bottom",50),("right",90)]:
        el = OxmlElement(f"w:{side}"); el.set(qn("w:w"), str(v)); el.set(qn("w:type"), "dxa"); mar.append(el)
    tblPr.append(mar)
    grid = t.find(qn("w:tblGrid"))
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w/2.54*1440)))
    for row in tbl.rows:
        for c, w in zip(row.cells, widths_cm):
            set_cell_width(c, w)
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        if header and i == 0:
            th = OxmlElement("w:tblHeader"); trPr.append(th)

def fill_cell(cell, text, size=9, bold=False, color=INK, align=None):
    cell.text = ""
    parts = text.split("\\n") if "\\n" in text else [text]
    p = cell.paragraphs[0]
    for k, part in enumerate(parts):
        if k > 0:
            p = cell.add_paragraph()
        add_inline(p, part, size=size, color=color, bold=bold)
        p.paragraph_format.space_after = Pt(1); p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.line_spacing = 1.08
        if align: p.alignment = align

def weights_for(rows, ncols):
    w = []
    for c in range(ncols):
        lens = [min(len(r[c]) if c < len(r) else 0, 160) for r in rows]
        mean = sum(lens)/max(1,len(lens)); mx = max(lens) if lens else 0
        w.append(max(7.0, 0.6*mean + 0.4*mx))
    return w

def data_table(rows, aligns, header_fill=NAVY, caption_text=None):
    ncols = len(rows[0])
    total = CONTENT_W
    w = weights_for(rows, ncols)
    widths = [total*x/sum(w) for x in w]
    def minw(k):
        words = []
        for r in rows:
            cell = r[k] if k < len(r) else ""
            words += re.findall(r"[^\s/]+", cell.replace("**", ""))
        longest = max((len(x) for x in words), default=4)
        return min(4.4, 0.23 * longest + 0.45)
    MINS = [minw(k) for k in range(ncols)]
    for _ in range(6):
        small = [k for k,x in enumerate(widths) if x < MINS[k]]
        if not small: break
        deficit = sum(MINS[k] - widths[k] for k in small)
        for k in small: widths[k] = MINS[k]
        big = [k for k in range(ncols) if k not in small]
        tot_big = sum(widths[k] for k in big)
        for k in big: widths[k] -= deficit * widths[k] / tot_big
    widths = [x * total / sum(widths) for x in widths]
    tbl = doc.add_table(rows=len(rows), cols=ncols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_props(tbl, widths, header=True)
    for i, r in enumerate(rows):
        for j in range(ncols):
            txt = r[j] if j < len(r) else ""
            cell = tbl.cell(i, j)
            al = WD_ALIGN_PARAGRAPH.RIGHT if (aligns[j] == "R" and i > 0) else None
            fsz = 8.5 if rows[0][0] in ("Unidade", "Identificação") else 9
            if i == 0:
                fill_cell(cell, txt, size=fsz, bold=True, color="FFFFFF", align=al)
                set_cell_shade(cell, header_fill)
            else:
                fill_cell(cell, txt, size=fsz, align=al)
                if i % 2 == 0: set_cell_shade(cell, ZEBRA)
    keep_rows = tbl.rows if len(rows) <= 10 else tbl.rows[-1:]
    for row in keep_rows:
        for cell in row.cells:
            for pp in cell.paragraphs: pp.paragraph_format.keep_with_next = True
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(6); sp.paragraph_format.keep_with_next = True
    return tbl

def parse_table(block):
    rows = []; aligns = None
    for ln in block:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if all(re.fullmatch(r":?-{3,}:?", c) for c in cells if c):
            aligns = ["R" if c.endswith(":") else "L" for c in cells]
            continue
        rows.append(cells)
    if aligns is None: aligns = ["L"]*len(rows[0])
    return rows, aligns

# ---------- figures ----------
FIG = {
 "f1": (f"{CHARTS}/v1-fig1-trilha-coortes.png", 16.0, "Figura 1 — Trilhas das duas coortes de 2027 (planejamento proposto, não executado)", "Coorte A: 2027 (1º), 2028 (2º), 2029 (3º). Coorte B: concluintes em 2027, com trilha de transição própria. Fonte: seção 5 e Anexo F do PDP."),
 "f2": (f"{CHARTS}/v1-fig2-ciclo-seis-movimentos.png", 13.5, "Figura 2 — Ciclo pedagógico em seis movimentos", "Desenho didático de cada encontro (seção 3.1). Não representa carga horária nem resultado."),
 "f3": (f"{CHARTS}/v1-fig3-avaliacao-recuperacao.png", 16.0, "Figura 3 — Planos de observação e recuperação paralela", "Instrumentos propostos (seção 12). A norma oficial de avaliação e recuperação está no PPC, pp. 38–40."),
 "f4": (f"{CHARTS}/v1-fig4-portoes-projeto.png", 16.0, "Figura 4 — Portões de um projeto da Agência-Escola (G0–G6)", "Seção 13.2. Cliente real não implica autorização para publicar, contratar, receber valores ou divulgar imagem."),
 "f5": (f"{CHARTS}/v1-fig5-governanca-portoes.png", 16.0, "Figura 5 — Portas de governança e estado atual", "Seção 19. Estados como registrados no documento; nenhuma etapa além de G1 está concluída."),
}
FIG_AFTER = [
 ("Proposta modular para adequação à grade real", ["f2"]),
 ("**Referência visual para DOCX:**", ["f1"]),
 ("O calendário trimestral, mínimo de instrumentos", ["f3"]),
 ("**Proibição expressa:**", ["f4"]),
 ("**RACI de referência:**", ["f5"]),
]
def maybe_figs(raw):
    for prefix, keys in FIG_AFTER:
        if raw.startswith(prefix):
            for k in keys:
                path, wcm, cap, note = FIG[k]
                picture(path, wcm, cap)
                caption(cap, note)

# ---------- header / footer ----------
hdr = sec.header
hp = hdr.paragraphs[0]; hp.text = ""
add_inline(hp, "PLANEJAMENTO PEDAGÓGICO TRIENAL 2027–2029  ·  CURSO TÉCNICO EM MARKETING  ·  VERSÃO 1.0  ·  NÃO HOMOLOGADO", size=7.5, color=GREY)
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
para_border(hp, "bottom", GOLD, sz=6, space=4)

ftr = sec.footer
fp = ftr.paragraphs[0]; fp.text = ""
fp.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT)
add_inline(fp, "EEB Araújo Figueiredo · Urubici/SC  |  20/11/2026  |  Circulação: Direção, até deliberação", size=7.5, color=GREY)
r = fp.add_run("\tPágina "); r.font.size = Pt(7.5); r.font.color.rgb = RGBColor.from_string(GREY)
add_field(fp, "PAGE", size=7.5)
r = fp.add_run(" de "); r.font.size = Pt(7.5); r.font.color.rgb = RGBColor.from_string(GREY)
add_field(fp, "NUMPAGES", size=7.5)
# first-page footer intentionally blank (cover)

cp = doc.core_properties
cp.title = "Planejamento Pedagógico Trienal — Curso Técnico em Marketing 2027–2029 — Versão 1.0 (não homologada)"
cp.author = "Araujo Innovation Lab"; cp.subject = "Proposta para deliberação da Direção"
cp.keywords = "PPC SED 00178627/2026; PDP V5.1; planejamento trienal; proposta não homologada"

# ---------- COVER ----------
L = lines
para("", after=40)
para("EEB ARAÚJO FIGUEIREDO  •  URUBICI, SANTA CATARINA", size=9.5, color=GOLD, bold=True, after=14)
t = doc.add_paragraph(style="Heading 1")
t.paragraph_format.space_after = Pt(10)
add_inline(t, "PLANEJAMENTO PEDAGÓGICO TRIENAL", size=28, color=NAVY)
para_border(t, "bottom", GOLD, sz=12, space=10)
para("CURSO TÉCNICO EM MARKETING — 2027–2029", size=13, color=GREY, after=18, font=HFONT, italic=True)
para("Conhecer · Experimentar · Criar · Avaliar · Evoluir", size=10.5, color=INK, after=16)
para("Agência-Escola | Aprendizagem aplicada | Comunicação e design | IA crítica | Projetos integradores | Evidências de evolução", size=9.5, color=INK, after=16)
para("Versão 1.0 — 20/11/2026 — Não homologada.", size=10.5, color=NAVY, bold=True, after=26)
badge = doc.add_paragraph()
add_inline(badge, "PROPOSTA PARA DELIBERAÇÃO DA DIREÇÃO — NÃO HOMOLOGADA", size=10, color="FFFFFF", bold=True)
badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
badge.paragraph_format.space_before = Pt(6); badge.paragraph_format.space_after = Pt(6)
shade(badge._p.get_or_add_pPr(), NAVY)
badge.paragraph_format.left_indent = Cm(0.2); badge.paragraph_format.right_indent = Cm(0.2)
para("Urubici/SC  ·  Período 2027–2029  ·  Versão 1.0", size=10, color=INK, align=WD_ALIGN_PARAGRAPH.CENTER, before=10, after=30)
callout(next(l for l in L if l.startswith("> **PROPOSTA INSTITUCIONAL")).lstrip("> ").strip(), size=9.5)
page_break()

# ---------- DOCUMENT CONTROL ----------
h = doc.add_paragraph(style="Heading 1"); add_inline(h, "Ficha de controle do documento", size=19, color=NAVY)
para("Metadados de controle. Nenhuma aprovação, assinatura ou homologação está registrada neste documento.", size=9.5, color=GREY, italic=True, after=10)
callout(next(l for l in L if l.startswith("> **PROPOSTA INSTITUCIONAL")).lstrip("> ").strip(), size=9.5)
callout(next(l for l in L if l.startswith("> **Hierarquia documental")).lstrip("> ").strip(), size=9.5)
ctrl = [
 ["Campo", "Registro"],
 ["Documento", "Planejamento Pedagógico Trienal — Curso Técnico em Marketing 2027–2029"],
 ["Versão / data", "1.0 · 20/11/2026 · proposta documental"],
 ["Status", "Proposta para deliberação da Direção — não homologada. Nenhuma assinatura ou aprovação registrada."],
 ["Fonte oficial", "PPC EMIEP Técnico em Marketing, SED/SC, processo 00178627/2026, assinado em 03/09/2026 (SHA-256 31fb64c5…f58, arquivo PPC-EMIEP-Marketing.pdf)."],
 ["Fonte metodológica", "Plano Diretor Pedagógico V5.1, commit af50db77a8722c9d45e5a3b4a268eca5c3e62dbb no branch principal (documentalmente promovido; não homologado)."],
 ["Hierarquia documental", "PPC oficial → PPP/PGE e regimentos (a confirmar) → PDP V5.1 → este planejamento → planos anuais e docentes (a detalhar após aceite)."],
 ["Expediente proposto", "Profª Letícia Meurer — docência titular (designação a confirmar). Prof. Coordenador José Alexandre F. de Santana — orientação e integração (designação a confirmar). Nomes reproduzidos da fonte: ver nota de governança no relatório de QA."],
 ["Público-alvo", "Direção escolar. Docentes serão consultados somente após aceite (PDP §12.1; Anexo I1)."],
 ["Portas de decisão", "Revisão documental (G1); deliberação da Direção (G2); pactuação docente e prontidão (G3); autorização por projeto (G4); acompanhamento (G5)."],
 ["Identidade visual", "Sem logotipo, brasão ou assinatura: nenhum ativo oficial validado. Cores de apoio provisórias, herdadas do PDP V5.1."],
]
data_table(ctrl, ["L","L"])
para(next(l for l in L if l.startswith("**Referência editorial:**")).strip(), size=8.5, color=GREY, italic=True, before=6)
para(next(l for l in L if l.startswith("**Créditos funcionais propostos para expediente:**")).strip(), size=8.5, color=INK, before=4)

# ---------- CONTENTS ----------
h = doc.add_paragraph(style="Heading 1"); add_inline(h, "Sumário", size=19, color=NAVY)
para("Numeração de página refere-se às páginas deste arquivo. Sumário gerado a partir das seções de nível 1.", size=9, color=GREY, italic=True, after=8)
TOC_ENTRIES = []  # filled below, placeholders preserved for two-pass stability

# ---------- BODY ----------
body = L[next(i for i,l in enumerate(L) if l.startswith('## ORIENTA')):]
blocks = []
i = 0
while i < len(body):
    ln = body[i]
    if ln.strip() == "":
        i += 1; continue
    if ln.startswith("## "):
        blocks.append(("h1", ln[3:].strip(), ln)); i += 1; continue
    if ln.startswith("# "):
        blocks.append(("h1", ln[2:].strip(), ln)); i += 1; continue
    if ln.startswith("### "):
        blocks.append(("h2", ln[4:].strip(), ln)); i += 1; continue
    if ln.startswith("> "):
        grp = []; raw = ln
        while i < len(body) and body[i].startswith(">"):
            grp.append(body[i].lstrip("> ").strip()); i += 1
        blocks.append(("quote", grp, raw)); continue
    if ln.startswith("|"):
        grp = []; raw = ln
        while i < len(body) and body[i].startswith("|"):
            grp.append(body[i]); i += 1
        blocks.append(("table", grp, raw)); continue
    if ln.startswith("- "):
        blocks.append(("bullet", ln[2:].strip(), ln)); i += 1; continue
    if re.match(r"^\d+\. ", ln):
        m = re.match(r"^(\d+)\. (.*)$", ln)
        blocks.append(("num", (m.group(1), m.group(2).strip()), ln)); i += 1; continue
    if ln.startswith("• "):
        blocks.append(("dot", ln[2:].strip(), ln)); i += 1; continue
    if ln.startswith("□ "):
        blocks.append(("box", ln[2:].strip(), ln)); i += 1; continue
    blocks.append(("p", ln.strip(), ln)); i += 1

# closing decision page is appended after section 17 (see end)
toc_items = []
first_h1 = True
for bi, (kind, val, raw) in enumerate(blocks):
    if kind == "h1":
        if first_h1:
            first_h1 = False
        p = doc.add_paragraph(style="Heading 1")
        p.paragraph_format.page_break_before = True
        add_inline(p, val, size=19, color=NAVY)
        toc_items.append(val)
    elif kind == "h2":
        p = doc.add_paragraph(style="Heading 2"); add_inline(p, val, size=13.5, color=NAVY)
    elif kind == "quote":
        for q in val: callout(q, size=9.5)
        sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(2)
    elif kind == "table":
        rows, aligns = parse_table(val)
        data_table(rows, aligns)
    elif kind == "bullet":
        p = doc.add_paragraph(style="List Bullet"); add_inline(p, val, size=10.5, color=INK)
        p.paragraph_format.space_after = Pt(4)
        nxt = blocks[bi+1][0] if bi+1 < len(blocks) else None
        if nxt == "bullet": p.paragraph_format.keep_with_next = True
    elif kind == "num":
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75); p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.space_after = Pt(4)
        add_inline(p, f"{val[0]}.\t", size=10.5, color=NAVY, bold=True)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.75))
        add_inline(p, val[1], size=10.5, color=INK)
    elif kind == "dot":
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6); p.paragraph_format.first_line_indent = Cm(-0.45)
        p.paragraph_format.space_after = Pt(4)
        add_inline(p, "•\t" + val, size=10.5, color=INK)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.6))
    elif kind == "box":
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.7); p.paragraph_format.first_line_indent = Cm(-0.7)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(0.7))
        add_inline(p, "☐\t" + val, size=10.5, color=INK)
    else:  # p
        p = doc.add_paragraph(); add_inline(p, val, size=10.5, color=INK)
        p.paragraph_format.space_after = Pt(6)
        if bi+1 < len(blocks) and blocks[bi+1][0] == "table":
            p.paragraph_format.keep_with_next = True
    maybe_figs(raw)

# (no closing page in V1: the decision minute is section 20 of the source)

# ---------- TOC (static, with page numbers from render) ----------
toc_anchor = None
# insert TOC paragraphs right after the "Sumário" heading: find that paragraph
def insert_after(par, new_par_el):
    par._p.addnext(new_par_el)
sumario_p = None
for p in doc.paragraphs:
    if p.text.strip() == "Sumário":
        sumario_p = p; break
cursor = sumario_p._p
all_toc = toc_items
for n, item in enumerate(all_toc, start=1):
    p = doc.add_paragraph()
    tabs = p.paragraph_format.tab_stops
    tabs.add_tab_stop(Cm(CONTENT_W), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0)
    pg = str(pagemap.get(item, "—"))
    add_inline(p, item, size=10, color=INK)
    r = p.add_run("\t" + pg); r.font.size = Pt(10); r.font.color.rgb = RGBColor.from_string(NAVY)
    cursor.addnext(p._p); cursor = p._p

# no-op for header-only mapping
doc.save(OUTDOCX)
json.dump(all_toc, open(OUTDOCX + ".toc.json", "w"), ensure_ascii=False, indent=1)
print("saved", OUTDOCX, "toc entries:", len(all_toc), "blocks:", len(blocks))
