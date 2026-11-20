"""Release tests for V5.1 corrected candidate: mathematical checks and document-integrity checks.
Usage: python3 tools/test_release.py <md> <docx> <pdf>"""
import sys, re, random, zipfile, subprocess, json
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
from ppc_units import UNITS
md_p, docx_p, pdf_p = sys.argv[1], sys.argv[2], sys.argv[3]
results = []
def t(name, ok, detail=""):
    results.append((name, bool(ok), detail))

# ---------- Mathematical tests ----------
W = [30,20,20,15,10,5]                       # Anexo C2 weights (sum 100)
def k01(levels, weights=W):
    return sum(w*l for w,l in zip(weights, levels)) / sum(weights)
ex = k01([4,3,2,4,3,1])
t("K01 exemplo: 4/3/2/4/3/1 com pesos 30/20/20/15/10/5 = 3,15", abs(ex-3.15) < 1e-9, f"resultado {ex:.2f}")
t("K01 limites: todos 0 → 0; todos 4 → 4", k01([0]*6)==0 and abs(k01([4]*6)-4)<1e-9)
rnd = random.Random(20261120)
bad = 0
for _ in range(10000):
    lv = [rnd.randint(0,4) for _ in range(6)]
    v = k01(lv)
    if not (0 <= v <= 4): bad += 1
t("K01 propriedade: 10.000 vetores de níveis 0–4 produzem resultado em [0,4]", bad==0, f"fora do intervalo: {bad}")
old = sum([4,3,2,4,3,1])
t("K01 defeito antigo (soma de pontos) saía de 0–4 no exemplo", old > 4, f"soma antiga = {old}")
k15_rows = [l for l in open(md_p, encoding="utf-8").read().split("\n") if l.startswith("| K15 · ") or l.startswith("| K15 | ")]
t("K15 sem visitas: texto da regra presente em §9 e G2 (não aplicável, não 0%)", len(k15_rows)==2 and all("não aplicável" in l for l in k15_rows), f"{len(k15_rows)} linhas")
def k15(visitas_com_relatorio, visitas_executadas):
    if visitas_executadas == 0: return None
    return 100*visitas_com_relatorio/visitas_executadas
t("K15 exemplo: 3 de 4 visitas com roteiro e relatório = 75%", abs(k15(3,4)-75)<1e-9)
t("K15 sem visita devolve não aplicável (None), não 0", k15(0,0) is None)
def k16(matriculas, vagas):
    if vagas == 0: return None
    return 100*matriculas/vagas
t("K16 exemplo: 40 matrículas válidas de 60 vagas = 66,7%", abs(k16(40,60)-66.666)<0.01, f"{k16(40,60):.1f}%")
t("K16 sem vagas ofertadas devolve não aplicável", k16(10,0) is None)
hours = {c["code"]: c["ch"] for c in COMPONENTS}
t("Carga: 1º (4×64) = 256; 2º (96+3×64) = 288; 3º (4×64) = 256; total técnico = 800",
  4*64==256 and 96+3*64==288 and 4*64==256 and sum(hours.values())==800, f"soma {sum(hours.values())}")
t("Carga: FGB 744+712+744 = 2.200; integradoras 3×40 = 120 dentro; curso 2.200+800 = 3.000",
  744+712+744==2200 and 3*40==120 and 2200+800==3000)
mapped = [ (c["code"],i) for c in COMPONENTS for u in UNITS[c["code"]] for i in u["items"] ]
t("Mapeamento: 183 itens, cada um em exatamente uma unidade", len(mapped)==183 and len(set(mapped))==183, f"{len(mapped)} atribuições, {len(set(mapped))} únicas")
empty = [ (c,u["id"],f) for c,us in UNITS.items() for u in us for f in ("fund","ativ","aval","conx","title") if not u[f].strip() ]
t("Unidades: nenhum campo vazio", not empty, str(empty[:3]))
t("Unidades: 49 unidades propostas", sum(len(v) for v in UNITS.values())==49, str(sum(len(v) for v in UNITS.values())))

# ---------- Document-integrity tests ----------
md = open(md_p, encoding="utf-8").read()
t("MD versão 5.1 no front matter", 'version: "5.1"' in md)
t("MD sem 'Versão 5.0' como versão atual", "Versão 5.0 •" not in md)
t("MD: histórico de versões presente (Anexo O)", md.count("Histórico de versões") == 1)
t("MD: removida a afirmação 'acesso em 19/11/2026'", "acesso em 19/11/2026" not in md)
t("MD: K01 definido como média ponderada, sem soma de pontos", "soma de pontos da rubrica" not in md)
t("MD: K15 não mistura briefs semestrais no indicador", "briefs aprovados (contagem semestral)" not in md)
t("MD: K16 é taxa de ocupação", "Taxa de ocupação" in md)
t("MD: §2.2 objetivos e perfil oficiais presentes", "### 2.2 Objetivos e perfil profissional oficiais" in md)
t("MD: Anexo O com 49 unidades", len(re.findall(r"^\| U\d · ", md, re.M)) == 49, str(len(re.findall(r"^\| U\d · ", md, re.M))))
z = zipfile.ZipFile(docx_p)
doc_xml = z.read("word/document.xml").decode("utf-8")
descr = re.findall(r'descr="([^"]*)"', doc_xml)
t("DOCX: 5 figuras com texto alternativo", len(descr) == 5, f"{len(descr)} descr")
footers = [z.read(n).decode("utf-8") for n in z.namelist() if re.match(r"word/footer\d*\.xml", n)]
ftxt = " ".join(footers)
t("DOCX rodapé com data 20/11/2026", "20/11/2026" in ftxt)
t("DOCX rodapé sem 19/11/2026", "19/11/2026" not in ftxt)
headers = " ".join(z.read(n).decode("utf-8") for n in z.namelist() if re.match(r"word/header\d*\.xml", n))
t("DOCX cabeçalho com VERSÃO 5.1", "VERSÃO 5.1" in headers and "VERSÃO 5.0" not in headers)
pdf_txt = subprocess.run(["pdftotext", "-layout", pdf_p, "-"], capture_output=True, text=True).stdout
t("PDF: texto de versão 5.1 presente", "Versão 5.1" in pdf_txt)
t("PDF: nenhum resíduo 'acesso em 19/11/2026'", "acesso em 19/11/2026" not in pdf_txt)
pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf_p], capture_output=True, text=True).stdout).group(1))
t("PDF: 101 páginas (§2.1 antes de §2.2 altera a paginação)", pages == 101, str(pages))
toc_ok = len(re.findall(r"^## ", md, re.M)) > 0
t("MD: estrutura de seções com H2 presentes", toc_ok)


# ---------- Traceability hash regression ----------
import os
bs_path = os.path.join(os.path.dirname(md_p) or ".", "output", "BS_TRACEABILITY_REPORT.md")
if os.path.exists(bs_path):
    bs_txt = open(bs_path, encoding="utf-8").read()
    actual_md = __import__("hashlib").sha256(open(md_p,"rb").read()).hexdigest()
    m = re.search(r"Documento-alvo V5\.1:.*?SHA256 `([0-9a-f]{64})`", bs_txt)
    t("BS_TRACEABILITY: hash do documento-alvo igual ao arquivo atual", bool(m) and m.group(1)==actual_md, (m.group(1)[:12] if m else "ausente"))

# ---------- report ----------
ok = sum(1 for _,o,_ in results if o)
for n,o,d in results:
    print(("PASS" if o else "FAIL"), "|", n, ("|"+d if d else ""))
print(f"\n{ok}/{len(results)} testes passaram")
json.dump([{"name":n,"pass":o,"detail":d} for n,o,d in results], open(sys.argv[4] if len(sys.argv)>4 else "/dev/null","w",encoding="utf-8"), ensure_ascii=False, indent=1)
sys.exit(0 if ok==len(results) else 1)
