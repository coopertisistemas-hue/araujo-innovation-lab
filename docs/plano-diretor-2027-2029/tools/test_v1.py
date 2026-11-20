"""Release tests for the Planejamento Pedagógico Trienal V1 edition.
Usage: python3 tools/test_v1.py <md> <docx> <pdf> <toc_map_json> <out_json>"""
import sys, re, zipfile, subprocess, json, hashlib
md_p, docx_p, pdf_p, map_p, out_p = sys.argv[1:6]
md = open(md_p, encoding="utf-8").read()
R = []
def t(name, ok, d=""): R.append((name, bool(ok), d))
# structure
secs = re.findall(r"^## (\d+) \|", md, re.M)
t("Estrutura: seções 1–21 presentes, em ordem", secs == [str(i) for i in range(1,22)], str(secs[-3:]))
annex = re.findall(r"^## ANEXO ([A-J]) \|", md, re.M)
t("Anexos: A–J presentes, em ordem", annex == list("ABCDEFGHIJ"), "".join(annex))
t("Divisor 'ANEXOS OPERACIONAIS PROPOSTOS' presente", "# ANEXOS OPERACIONAIS PROPOSTOS" in md)
# curriculum
comps = re.findall(r"^### (D\d\d) \|", md, re.M)
t("Componentes: 12 técnicos (D01–D12) em §10", comps == [f"D{i:02d}" for i in range(1,13)], str(len(comps)))
units = sorted(set(re.findall(r"\bD\d\d-U\d\b", md)))
t("Unidades: 49 referenciadas (D01–D12)", len(units) == 49, f"{len(units)} unidades")
hours = {"3.000 h": "3.000 h" in md, "2.200 h": "2.200 h" in md, "800 h": "800 h" in md, "120 h integradoras": "120 h" in md}
t("Cargas: 3.000 h, 2.200 h, 800 h e 120 h citadas", all(hours.values()), str(hours))
t("Cargas por série: 256 h, 288 h e 256 h", all(x in md for x in ["256 h", "288 h"]))
# cohorts
t("Duas coortes de 2027 distintas (A ingressante, B concluinte)", "COORTE A" in md.upper() and "COORTE B" in md.upper() and "Coorte B — concluintes em 2027" in md or "3º ANO CONCLUINTE" in md)
# formulas and status
t("K01 em média ponderada 0–4", "soma(peso × nível) ÷ soma(pesos válidos)" in md)
t("K15 não aplicável sem visitas", "NÃO APLICÁVEL" in md)
t("K16 é taxa de ocupação", "taxa de ocupação" in md.lower())
t("Status: NÃO HOMOLOGADA presente", "NÃO HOMOLOGADA" in md.upper())
def _positive_claims(text):
    out=[]
    for m in re.finditer(r"(?i)\b(aprovado pela direção|homologado pela direção|aprovada pela direção)\b", text):
        before = text[max(0, m.start()-8):m.start()].lower()
        if "não " not in before and "nao " not in before:
            out.append(m.group(0))
    return out
t("Nenhuma afirmação de aprovação alcançada (negações ignoradas)", not _positive_claims(md), str(_positive_claims(md)))
t("Dados sem números fictícios: 'SEM DADOS' presente", "SEM DADOS" in md)
names = {"Letícia Meurer": "Letícia Meurer", "José Alexandre": "José Alexandre"}
flagged = {}
for k, n in names.items():
    lines = [l for l in md.split("\n") if n in l]
    flagged[k] = all("a confirmar" in l.lower() or "a validar" in l.lower() for l in lines) and bool(lines)
t("Nomes de docentes só aparecem com 'a confirmar' (governança)", all(flagged.values()), str(flagged))
# document integrity
z = zipfile.ZipFile(docx_p)
doc = z.read("word/document.xml").decode("utf-8")
descr = re.findall(r'descr="([^"]*)"', doc)
t("DOCX: 5 figuras com texto alternativo", len(descr) == 5, f"{len(descr)}")
footers = " ".join(z.read(n).decode("utf-8") for n in z.namelist() if re.match(r"word/footer\d*\.xml", n))
headers = " ".join(z.read(n).decode("utf-8") for n in z.namelist() if re.match(r"word/header\d*\.xml", n))
t("DOCX: rodapé com 20/11/2026", "20/11/2026" in footers)
t("DOCX: cabeçalho com NÃO HOMOLOGADO", "NÃO HOMOLOGADO" in headers)
pdf_txt = subprocess.run(["pdftotext", pdf_p, "-"], capture_output=True, text=True).stdout
pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf_p], capture_output=True, text=True).stdout).group(1))
t("PDF: número de páginas registrado", pages > 0, str(pages))
tocmap = json.load(open(map_p, encoding="utf-8"))
t("Sumário: todas as entradas com número de página", all(v is not None for v in tocmap.values()), f"{len(tocmap)} entradas")
t("PDF: texto de expediente presente", "Créditos funcionais propostos" in pdf_txt or "Créditos funcionais" in pdf_txt)
ok = sum(1 for _,o,_ in R if o)
for n,o,d in R: print("PASS" if o else "FAIL", "|", n, ("|"+d if d else ""))
print(f"\n{ok}/{len(R)} testes passaram")
json.dump([{"name":n,"pass":o,"detail":d} for n,o,d in R], open(out_p,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
sys.exit(0 if ok==len(R) else 1)
