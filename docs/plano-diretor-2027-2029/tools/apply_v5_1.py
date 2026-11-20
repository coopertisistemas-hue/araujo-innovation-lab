"""Build the controlled document V5.1 from the preserved V5.0 source (not in place).
Usage: python3 tools/apply_v5_1.py <V5.0 md> <V5.1 md> <log json>"""
import sys, re, json
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
from ppc_units import UNITS
src, out, logp = sys.argv[1], sys.argv[2], sys.argv[3]
L = open(src, encoding="utf-8").read()
log = []
def sub(old, new, label):
    global L
    n = L.count(old); assert n == 1, (label, n)
    L = L.replace(old, new); log.append(label)

sub('version: "5.0"', 'version: "5.1"', "V51-01 front matter versão 5.1")
sub("Versão 5.0 • 19/11/2026 • Proposta para deliberação da Direção — NÃO homologada",
    "Versão 5.1 • 20/11/2026 • Proposta para deliberação da Direção — NÃO homologada. "
    "Histórico: V5.0 (19/11/2026), fonte preservada em `Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md`; "
    "V5.1 (20/11/2026): unidades de aprendizagem propostas e mapa dos 183 itens oficiais (Anexo O), portas de decisão separadas (§12.4), habilitação como pré-requisito de implantação (Anexo I4), K15 sujeito à deliberação.",
    "V51-02 linha de versão e histórico")
sub("| K15 · Visitas e parcerias | Visitas com roteiro e relatório ÷ visitas executadas × 100; briefs aprovados (contagem semestral) | Autorização, roteiro e relatório | Orientador + direção; Anual (decisão da revisão: ver DELTA_REPORT) |",
    "| K15 · Visitas e parcerias | Visitas com roteiro e relatório ÷ visitas executadas × 100; briefs aprovados (contagem semestral) | Autorização, roteiro e relatório | Orientador + direção; Anual (proposta sujeita à deliberação da Direção) |",
    "V51-03 K15 anual como proposta sujeita à deliberação")
sub("Os requisitos abaixo são transcritos do PPC (pp. 10–33). Um professor orientador só ministra um componente oficialmente se atender ao requisito e for designado formalmente.",
    "Os requisitos abaixo são transcritos do PPC (pp. 10–33). Um professor orientador só ministra um componente oficialmente se atender ao requisito e for designado formalmente.\n\n"
    "**Pré-requisito de implantação** (porta de prontidão operacional, não de revisão documental): a escola confirma o quadro docente habilitado e designado antes de qualquer oferta do componente.",
    "V51-04 I4 habilitação como pré-requisito de implantação")

gates = """### 12.4 Portas separadas de decisão

Esta proposta separa três portas, para que pendências operacionais não bloqueiem a revisão documental:

| Porta | Pergunta | Situação nesta versão |
| --- | --- | --- |
| 1. Revisão documental | A proposta está suficientemente fundamentada, coerente, rastreável e internamente consistente para análise independente e deliberação? | Atendida nesta versão: 183 itens oficiais mapeados (Anexo O), matriz-mestra (Anexo N), 28 requisitos do brainstorming rastreados (BS_TRACEABILITY_REPORT). |
| 2. Aprovação institucional | A Direção aprovou a proposta e suas condições? | Pendente. Decisão da Direção (§16). Nenhuma aprovação está presumida neste documento. |
| 3. Prontidão operacional | Docentes habilitados e designados, planos detalhados validados, recursos, salvaguardas de dados e autorizações estão prontos? | Pendente. Depende de habilitação (Anexo I4), inventário (Anexo J1), validação docente após o aceite (Anexo I1) e autorizações. Não bloqueia a porta 1. |

"""
sub("## 13 | Perfil de egresso por competência e matriz de progressão", gates + "## 13 | Perfil de egresso por competência e matriz de progressão", "V51-05 §12.4 portas separadas de decisão")

# Anexo O — proposed learning units and item map
lv = """## ANEXO O | Unidades de aprendizagem propostas e mapa dos 183 itens oficiais

Este anexo organiza os objetos oficiais do PPC em unidades de aprendizagem propostas. Cada item oficial recebe três marcas, separadas por nível de evidência:

- **N1 — Requisito oficial identificado e rastreável:** o item consta do PPC e foi transcrito no Anexo D (pp. indicadas).
- **N2 — Tratamento pedagógico proposto demonstrado:** o item está atribuído a uma unidade, com fundamento conceitual, atividade de laboratório representativa, evidência de avaliação e conexão interdisciplinar descritos nas tabelas abaixo. Não é plano de aula nem validação docente.
- **N3 — Plano detalhado pendente de validação docente:** objetivo por encontro e plano de aula, a cargo do docente habilitado e designado, após o aceite da Direção (Anexo I1). Esta marca se aplica a todos os itens.

Um mesmo item não aparece em duas unidades. Itens relacionados compartilham a atividade da unidade; não há um plano de aula por item.

"""
def rng(xs):
    xs = sorted(xs); out=[]; s=p=xs[0]
    for x in xs[1:]:
        if x==p+1: p=x; continue
        out.append((s,p)); s=p=x
    out.append((s,p))
    return ", ".join(f"D{{c}}.{a+1:02d}" if a==b else f"D{{c}}.{a+1:02d}–{b+1:02d}" for a,b in out)
body = []
for c in COMPONENTS:
    body.append(f"### O-{c['code']}. {c['name']} — unidades propostas ({len(c['objetos'])} itens)\n")
    body.append("| Unidade | Itens (identificação) | Fundamento conceitual | Atividade de laboratório representativa | Evidência de avaliação | Conexão interdisciplinar |")
    body.append("| --- | --- | --- | --- | --- | --- |")
    for u in UNITS[c["code"]]:
        ids = rng(u["items"]).replace("D{c}", c["code"])
        body.append(f"| {u['id']} · {u['title']} | {ids} | {u['fund']} | {u['ativ']} | {u['aval']} | {u['conx']} |")
    body.append("")
body.append("### O-Mapa. Mapa dos 183 itens (N1, N2, N3)\n")
body.append("| Identificação | Item oficial (PPC) | Origem | Unidade | N1 | N2 | N3 |")
body.append("| --- | --- | --- | --- | --- | --- | --- |")
n=0
for c in COMPONENTS:
    where = {}
    for u in UNITS[c["code"]]:
        for k in u["items"]: where[k] = u
    for k, o in enumerate(c["objetos"]):
        u = where[k]; n+=1
        origem = "habilidade" if o.endswith("[hab.]") else "objeto"
        label = o.replace(" [hab.]","")
        body.append(f"| {c['code']}.{k+1:02d} | {label} | {origem} | {u['id']} | sim | sim | pendente |")
assert n == 183, n
body.append("Origem: **objeto** = objeto do conhecimento do PPC; **habilidade** = item da coluna Habilidades do PPC. Componentes: D01 Administração de Marketing; D02 Design Gráfico; D03 Comportamento do Consumidor; D04 Legislação Aplicada; D05 Desenvolvimento de Produtos; D06 Marketing Digital; D07 Estratégia de Marketing; D08 Comunicação em Marketing; D09 Gestão de Eventos; D10 Promoção de Vendas; D11 Logística; D12 Práticas de Marketing. Unidades: títulos na tabela de cada componente acima.")
body.append("")
anex = lv + "\n".join(body) + "\n"
L = L.rstrip("\n") + "\n\n" + anex
log.append("V51-06 Anexo O: unidades propostas por componente e mapa dos 183 itens (N1, N2, N3)")
open(out, "w", encoding="utf-8").write(L)
json.dump(log, open(logp,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok edits", len(log), "lines", len(L.split("\n")))
