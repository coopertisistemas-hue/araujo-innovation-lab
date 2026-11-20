"""Second PPC reconciliation pass (v4): §12.2 institutional facts from PPC I, VIII, XI and
Anexo N (§3.1 traceability matrix) generated from the Anexo D cadernos and tools/ppc_data.py."""
import sys, re, json
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
src, out, logp = sys.argv[1], sys.argv[2], sys.argv[3]
L = open(src, encoding="utf-8").read()
log = []
def sub(old, new, label):
    global L
    n = L.count(old); assert n == 1, (label, n)
    L = L.replace(old, new); log.append(label)

sub("- O plano de ingressantes 2027–2029 não equivale ao plano de transição da turma já concluinte em 2027.",
    "- O plano de ingressantes 2027–2029 não equivale ao plano de transição da turma já concluinte em 2027.\n"
    "- Requisitos de acesso (PPC, VIII): estar matriculado na 1ª série do Ensino Médio, com Ensino Fundamental concluído. Prazo máximo de integralização: 5 anos (PPC, XI). Estágio: não há (PPC, I). Infraestrutura mínima (PPC, IX): biblioteca com acervo físico ou virtual específico, datashow, notebook e laboratórios de informática, redes, manutenção e tecnologias.",
    "V4-01 §12.2 requisitos de acesso, prazo, estágio e infraestrutura (PPC I, VIII, IX, XI)")

def field(block, label):
    m = re.search(r"\*\*%s:?\*\*:? ?(.*?)(?=\n\n|\Z)" % re.escape(label), block, re.S)
    return m.group(1).strip() if m else None

rows_all = []
for c in COMPONENTS:
    m = re.search(r"^### %s\. .*$" % c["code"], L, re.M)
    end = L.find("\n### ", m.end()); end2 = L.find("\n## ", m.end())
    ends = [e for e in (end, end2) if e != -1]
    blk = L[m.start(): min(ends)] if ends else L[m.start():]
    lab = field(blk, "Laboratório e tecnologia")
    prod = field(blk, "Produtos verificáveis")
    prova = field(blk, "Prova individual")
    rub = field(blk, "Rubrica")
    val = field(blk, "Validação")
    conx = field(blk, "Conexão horizontal")
    conxl = field(blk, "Conexão longitudinal")
    proj = field(blk, "Projeto de aprendizagem")
    seq = blk.split("**Sequência conceitual e laboratorial sugerida:**")[1].split("\n\n**")[0].strip() if "**Sequência conceitual" in blk else None
    E, A, N, P = "EXPLÍCITO_PPC", "APLICAÇÃO_DIDÁTICA", "ENRIQUECIMENTO_PROPOSTO", "PENDENTE_DE_VALIDAÇÃO"
    rows = [
      ("Série e carga oficial", f"{c['ano']}º ano · {c['ch']} h (PPC, Matriz, p. 4)", E),
      ("Objetos do conhecimento", f"Lista oficial no caderno {c['code']} ({len(c['objetos'])} itens; itens de habilidade marcados)", E),
      ("Página do PPC", f"pp. {c['pages']} (paginação textual do PPC)", E),
      ("Objetivo do encontro", "A definir no plano de aula do docente habilitado; não há objetivo por encontro no PPC nem no Markdown", P),
      ("Fundamento conceitual", c["objetivo"], E),
      ("Prática de laboratório", lab or "—", A),
      ("Ferramenta candidata", ((lab or "—").split(" Em todos os encontros")[0]) + " (ferramentas sujeitas a licença, idade e autorização)", A),
      ("Projeto associado", proj or "—", A),
      ("Tarefa individual", "Produção individual e defesa individual conforme a rubrica do caderno (ver Prova individual)", A),
      ("Prova", prova or "—", A),
      ("Rubrica", rub or "—", A),
      ("Segurança e IA", "Uso de IA somente com registro e checagem; protocolo no Anexo H1–H2 e matriz H5; dados e imagens somente com autorização", N),
      ("Conexão interdisciplinar", " ".join(x for x in [conx, conxl] if x) or "—", A),
      ("Evidência em portfólio", prod or "—", A),
    ]
    rows_all.append((c, rows))

anexo = ["## ANEXO N | Matriz-mestra de rastreabilidade curricular (PPC × Markdown)", "",
 "Esta matriz cumpre o §3.1 da matriz de brainstorming: uma tabela por componente técnico, com cada campo marcado como **EXPLÍCITO_PPC** (transcrição do PPC), **APLICAÇÃO_DIDÁTICA** (proposta do caderno do Anexo D), **ENRIQUECIMENTO_PROPOSTO** (acréscimo fora do PPC) ou **PENDENTE_DE_VALIDAÇÃO** (depende do docente habilitado). Os conteúdos são copiados dos cadernos do Anexo D e do PPC; nada foi acrescentado além do que já consta do documento.", ""]
for c, rows in rows_all:
    anexo += [f"### N-{c['code']}. {c['name']} — {c['ano']}º ano, {c['ch']} h", "",
              "| Campo | Conteúdo | Marcação |", "| --- | --- | --- |"]
    for f, v, mk in rows:
        vv = v.replace("|", "/").replace("\n", " ")
        anexo.append(f"| {f} | {vv} | {mk} |")
    anexo.append("")
L = L.rstrip("\n") + "\n\n" + "\n".join(anexo) + "\n"
log.append("V4-02 Anexo N matriz-mestra §3.1 (12 tabelas, 14 campos, marcação por campo)")
open(out, "w", encoding="utf-8").write(L)
json.dump(log, open(logp,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("applied", len(log), "edits; lines", len(L.split("\n")))
