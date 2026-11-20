"""Apply the PPC reconciliation (v3) to the committed V5 Markdown.
Usage: python3 tools/apply_v3.py <input_md> <output_md> <log_json>
Every official object line is generated from tools/ppc_data.py (PPC transcription)."""
import sys, re, json
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
src, out, logp = sys.argv[1], sys.argv[2], sys.argv[3]
L = open(src, encoding="utf-8").read()
log = []
def sub(old, new, label, count=1):
    global L
    n = L.count(old); assert n == count, (label, n)
    L = L.replace(old, new); log.append(label)

# 1) §15.1 — official assessment rules (PPC pp. 38–40)
sub("O documento oficial indica, para componentes com 2 aulas, **mínimo de 3 instrumentos**, e 3 ou mais aulas, **mínimo de 4 instrumentos**; a escola deverá conferir a norma vigente antes da implantação.",
    "O PPC (pp. 38–40) indica, para componentes com 1 aula, **mínimo de 2 instrumentos**; com 2 aulas, **mínimo de 3 instrumentos**; e com 3 ou mais aulas, **mínimo de 4 instrumentos**. A norma vigente deve ser reconfirmada pela escola antes da implantação.",
    "V3-01 §15.1 mínimo de instrumentos por aulas, conforme PPC")
sub("Não confundir “projetos” com dispensa de provas ou de evidência individual.",
    "Não confundir “projetos” com dispensa de provas ou de evidência individual.\n\n**Regras oficiais de aprovação e recuperação (PPC, pp. 38–40):** a avaliação é trimestral; o estudante é aprovado na série somente com frequência igual ou superior a 75% e média global igual ou superior a 6,0 (Portaria Normativa SED nº 2992, de 25/10/2024); a recuperação paralela é obrigatória e, se o professor aplicou três instrumentos, deve aplicar no mínimo três instrumentos de recuperação, seguidos de nova avaliação quando o nível de aprendizagem for insuficiente.",
    "V3-02 §15.1 regras oficiais de aprovação (75% e 6,0) e recuperação paralela (PPC)")

# 2) §2.1 — integrated pillar map (BS10)
pillar = """### 2.1 Mapa integrado dos pilares do curso (BS10)

O digital e a IA são destaques, mas não substituem os demais pilares. Cada pilar reúne componentes oficiais, ferramentas candidatas e entregas, e todos os pilares são exigidos no percurso.

| Pilar | Componentes oficiais | Fundamento exigido | Ferramentas candidatas | Entrega de referência |
| --- | --- | --- | --- | --- |
| Fundamentos, consumidor e direito | Administração de Marketing (1º); Comportamento do Consumidor (1º); Legislação Aplicada (1º) | Mercado, demanda, consumidor, contratos e direito do consumidor | Planilha SWOT, formulários institucionais, NotebookLM com fontes autorizadas | Diagnóstico, pesquisa ética e parecer de risco |
| Criação e comunicação | Design Gráfico (1º); Comunicação em Marketing (2º) | Tipografia, Gestalt, cor, grid, copywriting e roteiro | Penpot, Inkscape, GIMP, Canva Educação, Shotcut | Identidade, peças, vídeo e pitch |
| Estratégia e produto | Desenvolvimento de Produtos (2º); Estratégia de Marketing (2º) | Ciclo de produto, posicionamento, segmentação e orçamento | Canvas estratégico, prototipagem, planilhas de custo | Plano estratégico e MVP simulado |
| Digital e dados | Marketing Digital (2º) | Mídias, tráfego, segmentação e métricas | Analytics demonstrativo, planilhas, CRM simulado, IA supervisionada | Campanha simulada com painel e cálculo de KPI |
| Operação, eventos e relacionamento | Gestão de Eventos (3º); Promoção de Vendas (3º); Logística (3º); Práticas de Marketing (3º) | Planejamento operacional, merchandising, distribuição e rotina de agência | Kanban, planilhas, cronograma, dashboard | Plano de evento, plano promocional, mapa de distribuição e portfólio final |

SEO, CRM, analytics e automação são **ampliações pedagógicas** do Marketing Digital (§2), não objetos oficiais do componente.

"""
sub("## 3 | Referências externas e nivelamento transparente", pillar + "## 3 | Referências externas e nivelamento transparente", "V3-03 §2.1 mapa integrado dos pilares (BS10)")

# 3) Anexo D — official lists per component (PPC transcription)
for c in COMPONENTS:
    title_re = re.compile(r"^### %s\. .*$" % c["code"], re.M)
    m = title_re.search(L); assert m, c["code"]
    objs = []
    for o in c["objetos"]:
        if o.endswith("[hab.]"):
            objs.append(o.replace(" [hab.]", " (habilidade do PPC)"))
        else:
            objs.append(o)
    block = (f"\n\n**Habilitação do docente (PPC):** {c['habilitacao']}."
             f"\n\n**Objetivo de aprendizagem (PPC, pp. {c['pages']}):** {c['objetivo']}"
             f"\n\n**Objetos do conhecimento oficiais (PPC, transcrição sintética):** " + "; ".join(objs) + ".")
    L = L[:m.end()] + block + L[m.end():]
    log.append(f"V3-04 {c['code']} objetos e objetivo oficiais (PPC) incluídos no caderno")

# 4) Anexo B2 — per-subject plan for general components (BS08)
b2 = """### B2. Ficha de plano por componente geral (proposta; sujeita à validação docente após o aceite)

Cada componente da formação geral tem uma ficha por série, derivada do Anexo B. Não cria carga, não altera conteúdo oficial e não atribui aulas: o docente titular valida e pode recusar.

| Componente geral | Objetivo do projeto integrador | Produto (1º–3º ano) | Aprovação docente | Avaliação |
| --- | --- | --- | --- | --- |
| Língua Portuguesa e Literatura | Escrita persuasiva, narrativa de marca e redação técnica em projetos de comunicação | Textos de campanha, roteiros revisados e relatório com defesa argumentada | Docente de Língua Portuguesa valida o conteúdo e recebe crédito | Avaliação oficial do docente; o projeto não gera nota paralela |
| Inglês | Vocabulário e leitura de linguagem profissional de marketing | Glossário crítico, leitura de documentação e pitch bilíngue opcional | Docente de Inglês | Avaliação oficial do docente |
| Arte | Elementos visuais, direção de arte e curadoria com créditos | Cartaz autoral, referências culturais e curadoria de mostra | Docente de Arte | Avaliação oficial do docente |
| Educação Física | Comunicação esportiva e campanha de bem-estar sem exposição de estudantes | Pesquisa de hábitos, campanha e calendário de evento esportivo | Docente de Educação Física | Avaliação oficial do docente |
| Matemática | Porcentagens, gráficos, taxas e decisões numéricas em campanhas | Infográficos com fontes, orçamento simulado e análise de cenários | Docente de Matemática | Avaliação oficial do docente |
| História | História do consumo e da propaganda, memória local | Linha do tempo de campanhas e editorial de patrimônio com créditos | Docente de História | Avaliação oficial do docente |
| Geografia | Território, comércio local e distribuição | Mapeamento de demanda e análise espacial de ações | Docente de Geografia | Avaliação oficial do docente |
| Sociologia | Consumo, grupos, desigualdades e pesquisa ética | Pesquisa ética de comportamento e crítica social de campanhas | Docente de Sociologia | Avaliação oficial do docente |
| Filosofia | Ética de comunicação, argumentação e vieses algorítmicos | Debate estruturado sobre IA e responsabilidade | Docente de Filosofia | Avaliação oficial do docente |
| Biologia | Divulgação científica com validação docente | Campanha de sustentabilidade baseada em fontes; análise de comunicação de saúde sem aconselhamento | Docente de Biologia | Avaliação oficial do docente |
| Química | Comunicação de conceitos e segurança | Campanha de consumo consciente com evidências e mostra científica | Docente de Química | Avaliação oficial do docente |
| Física | Visualização de dados e divulgação de experimentos | Roteiro audiovisual explicativo e exposição interativa | Docente de Física | Avaliação oficial do docente |

"""
sub("## ANEXO C | Instrumentos prontos para utilização", b2 + "## ANEXO C | Instrumentos prontos para utilização", "V3-05 Anexo B2 ficha por componente geral (BS08)")

# 5) Anexo I4 — official teacher qualification per component (PPC)
rows = "\n".join(f"| {c['name']} | {c['habilitacao']} |" for c in COMPONENTS)
i4 = f"""### I4. Habilitação docente por componente, conforme o PPC (BS27)

Os requisitos abaixo são transcritos do PPC (pp. 10–33). Um professor orientador só ministra um componente oficialmente se atender ao requisito e for designado formalmente.

| Componente | Habilitação exigida pelo PPC |
| --- | --- |
{rows}

**Atenção:** o PPC exige, para Legislação Aplicada, formação em Direito ou Ciências Jurídicas e Sociais, e não a lista geral de graduações de gestão. Para Design Gráfico, a lista oficial não cita graduação específica em design. A escola deve confirmar essas condições ao inventariar os docentes (Anexo J1).

"""
sub("## ANEXO J | Infraestrutura, investimentos e estudo de viabilidade", i4 + "## ANEXO J | Infraestrutura, investimentos e estudo de viabilidade", "V3-06 Anexo I4 habilitação oficial por componente (BS27)")

# 6) §9 — K11 includes perceived value and motivation (BS19); K15 decision note
sub("| K11 · Satisfação e pertencimento | % respostas 4–5 em escala 1–5; apresentar taxa de resposta | Pesquisa anônima apropriada | Coordenação; Semestral |",
    "| K11 · Satisfação, percepção de valor e motivação | % respostas 4–5 em escala 1–5 para cada item (satisfação, percepção de valor, motivação), com linha de base no diagnóstico inicial; apresentar taxa de resposta | Pesquisa anônima apropriada | Coordenação; Semestral |",
    "V3-07 §9 K11 inclui percepção de valor e motivação com linha de base (BS19)")
sub("| K15 · Visitas e parcerias | Visitas com roteiro e relatório ÷ visitas executadas × 100; briefs aprovados | Autorização, roteiro e relatório | Orientador + direção; Anual |",
    "| K15 · Visitas e parcerias | Visitas com roteiro e relatório ÷ visitas executadas × 100; briefs aprovados (contagem semestral) | Autorização, roteiro e relatório | Orientador + direção; Anual (decisão da revisão: ver DELTA_REPORT) |",
    "V3-08 §9 K15 periodicidade anual com decisão registrada (K15)")

# 7) G2 — K11 and K15 text; add Responsável and Risco de viés columns (BS20)
sub("| K11 | Satisfação discente = respostas positivas (4–5) / respostas válidas ×100 | escuta com questionário anônimo e taxa resposta | semestral; citar viés de não resposta |",
    "| K11 | Satisfação discente, percepção de valor e motivação = respostas positivas (4–5) / respostas válidas ×100, por item | escuta com questionário anônimo, taxa resposta e linha de base | semestral; citar viés de não resposta |",
    "V3-09 G2 K11 inclui percepção de valor e motivação (BS19)")
sub("| K15 | Visitas com aprendizagem = visitas com roteiro e relatório / visitas executadas ×100 | atas e portfólios | anual; sem visitas = “não aplicável”, não 0% |",
    "| K15 | Visitas com aprendizagem = visitas com roteiro e relatório / visitas executadas ×100 | atas e portfólios | anual (ciclo de visitas é anual, §14.6); sem visitas = “não aplicável”, não 0% |",
    "V3-10 G2 K15 justificativa da periodicidade anual")
owners = {"K01":("Docentes; coordenação valida","dificuldade desigual dos instrumentos (G5)"),
 "K02":("Docente titular","avaliação de retorno mais fácil que a diagnóstica (G5)"),
 "K03":("Líder da equipe valida; docente audita","registro tardio de entregas"),
 "K04":("Docentes","ausente não é o mesmo que não avaliado"),
 "K05":("Coordenação + docentes","registro seletivo de encontros"),
 "K06":("Docentes","presença registrada sem contribuição real"),
 "K07":("Docentes","N pequeno de defesas"),
 "K08":("Docentes","uso de IA declarado de forma parcial"),
 "K09":("Docentes + avaliador externo quando possível","N pequeno de projetos"),
 "K10":("Orientador","adesão voluntária pode enviesar o resultado"),
 "K11":("Coordenação","não resposta superestima aceitação (G5)"),
 "K12":("Gestão","transferências não contadas"),
 "K13":("Gestão","registros de frequência inconsistentes"),
 "K14":("Coordenação","adesão voluntária, sem ranking punitivo"),
 "K15":("Orientador + direção","visitas não realizadas tratadas como 0%"),
 "K16":("Gestão","sazonalidade e campanhas pontuais")}
lines = L.split("\n")
hdr_i = next(i for i,l in enumerate(lines) if l.startswith("| ID | Indicador / fórmula | Recorte e fonte | Periodicidade e alerta |"))
lines[hdr_i] = "| ID | Indicador / fórmula | Recorte e fonte | Periodicidade e alerta | Responsável | Risco de viés |"
lines[hdr_i+1] = "| --- | --- | --- | --- | --- | --- |"
for i in range(hdr_i+2, len(lines)):
    m = re.match(r"^\| (K\d\d) \|", lines[i])
    if not m: break
    resp, risk = owners[m.group(1)]
    lines[i] = lines[i].rstrip() + f" {resp} | {risk} |"
L = "\n".join(lines)
log.append("V3-11 G2 acrescenta colunas Responsável e Risco de viés por indicador (BS20)")

# 8) Appendix M1 — PPC page ranges corrected to the PDF (Eventos p. 25 → Práticas p. 33)
sub("- **Páginas 26–32:** Eventos, Promoção, Logística, Práticas.", "- **Páginas 25–33:** Gestão de eventos, Promoção de vendas, Logística, Práticas de Marketing.", "V3-12 M1 faixa de páginas dos componentes corrigida ao PPC")

open(out, "w", encoding="utf-8").write(L)
json.dump(log, open(logp,"w",encoding="utf-8"), ensure_ascii=False, indent=1)
print("applied", len(log), "edits")
