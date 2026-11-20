import re, sys, json
SRC = sys.argv[1]; OUT = sys.argv[2]; LOG = sys.argv[3]
orig = open(SRC, encoding="utf-8").read().split("\n")
head = orig[:25]                       # lines 1–25 (front matter + cover quotes)
body = orig[25:]
# split body into H2 segments
segs = []; cur = None
for ln in body:
    if ln.startswith("## "):
        if cur: segs.append(cur)
        cur = [ln]
    else:
        cur.append(ln)
segs.append(cur)
def key(seg):
    h = seg[0]
    m = re.match(r"## (\d+) \|", h)
    if m: return (0, int(m.group(1)))
    m = re.match(r"## ANEXO ([A-M]) \|", h)
    if m: return (1, ord(m.group(1)))
    if h.startswith("## Orientação"): return (-1, 0)
    return (9, 0)
# Front part before first H2 is head; body segments are all H2-led (line 26 starts first H2)
assert all(key(s)[0] in (-1,0,1) for s in segs), [s[0] for s in segs]
reordered = sorted(segs, key=key)
new_body = [l for s in reordered for l in s]
text_lines = head + new_body
log = {"reorder": [s[0] for s in reordered]}

def rep(lines, old, new, label):
    joined = "\n".join(lines)
    n = joined.count(old)
    assert n == 1, (label, n)
    return joined.replace(old, new).split("\n")

L = "\n".join(text_lines)
def sub(old, new, label):
    global L
    n = L.count(old); assert n == 1, (label, n)
    L = L.replace(old, new); log.setdefault("replacements", []).append(label)

# ---- E2: §9 table aligned with K01–K16 dictionary (BS20 / G-BS-05)
old_t = L[L.index("| Indicador | Cálculo / descrição | Fonte | Responsável / periodicidade |"):]
old_t = old_t[:old_t.index("\n\nMetas e indicadores são propostas")]
new_t = """| Indicador | Cálculo / descrição | Fonte | Responsável / periodicidade |
| --- | --- | --- | --- |
| K01 · Aprendizagem por competência | Média das notas de rubrica (0–4) por competência, ano e trimestre | Rubricas individuais em provas e projetos | Docentes; coordenação valida; Trimestral |
| K02 · Progresso individual | Nota atual − nota diagnóstica na mesma rubrica (0–4) | Diagnóstico + prova comparável | Docente titular; Trimestral |
| K06 · Participação efetiva | Alunos presentes elegíveis com contribuição rastreável ÷ alunos presentes elegíveis × 100 | Diário de bordo; observação; versões | Docentes; Mensal |
| K03 · Entregas no prazo | Entregas realizadas no prazo ÷ entregas previstas elegíveis × 100 | Registros de projetos | Líder da equipe valida; docente audita; Mensal |
| K05 · Aplicação em laboratório | Encontros técnicos com artefato/evidência ÷ encontros técnicos ocorridos × 100 | Plano e diário de aula | Coordenação + docentes; Mensal |
| K04 · Portfólios completos | Alunos com itens mínimos aprovados ÷ alunos avaliáveis × 100 | Rubricas e arquivos | Docentes; Trimestral |
| K10 · Integração docente | Docentes com projeto formal registrado ÷ docentes convidados × 100 | Briefs coassinados | Orientador; Semestral |
| K11 · Satisfação e pertencimento | % respostas 4–5 em escala 1–5; apresentar taxa de resposta | Pesquisa anônima apropriada | Coordenação; Semestral |
| K13 e K12 · Frequência e permanência | Frequência: presenças registradas ÷ oportunidades oficiais × 100 (K13). Permanência: matriculados no fechamento ÷ matriculados elegíveis no início × 100 (K12) | Secretaria escolar | Gestão; frequência trimestral; permanência anual |
| K09 · Qualidade dos projetos | Projetos aprovados em 5 checks (briefing, fontes, ética, KPI e defesa) ÷ projetos avaliados × 100 | Checklist e banca | Docentes + avaliador externo quando possível; Trimestral |
| K08 · Competência em uso crítico de IA | Alunos com fonte e declaração válida ÷ alunos que usaram IA na atividade × 100 | Rubrica 0–4 e registro de atividade | Docentes; Trimestral |
| K07 · Defesa individual suficiente | Alunos com rubrica ≥ 2 na defesa individual ÷ alunos avaliados × 100 | Apresentação individual | Docentes; Trimestral |
| K15 · Visitas e parcerias | Visitas com roteiro e relatório ÷ visitas executadas × 100; briefs aprovados | Autorização, roteiro e relatório | Orientador + direção; Anual |
| K14 · Formação docente | Participantes com exercício demonstrado ÷ docentes inscritos × 100 | Listas e artefatos voluntários após aceite | Coordenação; Semestral |
| K16 · Interesse pelo curso | Matrículas/inscrições válidas por ano, sempre com contexto de capacidade e comparação | Secretaria, campanhas e capacidade instalada | Gestão; Anual |"""
sub(old_t, new_t, "E2 §9 KPI table aligned to K01–K16 (BS20, G-BS-05)")

# ---- E3: §4.1 ficha de aula (BS03, BS04; §3.2)
ficha = """### 4.1 Ficha de aula operacional replicável

Toda aula técnica registra, no mínimo, os campos abaixo. A distribuição de 115 minutos apresentada acima é **exemplo**, não grade escolar homologada. A tecnologia é meio para o objetivo de aprendizagem: a meta é de encontros com evidência aplicada (K05, ≥60% proposto após estudo de viabilidade), e não de 100% do tempo em tela.

| Campo | Preenchimento mínimo |
| --- | --- |
| Componente, série e item do PPC | Nome oficial do componente, série prevista na matriz e objeto da ementa atendido (página do PPC a confirmar na matriz-mestra) |
| Pergunta-motriz | Problema real ou hipótese a investigar, formulada antes da ferramenta |
| Objetivo observável | O que o estudante deverá demonstrar ao final do encontro |
| Bloco conceitual | Conceito, fundamento e critério de qualidade |
| Prática de laboratório | Atividade em grupo com ferramenta ou método, e alternativa offline ou institucional equivalente |
| Justificativa de decisão | Por que a escolha atende ao objetivo, além de como a ferramenta foi operada |
| Tarefa individual e atividade coletiva | Contribuição rastreável de cada estudante |
| Entregável, rubrica e feedback | Artefato avaliado com os critérios da rubrica e devolutiva registrada |
| Dados e autorizações | Dados, imagens e marcas usados somente com autorização; nenhum dado pessoal de estudante em ferramenta externa |

"""
sub("## 5 | Valorização dos professores", ficha + "## 5 | Valorização dos professores", "E3 §4.1 ficha de aula replicável (BS03, BS04)")

# ---- E5: §5.1 ciclo exemplificativo (BS07, BS08)
ciclo = """### 5.1 Ciclo exemplificativo de um projeto docente

Ilustração do fluxo mínimo com a Campanha de Leitura (Língua Portuguesa). Os passos seguem a seção 14.2 e não constituem compromisso de execução:

1. **Demanda:** o docente propõe o objetivo pedagógico (ampliar a participação voluntária na leitura) e o público.
2. **Entrevista:** a equipe entrevista o docente e estudantes-leitores, com consentimento, para delimitar problema e escopo.
3. **Diagnóstico:** pesquisa anônima de hábitos de leitura, com taxa de resposta registrada.
4. **Estratégia:** objetivo SMART, segmento, mensagens, canais institucionais e indicador de participação agregada.
5. **Peças:** cartaz, roteiro de vídeo e formulário, com crédito ao componente e ao docente responsável.
6. **Revisão docente:** validação do conteúdo disciplinar antes de qualquer divulgação.
7. **Apresentação:** pitch ao docente e à coordenação; publicação somente se autorizada.
8. **Mensuração e devolutiva:** relatório agregado ao docente e à direção, com limitações explícitas.

"""
sub("## 6 | Trilhas dos três anos", ciclo + "## 6 | Trilhas dos três anos", "E5 §5.1 ciclo exemplificativo (BS07, BS08)")

# ---- E6: §14.1 competição saudável e feedback (BS05)
sub("crítica construtiva e defesa de evidências; evitar ruído ou competição que humilhe.",
    "crítica construtiva e defesa de evidências, com competição saudável, reconhecimento simbólico e feedback respeitoso; evitar ruído ou competição que humilhe.",
    "E6 §14.1 competição saudável e reconhecimento simbólico (BS05)")

# ---- E4: §15.7 exemplos de itens por série (BS13; §3.4)
itens = """### 15.7 Exemplos pedagógicos de itens por série

Exemplos para orientar a elaboração das provas pelo docente titular. **Não são provas oficiais** e não substituem os instrumentos da SED/SC. Cada item se aplica tanto à prova sem IA quanto à prova com registro de uso de IA (seção 8).

| Série | Item de prova (exemplo) | Nível de domínio esperado (seção 13.2) | Critério de correção |
| --- | --- | --- | --- |
| 1º ano | A partir de um briefing escolar fictício, identifique o público-alvo, justifique a escolha e formule duas perguntas para a entrevista. | N1–N2 | Público específico, justificativa coerente com o briefing, perguntas abertas e pertinentes |
| 2º ano | Escolha dois canais para a campanha simulada, calcule a taxa de conversão de um funil fictício e explique um limite da métrica. | N2–N3 | Cálculo correto, canal justificado pelo público, limite da métrica explicitado |
| 3º ano | Defenda, em até cinco minutos, uma decisão de investimento ou operação para o cliente simulado, com riscos e um indicador de acompanhamento. | N3–N4 | Decisão fundamentada em dados, riscos identificados, indicador de controle e plano de contingência |

"""
sub("## 16 | Proposta de deliberação", itens + "## 16 | Proposta de deliberação", "E4 §15.7 exemplos de itens por série (BS13)")

# ---- E8: BS21 classification after benchmarking table (K1)
classif = """Classificação das referências: **normativo oficial** — CNCT/MEC; **curso técnico comparável** — SENAC-SC (800 h); **curso de aperfeiçoamento complementar** — SENAI-SP (30 h), que não equivale a curso técnico; **referencial internacional** — UNESCO. Nenhuma delas comprova desempenho superior de qualquer instituição.

"""
sub("### K2. Matriz de alinhamento por evidência", classif + "### K2. Matriz de alinhamento por evidência", "E8 Anexo K classificação das referências (BS21)")

# ---- E7: Anexo C4 autoavaliação e reflexão (BS19, BS06)
c4 = """### C4. Ficha de autoavaliação e reflexão individual (proposta)

Preenchida pelo estudante ao fim de cada trimestre e comparada à primeira versão do portfólio (antes/depois):

□ Competência em que mais avancei e a evidência que comprova o avanço

□ Dificuldade enfrentada e o que fiz para superá-la

□ Minha contribuição na equipe, com exemplos verificáveis

□ Uso de IA nesta atividade: onde usei, o que conferi e o que decidi sozinho

□ Feedback recebido e uma ação de melhoria para o próximo ciclo

"""
c3 = "Período e turmas; matrícula/frequência; instrumentos e cobertura; médias de rubricas e evolução; projetos e briefs; integração docente; aulas práticas efetivas; riscos e incidentes; feedback de estudantes; melhorias priorizadas; recursos necessários; decisões solicitadas. Anexar metodologia de cálculo, denominadores e evidências amostrais anonimizadas."
sub(c3, c3 + "\n\n" + c4.rstrip("\n") + "\n", "E7 Anexo C4 autoavaliação e reflexão (BS19, BS06)")

# ---- E9: Anexo H5 matriz de IA por componente (BS11)
h5 = """### H5. Matriz de uso de IA por componente (proposta)

Derivada dos laboratórios e validações de cada caderno do Anexo D. Cada uso exige registro, fonte admitida e checagem humana; nenhum componente torna o uso de IA obrigatório.

| Componente (série) | Caso de uso de IA admitido | Fonte admitida | Checagem humana obrigatória |
| --- | --- | --- | --- |
| Administração de Marketing (1º) | Levantar hipóteses para a SWOT | Leituras autorizadas e NotebookLM | Confrontar as hipóteses com os dados do briefing |
| Design Gráfico (1º) | Gerar alternativas de layout, com registro | Materiais de domínio público ou licenciados | Verificar licenças, formatos editáveis e legibilidade |
| Comportamento do Consumidor (1º) | Organizar teoria e roteiro de síntese | Referências autorizadas | Não tratar personas geradas como pessoas reais |
| Legislação Aplicada (1º) | Resumir textos oficiais | Fontes legais oficiais | Distinguir estudo de caso de consultoria jurídica; referenciar cada regra |
| Desenvolvimento de Produtos (2º) | Comparar hipóteses de produto | Pesquisas e testes registrados | Separar premissas de custo de orçamento real |
| Marketing Digital (2º) | Simular criativos e interpretar métricas | Demonstrativos e planilhas | Não confundir curtidas com resultado de negócio |
| Estratégia de Marketing (2º) | Apoiar cenários e priorização | Canvas e planilhas do grupo | Cada escolha parte de dado ou hipótese declarada |
| Comunicação em Marketing (2º) | Sugerir títulos e roteiros | Editor de texto e referências autorizadas | Verificar veracidade, CTA ético e acessibilidade |
| Gestão de Eventos (3º) | Apoiar checklist e cronograma | Formulários anônimos e planilhas | Confirmar autorização antes de qualquer divulgação |
| Promoção de Vendas (3º) | Cenários de preço hipotético | Planilhas de simulação | Oferta transparente, sem promessa enganosa |
| Logística (3º) | Simular cenários de custo | Planilhas e calculadora | Decisão combina custo, nível de serviço e sustentabilidade |
| Práticas de Marketing (3º) | Apoiar execução e análise crítica | Kanban, dashboard e fontes registradas | Responsável humano aprova cada entrega |

"""
sub("## ANEXO I | Formação docente", h5 + "## ANEXO I | Formação docente", "E9 Anexo H5 matriz de uso de IA por componente (BS11)")

open(OUT, "w", encoding="utf-8").write(L)
json.dump(log, open(LOG, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("ok lines", len(L.split("\n")), "orig", len(orig))
