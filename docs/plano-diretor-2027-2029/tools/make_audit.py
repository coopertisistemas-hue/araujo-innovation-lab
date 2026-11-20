"""Generate CURRICULAR_ALIGNMENT_AUDIT.md (V5.1) from the PPC transcription, the proposed units and the pre-revision screening.
Usage: python3 tools/make_audit.py <screening_json> <out_md>"""
import sys, json
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
from ppc_units import UNITS
gap = {g["code"]: g for g in json.load(open(sys.argv[1], encoding="utf-8"))}
out = []
A = out.append
A("# Matriz de auditoria curricular — PPC oficial × Markdown V5.1\n")
A("**Status documental:** auditoria objeto a objeto realizada com o PPC oficial. Cada um dos 183 itens tem três marcas separadas: **N1** (requisito oficial rastreável), **N2** (tratamento pedagógico proposto demonstrado, por unidade com atividade e evidência descritas) e **N3** (plano detalhado pendente de validação docente após o aceite). Não é certificação de aderência nem de operação.\n")
A("## 0. Identificação do PPC consultado\n")
A("- Arquivo: `PPC-EMIEP-Marketing.pdf` · 43 páginas PDF (42 textuais + página de assinaturas).")
A("- Processo SED 00178627/2026 · código 7F67J6IF · assinaturas de 03/09/2026 (CPFs mascarados no documento).")
A("- SHA-256: `31fb64c5547eae8e678df8f01cbda451676e1446a0d32fd00e51f24a5efa6f58`. Não há referência de hash anterior registrada; a identificação é feita pelo processo e código.\n")
A("## 1. Carga horária (tabela V do PPC, diurna)\n")
A("| Verificação | PPC | Markdown | Resultado |")
A("|---|---|---|---|")
A("| Total do curso | 3.000 h | 3.000 h | CONFERE |")
A("| Formação Geral Básica (CHT) | 1º 744 · 2º 712 · 3º 744 = 2.200 h (somas incluem 40 h de integradoras por série) | 2.200 h | CONFERE |")
A("| Atividades integradoras EPT | 40 h por série = 120 h, dentro da FGB | 120 h dentro da FGB, sem soma dupla | CONFERE |")
A("| Itinerário técnico | 1º 256 · 2º 288 · 3º 256 = 800 h | 800 h | CONFERE |")
A("| Matriz noturna | Mesma carga total; até 30% híbrido (Resolução CEE/SC 070/2025) | Referência é a diurna (§12.2) | DECLARADO |\n")
A("## 2. Resumo por componente\n")
A("| Código | Componente (PPC) | Série · carga · páginas | Itens | Unidades propostas | Conteúdo | Habilitação |")
A("|---|---|---|---|---|---|---|")
for c in COMPONENTS:
    cont = "DIRECT" + (" + COMPLEMENTARY (SEO, CRM, analytics, IA: §2)" if c["code"]=="D06" else "")
    hab = "REQUIRES CLARIFICATION (I4)" if c["code"] in ("D02","D04") else "DIRECT"
    A(f"| {c['code']} | {c['pdfname']} | {c['ano']}º · {c['ch']} h · pp. {c['pages']} | {len(c['objetos'])} | {len(UNITS[c['code']])} | {cont} | {hab} |")
A("")
A("## 3. Auditoria objeto a objeto (183 itens)\n")
A("| Item | Objeto oficial (PPC) | Origem | Triagem pré-revisão (heurística) | Unidade proposta | N1 | N2 | N3 |")
A("|---|---|---|---|---|---|---|---|")
tot = {"COBERTO":0,"PARCIAL":0,"AUSENTE":0}
for c in COMPONENTS:
    where = {}
    for u in UNITS[c["code"]]:
        for k in u["items"]: where[k] = u
    for k, o in enumerate(c["objetos"]):
        g = gap[c["code"]]["objects"][k]; assert g["obj"] == o
        tot[g["status"]] += 1
        origem = "habilidade" if o.endswith("[hab.]") else "objeto"
        u = where[k]
        A(f"| {c['code']}.{k+1:02d} | {o.replace(' [hab.]','')} | {origem} | {g['status']} ({g['ratio']}) | {u['id']} {u['title']} | sim | sim | pendente |")
A("")
A(f"**Totais:** 183 itens; N1 = 183; N2 = 183 (todos atribuídos a uma única unidade); N3 = 183 pendentes. Triagem pré-revisão: COBERTO {tot['COBERTO']}, PARCIAL {tot['PARCIAL']}, AUSENTE {tot['AUSENTE']}, indicando quais itens não tinham tratamento explícito antes desta revisão.\n")
A("## 4. Perfil profissional (PPC, III) × competências\n")
A("| Habilitação do PPC | Competências (§13.1) | Onde aparece |")
A("|---|---|---|")
A("| Projetar e implementar planos de marketing | C04 | D07, D12 |")
A("| Análise de vendas, preços e produtos | C07 | D10, D12, K01–K16 |")
A("| Projetos de comunicação, fidelização e relação com fornecedores | C05, C10 | D08, D09, D11, D03 (fornecedores) |")
A("| Marketing digital | C06 | D06 |")
A("| Apresentação de produtos no ponto de venda | C05, C06 | D10 |")
A("| Estudos de mercado | C03 | D03, D01 |")
A("| Socioemocionais: comunicação, diversidade, trabalho colaborativo | C11 | §14.3; C2; C4 |\n")
A("## 5. Achados e situação (V5.1)\n")
A("| # | Achado | Severidade | Situação |")
A("|---|---|---|---|")
rows = [
 ("1","Objetos oficiais não tinham tratamento explícito (triagem: 67 AUSENTE, 71 PARCIAL).","MAJOR","RESOLVIDO documentalmente: cada item tem unidade, atividade e evidência (N2)."),
 ("2","Objetivo por encontro e plano de aula não existem.","OPERACIONAL","N3 pendente após aceite. Não bloqueia revisão documental (§12.4)."),
 ("3","Habilitação de Legislação Aplicada: Direito ou Ciências Jurídicas e Sociais.","OPERACIONAL","Pré-requisito de implantação (I4). Confirmar quadro docente."),
 ("4","Habilitação de Design Gráfico: lista oficial sem graduação específica em design.","OPERACIONAL","Pré-requisito de implantação (I4). Decisão da escola."),
 ("5","Regras de aprovação e recuperação ausentes.","MAJOR","RESOLVIDO em §15.1 (PPC, pp. 38–40)."),
 ("6","Mínimos de instrumentos por aulas incompletos.","MINOR","RESOLVIDO (§15.1). Norma vigente a reconfirmar."),
 ("7","Acesso, prazo de 5 anos, estágio e infraestrutura ausentes em §12.2.","MINOR","RESOLVIDO (§12.2)."),
 ("8","Faixa de páginas de Eventos a Práticas (M1): 26–32, no PPC 25–33.","MINOR","RESOLVIDO (M1)."),
 ("9","Habilidades do PPC mapeadas por unidade, não por competência individual.","MINOR","Método declarado; 19 itens de habilidade marcados."),
 ("10","'Gestão de eventos' (PPC) × 'Gestão de Eventos' (MD).","MINOR","Sem efeito curricular."),
 ("11","Matriz noturna não auditada item a item.","MINOR","DECLARADO (§12.2)."),
 ("12","Anexo B: formação geral segue o Currículo Base; o PPC não traz seus objetos.","INFO","Mantido como proposta sinalizada (PPC, p. 34)."),
 ("13","Autoria das unidades: as unidades e a atribuição dos itens foram elaboradas pelo proponente, não pelo PPC.","MAJOR (auditoria)","Para o auditor independente verificar a pertinência de cada atribuição."),
]
for r in rows: A("| " + " | ".join(r) + " |")
A("")
A("**Classificação final (V5.1):** conteúdo DIRECT nos 12 componentes; COMPLEMENTARY em Marketing Digital; REQUIRES CLARIFICATION em habilitação de Design Gráfico e Legislação Aplicada (pré-requisito operacional); GAP de conteúdo documental: nenhum; N3 pendente para todos os itens (operacional).")
open(sys.argv[2], "w", encoding="utf-8").write("\n".join(out) + "\n")
print("audit written", sys.argv[2], tot)
