# DELTA_REPORT — V5 publicado (commit a795e98) → V5 revisado (v4)

| Item | Antes (a795e98) | Depois (v4) |
|---|---|---|
| Markdown canônico SHA256 | `875ec13184d20c1187e80b945b2c8040240e46f276fe9662a69a1de94679dfa3` | `8c6500805940b32894d7b413b1ff1038f6f79b027baf4d207ecb18fb378df0ba` |
| Linhas do Markdown | 1.116 | 1.478 |
| Páginas do DOCX/PDF | 58 | 78 |
| Rótulo de versão | 5.0 | 5.0 (ver item 3) |

## 1. Restauração do arquivo de trabalho (segurança)
O arquivo `Plano_Diretor_…_V5.md` estava sobrescrito pela cópia manual (SHA256 `036d5f61…`, versão anterior à reordenação). A versão transferida foi preservada como evidência e o arquivo foi restaurado para o commit a795e98 (`875ec131…`) com `git checkout --`, sem reset e sem perda de histórico. A matriz BS transferida é idêntica à versionada (`94ca0f38…`).

## 2. Verificação contra o PPC oficial
- PPC: `PPC-EMIEP-Marketing.pdf`, 43 páginas, processo SED 00178627/2026, código 7F67J6IF. SHA256 `31fb64c5…f58`. Sem referência de hash anterior para comparação.
- Carga horária conferida na tabela V do PPC (total 3.000; FGB 2.200 com 120 h integradoras dentro; técnicas 800).
- 183 objetos oficiais e 12 cargas: ver CURRICULAR_ALIGNMENT_AUDIT.md.

## 3. Alterações de conteúdo

| Ref. | Local | Tipo | Requisito | Justificativa |
|---|---|---|---|---|
| V3-01 | §15.1 | Substituição de frase | BS13 | Mínimos de instrumentos (2 para 1 aula; 3 para 2 aulas; 4 para 3+) conferidos no PPC, pp. 38–40. |
| V3-02 | §15.1 | Acréscimo | BS13 | Regras oficiais: 75% de frequência, média 6,0, recuperação paralela (PPC, pp. 38–40). |
| V3-03 | §2.1 (novo) | Inserção | BS10 | Mapa integrado em cinco pilares. |
| V3-04 | Anexo D (12 cadernos) | Inserção em cada caderno | BS01 | Habilitação (PPC), objetivo de aprendizagem e objetos oficiais transcritos; itens de habilidade marcados. |
| V3-05 | Anexo B2 (novo) | Inserção | BS08 | Ficha por componente geral (12), derivada do Anexo B; aprovação docente pendente. |
| V3-06 | Anexo I4 (novo) | Inserção | BS27 | Habilitação oficial por componente (PPC). |
| V3-07 | §9, K11 | Substituição de linha | BS19 | Satisfação, percepção de valor e motivação, com linha de base. |
| V3-08 | §9, K15 | Substituição de linha | BS20 | Periodicidade anual com decisão registrada. |
| V3-09 | G2, K11 | Substituição de linha | BS19 | Mesmo texto de V3-07. |
| V3-10 | G2, K15 | Acréscimo de justificativa | BS20 | Ciclo de visitas anual (§14.6). |
| V3-11 | G2 (tabela) | Acréscimo de colunas | BS20 | Colunas **Responsável** e **Risco de viés** para os 16 indicadores. |
| V3-12 | M1 | Correção de faixa | BS01 | Eventos a Práticas: pp. 25–33 (PPC), antes 26–32. |
| V4-01 | §12.2 | Acréscimo | BS01 | Requisitos de acesso, prazo de 5 anos, estágio (não há) e infraestrutura (PPC I, VIII, IX, XI). |
| V4-02 | Anexo N (novo) | Inserção | BS01, §3.1 | Matriz-mestra: 12 tabelas × 14 campos, cada campo marcado (EXPLÍCITO_PPC, APLICAÇÃO_DIDÁTICA, ENRIQUECIMENTO_PROPOSTO, PENDENTE_DE_VALIDAÇÃO). Gerada a partir dos cadernos e do PPC. |

Anteriormente (commit a795e98): reordenação de §§1–17 antes de A–M; E2–E11. Entre os itens E2–E11, a alteração das linhas do §14.2 G1 e §14.5 foi mantida.

**Linhas originais substituídas nesta revisão:** 22 (diff contra a795e98), todas listadas acima (cabeçalho e linhas do G2, duas linhas do §9 e a frase do §15.1). Nenhuma outra linha original foi perdida.

**Carga horária e ementa oficiais:** nenhuma alteração. Todos os totais (800, 2.200, 120, 3.000) foram conferidos e permanecem os mesmos.

## 4. Decisão de periodicidade de K15
**Decisão:** K15 (visitas com aprendizagem) permanece **anual**. Motivo: o ciclo de visitas é anual por desenho (§14.6: 1º, 2º e 3º anos); medir por semestre produz denominadores muito pequenos, o que a própria regra G5 proíbe. A contagem semestral de briefs aprovados fica registrada como informação complementar. **Decisão a confirmar** pela escola, caso prefira semestral.

## 5. Versionamento
O rótulo continua 5.0. Recomendação: publicar esta revisão como **5.1** antes de circular para a Direção, para que a versão lida corresponda a este delta. Decisão do Master.

## 6. Reprodução
`tools/ppc_data.py` (transcrição do PPC), `tools/apply_v3.py` e `tools/apply_v4.py` (alterações aplicadas a partir da a795e98), `tools/build_docx.py`, `tools/pagemap.py`, `tools/qa_source.py`, `tools/ppc_gap.py` (triagem). Comandos em QA_REPORT.md.
