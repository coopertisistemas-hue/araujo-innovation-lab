# DELTA_REPORT — candidato corrigido V5.1 (contra o candidato 2de5f71)

| Item | 2de5f71 (candidato auditado) | Candidato corrigido |
|---|---|---|
| Markdown V5.1 | `729971e9e71bf96572ce77bb49362e58085b9ada2b560d2bb6e2fcecc0ed14b7` | `13d5bc05d0026486f230076901e1dedbda83f6cff52e452c485ba8fd915dda23` |
| DOCX | arquivado em `output/historico-v5.1-2de5f71/` | `1fcb818ed63d9c69bc29699b8b1d615af2579d37e77fdbc57ffa8db205a2f0e8` |
| PDF | arquivado em `output/historico-v5.1-2de5f71/` | `d3f92e6a45b30ffc0e4ed695be015822af46610f4f35c50106fe5b323045b11e` (100 páginas) |
| Unidades propostas | 48 | 49 (D01 U4 desdobrada) |
| Testes | não havia | 31/31 (`tests/test_results.json`) |

## Correções obrigatórias

| Correção | Antes | Depois | Teste |
|---|---|---|---|
| **K01** (bloqueante) | "soma de pontos ÷ avaliações (0–4)", incoerente com a escala | Média ponderada dos níveis: Σ(peso × nível) ÷ Σ(peso); exemplo 4/3/2/4/3/1 com pesos 30/20/20/15/10/5 = **3,15**. Em §9 e no dicionário G2 | 10.000 vetores de níveis 0–4 ficam em [0,4]; a soma antiga daria 17 no exemplo |
| **K15** (interpretação; ver pendências) | Anual, mas misturava briefs semestrais no indicador; sem regra para ausência de visitas | Anual; n/N = visitas com roteiro e relatório ÷ visitas executadas × 100; sem visitas: não aplicável, nunca 0%; briefs aprovados contados à parte | 3 de 4 = 75%; sem visita devolve não aplicável |
| **K16** (mensurabilidade) | "Matrículas/inscrições válidas por ano": contagem sem base | Taxa de ocupação = matrículas válidas ÷ vagas ofertadas × 100, com contexto | 40 de 60 = 66,7%; sem vagas devolve não aplicável |
| **Data do rodapé** | 19/11/2026 | 20/11/2026 (data da versão 5.1) | rodapé do DOCX contém 20/11/2026 e não contém 19/11/2026 |

## Correções adicionais (achados da auditoria)

| Item | Mudança |
|---|---|
| D01 U4 superdimensionada | Dividida em U4 (financeiro, pessoas e endomarketing: itens 17–20) e U5 (qualidade e melhoria contínua: itens 21–30), com atividade e evidência para cada uma |
| Objetivos e perfil oficiais ausentes | Novo §2.2 com objetivo geral, objetivos específicos, perfil profissional e certificações intermediárias (PPC II–IV) |
| Acesso às referências | "acesso em 19/11/2026" → "consulta de 19/11/2026; reconfirmação pendente em 5 links" |
| Rótulo de transcrição | Dois itens de habilidade passam a "paráfrase de habilidade do PPC" |

## Ajustes de layout
- Capa: histórico de versões movido para o Anexo O; a linha de versão ficou curta.
- Tabelas de unidades e do mapa de itens em 8,5 pt; coluna de marcações mais larga, sem quebra de "PENDENTE_DE_VALIDAÇÃO".
- Versão nos cabeçalhos, rodapés e ficha de controle: 5.1.

## Não alterado
Cargas horárias oficiais, lista de componentes, texto oficial transcrito, §12.4 (portas separadas), Anexos N e O (exceto unidades alteradas), 28 requisitos do brainstorming e sua rastreabilidade.

## Pendências (fora do escopo congelado)
- Reconfirmar as 5 referências externas não confirmadas (3 UNESCO, 2 CNCT).
- Ajustes leves de sobreposição em D05 U2, D06 U2 e D09 U3 (N2 levemente superestimado): sinalizados, não corrigidos.
- Confirmação institucional de habilitação para Design Gráfico e Legislação Aplicada; validação docente após o aceite.
