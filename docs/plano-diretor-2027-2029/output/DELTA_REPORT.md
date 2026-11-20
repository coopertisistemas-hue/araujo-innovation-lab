# DELTA_REPORT — V5.0 (c9696f2) → V5.1 (controlado)

| Item | V5.0 (preservada) | V5.1 (controlada) |
|---|---|---|
| Markdown | `8c6500805940b32894d7b413b1ff1038f6f79b027baf4d207ecb18fb378df0ba` | `729971e9e71bf96572ce77bb49362e58085b9ada2b560d2bb6e2fcecc0ed14b7` |
| Arquivo | `…_COMPLETO_V5.md` (intacto) | `…_COMPLETO_V5_1.md` (novo) |
| Versão | 5.0 | 5.1 (20/11/2026) |
| Páginas DOCX/PDF | 78 | 100 |
| Artefatos V5.0 | `output/historico-v5/` | `output/` |

## 1. Preservação
A fonte V5.0 não foi alterada. Os artefatos V5.0 (DOCX, PDF, QA, checksums, auditoria, rastreabilidade, delta e as PNGs de prova) estão em `output/historico-v5/`, movidos com `git mv`.

## 2. Alterações da V5.1 (script `tools/apply_v5_1.py`)

| Ref. | Local | Alteração | Requisito / justificativa |
|---|---|---|---|
| V51-01 | Front matter | versão 5.1 | Controle de versão |
| V51-02 | Linha de versão | Versão 5.1 e histórico de versões | Rastreabilidade entre V5.0 e V5.1 |
| V51-03 | §9, K15 | "proposta sujeita à deliberação da Direção" | K15 anual mantido como proposta |
| V51-04 | Anexo I4 | Habilitação como pré-requisito de implantação | Porta 3, não porta 1 |
| V51-05 | §12.4 (novo) | Portas separadas: revisão documental, aprovação, prontidão operacional | Separação de gates |
| V51-06 | Anexo O (novo) | Unidades de aprendizagem propostas (3 a 5 por componente, 48 no total) e mapa dos 183 itens com N1, N2, N3 | Rastreabilidade item a item sem plano de aula por item |

**Anexo B2:** a legenda que separava o título da tabela foi corrigida por `keep_with_next` no construtor do DOCX (intro + tabela mantidos juntos).

## 3. Conteúdos sem alteração
Carga horária oficial, lista de componentes, objetos do Anexo D, Anexo N e demais seções permanecem como na V5.0.

## 4. Decisões registradas
- **K15 anual:** proposta, sujeita à deliberação da Direção. Motivo: ciclo de visitas anual (§14.6) e regra de pequeno N (G5).
- **Unidades de aprendizagem:** proposta do proponente. Qualquer atribuição de item a outra unidade deve ser validada pelo docente habilitado após o aceite.

## 5. Reprodução
`tools/ppc_data.py`, `tools/ppc_units.py`, `tools/apply_v5_1.py`, `tools/make_audit.py`, `tools/build_docx.py`, `tools/pagemap.py`, `tools/qa_source.py`. Comandos em QA_REPORT.md.
