# QA_REPORT — Plano Diretor Pedagógico Marketing 2027–2029, V5.1 (proposta, não homologada)

**Status:** PUBLICADO PARA AUDITORIA INDEPENDENTE — NÃO CERTIFICADO; NÃO APROVADO PELA DIREÇÃO.
**Porta 1 (revisão documental):** ATENDIDA. **Porta 2 (aprovação institucional):** PENDENTE. **Porta 3 (prontidão operacional):** PENDENTE.
**Recomendação:** READY_FOR_INDEPENDENT_AUDIT, para a porta 1 apenas (BS_TRACEABILITY_REPORT.md).

## Hashes
- Markdown V5.1 (canônico): `729971e9e71bf96572ce77bb49362e58085b9ada2b560d2bb6e2fcecc0ed14b7`
- DOCX V5.1: `0ca844ac9ff674806edfa005415ad041cae73f39dde659a7f832bd09e0769e7c`
- PDF V5.1: `859b5b61c0b3c5476f3411a665fa4894a43f16502d28a127f72c393a6630efa4`
- Demais: SHA256SUMS.txt. Artefatos V5.0: `output/historico-v5/`.

## Dimensões
- PAGE_COUNT: **100** (A4). PNGs de prova: 100 em `output/proof-png/` (locais, não versionadas).
- Sumário gerado em duas passagens; números de página estáveis.

## Verificações executadas
1. **Preservação de texto** (`tools/qa_source.py`): 3.952 itens do Markdown V5.1 presentes no DOCX; 0 ausentes. Tabelas: nenhuma faltante.
2. **Unidades:** 183 de 183 itens atribuídos a uma única unidade; nenhuma duplicata; nenhum campo de unidade vazio (script `tools/ppc_units.py`, verificado).
3. **Cargas e PPC:** tabela V conferida (3.000 = 2.200 + 800; 120 h dentro da FGB; técnicas 256 + 288 + 256).
4. **Diferença contra V5.0 (c9696f2):** nenhuma linha original alterada além das listadas em DELTA_REPORT.md (V51-01 a V51-06 e os ajustes de K15 e I4).
5. **Layout:** páginas alteradas inspecionadas (§12.4, I4, Anexo O, unidades D01–D12, mapa de 183 itens). Anexo B2: intro mantido com a tabela (`keep_with_next`).
6. **Links:** CNCT 403; gov.br 429; UNESCO sem resposta (3). Demais 200. Não confirmados.

## Não executado
- Validação OOXML no Word; renderização por LibreOffice 24.2.
- Comparação do hash do PPC com referência anterior (não registrada).
- Validação docente das unidades e atribuições (porta 3).

## Reprodução
```bash
cd docs/plano-diretor-2027-2029
python3 tools/apply_v5_1.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md /tmp/v51.json
python3 tools/make_audit.py <screening.json> output/CURRICULAR_ALIGNMENT_AUDIT.md
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md /tmp/p1.docx output/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/p1.docx && python3 tools/pagemap.py /tmp/p1.pdf /tmp/p1.docx.toc.json /tmp/pagemap.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx output/assets /tmp/pagemap.json
soffice --headless --convert-to pdf --outdir output output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx
python3 tools/qa_source.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx
```
`apply_v5_1.py` não é idempotente: rodar sobre o próprio resultado falha nas asserções, de propósito.
