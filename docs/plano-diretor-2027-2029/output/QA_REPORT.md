# QA_REPORT — candidato de fechamento (ordem §2.1 → §2.2)

**Status:** publicado para auditoria independente — não aprovado pela Direção; não certificado operacionalmente.
**Recomendação:** READY_FOR_INDEPENDENT_AUDIT para a porta documental; promoção para main pendente de autorização do Master.

## Hashes
- Markdown V5.1: `48f5b2929583e73277759d68314a43528cf0e9961bd551f3e0a761815e6f052e`
- DOCX: `46d2ca8d00d446505ed446cfad1f18729e341759c753791f35f76545347b9853`
- PDF: `fa4a1cddd497b24afda7fa2a39b967b02e5d81f08cb95e6e63e121e79b28fbcf` (101 páginas)
- Demais: SHA256SUMS.txt.

## Testes (tests/test_results.json)
- **32/32 aprovados**, incluindo a regressão de hash do BS_TRACEABILITY_REPORT (hash do documento-alvo igual ao arquivo atual).
- Página esperada: 101 (ajustada, pois a reordenação de §2.1/§2.2 alterou a paginação).

## Verificações adicionais
- `tools/qa_source.py`: 3.970 itens do Markdown presentes no DOCX; 0 ausentes.
- TOC: estável em duas passagens.
- Seções: §2.1 antes de §2.2 no PDF (verificado pelo texto do PDF).
- Páginas com pouco texto: 19, 52, 66 — fins de seção.
- PNGs de prova: 101 em `output/proof-png/` (locais, não versionados).

## Limitações
- Sem validação OOXML no Word; renderização por LibreOffice 24.2.
- Referências: 5 de 13 links não confirmados (3 UNESCO, 2 CNCT).

## Reprodução
Ver comandos em `tools/` e nos relatórios anteriores; o gerador (`apply_v5_1.py`) reconstrói o Markdown a partir do V5.0 preservado.
