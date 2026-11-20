# QA_REPORT — candidato corrigido V5.1

**Status:** publicado para auditoria independente — não auditado por terceiros; não aprovado pela Direção; não certificado operacionalmente.
**Recomendação:** READY_FOR_INDEPENDENT_AUDIT, apenas para a porta documental (ver BS_TRACEABILITY_REPORT.md).

## Hashes
- Markdown V5.1: `13d5bc05d0026486f230076901e1dedbda83f6cff52e452c485ba8fd915dda23`
- DOCX: `1fcb818ed63d9c69bc29699b8b1d615af2579d37e77fdbc57ffa8db205a2f0e8`
- PDF: `d3f92e6a45b30ffc0e4ed695be015822af46610f4f35c50106fe5b323045b11e`
- Demais: SHA256SUMS.txt.

## Testes (tests/test_results.json, `tools/test_release.py`)
- **32/32 aprovados.**
- Matemática: K01 (exemplo 3,15; limites; 10.000 vetores em [0,4]; defeito antigo quantificado); K15 (75%; não aplicável); K16 (66,7%; não aplicável); carga (256, 288, 256, 800; 744+712+744 = 2.200; 120 dentro; 3.000); mapeamento (183 itens, uma unidade cada; 49 unidades; nenhum campo vazio).
- Integridade de documento: versão 5.1 no front matter e cabeçalho; sem "Versão 5.0" atual; histórico no Anexo O; afirmação "acesso em 19/11/2026" removida; K01 sem "soma de pontos"; 5 figuras com texto alternativo; rodapé 20/11/2026; PDF com 100 páginas.

## Verificações adicionais
- `tools/qa_source.py`: 3.970 itens do Markdown presentes no DOCX; 0 ausentes.
- TOC: números de página estáveis em duas passagens (100 páginas).
- Páginas com pouco texto: 18, 51, 65 — fins de seção, intencionais.
- Renderização: LibreOffice 24.2; PNGs de prova locais (100) em `output/proof-png/`.

## Limitações
- Não há validação OOXML no Word.
- Não há revisão visual de cada página em alta resolução nesta rodada; a verificação por montagem e os testes automáticos cobrem os pontos alterados.
- Referências: 5 de 13 links não confirmados (3 UNESCO sem resposta; 2 CNCT com 403 e 429).
- K15: a interpretação (anual, n/N, briefs fora do indicador) foi aplicada na ausência do texto da diretriz de correções; ver DELTA_REPORT.

## Reprodução
```bash
python3 tools/apply_v5_1.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md /tmp/v51.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md /tmp/p1.docx output/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/p1.docx && python3 tools/pagemap.py /tmp/p1.pdf /tmp/p1.docx.toc.json /tmp/pagemap.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx output/assets /tmp/pagemap.json
soffice --headless --convert-to pdf --outdir output output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx
python3 tools/qa_source.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx
python3 tools/test_release.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5_1.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.docx output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5_1.pdf output/tests/test_results.json
```
