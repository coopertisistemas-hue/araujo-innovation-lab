# QA_REPORT — Planejamento Pedagógico Trienal V1 (delta final E1–E3)

**Status:** publicado para auditoria independente — não aprovado pela Direção; não homologado; não certificado operacionalmente. Sem merge para main.
**Recomendação:** READY_FOR_INDEPENDENT_AUDIT.

## Hashes
- Fonte canônica: `bcca4e371318f8631af91a24436c973045215dd4c3d12162a76c6fa5c2dac3ab`
- DOCX: `70d3553892184c07dce705eedd5a28f2159e0b12625c3d3f2703c993760782f0`
- PDF: `cb31c00ebad654ab448b9611c225ae6b328a24aac4b92f6c124ff539eea7e1a9` (50 páginas, A4)
- Demais: SHA256SUMS.txt.

## Testes (tests/test_v1_results.json)
- **23/23 aprovados.** Inclui: nenhuma pessoa nomeada no MD, DOCX e PDF (E1); expediente em funções; 21 seções; anexos A–J; 12 componentes; 49 unidades; cargas 3.000 / 2.200 / 800 / 120 h e 256 / 288 / 256 h; duas coortes; K01, K15 e K16 conforme o PDP; status não homologado; nenhuma afirmação positiva de aprovação; SEM DADOS; 5 figuras com texto alternativo; cabeçalho e rodapé; sumário com 33 entradas resolvidas; PDF com 50 páginas.

## Verificações adicionais
- `tools/qa_source.py`: 1.224 itens da fonte presentes no DOCX; 0 ausentes.
- Sumário em duas passagens estável; "Sumário" agora começa em página própria (p. 3); nenhuma entrada sem número.
- Inspeção em alta resolução (110 dpi): páginas 2, 8, 10, 12, 13, 24, 25, 27 e 36; ciclo de seis movimentos (E3) legível; §6.3 (E2) corrigido; sem recortes nas tabelas densas.
- Folhas de contato das 50 páginas em 45 dpi (inspeção anterior a este delta).

## Correção feita durante a inspeção
- O título "Sumário" estava no pé da ficha de controle (p. 2). Adicionada quebra de página antes do sumário; o sumário passou a ocupar a p. 3. Entradas e numeração revalidadas.

## Limitações
- Sem validação OOXML no Word.
- Rótulos das figuras 3–5 pequenos em impressão (registrado em EDITORIAL_NOTES).

## Reprodução
```bash
cd docs/plano-diretor-2027-2029
python3 tools/make_v1_diagrams.py output/planejamento-trienal-v1/assets
python3 tools/build_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md /tmp/g1.docx output/planejamento-trienal-v1/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/g1.docx && python3 tools/pagemap.py /tmp/g1.pdf /tmp/g1.docx.toc.json /tmp/gm1.json
python3 tools/build_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx output/planejamento-trienal-v1/assets /tmp/gm1.json
soffice --headless --convert-to pdf --outdir output/planejamento-trienal-v1 output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx
python3 tools/test_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.pdf /tmp/gm1.json output/planejamento-trienal-v1/tests/test_v1_results.json
```
