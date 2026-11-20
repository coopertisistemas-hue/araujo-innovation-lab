# QA_REPORT — Planejamento Pedagógico Trienal V1 (edição premium)

**Status:** publicado para auditoria independente — não aprovado pela Direção; não homologado; não certificado operacionalmente.
**Recomendação:** READY_FOR_INDEPENDENT_AUDIT (documentário), com as decisões E1–E5 de EDITORIAL_NOTES.md registradas para o auditor e para o Master.

## Hashes
- Fonte canônica: `31ba5f61fe4c5c16e9cb00dd44c82065d7a429777692415df8b4bd4e1539fa80`
- DOCX: `0788524d491cb8437e7663e0f533d22237a774049f857e67d8ab4eb3b2654aa8`
- PDF: `42f4eca84be299e8abd101b55334828c48a0dab14331f2e31736c07e158d90ad` (50 páginas, A4)
- Demais: SHA256SUMS.txt.

## Testes (tests/test_v1_results.json, `tools/test_v1.py`)
- **21/21 aprovados**: estrutura (21 seções, anexos A–J, divisor); 12 componentes; 49 unidades referenciadas; cargas 3.000 / 2.200 / 800 / 120 h e 256 / 288 / 256 h; duas coortes; K01 em média ponderada; K15 não aplicável; K16 taxa de ocupação; status NÃO HOMOLOGADA; nenhuma afirmação positiva de aprovação; SEM DADOS presente; nomes de docentes sempre com "a confirmar"; 5 figuras com texto alternativo; cabeçalho e rodapé; sumário com 33 entradas resolvidas; PDF com 50 páginas.

## Verificações adicionais
- `tools/qa_source.py`: 1.224 itens presentes; 0 ausentes.
- Sumário: estável em duas passagens.
- Inspeção visual: 50 páginas em cinco folhas de contato a 45 dpi, mais PNGs de prova em `proof-png/` (locais, não versionados).

## Limitações
- Sem validação OOXML no Word; renderização LibreOffice 24.2.
- Sem conferência de página a página em alta resolução.

## Reprodução
```bash
cd docs/plano-diretor-2027-2029
python3 tools/make_v1_diagrams.py output/planejamento-trienal-v1/assets
python3 tools/build_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md /tmp/f1.docx output/planejamento-trienal-v1/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/f1.docx && python3 tools/pagemap.py /tmp/f1.pdf /tmp/f1.docx.toc.json /tmp/fm1.json
python3 tools/build_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx output/planejamento-trienal-v1/assets /tmp/fm1.json
soffice --headless --convert-to pdf --outdir output/planejamento-trienal-v1 output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx
python3 tools/test_v1.py PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.md output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.docx output/planejamento-trienal-v1/PLANEJAMENTO_PEDAGOGICO_TRIENAL_MARKETING_ARAUJO_2027-2029_PROPOSTA_V1.pdf /tmp/fm1.json output/planejamento-trienal-v1/tests/test_v1_results.json
```
