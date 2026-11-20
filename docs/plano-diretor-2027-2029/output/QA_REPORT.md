# QA report — Plano Diretor Pedagógico Marketing 2027–2029, V5 revisado (proposta, não homologada)

**Status:** PUBLISHED FOR INDEPENDENT REVIEW — NOT CERTIFIED, NOT APPROVED BY SCHOOL DIRECTION.
**Recomendação de auditoria:** HOLD (ver BS_TRACEABILITY_REPORT.md: 5 requisitos PARCIAIS; PPC oficial ausente).

## Hashes
- Markdown canônico: `875ec13184d20c1187e80b945b2c8040240e46f276fe9662a69a1de94679dfa3` (1.116 linhas; V5 anterior: `036d5f61…`)
- Matriz BS de entrada: `94ca0f389802668de289bd60b5371dc656ae8f3877bbca88f8d25105aa9f2714`
- DOCX: ver SHA256SUMS.txt · PDF: ver SHA256SUMS.txt
- PAGE_COUNT: 58 (A4) · TOC estável em duas passagens.

## Verificações executadas nesta rodada
1. **Preservação de conteúdo** (`tools/qa_source.py`): 1.514 itens de texto do Markdown revisado encontrados no DOCX; 0 ausentes. 41 tabelas do Markdown presentes no DOCX (multiset de contagem de linhas: nenhuma ausente).
2. **Diferença de linhas entre V5 anterior e revisado:** 15 linhas originais não aparecem mais — 12 linhas antigas do painel §9, 1 cláusula do §14.1 e 2 linhas reescritas (§14.2 G1 e §14.5). Todas as outras linhas são reordenação ou inserção, conforme DELTA_REPORT.md.
3. **Aritmética curricular interna:** 800 h técnicas (64×11 + 96), 256 + 288 + 256 = 800, 2.200 + 800 = 3.000, 120 h dentro de 2.200, Anexo E = 64 h. Não verificada contra o PPC.
4. **Layout:** TOC com números de página estável entre duas passagens; 3 páginas com menos de 90 palavras (p. 17: figura de governança; p. 45: tabela de Anexo G; p. 57: final da lista de regras editoriais), todas intencionais.
5. **Inspeção visual:** todas as 58 páginas geradas em PNG (`output/proof-png/`, local). Páginas alteradas inspecionadas em montagem (§4.1, §5.1, §9, §15.5–15.7, Anexo C, Anexo H5, Anexo K). Cabeçalhos, rodapés, quebras de página e tabelas com cabeçalho repetido conferidos.
6. **Hashes e links:** links de benchmarks não reconfirmados nesta rodada: 3 UNESCO sem resposta, 2 CNCT (403 e 429). Os demais responderam 200.

## Não executado
- Validação OOXML no Word (`validate.py`). A renderização usou LibreOffice 24.2.
- Conferência item a item com o PPC: o PDF oficial não está disponível (ver CURRICULAR_ALIGNMENT_AUDIT.md).

## Reprodução
```bash
python3 tools/make_charts.py output/assets            # figuras (opcional; já versionadas)
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md /tmp/pass1.docx output/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/pass1.docx
python3 tools/pagemap.py /tmp/pass1.pdf /tmp/pass1.docx.toc.json /tmp/pagemap.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx output/assets /tmp/pagemap.json
soffice --headless --convert-to pdf --outdir output output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx
python3 tools/qa_source.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx
```
A revisão de conteúdo foi aplicada por `tools/edit_md.py` sobre o Markdown do commit 4e82c5d. O script não é idempotente: aplicá-lo novamente falha nas asserções, o que é intencional.
