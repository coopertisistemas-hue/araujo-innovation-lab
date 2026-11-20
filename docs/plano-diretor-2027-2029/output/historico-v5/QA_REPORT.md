# QA_REPORT — Plano Diretor Pedagógico Marketing 2027–2029, V5 revisado (v4)

**Status:** PUBLICADO PARA AUDITORIA INDEPENDENTE — NÃO CERTIFICADO; NÃO APROVADO PELA DIREÇÃO.
**Recomendação:** **HOLD** (ver BS_TRACEABILITY_REPORT.md: G-BS-02 PARCIAL; objetivo por encontro e atividade por objeto pendentes).

## Hashes
- Markdown canônico: `8c6500805940b32894d7b413b1ff1038f6f79b027baf4d207ecb18fb378df0ba`
- DOCX: `07add14dee14bed37de35f59bc3ad59f6e0bdfc501cfe6edf8cca695f4fb5142`
- PDF: `1150ab3470e30f7c1c1948988117862f983be28cf9dbcd5548c09809f6f49b63`
- Demais hashes: SHA256SUMS.txt.

## Dimensões
- PAGE_COUNT: **78** (A4). PNGs de prova locais: 78 (`output/proof-png/`, não versionados).
- Sumário (TOC): gerado em duas passagens; números de página estáveis.

## Verificações executadas
1. **Preservação de conteúdo** (`tools/qa_source.py`): 2.269 itens de texto do Markdown encontrados no DOCX; 0 ausentes. 56 tabelas do Markdown presentes (contagem de linhas, sem faltantes).
2. **Diferença contra a versão publicada (a795e98):** 22 linhas originais substituídas, todas listadas em DELTA_REPORT.md. Nenhuma outra linha perdida.
3. **Aritmética e PPC:** tabela V do PPC conferida (3.000 = 2.200 + 800; 120 h integradoras dentro da FGB; 1º 256 · 2º 288 · 3º 256 h). Os 183 objetos e os 12 componentes foram auditados (CURRICULAR_ALIGNMENT_AUDIT.md).
4. **Layout:** inspeção visual das páginas alteradas (§2.1, §9, §15, Anexo B2, C4, D01–D12, G2, H5, I4, N). Páginas com pouco texto (17, 19, 50, 64) são fins de seção. Um cabeçalho de tabela fica em página seguinte (B2), cosmético.
5. **Cabeçalhos, rodapés, quebras de página e tabelas com cabeçalho repetido:** conferidos nas páginas alteradas.

## Não executado
- Validação OOXML no Word. A renderização usou LibreOffice 24.2.
- Comparação do hash do PPC com referência anterior: não há referência registrada.
- Links de benchmarks: 3 UNESCO sem resposta; 2 CNCT (403 e 429). Os demais responderam 200.
- Confirmação, com o docente habilitado, da atividade de cada objeto e do objetivo por encontro (PENDENTE_DE_VALIDAÇÃO).

## Reprodução
```bash
cd docs/plano-diretor-2027-2029
python3 tools/apply_v3.py <md_commit_a795e98> /tmp/v3.md /tmp/v3.json
python3 tools/apply_v4.py /tmp/v3.md Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md /tmp/v4.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md /tmp/p1.docx output/assets
soffice --headless --convert-to pdf --outdir /tmp /tmp/p1.docx
python3 tools/pagemap.py /tmp/p1.pdf /tmp/p1.docx.toc.json /tmp/pagemap.json
python3 tools/build_docx.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx output/assets /tmp/pagemap.json
soffice --headless --convert-to pdf --outdir output output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx
python3 tools/qa_source.py Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md output/PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx
```
`apply_v3.py` e `apply_v4.py` não são idempotentes: rodar de novo sobre o próprio resultado falha nas asserções, de propósito.
