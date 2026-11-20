# QA report — Plano Diretor Pedagógico Marketing 2027–2029, v5.0 (proposta, não homologada)

SOURCE_SHA256: 036d5f61382a3b191fe4560e59006eb713402238e28579c33c07e2dfb9e96d6e  (Plano_Diretor_Pedagogico_Col_Araujo_2027-2029_COMPLETO_V5.md, 1038 lines; not modified)
REPO_REFERENCE: coopertisistemas-hue/araujo-innovation-lab — STYLEGUIDE.md, templates/executive-document-template.md (read only)
BRAND_ASSETS_VERIFIED: none. branding/logo, branding/palette, branding/manual, branding/templates are empty. No HEX values and no logo validated. Provisional tokens only: navy #1F2F4D, gold #B8975A, warm grey #8C8279, cream #F4F0E8. No logo, crest or seal used.
OUTPUT_DOCX: PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.docx
OUTPUT_PDF: PLANO_DIRETOR_PEDAGOGICO_MARKETING_ARAUJO_2027-2029_PROPOSTA_V5.pdf (LibreOffice render)
PAGE_COUNT: 58 (A4)

SECTION_COVERAGE: all of sections 1–17 and annexes A–M present in source order. 12 technical subject plans (Anexo A, D01–D12), KPI dictionary K01–K16 (G2), 38 source tables, 16 KPIs, references (M2, 17). Added, not in source: cover, document control sheet, static TOC, five figures, closing decision page (after §17).

RENDER_QA: all 58 pages rasterized (PNG at 110 dpi, output/proof-png/) and inspected in contact sheets. Fixed during QA: code column wrapping (13.1), word-breaking headers (G3, RACI, 15.5), RACI legend orphaned, single-bullet page. Last render: no blank pages, TOC page numbers stable across two passes. Not run: Word-native open or OOXML validation.

TABLE_QA: 38 of 38 source tables present in order by row count. Column widths sized from longest word. Header rows repeat across pages.

FACTUAL_QA:
- Source-text check: 1390 text items from the MD found in the DOCX, 0 missing.
- Hours: 2.200 FGB + 800 técnico = 3.000. 120 h integradoras inside 2.200, not added again. 2.080 = 2.200 − 120 is derived and labelled as such in figure 1.
- Technical components: 1º 256 h, 2º 288 h, 3º 256 h = 800 h. Annex E blocks 12+14+12+12+14 = 64 h.
- Figure 4 values match §13.3 exactly (checked programmatically).
- Charts labelled PROPOSED TARGETS / NOT OBSERVED RESULTS or as planned, not observed.
- External links (§M2, §17), tested with curl on 2026-11-20: 200 for github, senac (x2), google support (x2), senai, sc.senac matriz. 403 for cnct.mec.gov.br (bot block). 429 for the gov.br CNCT PDF (rate limit). No response (000) for the three UNESCO URLs. The last four are unconfirmed, not confirmed broken.

STATUS: fully authored and rendered. Pending institutional approval. Not CERTIFIED. No approval, signature or homologation is implied anywhere in the document.

LIMITATIONS:
1. No verified brand assets. Provisional tokens only; no logo; the header carries text, not a logo.
2. Fonts: Georgia (headings) and Calibri (body) are set in the DOCX. In this VM the proof PDF substitutes DejaVu Serif for Georgia (wider, so Word should fit at least as well) and Carlito for Calibri (metric-compatible). Word pagination may differ slightly.
3. TOC is static: page numbers were computed from the proof render. Regenerate after any edit.
4. Figure 4 is a heatmap, not the line chart suggested in §13.3. The source does not map the 12 competencies to its four groupings, so a line chart would require inventing that mapping.
5. Source YAML front matter (lines 1–8) is not reproduced verbatim. Its information appears in the control sheet (version, date, status, PPC reference, institution in the footer).
6. Path discrepancy: the first paste pointed to /home/ubuntu/connect/araujo-innovation-lab/..., which does not exist. The same-named file in the working directory was used (SHA256 above).
7. Source flags to verify before implementation, preserved as written: PPC minimum instrument counts (§15.1), NotebookLM age, licensing and privacy terms (Anexo H2.6), SENAC matrix version (§2). No legally inconsistent passage was found that needed a silent change.
8. Charts were drawn with matplotlib from source values; the alt text is set on each image.

NEXT_APPROVAL_GATE: Direção deliberation, §16 (PROPOSTA_EM_ANÁLISE → APROVADO / APROVADO_COM_RESSALVAS / REVISÃO_SOLICITADA / NÃO_APROVADO). Teacher consultation (Anexo I1) only after acceptance.

## Publication and local evidence (added at repository sync)
- PROOF_PNG_LOCAL: docs/plano-diretor-2027-2029/output/proof-png/ — 58 page PNGs at 110 dpi (≈9.4 MB). Kept local, not committed, to avoid bulk binaries; regenerate from the committed PDF with `pdftoppm -r 110 -png <pdf> page`.
- Committed for review: source Markdown (unchanged), DOCX, PDF, QA_REPORT.md, SHA256SUMS.txt, assets/ (five charts).
- Status: PUBLISHED FOR INDEPENDENT REVIEW, NOT CERTIFIED OR APPROVED BY SCHOOL DIRECTION.
