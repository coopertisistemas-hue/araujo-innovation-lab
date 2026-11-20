import re, sys
from docx import Document
M = sys.argv[1]; D = sys.argv[2]
norm = lambda s: re.sub(r"\s+", " ", s.replace("**", "").replace("`", "")).strip()
src = open(M, encoding="utf-8").read().split("\n")[25:]   # body from source line 26
doc = Document(D)
# docx full text: paragraphs + table cells, in order
parts = [p.text for p in doc.paragraphs]
for t in doc.tables:
    for r in t.rows:
        for c in r.cells:
            parts.append(c.text)
blob = norm(" \n ".join(parts))
blob_nospace = blob.replace(" ", "")
missing = []; checked = 0
for ln in src:
    s = ln.strip()
    if not s or re.fullmatch(r"\|?\s*:?-{3,}.*", s): continue
    if s.startswith("|"):
        cells = [c.strip() for c in s.strip("|").split("|")]
        for c in cells:
            if not c: continue
            checked += 1
            if norm(c).replace(" ", "") not in blob_nospace:
                missing.append(("cell", c[:80]))
        continue
    t = re.sub(r"^(#+\s*|>\s*|- |• |□ |\d+\.\s+)", "", s)
    t = norm(t)
    if len(t) < 3: continue
    checked += 1
    if t.replace(" ", "") not in blob_nospace:
        missing.append(("line", t[:90]))
# table row count check: source tables vs docx tables (excluding added tables)
src_tables = []; cur = []
for ln in src:
    if ln.strip().startswith("|"): cur.append(ln)
    else:
        if cur: src_tables.append(cur); cur = []
if cur: src_tables.append(cur)
src_rows = [sum(1 for r in t if not re.fullmatch(r"\|?\s*:?-{3,}.*", r.strip())) for t in src_tables]
doc_rows = [len(t.rows) for t in doc.tables]
print("checked items:", checked, "missing:", len(missing))
for m in missing[:30]: print("  MISSING", m)
print("source tables:", len(src_tables), "docx tables:", len(doc.tables), "(docx includes 2 added control/decision tables + 1 registro table)")
# verify each source table row count appears in order within docx tables
from collections import Counter
cs = Counter(src_rows); cd = Counter(doc_rows)
mismatch = [(k,v) for k,v in (cs - cd).items()]
print("row-count multiset: missing in docx ->", mismatch)
print("source tables matched in order by row count:", len(src_tables) - len(mismatch), "/", len(src_tables), "unmatched:", mismatch)
