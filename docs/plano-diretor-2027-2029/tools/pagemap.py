import subprocess, json, re, sys
pdf, toc_json, out = sys.argv[1], sys.argv[2], sys.argv[3]
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
import unicodedata
norm = lambda s: re.sub(r"\s+", " ", unicodedata.normalize("NFKD", s.replace("-", "")).encode("ascii","ignore").decode()).strip()
pages = [norm(subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), pdf, "-"], capture_output=True, text=True).stdout) for i in range(1, n+1)]
toc_page = next(i for i, t in enumerate(pages) if norm("Sumário") in t and norm("Numeração de página") in t)
items = json.load(open(toc_json))
res = {}
for it in items:
    key = norm(it)[:40]
    found = None
    for i in range(toc_page+1, n):
        # heading text appears at page start region
        if key in pages[i]:
            found = i+1; break
    res[it] = found
    print(found, it[:70])
json.dump(res, open(out, "w"), ensure_ascii=False, indent=1)
print("pages:", n, "toc_page:", toc_page+1, "missing:", [k for k,v in res.items() if v is None])
