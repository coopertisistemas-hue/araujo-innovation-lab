import subprocess, json, re, sys
pdf, toc_json, out = sys.argv[1], sys.argv[2], sys.argv[3]
n = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout).group(1))
norm = lambda s: re.sub(r"\s+", " ", s).strip()
pages = [norm(subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"], capture_output=True, text=True).stdout) for i in range(1, n+1)]
toc_page = next(i for i, t in enumerate(pages) if "Sumário" in t and "Numeração de página" in t)
items = json.load(open(toc_json))
res = {}
for it in items:
    key = norm(it)
    found = None
    for i in range(toc_page+1, n):
        # heading text appears at page start region
        if key[:60] in pages[i]:
            found = i+1; break
    res[it] = found
    print(found, it[:70])
json.dump(res, open(out, "w"), ensure_ascii=False, indent=1)
print("pages:", n, "toc_page:", toc_page+1, "missing:", [k for k,v in res.items() if v is None])
