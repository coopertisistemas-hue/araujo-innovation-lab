import re, sys, json, unicodedata
sys.path.insert(0, __file__.rsplit("/",1)[0])
from ppc_data import COMPONENTS
def norm(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode().lower()
    return re.sub(r"[^a-z0-9 ]+"," ", s)
STOP = set("para pelo pela com das dos nas nos uma uns umas que sobre entre como seus suas sem aos ao por mais num numa esta este isto seu sua the and dos das dos ou".split())
def toks(label):
    label = label.replace("[hab.]","")
    return [t for t in norm(label).split() if len(t) >= 4 and t not in STOP]
def blocks(md, comp):
    L = md.split("\n")
    out = []
    for i,l in enumerate(L):
        if l.startswith("### ") and (l.startswith("### "+comp["name"]+" |") or re.match(r"### %s\. "%comp["code"], l)):
            j = i+1
            while j < len(L) and not (L[j].startswith("### ") or L[j].startswith("## ")): j += 1
            out.append("\n".join(L[i:j]))
        if l.startswith("## ANEXO E") and comp["code"]=="D02":
            j=i+1
            while j < len(L) and not L[j].startswith("## "): j+=1
            out.append("\n".join(L[i:j]))
    return "\n".join(out)
if __name__ == "__main__":
    md = open(sys.argv[1], encoding="utf-8").read()
    res = []
    for c in COMPONENTS:
        b = norm(blocks(md, c))
        row = {"code": c["code"], "name": c["name"], "objects": []}
        for o in c["objetos"]:
            t = toks(o)
            hit = [x for x in t if re.search(r"\b"+re.escape(x), b)]
            r = len(hit)/max(1,len(t))
            status = "COBERTO" if r >= 0.6 else ("PARCIAL" if r >= 0.3 else "AUSENTE")
            row["objects"].append({"obj": o, "ratio": round(r,2), "status": status})
        res.append(row)
    json.dump(res, open(sys.argv[2],"w",encoding="utf-8"), ensure_ascii=False, indent=1)
    tot = {"COBERTO":0,"PARCIAL":0,"AUSENTE":0}
    for r in res:
        for o in r["objects"]: tot[o["status"]] += 1
    print("pre-fix totals:", tot)
    for r in res:
        miss = [o["obj"] for o in r["objects"] if o["status"]!="COBERTO"]
        print(f"{r['code']} {r['name']}: {sum(1 for o in r['objects'] if o['status']=='COBERTO')}/{len(r['objects'])} cobertos; não cobertos: {len(miss)}")
