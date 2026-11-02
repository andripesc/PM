import json, sys
from engine import *
from optimize import unserial
dec = unserial(json.load(open(sys.argv[1] if len(sys.argv)>1 else "solution.json"))["decisions"])
r = evaluate(dec)
print(summary(r))
S, T, P = r["sched"], r["tasks"], r["preds"]
def chain(t, depth=0):
    s = S[t]
    print("  "*depth + f"{t} {wd_date(s['start'])} -> {wd_date(max(s['finish'],s['start']))} dur {T[t]['dur']}")
    # binding preds: finish+1 == start
    for p in P[t]:
        if S[p]["finish"] + 1 == s["start"]:
            chain(p, depth+1); return
    # delivery binding?
    for (pkg, lot), info in r["deliveries"].items():
        if next_wd_after(info["delivery"]) == s["start"] and t in [("F46" if lot==1 else "F47") if (pkg=="Generators" and dec["gen_disassembled"][lot]) else a for a in NEEDS[pkg][lot]]:
            print("  "*(depth+1) + f"DELIVERY {pkg} lot {lot}: PO {info['po_date']} ready {info['ready']} delivered {info['delivery']} ({info['mode']}, exp={info['expedite']})")
            return
    # resource wait
    print("  "*(depth+1) + f"(resource wait / SNET; link-est={max([S[p]['finish']+1 for p in P[t]] or [0])})")
for m in ("MS_HALL_A","MS_HALL_B","MS_FPA"):
    print("=====", m); chain(m)
