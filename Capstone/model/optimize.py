"""Hill-climbing search over crashing, expediting, transport mode and generator shipping decisions (max NPV)."""
import json, copy, sys, datetime as dt
from engine import *

def score(dec):
    r = evaluate(dec)
    pen = 0.0
    if not r["feasible"]:
        pen = 1e12 + (r["cf"]["fpa"] - FIN["deadline"]).days * 1e9
    return r["npv"] - pen, r

def neighbours(dec):
    # crashing: every integer duration between min and normal
    for aid, a in ACTS.items():
        if a["mindur"]:
            cur = dec["crash"].get(aid, a["dur"])
            for d in range(a["mindur"], a["dur"] + 1):
                if d != cur:
                    n = copy.deepcopy(dec); n["crash"][aid] = d
                    yield f"crash {aid} -> {d} wd", n
    for pkg, pos in dec["pos"].items():
        region = PK[pkg]["origin"]
        for i, po in enumerate(pos):
            n = copy.deepcopy(dec); n["pos"][pkg][i]["expedite"] = not po["expedite"]
            yield f"{'expedite' if not po['expedite'] else 'un-expedite'} {pkg} lot {po['lot']}", n
            for (rg, m) in TR:
                if rg != region or m == po["mode"]: continue
                dis = dec["gen_disassembled"][po["lot"]] if pkg == "Generators" else False
                if shipment_charge(pkg, 1, m, dis) is None: continue
                n = copy.deepcopy(dec); n["pos"][pkg][i]["mode"] = m
                yield f"{pkg} lot {po['lot']} by {m}", n
    for k in (1, 2):
        n = copy.deepcopy(dec); n["gen_disassembled"][k] = not dec["gen_disassembled"][k]
        if n["gen_disassembled"][k]:
            # when disassembled, air becomes possible; keep mode as is (surface) here, separate move handles air
            pass
        else:
            po = n["pos"]["Generators"][k - 1]
            if shipment_charge("Generators", 1, po["mode"], False) is None:
                po["mode"] = "Road (truck)"
        yield f"generators lot {k} {'disassembled' if n['gen_disassembled'][k] else 'assembled'}", n

def climb(dec, log):
    best, r = score(dec)
    log.append(dict(step=0, move="baseline", npv_k=r["npv"] / 1000, hall_a=str(r["cf"]["hall_a"]), hall_b=str(r["cf"]["hall_b"]), fpa=str(r["cf"]["fpa"]), d=r["d"], feasible=r["feasible"]))
    step = 0
    while True:
        step += 1
        cands = []
        for name, n in neighbours(dec):
            try:
                sc, rr = score(n)
            except AssertionError as e:
                continue
            cands.append((sc, name, n, rr))
        cands.sort(key=lambda x: -x[0])
        if not cands or cands[0][0] <= best + 1.0:
            break
        best, name, dec, r = cands[0]
        log.append(dict(step=step, move=name, npv_k=r["npv"] / 1000, hall_a=str(r["cf"]["hall_a"]), hall_b=str(r["cf"]["hall_b"]), fpa=str(r["cf"]["fpa"]), d=r["d"], feasible=r["feasible"]))
        print(step, name, round(r["npv"] / 1e6, 2), r["cf"]["hall_a"], r["cf"]["hall_b"], r["cf"]["fpa"], r["d"], flush=True)
    return dec, best

def is_relax(name, dec):
    if name.startswith("un-expedite"): return True
    if name.startswith("crash"):
        aid = name.split()[1]; d = int(name.split("-> ")[1].split()[0])
        return d > dec["crash"].get(aid, ACTS[aid]["dur"])
    if " by " in name:
        pkg = name.split(" lot ")[0]; lot = int(name.split(" lot ")[1].split()[0]); mode = name.split(" by ")[1]
        cur = dec["pos"][pkg][lot - 1]["mode"]
        return TR[(PK[pkg]["origin"], mode)]["rate"] < TR[(PK[pkg]["origin"], cur)]["rate"]
    return False

def prune(dec, log):
    """Relax acceleration that does not pay: best-improvement over un-crash / un-expedite / cheaper-mode moves."""
    best, r = score(dec)
    step = 0
    while True:
        step += 1
        cands = []
        for name, n in neighbours(dec):
            if not is_relax(name, dec): continue
            try:
                sc, rr = score(n)
            except AssertionError:
                continue
            cands.append((sc, name, n, rr))
        cands.sort(key=lambda x: -x[0])
        if not cands or cands[0][0] < best - 1.0:
            break
        best, name, dec, r = cands[0]
        log.append(dict(step=f"relax{step}", move=name, npv_k=r["npv"] / 1000, hall_a=str(r["cf"]["hall_a"]), hall_b=str(r["cf"]["hall_b"]), fpa=str(r["cf"]["fpa"]), d=r["d"], feasible=r["feasible"]))
        print("relax", name, round(r["npv"] / 1e6, 3), r["cf"]["hall_a"], r["cf"]["hall_b"], r["cf"]["fpa"], flush=True)
    return dec, best

def max_accel():
    dec = default_decisions()
    for aid, a in ACTS.items():
        if a["mindur"]:
            dec["crash"][aid] = a["mindur"]
    for pkg, pos in dec["pos"].items():
        region = PK[pkg]["origin"]
        fastest = min((m for (rg, m) in TR if rg == region and shipment_charge(pkg, 1, m, False) is not None), key=lambda m: TR[(region, m)]["cd"])
        for po in pos:
            po["expedite"] = True; po["mode"] = fastest
    return dec

def serial(dec):
    return dict(crash=dec["crash"], pos=dec["pos"], gen_disassembled={str(k): v for k, v in dec["gen_disassembled"].items()},
                jit=dec["jit"], po_offset={f"{k[0]}|{k[1]}": v for k, v in dec["po_offset"].items()})

def unserial(d):
    return dict(crash={k: int(v) for k, v in d["crash"].items()}, pos=d["pos"], gen_disassembled={int(k): v for k, v in d["gen_disassembled"].items()},
                jit=d["jit"], po_offset={(k.split("|")[0], int(k.split("|")[1])): v for k, v in d["po_offset"].items()})

if __name__ == "__main__":
    results = []
    # start 1: baseline (no acceleration) -> climb
    log1 = []
    dec1, best1 = climb(default_decisions(), log1)
    dec1, best1 = prune(dec1, log1)
    dec1, best1 = climb(dec1, log1)
    results.append((best1, dec1, log1, "baseline-climb"))
    # start 2: maximum acceleration -> relax -> climb
    log2 = []
    dec2 = max_accel()
    sc, r = score(dec2)
    log2.append(dict(step="max-accel", move="all crashed to minimum, all POs expedited, fastest transport", npv_k=r["npv"] / 1000, hall_a=str(r["cf"]["hall_a"]), hall_b=str(r["cf"]["hall_b"]), fpa=str(r["cf"]["fpa"]), d=r["d"], feasible=r["feasible"]))
    print("max-accel", round(r["npv"] / 1e6, 3), r["cf"]["hall_a"], r["cf"]["hall_b"], r["cf"]["fpa"], flush=True)
    for rnd in range(2):
        dec2, best2 = prune(dec2, log2)
        dec2, best2 = climb(dec2, log2)
    results.append((best2, dec2, log2, "max-accel-relax"))
    results.sort(key=lambda x: -x[0])
    best, dec, log, name = results[0]
    print("BEST START:", name, [(n, round(b / 1e6, 3)) for b, _, _, n in results])
    json.dump(dict(decisions=serial(dec), log=log, start=name, alternatives=[(n, b / 1000) for b, _, _, n in results]), open("solution.json", "w"), indent=1, default=str)
    r = evaluate(dec)
    print("FINAL", summary(r))
