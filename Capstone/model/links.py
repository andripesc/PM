"""Derive explicit resource-sequencing links so that a link-only CPM (as MS Project computes it, without leveling)
reproduces the resource-feasible schedule produced by the engine."""
from collections import defaultdict
from engine import *

def shares_resource(tasks, a, b):
    ra = {r for r, u in tasks[a]["units"]}
    rb = {r for r, u in tasks[b]["units"]}
    if ra & rb:
        return True
    return a in CORRIDOR_ACTS and b in CORRIDOR_ACTS

def derive_links(dec, r):
    """Returns list of (pred, succ, lag_wd) extra links and the verified link-only schedule."""
    tasks, sched, preds, snet = r["tasks"], r["sched"], r["preds"], r["snet"]
    extra = []
    for it in range(20):
        # link-only CPM with current extra links
        ex = [(p, s) for p, s, lag in extra if lag == 0]
        lagged = [(p, s, lag) for p, s, lag in extra if lag > 0]
        r2 = schedule(dec, snet, ex, resource_check=False)
        # apply lagged links manually: not supported by schedule(); emulate via snet
        s2 = r2["sched"]
        mismatch = [t for t in tasks if s2[t]["start"] != sched[t]["start"]]
        if not mismatch:
            return extra, r2
        # fix the earliest mismatching task
        mismatch.sort(key=lambda t: sched[t]["start"])
        j = mismatch[0]
        sj = sched[j]["start"]
        cands = [i for i in tasks if i != j and tasks[i]["dur"] > 0 and sched[i]["finish"] == sj - 1 and shares_resource(tasks, i, j)]
        if not cands:
            cands = [i for i in tasks if i != j and tasks[i]["dur"] > 0 and sched[i]["finish"] < sj and shares_resource(tasks, i, j)]
            cands.sort(key=lambda i: -sched[i]["finish"])
            cands = cands[:1]
        if not cands:
            raise RuntimeError(f"cannot explain start of {j}")
        for i in cands:
            lag = sj - sched[i]["finish"] - 1
            if lag == 0:
                extra.append((i, j, 0))
            else:
                # a lagged link: emulate by a start-no-earlier-than (kept as explicit lag in MS Project)
                extra.append((i, j, lag))
                snet = dict(snet); snet[j] = max(snet.get(j, 0), sj)
    raise RuntimeError("could not converge")
