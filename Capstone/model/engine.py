"""Integrated schedule / procurement / cash-flow engine for the Aurelia capstone case.

Calendar: Mon-Fri, 8 h/day, no holidays. Working-day index 0 = NTP (Mon 11 Jan 2027).
All money in EUR inside the engine (outputs converted to kEUR where stated).
"""
import math, datetime as dt, copy
from collections import defaultdict
from data import load

DATA = load()
ACTS, RES, PK, TR, FIN = DATA["acts"], DATA["res"], DATA["pk"], DATA["transport"], DATA["fin"]
NTP = FIN["ntp"]
HORIZON = 900  # working days ~ 3.5 years

# ---------------------------------------------------------------- calendar
_WD = []
_d = NTP
while len(_WD) < HORIZON:
    if _d.weekday() < 5:
        _WD.append(_d)
    _d += dt.timedelta(days=1)
_WD_INDEX = {d: i for i, d in enumerate(_WD)}

def wd_date(i):
    return _WD[i]

def next_wd_after(date):
    """index of first working day strictly after calendar date"""
    d = date + dt.timedelta(days=1)
    while d.weekday() >= 5:
        d += dt.timedelta(days=1)
    return _WD_INDEX[d]

def last_wd_on_or_before(date):
    d = date
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return _WD_INDEX[d]

# ---------------------------------------------------------------- network definition (team assumptions, see Charter annex)
# (pred, succ, type, lag_wd)  -- FS unless stated
LINKS = [
    # --- sequencing policy: Train 1 / Hall A engineering packages are released before the Train 2 / Hall B packages
    ("F09", "F10"), ("F13", "F14"), ("F15", "F16"),
    # --- P1 engineering
    ("F01", "F02"), ("F01", "F03"), ("F02", "F04"), ("F04", "F05"), ("F05", "F06"),
    ("F04", "F07"), ("F07", "F08-1"), ("F08-1", "F09"), ("F08-1", "F10"),
    ("F04", "F11"), ("F11", "F12-1"), ("F12-1", "F13"), ("F12-1", "F14"),
    ("F05", "F15"), ("A03", "F15"), ("F05", "F16"), ("A03", "F16"),
    # --- P1 building conversion
    ("F05", "F17"), ("F17", "F18"), ("F18", "F19"), ("F03", "F19"), ("F06", "F19"),
    ("F19", "F20"), ("F20", "F21"), ("F02", "F21"),
    # --- Train 1 / Hall A
    ("F21", "F22"), ("F09", "F22"), ("F22", "F23"), ("F22", "F24"),
    ("F21", "F25"), ("F13", "F25"), ("F25", "F26"),
    ("F21", "F27"), ("F15", "F27"), ("F27", "F28"), ("F28", "F29"),
    # --- Train 2 / Hall B
    ("F21", "F30"), ("F10", "F30"), ("F30", "F31"), ("F30", "F32"),
    ("F21", "F33"), ("F14", "F33"), ("F33", "F34"),
    ("F21", "F35"), ("F16", "F35"), ("F35", "F36"), ("F36", "F37"),
    # --- shared systems and tests
    ("F22", "F38"), ("F27", "F38"),
    ("F23", "F39"), ("F24", "F39"), ("F38", "F39"), ("F26", "F40"), ("F38", "F40"),
    ("F29", "F41"), ("F39", "F41"), ("F40", "F41"),
    ("F31", "F42"), ("F32", "F42"), ("F38", "F42"), ("F34", "F43"), ("F38", "F43"),
    ("F37", "F44"), ("F42", "F44"), ("F43", "F44"),
    ("F41", "F45"), ("F44", "F45"), ("A12", "F45"),
    # --- P2
    ("A01", "A02"), ("A02", "A03"), ("A03", "A04"), ("A04", "A05-1"), ("A04", "A06-1"),
    ("A04", "A05-2"), ("A04", "A06-2"), ("A03", "A07"),
    ("F29", "A08"), ("F39", "A08"), ("F40", "A08"),
    ("A08", "A09"), ("F41", "A09"), ("A07", "A09"),
    ("F37", "A10"), ("F42", "A10"), ("F43", "A10"),
    ("A10", "A11"), ("F44", "A11"), ("A07", "A11"),
    ("A09", "A12"), ("A11", "A12"),
    # --- P3
    ("O01", "O02"), ("O02", "O03"), ("O02", "O04"), ("O04", "O05"), ("O04", "O06"), ("O04", "O07"),
    ("O03", "O08"), ("O05", "O08"), ("O06", "O08"), ("O07", "O08"),
    ("O05", "O09"), ("F24", "O09"), ("F26", "O09"),
    ("O06", "O10"), ("A08", "O10"),
    ("O08", "O11"), ("O09", "O11"), ("O10", "O11"), ("F41", "O11"), ("A09", "O11"),
    ("O05", "O12"), ("O06", "O12"),
    ("O12", "O13"), ("F32", "O13"), ("F34", "O13"), ("A10", "O13"),
    ("O13", "O14"), ("F44", "O14"), ("A11", "O14"),
    ("F45", "O15"), ("O14", "O15"), ("O11", "O15"),
    # --- milestones
    ("F41", "MS_HALL_A"), ("A09", "MS_HALL_A"), ("O11", "MS_HALL_A"), ("O08", "MS_HALL_A"),
    ("F44", "MS_HALL_B"), ("A11", "MS_HALL_B"), ("O14", "MS_HALL_B"),
    ("F45", "MS_FPA"), ("O15", "MS_FPA"), ("MS_HALL_A", "MS_FPA"), ("MS_HALL_B", "MS_FPA"),
]
# links added only when generators are shipped disassembled
LINKS_DIS = {1: [("F22", "F46"), ("F46", "F24")], 2: [("F30", "F47"), ("F47", "F32")]}

# delivery requirements: package -> lot -> list of installing activities (delivery must precede start)
NEEDS = {
    "Power transformers": {1: ["F23"], 2: ["F31"]},
    "Generators": {1: ["F24"], 2: ["F32"]},            # replaced by F46/F47 when disassembled
    "Switchgear + UPS/batteries": {1: ["F23", "F24"], 2: ["F31", "F32"]},
    "Cooling train equipment": {1: ["F26"], 2: ["F34"]},
    "Busway / PDU / racks": {1: ["F28"], 2: ["F36"]},
    "Shared controls + fibre": {1: ["F38"]},
    "AI compute capacity": {1: ["A08"], 2: ["A10"]},
    "High-speed network + storage": {1: ["A09"], 2: ["A11"]},
}
# pseudo-resource: single heavy-equipment corridor (team assumption, max 2 concurrent heavy-equipment installation fronts)
CORRIDOR_ACTS = ["F23", "F26", "F31", "F34", "F46", "F47"]
CORRIDOR_MAX = 2

MILESTONES = {"MS_HALL_A": "Hall A commercial service", "MS_HALL_B": "Hall B commercial service", "MS_FPA": "Final Program Acceptance"}

# ---------------------------------------------------------------- decisions
def default_decisions():
    """Baseline: normal durations, no expediting, cheapest transport, separate POs per lot in one ordering event."""
    d = dict(crash={}, pos={}, gen_disassembled={1: False, 2: False}, jit=True, po_offset={})
    for name, p in PK.items():
        region = p["origin"]
        cheapest = min((m for (r, m) in TR if r == region and (TR[(r, m)]["oog_teu"] is not None or p["oog_pieces"] == 0)),
                       key=lambda m: TR[(region, m)]["rate"])
        d["pos"][name] = [dict(lot=k, expedite=False, mode=cheapest) for k in range(1, p["lots"] + 1)]
    return d

def crash_units(act, d):
    """resource units of a catalogue activity at chosen duration d (workbook note 4)."""
    N = act["dur"]
    out = []
    for r, u in act["work"]:
        if r in act["crash"] and d < N:
            u2 = math.ceil(u * (1 + (N / d - 1) / act["p"]) * 100 - 1e-9) / 100
            out.append((r, u2))
        else:
            out.append((r, u))
    return out

# ---------------------------------------------------------------- task table
def build_tasks(dec):
    """Instantiate the task list (catalogue activities + PO occurrences + milestones)."""
    tasks = {}
    def add(tid, base, **kw):
        a = copy.deepcopy(ACTS[base]) if base else dict(id=tid, name=MILESTONES.get(tid, tid), dur=0, work=[], mat=[], crash=[], p=None, area="Program", loc="Program", mindur=None, note="")
        a["id"] = tid; a["base"] = base
        a.update(kw)
        tasks[tid] = a
    for aid in ACTS:
        if aid in ("F08", "F12", "A05", "A06", "F46", "F47"):
            continue
        add(aid, aid)
    add("F08-1", "F08", name=ACTS["F08"]["name"] + " - ordering event 1 (all power packages, lots 1 and 2 as separate POs)")
    add("F12-1", "F12", name=ACTS["F12"]["name"] + " - ordering event 1 (cooling lots 1 and 2 as separate POs)")
    for k in (1, 2):
        add(f"A05-{k}", "A05", name=ACTS["A05"]["name"] + f" - PO {k} (Hall {'AB'[k-1]} lot)", mat=[("Network + storage lot", 1)])
        add(f"A06-{k}", "A06", name=ACTS["A06"]["name"] + f" - PO {k} (Hall {'AB'[k-1]} lot)", mat=[("AI compute lot", 1)])
        if dec["gen_disassembled"][k]:
            add("F46" if k == 1 else "F47", "F46" if k == 1 else "F47")
    for m in MILESTONES:
        add(m, None)
    # apply crashing
    for tid, t in tasks.items():
        base = t["base"]
        if base in dec["crash"] and dec["crash"][base] < t["dur"]:
            d = dec["crash"][base]
            assert t["mindur"] and d >= t["mindur"], (tid, d)
            t["units"] = crash_units(ACTS[base], d)
            t["dur"] = d
        else:
            t["units"] = list(t["work"])
        # daily cost
        work_cost = sum(u * 8 * RES[r]["rate"] for r, u in t["units"]) * t["dur"]
        mat_cost = sum(q * RES[r]["rate"] for r, q in t["mat"])
        t["cost"] = work_cost + mat_cost
        t["work_cost"] = work_cost; t["mat_cost"] = mat_cost
        t["cost_per_day"] = t["cost"] / t["dur"] if t["dur"] else 0.0
    return tasks

def links_for(dec):
    L = [(p, s) for p, s in LINKS]
    for k in (1, 2):
        if dec["gen_disassembled"][k]:
            L += LINKS_DIS[k]
    return L

# ---------------------------------------------------------------- procurement
def po_event_for(pkg, lot):
    p = PK[pkg]
    oa = p["order_act"]
    if oa in ("F08", "F12"):
        return oa + "-1"
    if oa in ("A05", "A06"):
        return f"{oa}-{lot}"
    return None  # busway (F15/F16 end) and controls (F05 end)

def shipment_charge(pkg, lots, mode, disassembled):
    p = PK[pkg]; tr = TR[(p["origin"], mode)]
    vol = p["vol_lot"] * lots; wt = p["wt_lot"] * lots; oog = p["oog_pieces"] * lots
    if pkg == "Generators" and disassembled:
        vol = DATA["gen_dis"]["vol"] * lots; wt = DATA["gen_dis"]["wt"] * lots; oog = 0
    if tr["air"]:
        if oog:
            return None
        teu = math.ceil(max(vol / 33, wt / 5.5) - 1e-9)
    else:
        if oog and tr["oog_teu"] is None:
            return None
        teu = math.ceil(vol / 33 - 1e-9) + (oog * tr["oog_teu"] if oog else 0)
    return teu * tr["rate"], teu, tr["cd"]

def compute_deliveries(dec, tasks, sched, only=None):
    """Given PO activity finish dates, apply the single-production-line rule. Returns per package/lot info."""
    out = {}
    for pkg, p in PK.items():
        if only and pkg not in only:
            continue
        pos = dec["pos"][pkg]
        # PO date per PO (end of ordering activity, or chosen date)
        po_dates = []
        for po in pos:
            ev = po_event_for(pkg, po["lot"])
            if ev:
                if ev not in sched:
                    continue  # later PO not yet scheduled (lots are served in lot order)
                fin = sched[ev]["finish"]
                date = wd_date(fin)
            else:
                # team-chosen order date: end of the specifying activity (+ optional offset in wd)
                spec = {"Busway / PDU / racks": {1: "F15", 2: "F16"}, "Shared controls + fibre": {1: "F05"}}[pkg][po["lot"]]
                if spec not in sched:
                    continue
                off = dec["po_offset"].get((pkg, po["lot"]), 0)
                date = wd_date(sched[spec]["finish"] + off)
                ev = spec
            po_dates.append((date, po, ev))
        po_dates.sort(key=lambda x: (x[1]["lot"]))  # line serves lots in lot order (PO 1 placed first by construction)
        line_free = None
        for date, po, ev in po_dates:
            mfg = p["mfg_wk"] * 7 * (p["exp_factor"] if po["expedite"] else 1.0)
            mfg_cd = int(round(mfg))
            start = date if line_free is None else max(date, line_free)
            ready = start + dt.timedelta(days=mfg_cd)
            line_free = ready
            dis = dec["gen_disassembled"][po["lot"]] if pkg == "Generators" else False
            sc = shipment_charge(pkg, 1, po["mode"], dis)
            assert sc is not None, f"{pkg} cannot fly assembled"
            charge, teu, tcd = sc
            delivery = ready + dt.timedelta(days=tcd)
            premium = p["lot_value"] * 1000 * p["exp_premium"] if po["expedite"] else 0.0
            out[(pkg, po["lot"])] = dict(pkg=pkg, lot=po["lot"], po_event=ev, po_date=date, mfg_cd=mfg_cd, ready=ready,
                                         mode=po["mode"], transport_cd=tcd, delivery=delivery, teu=teu, transport_cost=charge,
                                         expedite=po["expedite"], premium=premium, disassembled=dis,
                                         value=p["lot_value"] * 1000)
    return out

# ---------------------------------------------------------------- scheduling (serial SGS, resource constrained)
def schedule(dec, snet=None, extra_links=None, resource_check=True):
    """Returns sched dict tid -> {start, finish} (working-day indices; finish = last working day, milestones finish=start-1 style)."""
    tasks = build_tasks(dec)
    links = links_for(dec) + (extra_links or [])
    preds = defaultdict(list)
    for p, s in links:
        preds[s].append(p)
    # delivery requirements expressed as (pkg,lot)->installing activities
    need_map = defaultdict(list)  # tid -> [(pkg,lot)]
    for pkg, lots in NEEDS.items():
        for lot, acts in lots.items():
            for a in acts:
                if pkg == "Generators" and dec["gen_disassembled"][lot]:
                    a = "F46" if lot == 1 else "F47"
                if a in tasks:
                    need_map[a].append((pkg, lot))
    # po events that feed each delivery
    def po_event_tid(pkg, lot):
        ev = po_event_for(pkg, lot)
        if ev: return ev
        return {"Busway / PDU / racks": {1: "F15", 2: "F16"}, "Shared controls + fibre": {1: "F05"}}[pkg][lot]
    for a, lst in need_map.items():
        for pkg, lot in lst:
            # the other PO of the same package may also constrain (line rule) -> add all POs of the package as preds
            for l2 in range(1, lot + 1):  # lots are manufactured in lot order on the single line
                ev = po_event_tid(pkg, l2)
                if ev not in preds[a]:
                    preds[a].append(ev)
    # priority: latest start from an unconstrained backward pass (approximate lead times as wd)
    order = topo(tasks, preds)
    # approximate supplier lead times (calendar days -> working days) for the priority pass only
    lead_wd = {}
    for pkg, p in PK.items():
        cum = 0
        for po in sorted(dec["pos"][pkg], key=lambda x: x["lot"]):
            mfg = int(round(p["mfg_wk"] * 7 * (p["exp_factor"] if po["expedite"] else 1.0)))
            cum += mfg
            lead_wd[(pkg, po["lot"])] = int(math.ceil((cum + TR[(p["origin"], po["mode"])]["cd"]) * 5 / 7))
    es = {}
    for tid in order:
        e = 0
        for p in preds[tid]:
            lag = 0
            for pkg, lot in need_map.get(tid, []):
                if p == po_event_tid(pkg, lot):
                    lag = max(lag, lead_wd[(pkg, lot)])
            e = max(e, es[p] + tasks[p]["dur"] + lag)
        if snet and tid in snet:
            e = max(e, snet[tid])
        es[tid] = e
    succs = defaultdict(list)
    for s, ps in preds.items():
        for p in ps:
            succs[p].append(s)
    T = max(es[t] + tasks[t]["dur"] for t in tasks)
    ls = {}
    for tid in reversed(order):
        l = T - tasks[tid]["dur"]
        if tid in MILESTONES:
            l = es[tid]  # each commercial-service / acceptance milestone is treated as its own deadline (ASAP)
        for s in succs[tid]:
            l = min(l, ls[s] - tasks[tid]["dur"])
        ls[tid] = l
    # serial SGS
    usage = defaultdict(lambda: [0.0] * HORIZON)
    sched = {}
    deliveries = {}
    remaining = set(tasks)
    while remaining:
        ready = [t for t in remaining if all(p in sched for p in preds[t])]
        ready.sort(key=lambda t: (ls[t], t))
        tid = ready[0]
        t = tasks[tid]
        est = 0
        for p in preds[tid]:
            est = max(est, sched[p]["finish"] + 1)
        if snet and tid in snet:
            est = max(est, snet[tid])
        # deliveries
        if tid in need_map:
            dl = compute_deliveries(dec, tasks, sched, only={pkg for pkg, lot in need_map[tid]})
            for pkg, lot in need_map[tid]:
                info = dl[(pkg, lot)]
                est = max(est, next_wd_after(info["delivery"]))
        dur = t["dur"]
        if dur == 0:
            sched[tid] = dict(start=est, finish=est - 1)
            remaining.discard(tid); continue
        reqs = list(t["units"])
        if tid in CORRIDOR_ACTS:
            reqs.append(("__corridor__", 1.0))
        s = est
        if resource_check:
            while True:
                ok = True
                for r, u in reqs:
                    mx = CORRIDOR_MAX if r == "__corridor__" else RES[r]["max"]
                    arr = usage[r]
                    for k in range(s, s + dur):
                        if arr[k] + u > mx + 1e-9:
                            ok = False; s = k + 1; break
                    if not ok: break
                if ok: break
        for r, u in reqs:
            arr = usage[r]
            for k in range(s, s + dur):
                arr[k] += u
        sched[tid] = dict(start=s, finish=s + dur - 1)
        remaining.discard(tid)
    deliveries = compute_deliveries(dec, tasks, sched)
    return dict(tasks=tasks, sched=sched, preds=preds, deliveries=deliveries, usage=usage, es=es, ls=ls, links=links)

def topo(tasks, preds):
    seen, order = set(), []
    def visit(t, stack=()):
        if t in seen: return
        assert t not in stack, ("cycle", t, stack)
        for p in preds[t]:
            visit(p, stack + (t,))
        seen.add(t); order.append(t)
    for t in tasks: visit(t)
    return order

def jit_schedule(dec, extra_links=None, resource_check=True):
    """Iterate SNET dates for compute / network POs so that lots are delivered just in time."""
    snet = {}
    r = schedule(dec, snet, extra_links, resource_check)
    if not dec.get("jit", True):
        return r, snet
    for it in range(6):
        changed = False
        for pkg, oa in (("AI compute capacity", "A06"), ("High-speed network + storage", "A05")):
            for lot in (1, 2):
                tid = f"{oa}-{lot}"
                inst = NEEDS[pkg][lot][0]
                # facility-driven earliest start of the installing activity (ignoring this delivery)
                others = [r["sched"][p]["finish"] + 1 for p in r["preds"][inst] if p not in (f"{oa}-1", f"{oa}-2")]
                need_start = max(others) if others else 0
                info = r["deliveries"][(pkg, lot)]
                lead = info["mfg_cd"] + info["transport_cd"]
                po_end_date = wd_date(need_start) - dt.timedelta(days=1 + lead)
                po_end = last_wd_on_or_before(po_end_date)
                dur = r["tasks"][tid]["dur"]
                target = po_end - dur + 1
                lo = max(r["sched"][p]["finish"] + 1 for p in r["preds"][tid])
                target = max(target, lo)
                if snet.get(tid) != target:
                    snet[tid] = target; changed = True
        if not changed:
            break
        r = schedule(dec, snet, extra_links, resource_check)
    return r, snet

# ---------------------------------------------------------------- cash flows
def month_index(date):
    return (date.year - NTP.year) * 12 + (date.month - 1)

NMONTHS = month_index(FIN["appraisal_end"]) + 1  # Jan 2027 .. Dec 2037 = 132

def month_start(m):
    return dt.date(NTP.year + m // 12, m % 12 + 1, 1)

def month_end(m):
    nm = m + 1
    return dt.date(NTP.year + nm // 12, nm % 12 + 1, 1) - dt.timedelta(days=1)

def cashflows(r):
    tasks, sched, dl = r["tasks"], r["sched"], r["deliveries"]
    capex = [0.0] * NMONTHS
    capex_act = [0.0] * NMONTHS; capex_lots = [0.0] * NMONTHS; capex_exp = [0.0] * NMONTHS; capex_tr = [0.0] * NMONTHS; capex_oh = [0.0] * NMONTHS
    for tid, t in tasks.items():
        if t["dur"] == 0: continue
        s = sched[tid]
        for k in range(s["start"], s["finish"] + 1):
            m = month_index(wd_date(k))
            if t["base"] in ("A05", "A06"):
                capex_lots[m] += t["mat_cost"] / t["dur"]
                capex_act[m] += t["work_cost"] / t["dur"]
            else:
                capex_act[m] += t["cost_per_day"]
    for info in dl.values():
        capex_exp[month_index(info["po_date"])] += info["premium"]
        capex_tr[month_index(info["delivery"])] += info["transport_cost"]
    fpa = wd_date(sched["MS_FPA"]["finish"])  # last working day of final acceptance activity
    hall_a = wd_date(sched["MS_HALL_A"]["finish"]) + dt.timedelta(days=1)  # commercial service date = day after last condition
    hall_b = wd_date(sched["MS_HALL_B"]["finish"]) + dt.timedelta(days=1)
    f17 = wd_date(sched["F17"]["start"])
    def days_overlap(a, b, m):
        lo, hi = max(a, month_start(m)), min(b, month_end(m))
        return max(0, (hi - lo).days + 1)
    for m in range(NMONTHS):
        capex_oh[m] += (FIN["oh_pmo"] + FIN["oh_ops"]) * 1000 * 12 / 365 * days_overlap(NTP, fpa, m)
        capex_oh[m] += FIN["oh_site"] * 1000 * 12 / 365 * days_overlap(f17, fpa, m)
    rev = [0.0] * NMONTHS; elec = [0.0] * NMONTHS
    rev_year = FIN["gpu_per_hall"] * FIN["hours"] * FIN["rate_gpuh"]
    elec_year = FIN["hall_mw"] * 1000 * FIN["pue"] * FIN["hours"] * FIN["elec_price"]
    for m in range(NMONTHS):
        for cod in (hall_a, hall_b):
            n = days_overlap(cod, FIN["appraisal_end"], m)
            rev[m] += rev_year / 365 * n
            elec[m] += elec_year / 365 * n
    for m in range(NMONTHS):
        capex[m] = capex_act[m] + capex_lots[m] + capex_exp[m] + capex_tr[m] + capex_oh[m]
    fcff = [rev[m] - elec[m] - capex[m] for m in range(NMONTHS)]
    return dict(capex=capex, capex_act=capex_act, capex_lots=capex_lots, capex_exp=capex_exp, capex_tr=capex_tr, capex_oh=capex_oh,
                rev=rev, elec=elec, fcff=fcff, fpa=fpa, hall_a=hall_a, hall_b=hall_b)

GRACE_END_M = month_index(FIN["grace_end"])          # Mar 2028
FIRST_REPAY_M = GRACE_END_M + 1                       # Apr 2028
N_REPAY = NMONTHS - FIRST_REPAY_M                     # 117

def wacc(d):
    return d * FIN["rd"] + (1 - d) * FIN["re"]

def npv(fcff, w):
    return sum(f / (1 + w) ** ((m + 0.5) / 12) for m, f in enumerate(fcff))

def ledger(cf, d):
    """Monthly debt ledger for debt share d of CAPEX. Returns per-month dicts and annual DSCR."""
    D = d * sum(cf["capex"])
    P = D / N_REPAY
    bal = 0.0
    rows = []
    for m in range(NMONTHS):
        draw = d * cf["capex"][m]
        bal_after_draw = bal + draw
        interest = bal_after_draw * FIN["rd"] / 12
        princ = min(P, bal_after_draw) if m >= FIRST_REPAY_M else 0.0
        bal_end = bal_after_draw - princ
        rows.append(dict(m=m, draw=draw, interest=interest, principal=princ, bal_end=bal_end,
                         cfads=cf["rev"][m] - cf["elec"][m]))
        bal = bal_end
    years = {}
    for row in rows:
        if row["m"] < FIRST_REPAY_M: continue
        y = NTP.year + row["m"] // 12
        a = years.setdefault(y, dict(cfads=0.0, ds=0.0, interest=0.0, principal=0.0))
        a["cfads"] += row["cfads"]; a["interest"] += row["interest"]; a["principal"] += row["principal"]; a["ds"] += row["interest"] + row["principal"]
    for y, a in years.items():
        a["dscr"] = a["cfads"] / a["ds"] if a["ds"] > 0 else float("inf")
    return dict(D=D, P=P, rows=rows, years=years, final_balance=rows[-1]["bal_end"])

def max_debt_share(cf, dscr_min=None, step=0.005):
    """Largest d with annual DSCR >= covenant in every test year and full repayment by maturity."""
    dscr_min = dscr_min or FIN["dscr_min"]
    # DSCR is linear in 1/d, so compute the unit ledger at d=1 and scale
    L = ledger(cf, 1.0)
    dmax = 1.0
    for y, a in L["years"].items():
        if a["ds"] > 0:
            dmax = min(dmax, a["cfads"] / (dscr_min * a["ds"]))
    dmax = max(0.0, math.floor(dmax / step) * step)
    # verify explicitly (cap effects)
    while dmax > 0:
        Lx = ledger(cf, dmax)
        if all(a["dscr"] >= dscr_min - 1e-9 for a in Lx["years"].values()) and abs(Lx["final_balance"]) < 1.0:
            break
        dmax -= step
    return dmax

def evaluate(dec, extra_links=None, resource_check=True, dscr_min=None):
    r, snet = jit_schedule(dec, extra_links, resource_check)
    cf = cashflows(r)
    d = max_debt_share(cf, dscr_min)
    w = wacc(d)
    v = npv(cf["fcff"], w)
    r.update(cf=cf, d=d, wacc=w, npv=v, snet=snet, dec=dec)
    r["feasible"] = cf["fpa"] <= FIN["deadline"]
    return r

def summary(r):
    cf = r["cf"]
    return dict(npv_k=r["npv"] / 1000, d=r["d"], wacc=r["wacc"], hall_a=cf["hall_a"], hall_b=cf["hall_b"], fpa=cf["fpa"],
                capex_k=sum(cf["capex"]) / 1000, feasible=r["feasible"])

if __name__ == "__main__":
    dec = default_decisions()
    r = evaluate(dec)
    print(summary(r))
    for tid in sorted(r["sched"], key=lambda t: r["sched"][t]["start"]):
        s = r["sched"][tid]
        print(tid, wd_date(s["start"]), wd_date(max(s["finish"], s["start"])), r["tasks"][tid]["dur"])
    for k, v in r["deliveries"].items():
        print(k, v["po_date"], v["ready"], v["delivery"], v["mode"], v["teu"], v["transport_cost"])
    for y, a in ledger(cf := r["cf"], r["d"])["years"].items():
        print(y, round(a["dscr"], 2), round(a["cfads"] / 1e6, 1), round(a["ds"] / 1e6, 1))
