"""Export the optimised integrated master schedule to MSPDI XML (opens directly in MS Project / ProjectLibre)."""
import json, sys, datetime as dt
import mpxj, jpype
jpype.startJVM()
from org.mpxj import (UnitOfMeasure, ProjectFile, Duration, TimeUnit, RelationType, TaskType, ResourceType, AccrueType, ConstraintType, Rate, Relation, Availability, CostRateTable, CostRateTableEntry)
from org.mpxj.common import LocalDateTimeHelper
from jpype import JArray, JObject
from org.mpxj.writer import UniversalProjectWriter, FileFormat
from java.time import LocalDateTime, LocalTime
from java.lang import Double

from engine import *
from links import derive_links
from optimize import unserial

def ldt(date, hour=8, minute=0):
    return LocalDateTime.of(date.year, date.month, date.day, hour, minute)

def build(solution_path, out_path):
    sol = json.load(open(solution_path))
    dec = unserial(sol["decisions"])
    r = evaluate(dec)
    extra, r2 = derive_links(dec, r)
    tasks, sched, preds, dl = r["tasks"], r["sched"], r["preds"], r["deliveries"]

    pf = ProjectFile()
    cal = pf.addDefaultBaseCalendar()
    pf.setDefaultCalendar(cal)
    props = pf.getProjectProperties()
    props.setName("Aurelia European AI Datacenter Conversion Program - Integrated Master Schedule"); props.setProjectTitle("Aurelia AI Datacenter Conversion - Program Master Plan")
    props.setStartDate(ldt(NTP))
    props.setCurrencySymbol("€")
    props.setDefaultDurationUnits(TimeUnit.DAYS)
    props.setDefaultTaskType(TaskType.FIXED_DURATION)
    # ---- resources
    rmap = {}
    for name, rs in RES.items():
        res = pf.addResource()
        res.setName(name)
        NA0, NA1 = LocalDateTimeHelper.START_DATE_NA, LocalDateTimeHelper.END_DATE_NA
        if rs["type"] == "Work":
            res.setType(ResourceType.WORK)
            res.getAvailability().add(Availability(NA0, NA1, Double(rs["max"] * 100.0)))
            rates = JArray(Rate)([Rate(float(rs["rate"]), TimeUnit.HOURS), Rate(0.0, TimeUnit.HOURS)])
        else:
            res.setType(ResourceType.MATERIAL)
            uom = UnitOfMeasure.Builder(pf).name(str(rs["label"])).abbreviation(str(rs["label"])).build()
            pf.getUnitsOfMeasure().add(uom); res.setUnitOfMeasure(uom)
            rates = JArray(Rate)([Rate(float(rs["rate"]), TimeUnit.HOURS), Rate(0.0, TimeUnit.HOURS)])
        crt = CostRateTable(); crt.add(CostRateTableEntry(NA0, NA1, Double(0.0), rates)); res.setCostRateTable(0, crt)
        res.setAccrueAt(AccrueType.PRORATED)
        res.setNotes(str(rs["note"] or ""))
        rmap[name] = res

    # ---- WBS
    root = pf.addTask(); root.setName("Aurelia European AI Datacenter Conversion Program"); root.setText(1, "PROGRAM")
    def summary(parent, name, text2=""):
        t = parent.addTask(); t.setName(name); t.setText(2, text2); return t
    tmap = {}
    ms_sum = summary(root, "Program milestones", "Program")
    p1 = summary(root, "P1 - Facility Conversion", "P1")
    p1_eng = summary(p1, "Engineering and design", "P1")
    p1_proc = summary(p1, "Facility equipment procurement and logistics", "P1")
    p1_conv = summary(p1, "Building conversion", "P1")
    p1_pt1 = summary(p1, "Power Train 1", "P1"); p1_ct1 = summary(p1, "Cooling Train 1", "P1"); p1_ha = summary(p1, "Data Hall A", "P1")
    p1_pt2 = summary(p1, "Power Train 2", "P1"); p1_ct2 = summary(p1, "Cooling Train 2", "P1"); p1_hb = summary(p1, "Data Hall B", "P1")
    p1_sh = summary(p1, "Shared systems and whole-plant integrated test", "P1")
    p2 = summary(root, "P2 - AI Platform", "P2")
    p2_arch = summary(p2, "Architecture and specification", "P2"); p2_proc = summary(p2, "AI platform procurement and logistics", "P2")
    p2_sw = summary(p2, "Platform software", "P2"); p2_ha = summary(p2, "Hall A deployment", "P2"); p2_hb = summary(p2, "Hall B deployment", "P2"); p2_int = summary(p2, "Cross-hall integration", "P2")
    p3 = summary(root, "P3 - Operations Readiness", "P3")
    p3_org = summary(p3, "Organization, procedures and shift arrangements", "P3"); p3_rec = summary(p3, "Recruitment", "P3")
    p3_h1 = summary(p3, "First hall readiness (Hall A)", "P3"); p3_h2 = summary(p3, "Second hall readiness (Hall B)", "P3"); p3_ho = summary(p3, "Operations handover", "P3")
    group = {}
    for a in ["F01", "F02", "F03", "F04", "F05", "F06", "F07", "F09", "F10", "F11", "F13", "F14", "F15", "F16"]: group[a] = p1_eng
    for a in ["F08-1", "F12-1"]: group[a] = p1_proc
    for a in ["F17", "F18", "F19", "F20", "F21"]: group[a] = p1_conv
    for a in ["F22", "F46", "F23", "F24", "F39"]: group[a] = p1_pt1
    for a in ["F25", "F26", "F40"]: group[a] = p1_ct1
    for a in ["F27", "F28", "F29", "F41"]: group[a] = p1_ha
    for a in ["F30", "F47", "F31", "F32", "F42"]: group[a] = p1_pt2
    for a in ["F33", "F34", "F43"]: group[a] = p1_ct2
    for a in ["F35", "F36", "F37", "F44"]: group[a] = p1_hb
    for a in ["F38", "F45"]: group[a] = p1_sh
    for a in ["A01", "A02", "A03", "A04"]: group[a] = p2_arch
    for a in ["A05-1", "A06-1", "A05-2", "A06-2"]: group[a] = p2_proc
    group["A07"] = p2_sw
    for a in ["A08", "A09"]: group[a] = p2_ha
    for a in ["A10", "A11"]: group[a] = p2_hb
    group["A12"] = p2_int
    for a in ["O01", "O02", "O03", "O08"]: group[a] = p3_org
    for a in ["O04", "O05", "O06", "O07", "O12"]: group[a] = p3_rec
    for a in ["O09", "O10", "O11"]: group[a] = p3_h1
    for a in ["O13", "O14"]: group[a] = p3_h2
    group["O15"] = p3_ho

    # NTP milestone
    ntp = ms_sum.addTask(); ntp.setName("Notice to Proceed (NTP)"); ntp.setMilestone(True); ntp.setDuration(Duration.getInstance(0, TimeUnit.DAYS))
    ntp.setStart(ldt(NTP)); ntp.setFinish(ldt(NTP)); ntp.setText(1, "NTP"); ntp.setText(2, "Program")
    ntp.setType(TaskType.FIXED_DURATION); ntp.setEffortDriven(False)
    order = sorted([t for t in tasks if t not in MILESTONES], key=lambda t: (sched[t]["start"], t))
    for tid in order:
        t = tasks[tid]; s = sched[tid]
        parent = group[tid]
        mt = parent.addTask()
        base = t["base"]
        label = f"{base} {t['name']}" if base else t["name"]
        mt.setName(label)
        mt.setText(1, tid); mt.setText(2, {"F": "P1", "A": "P2", "O": "P3"}[tid[0]]); mt.setText(3, t["loc"])
        mt.setDuration(Duration.getInstance(t["dur"], TimeUnit.DAYS))
        mt.setType(TaskType.FIXED_DURATION); mt.setEffortDriven(False)
        mt.setStart(ldt(wd_date(s["start"]))); mt.setFinish(ldt(wd_date(s["finish"]), 17))
        notes = []
        if base in dec["crash"] and dec["crash"][base] < ACTS[base]["dur"]:
            notes.append(f"CRASHED: normal {ACTS[base]['dur']} wd -> {t['dur']} wd (minimum {ACTS[base]['mindur']} wd). Crash-resource units = u x (1 + ({ACTS[base]['dur']}/{t['dur']} - 1)/{ACTS[base]['p']}) rounded up to the whole percent.")
        if t["note"]:
            notes.append(str(t["note"]))
        work_total = 0.0
        for rname, u in t["units"]:
            a = mt.addResourceAssignment(rmap[rname])
            a.setUnits(Double(round(u * 100, 2)))
            w = u * 8 * t["dur"]
            a.setWork(Duration.getInstance(w, TimeUnit.HOURS)); work_total += w
            a.setCost(Double(w * RES[rname]["rate"]))
        for rname, q in t["mat"]:
            a = mt.addResourceAssignment(rmap[rname])
            a.setUnits(Double(q))
            a.setCost(Double(q * RES[rname]["rate"]))
        mt.setWork(Duration.getInstance(work_total, TimeUnit.HOURS))
        fixed = 0.0
        # expedite premiums paid at order placement -> fixed cost on the ordering activity
        for info in dl.values():
            if info["po_event"] == tid and info["premium"] > 0:
                fixed += info["premium"]
                notes.append(f"Fixed cost includes expedite premium {info['pkg']} lot {info['lot']}: EUR {info['premium']:,.0f} (paid at order placement).")
        if fixed:
            mt.setFixedCost(Double(fixed)); mt.setFixedCostAccrual(AccrueType.END)
        mt.setCost(Double(t["cost"] + fixed))
        if tid in r["snet"] and r["snet"][tid] == s["start"] and tid.startswith("A0"):
            mt.setConstraintType(ConstraintType.START_NO_EARLIER_THAN); mt.setConstraintDate(ldt(wd_date(s["start"])))
            notes.append("Order date chosen by the team (just-in-time policy): the lot is paid with this activity, so it is placed as late as possible while delivering before the installing/testing activity needs it.")
        if notes:
            mt.setNotes("\n".join(notes))
        tmap[tid] = mt
    # milestones
    for mid, label in MILESTONES.items():
        mt = ms_sum.addTask(); mt.setName(label); mt.setMilestone(True); mt.setDuration(Duration.getInstance(0, TimeUnit.DAYS))
        d = wd_date(sched[mid]["start"] - 1) if sched[mid]["start"] > 0 else NTP
        mt.setStart(ldt(d, 17)); mt.setFinish(ldt(d, 17)); mt.setText(1, mid); mt.setText(2, "Program")
        mt.setType(TaskType.FIXED_DURATION); mt.setEffortDriven(False)
        if mid != "MS_FPA":
            mt.setNotes("Commercial service begins the calendar day after this milestone (last condition met).")
        tmap[mid] = mt
    # procurement chains
    proc_parent = {"F08": p1_proc, "F12": p1_proc, "A05": p2_proc, "A06": p2_proc}
    for (pkg, lot), info in sorted(dl.items(), key=lambda kv: (kv[1]["po_date"], kv[0])):
        parent = p2_proc if PK[pkg]["origin"] == "East Asia" else p1_proc
        po_tid = info["po_event"]
        pre = tmap[po_tid]
        tag = f"{pkg} - lot {lot}"
        mfg = parent.addTask(); mfg.setName(f"Manufacture {tag}{' (expedited)' if info['expedite'] else ''}")
        mfg.setDuration(Duration.getInstance(info["mfg_cd"], TimeUnit.ELAPSED_DAYS)); mfg.setType(TaskType.FIXED_DURATION); mfg.setEffortDriven(False)
        mstart = ldt(info["ready"] - dt.timedelta(days=info["mfg_cd"]), 17); mfg.setStart(mstart)
        # if the line was busy, manufacturing starts when the previous lot is ready -> lag
        lag_days = (info["ready"] - dt.timedelta(days=info["mfg_cd"]) - info["po_date"]).days
        mfg.setFinish(ldt(info["ready"], 17))
        rel = Relation.Builder().predecessorTask(pre).type(RelationType.FINISH_START).lag(Duration.getInstance(lag_days, TimeUnit.ELAPSED_DAYS))
        mfg.addPredecessor(rel)
        n = f"PO date = end of {po_tid}. Single production line: lot ready = max(PO date, previous lot ready) + {info['mfg_cd']} cd{' (expedited x' + str(PK[pkg]['exp_factor']) + ')' if info['expedite'] else ''}."
        if lag_days: n += f" Lag {lag_days} ed = waiting for the previous lot of this package on the supplier line."
        mfg.setNotes(n); mfg.setText(1, f"PO-{pkg[:12]}-{lot}"); mfg.setText(2, "Procurement")
        trn = parent.addTask(); trn.setName(f"Transport {tag} by {info['mode']}{' - disassembled' if info['disassembled'] else ''}")
        trn.setDuration(Duration.getInstance(info["transport_cd"], TimeUnit.ELAPSED_DAYS)); trn.setType(TaskType.FIXED_DURATION); trn.setEffortDriven(False)
        trn.setStart(ldt(info["ready"], 17)); trn.setFinish(ldt(info["delivery"], 17))
        trn.addPredecessor(Relation.Builder().predecessorTask(mfg).type(RelationType.FINISH_START))
        trn.setFixedCost(Double(info["transport_cost"])); trn.setFixedCostAccrual(AccrueType.END); trn.setCost(Double(info["transport_cost"]))
        trn.setNotes(f"{info['teu']:.0f} TEU x EUR {TR[(PK[pkg]['origin'], info['mode'])]['rate']:,.0f} = EUR {info['transport_cost']:,.0f}, paid on delivery."); trn.setText(2, "Procurement")
        dlv = parent.addTask(); dlv.setName(f"Delivered on site - {tag}"); dlv.setMilestone(True); dlv.setDuration(Duration.getInstance(0, TimeUnit.DAYS))
        dlv.setStart(ldt(info["delivery"], 17)); dlv.setFinish(ldt(info["delivery"], 17)); dlv.setType(TaskType.FIXED_DURATION); dlv.setEffortDriven(False)
        dlv.addPredecessor(Relation.Builder().predecessorTask(trn).type(RelationType.FINISH_START)); dlv.setText(2, "Procurement")
        acts = NEEDS[pkg][lot]
        if pkg == "Generators" and info["disassembled"]:
            acts = ["F46" if lot == 1 else "F47"]
        for a in acts:
            tmap[a].addPredecessor(Relation.Builder().predecessorTask(dlv).type(RelationType.FINISH_START))
    # logical links
    base_links = set()
    for p, s in r["links"]:
        if p in tmap and s in tmap:
            tmap[s].addPredecessor(Relation.Builder().predecessorTask(tmap[p]).type(RelationType.FINISH_START)); base_links.add((p, s))
    for t in ("F01", "A01", "O01"):
        tmap[t].addPredecessor(Relation.Builder().predecessorTask(ntp).type(RelationType.FINISH_START))
    # resource-sequencing links derived from the resource-feasible schedule
    for p, s, lag in extra:
        if (p, s) in base_links: continue
        rel = Relation.Builder().predecessorTask(tmap[p]).type(RelationType.FINISH_START)
        if lag:
            rel = rel.lag(Duration.getInstance(lag, TimeUnit.DAYS))
        tmap[s].addPredecessor(rel)
        old = tmap[s].getNotes() or ""
        tmap[s].setNotes((old + "\n" if old else "") + f"Resource-sequencing link from {p}{' (+' + str(lag) + ' wd)' if lag else ''}: shared crew/team at Max. Units; no automatic leveling used.")
    # summary dates
    for st in [root, ms_sum, p1, p1_eng, p1_proc, p1_conv, p1_pt1, p1_ct1, p1_ha, p1_pt2, p1_ct2, p1_hb, p1_sh, p2, p2_arch, p2_proc, p2_sw, p2_ha, p2_hb, p2_int, p3, p3_org, p3_rec, p3_h1, p3_h2, p3_ho]:
        st.setType(TaskType.FIXED_DURATION)
    UniversalProjectWriter(FileFormat.MSPDI).write(pf, out_path)
    return dict(extra=extra, r=r)

if __name__ == "__main__":
    res = build(sys.argv[1] if len(sys.argv) > 1 else "solution.json", sys.argv[2] if len(sys.argv) > 2 else "../deliverables/Aurelia_Program_Master_Plan.xml")
    print("extra resource links:", len(res["extra"]))
    for e in res["extra"]: print("  ", e)
