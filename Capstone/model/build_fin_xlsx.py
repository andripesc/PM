"""Build the Financial Model workbook (live formulas) from the optimised solution."""
import json, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment
from engine import *
from links import derive_links
from optimize import unserial

BLUE = Font(name="Arial", size=10, color="0000FF")
BLACK = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
GREEN = Font(name="Arial", size=10, color="008000")
H1 = Font(name="Arial", size=14, bold=True)
H2 = Font(name="Arial", size=11, bold=True)
YELLOW = PatternFill("solid", fgColor="FFFF00")
GREY = PatternFill("solid", fgColor="D9D9D9")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
KEUR = '#,##0;(#,##0);-'
KEUR1 = '#,##0.0;(#,##0.0);-'
PCT = '0.0%'
DATEF = 'dd-mmm-yyyy'

def style_ws(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font == Font():
                c.font = BLACK

def build(solution_path, out_path):
    sol = json.load(open(solution_path))
    dec = unserial(sol["decisions"])
    r = evaluate(dec)
    extra, _ = derive_links(dec, r)
    tasks, sched, dl, cf = r["tasks"], r["sched"], r["deliveries"], r["cf"]
    wb = Workbook()

    # ------------------------------------------------------------------ README
    ws = wb.active; ws.title = "README"
    rows = [
        ("Aurelia AI Infrastructure SpA - European AI Datacenter Conversion Program", H1),
        ("Financial Model: monthly unlevered FCFF, debt/equity optimisation, WACC, amortisation ledger, annual DSCR and NPV", H2),
        ("", BLACK),
        ("How the workbook is organised", H2),
        ("Inputs: all case inputs (blue) taken unchanged from the instructor workbook 'Datacenter_Data_Updated.xlsx'; team levers in yellow.", BLACK),
        ("Schedule: every activity of the Integrated Master Schedule (same dates as the MS Project file) with its resource cost; cost accrues prorated over its working days.", BLACK),
        ("Procurement: purchase orders, supplier lead-time rule (single production line), expediting, transport mode, delivery vs required-on-site dates and the related costs.", BLACK),
        ("ActivityCost_M: activity cost spread over calendar months with NETWORKDAYS (prorated accrual, Mon-Fri calendar).", BLACK),
        ("Monthly: CAPEX build-up, overheads, revenue, electricity, unlevered FCFF, cumulative cash flow, discount factors and present values.", BLACK),
        ("Financing: D/E optimisation table (every mix evaluated on the same FCFF at its own WACC), DSCR feasibility test, selection of the highest feasible NPV and covenant sensitivities (1.50x / 1.80x).", BLACK),
        ("Debt_Ledger: monthly borrowing and repayment ledger for the selected mix (draws pro rata to CAPEX, interest on drawn balance, equal monthly principal Apr 2028 - Dec 2037) and the annual post-grace DSCR test.", BLACK),
        ("Summary: required outputs (NPV, WACC, D/E, service dates, Final Program Acceptance, CAPEX, acceleration and procurement decisions, compliance checks).", BLACK),
        ("Resources: peak concurrent units per resource versus Max. Units (resource-feasibility verification, values exported from the schedule engine; cross-check in MS Project > Resource Usage).", BLACK),
        ("Sensitivity: NPV sensitivity to selling rate, electricity price, equity cost and a commercial-service delay.", BLACK),
        ("", BLACK),
        ("Colour legend", H2),
        ("Blue font = hard-coded input from the case workbook; black = formula; green = link to another sheet; yellow fill = team decision lever / key assumption.", BLACK),
        ("", BLACK),
        ("Modelling conventions (team assumptions, documented in the Program Charter)", H2),
        ("1. Calendar: Mon-Fri, 8 h/day, no holidays; working-day costs accrue prorated (cost per working day x NETWORKDAYS in month).", BLACK),
        ("2. Overheads prorated by calendar day (rate x 12 / 365) from NTP (PMO, ops-readiness management) or from the start of F17 (site services) up to and including the Final Program Acceptance date.", BLACK),
        ("3. Revenue and electricity per hall prorated by calendar day (annual amount / 365) from the commercial-service date (day after the last condition is met) to 31 Dec 2037; no utilisation, availability, inflation, tax or terminal value.", BLACK),
        ("4. Discounting: monthly mid-period, DF = (1 + WACC)^(-(m + 0.5)/12), m = 0 for Jan 2027 (switch on Inputs allows the nominal-monthly alternative (1 + WACC/12)^-(m+0.5)).", BLACK),
        ("5. WACC = d x rD + (1 - d) x rE with d = Debt / total program CAPEX (CAPEX = activity costs incl. equipment lots, expediting, transport and overheads up to Final Program Acceptance). No tax shield.", BLACK),
        ("6. Debt draws each month = d x CAPEX of the month (D + E = CAPEX, pro rata); interest on the drawn balance including the month's draw; interest during grace is funded by the sponsor and is not in FCFF.", BLACK),
        ("7. Principal: equal monthly instalments = total committed loan / 117 from Apr 2028 to Dec 2037, capped by the drawn balance; full repayment by maturity is verified.", BLACK),
        ("8. DSCR (annual, 2028 = Apr-Dec, then 2029-2037) = CFADS / (interest + principal), with CFADS = revenue - electricity (operating cash flow; CAPEX is financed by the D/E draws). Covenant: >= 1.65x in every test year.", BLACK),
        ("9. The highest feasible NPV is obtained with the largest debt share that satisfies the DSCR covenant in every year, because FCFF is identical for all mixes and rD < rE.", BLACK),
    ]
    for i, (t, f) in enumerate(rows, 1):
        ws.cell(i, 1, t).font = f
    ws.column_dimensions["A"].width = 160

    # ------------------------------------------------------------------ Inputs
    wi = wb.create_sheet("Inputs")
    wi["A1"] = "INPUTS (case data unchanged; yellow = team lever)"; wi["A1"].font = H1
    inputs = [
        ("Program NTP", NTP, DATEF, "Revenue and Finance!B5"),
        ("Financial appraisal end", FIN["appraisal_end"], DATEF, "Revenue and Finance!B6"),
        ("Final Program Acceptance deadline (Month 32)", FIN["deadline"], DATEF, "EDATE(NTP, 32) - Revenue and Finance!B7"),
        ("GPU-eq per hall", FIN["gpu_per_hall"], '#,##0', "Revenue and Finance!B9"),
        ("Hours per year", FIN["hours"], '#,##0', "Revenue and Finance!B11"),
        ("Realised selling rate [EUR/GPU-h]", FIN["rate_gpuh"], '0.00', "Revenue and Finance!B10"),
        ("IT design load per hall [MW]", FIN["hall_mw"], '0.0', "Revenue and Finance!B8"),
        ("PUE", FIN["pue"], '0.00', "Revenue and Finance!B14"),
        ("Electricity price [EUR/kWh]", FIN["elec_price"], '0.000', "Revenue and Finance!B13"),
        ("Overhead - PMO [kEUR/month]", FIN["oh_pmo"], KEUR, "Revenue and Finance!B20"),
        ("Overhead - General site services [kEUR/month]", FIN["oh_site"], KEUR, "Revenue and Finance!B21"),
        ("Overhead - Operations-readiness management [kEUR/month]", FIN["oh_ops"], KEUR, "Revenue and Finance!B22"),
        ("Debt cost rD", FIN["rd"], PCT, "Financing Inputs!B5 = 3.3% + 2.7%"),
        ("Equity cost rE", FIN["re"], PCT, "Financing Inputs!B6"),
        ("Minimum DSCR (covenant)", FIN["dscr_min"], '0.00"x"', "Financing Inputs!B7"),
        ("DSCR sensitivity - lower", FIN["dscr_lo"], '0.00"x"', "Financing Inputs!B8"),
        ("DSCR sensitivity - higher", FIN["dscr_hi"], '0.00"x"', "Financing Inputs!B9"),
        ("Principal grace end", FIN["grace_end"], DATEF, "Financing Inputs!B10"),
        ("Debt maturity", FIN["maturity"], DATEF, "Financing Inputs!B11"),
        ("Discounting convention (1 = effective annual (1+WACC)^(-m/12); 0 = nominal monthly (1+WACC/12)^-m)", 1, '0', "Team assumption; Financing Inputs!B18 prescribes monthly mid-period discounting"),
        ("Commercial-service delay sensitivity [months, both halls]", 0, '0', "Team lever for sensitivity only (0 in the base case)"),
    ]
    wi.append([]); wi.append(["Input", "Value", "Unit / source"])
    for c in wi[3]: c.font = BOLD
    names = {}
    for i, (lab, val, fmt, src) in enumerate(inputs, 4):
        wi.cell(i, 1, lab); c = wi.cell(i, 2, val); c.number_format = fmt; c.font = BLUE; wi.cell(i, 3, src)
        names[lab] = f"Inputs!$B${i}"
    wi.cell(23, 2).fill = YELLOW; wi.cell(24, 2).fill = YELLOW
    N = lambda k: names[k]
    row = 26
    wi.cell(row, 1, "Derived").font = H2
    derived = [
        ("Full-year revenue per hall in service [kEUR/yr]", f"={N('GPU-eq per hall')}*{N('Hours per year')}*{N('Realised selling rate [EUR/GPU-h]')}/1000", KEUR),
        ("Full-year electricity per hall in service [kEUR/yr]", f"={N('IT design load per hall [MW]')}*1000*{N('PUE')}*{N('Hours per year')}*{N('Electricity price [EUR/kWh]')}/1000", KEUR),
        ("First principal repayment month", f"=DATE(YEAR(EDATE({N('Principal grace end')},1)),MONTH(EDATE({N('Principal grace end')},1)),1)", DATEF),
        ("Number of principal instalments (Apr 2028 - Dec 2037)", f"=(YEAR({N('Debt maturity')})-YEAR(B29))*12+MONTH({N('Debt maturity')})-MONTH(B29)+1", '0'),
        ("Hall A commercial-service date (schedule)", "=Schedule!$B$3", DATEF),
        ("Hall B commercial-service date (schedule)", "=Schedule!$B$4", DATEF),
        ("Final Program Acceptance date (schedule)", "=Schedule!$B$5", DATEF),
        ("Start of F17 (site services overhead start)", "=Schedule!$B$6", DATEF),
        ("Hall A commercial-service date used in the model", f"=EDATE(B31,{N('Commercial-service delay sensitivity [months, both halls]')})", DATEF),
        ("Hall B commercial-service date used in the model", f"=EDATE(B32,{N('Commercial-service delay sensitivity [months, both halls]')})", DATEF),
        ("Final Program Acceptance used in the model", f"=EDATE(B33,{N('Commercial-service delay sensitivity [months, both halls]')})", DATEF),
        ("Deadline check (Final Program Acceptance <= Month 32)", f'=IF(B37<={N("Final Program Acceptance deadline (Month 32)")},"OK","LATE")', '@'),
        ("Selected debt share d = D / CAPEX", "=Financing!$C$5", PCT),
        ("WACC at selected d", f"=B39*{N('Debt cost rD')}+(1-B39)*{N('Equity cost rE')}", PCT),
    ]
    for i, (lab, f, fmt) in enumerate(derived, 27):
        wi.cell(i, 1, lab); c = wi.cell(i, 2, f); c.number_format = fmt
        if "Schedule!" in f or "Financing!" in f: c.font = GREEN
        names[lab] = f"Inputs!$B${i}"
    wi.column_dimensions["A"].width = 70; wi.column_dimensions["B"].width = 16; wi.column_dimensions["C"].width = 60

    # ------------------------------------------------------------------ Schedule
    wsch = wb.create_sheet("Schedule")
    wsch["A1"] = "INTEGRATED MASTER SCHEDULE - activities, dates and resource costs (identical to the MS Project plan)"; wsch["A1"].font = H1
    wsch["A3"] = "Hall A commercial service (day after last condition)"; wsch["B3"] = cf["hall_a"]; wsch["B3"].number_format = DATEF; wsch["B3"].font = BLUE
    wsch["A4"] = "Hall B commercial service (day after last condition)"; wsch["B4"] = cf["hall_b"]; wsch["B4"].number_format = DATEF; wsch["B4"].font = BLUE
    wsch["A5"] = "Final Program Acceptance (end of O15 / F45)"; wsch["B5"] = cf["fpa"]; wsch["B5"].number_format = DATEF; wsch["B5"].font = BLUE
    wsch["A6"] = "Start of F17 Install Temporary Facilities"; wsch["B6"] = wd_date(sched["F17"]["start"]); wsch["B6"].number_format = DATEF; wsch["B6"].font = BLUE
    wsch["C3"] = "Dates exported from the schedule engine / MS Project file (Program milestones)."
    hdr = ["ID", "Activity", "Project", "Location", "Start", "Finish", "Duration [wd]", "Normal duration [wd]", "Crashed?", "Resource assignments (units)",
           "Material assignments", "Work cost [kEUR]", "Material / lot cost [kEUR]", "Total cost [kEUR]", "Cost per working day [kEUR]", "NETWORKDAYS check"]
    wsch.append([]); wsch.append(hdr)
    for c in wsch[8]: c.font = BOLD; c.fill = GREY
    order = sorted([t for t in tasks if t not in MILESTONES], key=lambda t: (sched[t]["start"], t))
    first = 9
    for i, tid in enumerate(order):
        t = tasks[tid]; s = sched[tid]; rr = first + i
        base = t["base"]
        wsch.cell(rr, 1, tid); wsch.cell(rr, 2, t["name"]); wsch.cell(rr, 3, {"F": "P1", "A": "P2", "O": "P3"}[tid[0]]); wsch.cell(rr, 4, t["loc"])
        c = wsch.cell(rr, 5, wd_date(s["start"])); c.number_format = DATEF; c.font = BLUE
        c = wsch.cell(rr, 6, wd_date(s["finish"])); c.number_format = DATEF; c.font = BLUE
        wsch.cell(rr, 7, t["dur"]).font = BLUE; wsch.cell(rr, 8, ACTS[base]["dur"]).font = BLUE
        wsch.cell(rr, 9, "yes" if t["dur"] < ACTS[base]["dur"] else "no")
        wsch.cell(rr, 10, ", ".join(f"{rn} {u*100:.0f}%" for rn, u in t["units"]))
        wsch.cell(rr, 11, ", ".join(f"{rn} x {q:g} {RES[rn]['label']}" for rn, q in t["mat"]))
        c = wsch.cell(rr, 12, round(t["work_cost"] / 1000, 3)); c.number_format = KEUR1; c.font = BLUE
        c = wsch.cell(rr, 13, round(t["mat_cost"] / 1000, 3)); c.number_format = KEUR1; c.font = BLUE
        c = wsch.cell(rr, 14, f"=L{rr}+M{rr}"); c.number_format = KEUR1
        c = wsch.cell(rr, 15, f"=N{rr}/G{rr}"); c.number_format = '#,##0.000'
        c = wsch.cell(rr, 16, f'=IF(NETWORKDAYS(E{rr},F{rr})=G{rr},"OK","CHECK")')
    last = first + len(order) - 1
    wsch.cell(last + 2, 1, "Total"); wsch.cell(last + 2, 12, f"=SUM(L{first}:L{last})").number_format = KEUR; wsch.cell(last + 2, 13, f"=SUM(M{first}:M{last})").number_format = KEUR; wsch.cell(last + 2, 14, f"=SUM(N{first}:N{last})").number_format = KEUR
    wsch.cell(last + 3, 1, "Work cost = units x 8 h x working days x standard rate; material cost = quantity x unit price (Resources sheet of the case workbook). Crashed activities use the crash-resource units of workbook note 4. Values exported from the schedule engine; MS Project computes the same figures from the assignments.")
    for col, w in zip("ABCDEFGHIJKLMNOP", [8, 70, 8, 12, 12, 12, 12, 12, 9, 60, 50, 14, 16, 14, 16, 12]):
        wsch.column_dimensions[col].width = w
    SCH_FIRST, SCH_LAST = first, last

    # ------------------------------------------------------------------ Procurement
    wp = wb.create_sheet("Procurement")
    wp["A1"] = "PROCUREMENT AND LOGISTICS PLAN - purchase orders, lead times, expediting, transport, delivery vs required-on-site"; wp["A1"].font = H1
    hdr = ["Package", "Lot", "Ordering activity / PO event", "PO date (end of ordering activity)", "Lot value [kEUR]", "Expedited?", "Manufacturing [cd]", "Line free from previous lot",
           "Ready for shipment", "Transport mode", "Transport [cd]", "Delivered on site", "Required on site before (activity)", "Required-on-site date (activity start)", "Float [cd]",
           "TEU charged", "Transport cost [kEUR] (paid at delivery)", "Expedite premium [kEUR] (paid at PO)", "Payment of equipment"]
    wp.append([]); wp.append(hdr)
    for c in wp[3]: c.font = BOLD; c.fill = GREY
    rr = 4
    for (pkg, lot), info in sorted(dl.items(), key=lambda kv: (kv[1]["po_date"], kv[0])):
        acts = NEEDS[pkg][lot]
        if pkg == "Generators" and info["disassembled"]:
            acts = ["F46" if lot == 1 else "F47"]
        req = min(sched[a]["start"] for a in acts)
        prev_ready = info["ready"] - dt.timedelta(days=info["mfg_cd"])
        vals = [pkg, lot, info["po_event"], info["po_date"], info["value"] / 1000, "yes" if info["expedite"] else "no", info["mfg_cd"],
                prev_ready if prev_ready > info["po_date"] else None, info["ready"], info["mode"] + (" (disassembled)" if info["disassembled"] else ""), info["transport_cd"], info["delivery"],
                " / ".join(acts), wd_date(req), None, info["teu"], info["transport_cost"] / 1000, info["premium"] / 1000,
                "paid with the ordering activity (lot material resource)" if PK[pkg]["origin"] == "East Asia" else "paid as installed (material resource of the installing activity)"]
        for j, v in enumerate(vals, 1):
            c = wp.cell(rr, j, v)
            if j in (4, 8, 9, 12, 14): c.number_format = DATEF
            if j in (5, 17, 18): c.number_format = KEUR1
            if j in (4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 16, 17, 18): c.font = BLUE
        wp.cell(rr, 15, f"=N{rr}-L{rr}-1")
        rr += 1
    PROC_FIRST, PROC_LAST = 4, rr - 1
    wp.cell(rr + 1, 1, "Lead-time rule (case): lot ready = max(PO date, previous lot ready on the supplier's single line) + manufacturing time (x expedite factor, rounded to whole days); delivery = ready + transport duration; float = required-on-site date - delivery - 1.")
    wp.cell(rr + 2, 1, "Transport charge = TEU x rate; surface TEU = ROUNDUP(volume/33) + out-of-gauge pieces x TEU each; air TEU = ROUNDUP(MAX(volume/33, weight/5.5)). Expedite premium = premium % x value of the lots in the PO, paid at order placement.")
    for col, w in zip("ABCDEFGHIJKLMNOPQRS", [28, 5, 26, 14, 12, 9, 12, 14, 14, 24, 10, 14, 22, 16, 9, 8, 14, 14, 50]):
        wp.column_dimensions[col].width = w

    # ------------------------------------------------------------------ ActivityCost_M (grid)
    wg = wb.create_sheet("ActivityCost_M")
    wg["A1"] = "ACTIVITY COST BY MONTH [kEUR] = cost per working day x NETWORKDAYS(overlap of activity and month)"; wg["A1"].font = H1
    NG = 40  # Jan 2027 .. Apr 2030
    wg["A3"] = "Month start"; wg["A4"] = "Month end"; wg["A5"] = "ID"
    for m in range(NG):
        col = 5 + m
        c = wg.cell(3, col, month_start(m)); c.number_format = 'mmm-yy'; c.font = BLUE
        c = wg.cell(4, col, f"=EOMONTH({L(col)}3,0)"); c.number_format = DATEF
        wg.cell(5, col, f"={L(col)}3").number_format = 'mmm-yy'
    wg["B5"] = "Start"; wg["C5"] = "Finish"; wg["D5"] = "Cost/wd"
    for c in wg[5]: c.font = BOLD; c.fill = GREY
    for i in range(SCH_FIRST, SCH_LAST + 1):
        gr = 6 + (i - SCH_FIRST)
        wg.cell(gr, 1, f"=Schedule!A{i}").font = GREEN
        wg.cell(gr, 2, f"=Schedule!E{i}").number_format = DATEF; wg.cell(gr, 2).font = GREEN
        wg.cell(gr, 3, f"=Schedule!F{i}").number_format = DATEF; wg.cell(gr, 3).font = GREEN
        wg.cell(gr, 4, f"=Schedule!O{i}").number_format = '#,##0.000'; wg.cell(gr, 4).font = GREEN
        for m in range(NG):
            col = L(5 + m)
            wg.cell(gr, 5 + m, f"=IF(OR($C{gr}<{col}$3,$B{gr}>{col}$4),0,$D{gr}*NETWORKDAYS(MAX($B{gr},{col}$3),MIN($C{gr},{col}$4)))").number_format = KEUR1
    G_FIRST, G_LAST = 6, 6 + (SCH_LAST - SCH_FIRST)
    tot = G_LAST + 1
    wg.cell(tot, 1, "Total [kEUR]").font = BOLD
    for m in range(NG):
        col = L(5 + m)
        wg.cell(tot, 5 + m, f"=SUM({col}{G_FIRST}:{col}{G_LAST})").number_format = KEUR
    wg.cell(tot + 1, 1, "Check: grid total = Schedule total"); wg.cell(tot + 1, 5, f'=IF(ABS(SUM(E{tot}:{L(4+NG)}{tot})-Schedule!N{SCH_LAST+2})<0.5,"OK","CHECK")')
    wg.cell(tot + 2, 1, "Check: all activities finish within the grid"); wg.cell(tot + 2, 5, f'=IF(MAX(C{G_FIRST}:C{G_LAST})<={L(4+NG)}4,"OK","EXTEND GRID")')
    wg.column_dimensions["A"].width = 10; wg.column_dimensions["B"].width = 12; wg.column_dimensions["C"].width = 12
    GRID_TOT_ROW = tot

    # ------------------------------------------------------------------ Monthly
    wm = wb.create_sheet("Monthly")
    wm["A1"] = "MONTHLY UNLEVERED PROJECT CASH FLOW (FCFF) [kEUR] - identical for every financing mix"; wm["A1"].font = H1
    labels = ["Month start", "Month end", "m (mid-period index m+0.5)", "Days in month",
              "Activity costs incl. equipment and lots (prorated)", "Expedite premiums (paid at PO)", "Transport costs (paid at delivery)",
              "Overhead - PMO", "Overhead - general site services", "Overhead - ops-readiness management",
              "Total CAPEX", "Cumulative CAPEX",
              "Hall A days in service", "Hall B days in service", "Revenue Hall A", "Revenue Hall B", "Electricity Hall A", "Electricity Hall B",
              "Total revenue", "Total electricity", "CFADS (revenue - electricity)",
              "Unlevered FCFF", "Cumulative FCFF", "Discount factor at selected WACC", "PV of FCFF", "Cumulative PV"]
    for i, lab in enumerate(labels, 3):
        wm.cell(i, 1, lab)
    for i in (11, 19, 20, 21, 22, 23, 25, 26): wm.cell(i, 1).font = BOLD
    FIRSTC = 2
    for m in range(NMONTHS):
        col = L(FIRSTC + m); c = FIRSTC + m
        wm.cell(3, c, month_start(m)).number_format = 'mmm-yy'; wm.cell(3, c).font = BLUE
        wm.cell(4, c, f"=EOMONTH({col}3,0)").number_format = DATEF
        wm.cell(5, c, m).font = BLUE
        wm.cell(6, c, f"=DAY({col}4)")
        if m < NG:
            gcol = L(5 + m)
            wm.cell(7, c, f"=ActivityCost_M!{gcol}{GRID_TOT_ROW}").font = GREEN
            wm.cell(8, c, f"=SUMIFS(Procurement!$R${PROC_FIRST}:$R${PROC_LAST},Procurement!$D${PROC_FIRST}:$D${PROC_LAST},\">=\"&{col}3,Procurement!$D${PROC_FIRST}:$D${PROC_LAST},\"<=\"&{col}4)").font = GREEN
            wm.cell(9, c, f"=SUMIFS(Procurement!$Q${PROC_FIRST}:$Q${PROC_LAST},Procurement!$L${PROC_FIRST}:$L${PROC_LAST},\">=\"&{col}3,Procurement!$L${PROC_FIRST}:$L${PROC_LAST},\"<=\"&{col}4)").font = GREEN
        else:
            wm.cell(7, c, 0); wm.cell(8, c, 0); wm.cell(9, c, 0)
        fpa = N("Final Program Acceptance used in the model"); f17 = N("Start of F17 (site services overhead start)"); ntp = N("Program NTP")
        wm.cell(10, c, f"={N('Overhead - PMO [kEUR/month]')}*12/365*MAX(0,MIN({fpa},{col}4)-MAX({ntp},{col}3)+1)")
        wm.cell(11, c, f"={N('Overhead - General site services [kEUR/month]')}*12/365*MAX(0,MIN({fpa},{col}4)-MAX({f17},{col}3)+1)")
        wm.cell(12, c, f"={N('Overhead - Operations-readiness management [kEUR/month]')}*12/365*MAX(0,MIN({fpa},{col}4)-MAX({ntp},{col}3)+1)")
        wm.cell(13, c, f"=SUM({col}7:{col}12)")
        wm.cell(14, c, f"={col}13" if m == 0 else f"={L(c-1)}14+{col}13")
        ha = N("Hall A commercial-service date used in the model"); hb = N("Hall B commercial-service date used in the model"); ae = N("Financial appraisal end")
        wm.cell(15, c, f"=MAX(0,MIN({ae},{col}4)-MAX({ha},{col}3)+1)")
        wm.cell(16, c, f"=MAX(0,MIN({ae},{col}4)-MAX({hb},{col}3)+1)")
        rev = N("Full-year revenue per hall in service [kEUR/yr]"); el = N("Full-year electricity per hall in service [kEUR/yr]")
        wm.cell(17, c, f"={rev}/365*{col}15"); wm.cell(18, c, f"={rev}/365*{col}16")
        wm.cell(19, c, f"={el}/365*{col}15"); wm.cell(20, c, f"={el}/365*{col}16")
        wm.cell(21, c, f"={col}17+{col}18"); wm.cell(22, c, f"={col}19+{col}20"); wm.cell(23, c, f"={col}21-{col}22")
        wm.cell(24, c, f"={col}21-{col}22-{col}13")
        wm.cell(25, c, f"={col}24" if m == 0 else f"={L(c-1)}25+{col}24")
        w = N("WACC at selected d"); conv = N("Discounting convention (1 = effective annual (1+WACC)^(-m/12); 0 = nominal monthly (1+WACC/12)^-m)")
        wm.cell(26, c, f"=IF({conv}=1,(1+{w})^(-({col}5+0.5)/12),(1+{w}/12)^(-({col}5+0.5)))")
        wm.cell(27, c, f"={col}24*{col}26")
        wm.cell(28, c, f"={col}27" if m == 0 else f"={L(c-1)}28+{col}27")
        for rr_ in range(7, 29):
            wm.cell(rr_, c).number_format = KEUR if rr_ not in (26,) else '0.0000'
        wm.cell(15, c).number_format = '0'; wm.cell(16, c).number_format = '0'
    LASTC = L(FIRSTC + NMONTHS - 1)
    wm.cell(1, 1).font = H1
    # totals column
    tc = FIRSTC + NMONTHS + 1
    wm.cell(2, tc, "Total 2027-2037").font = BOLD
    for rr_ in list(range(7, 14)) + list(range(15, 25)) + [27]:
        wm.cell(rr_, tc, f"=SUM(B{rr_}:{LASTC}{rr_})").number_format = KEUR; wm.cell(rr_, tc).font = BOLD
    wm.cell(30, 1, "NPV at selected WACC [kEUR] = sum of PV").font = BOLD; wm.cell(30, 2, f"=SUM(B27:{LASTC}27)").number_format = KEUR; wm.cell(30, 2).font = BOLD
    wm.cell(31, 1, "Total CAPEX [kEUR]"); wm.cell(31, 2, f"=SUM(B13:{LASTC}13)").number_format = KEUR
    wm.cell(32, 1, "Payback month (first month with cumulative FCFF >= 0)"); wm.cell(32, 2, f'=IFERROR(INDEX(B3:{LASTC}3,MATCH(TRUE,INDEX(B25:{LASTC}25>=0,0),0)),"n/a")'); wm.cell(32, 2).number_format = 'mmm-yy'
    wm.cell(33, 1, "Undiscounted cumulative FCFF at horizon [kEUR]"); wm.cell(33, 2, f"={LASTC}25").number_format = KEUR
    wm.column_dimensions["A"].width = 52
    wm.freeze_panes = "B7"
    MROW = dict(capex=13, cfads=23, fcff=24, m=5, mstart=3, mend=4)

    # ------------------------------------------------------------------ Financing (D/E table)
    wf = wb.create_sheet("Financing")
    wf["A1"] = "DEBT / EQUITY OPTIMISATION - same FCFF, each mix discounted at its own WACC; feasible = annual DSCR >= covenant in every test year and full repayment by maturity"; wf["A1"].font = H1
    wf["A3"] = "Covenant used for selection"; wf["B3"] = f"={N('Minimum DSCR (covenant)')}"; wf["B3"].number_format = '0.00"x"'; wf["B3"].font = GREEN; wf["B3"].fill = YELLOW
    wf["A4"] = "Selected debt share (highest feasible d, step 0.5%)"; wf["A5"] = "d = D / CAPEX"; wf["C5"] = "=MAX(IF(1,0),SUMPRODUCT(MAX((J11:J211=1)*A11:A211)))"; wf["C5"].number_format = PCT; wf["C5"].font = BOLD; wf["C5"].fill = YELLOW
    wf["D5"] = "<- formula picks the largest feasible d in the table; override here only for sensitivity (also drives Debt_Ledger and Monthly discount factors)"
    wf["A6"] = "Total CAPEX [kEUR]"; wf["B6"] = "=Monthly!B31"; wf["B6"].number_format = KEUR; wf["B6"].font = GREEN
    wf["A7"] = "Committed debt D [kEUR]"; wf["B7"] = "=C5*B6"; wf["B7"].number_format = KEUR
    wf["A8"] = "Equity E [kEUR]"; wf["B8"] = "=B6-B7"; wf["B8"].number_format = KEUR
    years = list(range(2028, 2038))
    hdr = ["d (Debt/CAPEX)", "WACC", "NPV [kEUR]"] + [f"DSCR {y}" + (" (Apr-Dec)" if y == 2028 else "") for y in years] + ["Min DSCR", "Feasible (1 = yes)", "Final balance [kEUR]"]
    # layout: A d, B WACC, C NPV, D..M DSCR 2028..2037 (10 cols), N min, O feasible, P final balance -> but formula above references J; recompute columns
    cols = {"d": 1, "wacc": 2, "npv": 3}
    for k, y in enumerate(years): cols[y] = 4 + k
    cols["min"] = 4 + len(years); cols["feas"] = cols["min"] + 1; cols["bal"] = cols["feas"] + 1
    for j, h in enumerate(hdr, 1):
        wf.cell(10, j, h).font = BOLD; wf.cell(10, j).fill = GREY
    wf["C5"] = f"=SUMPRODUCT(MAX(({L(cols['feas'])}11:{L(cols['feas'])}211=1)*A11:A211))"
    # unit debt-service per year (d = 100%) from Debt_Ledger unit columns
    for i in range(201):
        rr = 11 + i; d = i * 0.005
        wf.cell(rr, 1, round(d, 3)).number_format = PCT; wf.cell(rr, 1).font = BLUE
        wf.cell(rr, 2, f"=A{rr}*{N('Debt cost rD')}+(1-A{rr})*{N('Equity cost rE')}").number_format = PCT
        conv = N("Discounting convention (1 = effective annual (1+WACC)^(-m/12); 0 = nominal monthly (1+WACC/12)^-m)")
        wf.cell(rr, 3, f"=IF({conv}=1,SUMPRODUCT(Monthly!$B$24:${LASTC}$24,(1+B{rr})^(-(Monthly!$B$5:${LASTC}$5+0.5)/12)),SUMPRODUCT(Monthly!$B$24:${LASTC}$24,(1+B{rr}/12)^(-(Monthly!$B$5:${LASTC}$5+0.5))))").number_format = KEUR
        for y in years:
            # DSCR_y(d) = CFADS_y / (d x unit debt service_y); unit DS from Debt_Ledger annual table (d = 100%)
            wf.cell(rr, cols[y], f"=IF(A{rr}=0,\"n/a\",Debt_Ledger!$C${20 + y - 2028}/(A{rr}*Debt_Ledger!$H${20 + y - 2028}))").number_format = '0.00'
        wf.cell(rr, cols["min"], f"=IF(A{rr}=0,\"n/a\",MIN({L(cols[2028])}{rr}:{L(cols[2037])}{rr}))").number_format = '0.00'
        wf.cell(rr, cols["feas"], f"=IF(A{rr}=0,1,IF(AND({L(cols['min'])}{rr}>=$B$3,ABS({L(cols['bal'])}{rr})<0.5),1,0))")
        wf.cell(rr, cols["bal"], f"=A{rr}*Debt_Ledger!$I$17").number_format = KEUR
    wf.cell(213, 1, "Selected row (highest feasible NPV)").font = BOLD
    wf.cell(214, 1, f"=C5").number_format = PCT; wf.cell(214, 2, f"=C5*{N('Debt cost rD')}+(1-C5)*{N('Equity cost rE')}").number_format = PCT
    wf.cell(214, 3, f"=INDEX(C11:C211,MATCH(C5,A11:A211,0))").number_format = KEUR; wf.cell(214, 3).font = BOLD
    wf.cell(214, cols["min"], f"=INDEX({L(cols['min'])}11:{L(cols['min'])}211,MATCH(C5,A11:A211,0))").number_format = '0.00'
    # covenant sensitivities
    wf.cell(216, 1, "Covenant sensitivity (re-optimised d)").font = H2
    wf.cell(217, 1, "Covenant"); wf.cell(217, 2, "Max feasible d"); wf.cell(217, 3, "WACC"); wf.cell(217, 4, "NPV [kEUR]")
    for c in wf[217]: c.font = BOLD
    for k, (lab, ref) in enumerate([("Lower 1.50x", N("DSCR sensitivity - lower")), ("Base 1.65x", N("Minimum DSCR (covenant)")), ("Higher 1.80x", N("DSCR sensitivity - higher"))]):
        rr = 218 + k
        wf.cell(rr, 1, f"={ref}").number_format = '0.00"x"'; wf.cell(rr, 1).font = GREEN
        mn = L(cols["min"]); bal = L(cols["bal"])
        wf.cell(rr, 2, f"=SUMPRODUCT(MAX(((A11:A211=0)+(({mn}11:{mn}211>=A{rr})*(ABS({bal}11:{bal}211)<0.5))>0)*A11:A211))").number_format = PCT
        wf.cell(rr, 3, f"=B{rr}*{N('Debt cost rD')}+(1-B{rr})*{N('Equity cost rE')}").number_format = PCT
        wf.cell(rr, 4, f"=INDEX(C11:C211,MATCH(B{rr},A11:A211,0))").number_format = KEUR
    wf.cell(222, 1, "Note: DSCR is linear in 1/d (debt service scales with d while CFADS does not), so the table uses the unit ledger (d = 100%) on Debt_Ledger and scales it; the explicit ledger for the selected d is shown on Debt_Ledger and must give the same DSCR values (check cell on Debt_Ledger).")
    wf.column_dimensions["A"].width = 44
    for j in range(2, 18): wf.column_dimensions[L(j)].width = 14
    wf.freeze_panes = "B11"

    # ------------------------------------------------------------------ Debt ledger
    wl = wb.create_sheet("Debt_Ledger")
    wl["A1"] = "MONTHLY BORROWING AND REPAYMENT LEDGER [kEUR] - selected mix (left) and unit ledger at d = 100% (used by the D/E table)"; wl["A1"].font = H1
    wl["A3"] = "Selected d"; wl["B3"] = "=Financing!C5"; wl["B3"].number_format = PCT; wl["B3"].font = GREEN
    wl["A4"] = "Total CAPEX [kEUR]"; wl["B4"] = "=Monthly!B31"; wl["B4"].number_format = KEUR; wl["B4"].font = GREEN
    wl["A5"] = "Committed loan D [kEUR]"; wl["B5"] = "=B3*B4"; wl["B5"].number_format = KEUR
    wl["A6"] = "Monthly principal instalment (D / instalments)"; wl["B6"] = f"=B5/{N('Number of principal instalments (Apr 2028 - Dec 2037)')}"; wl["B6"].number_format = KEUR1
    wl["A7"] = "Monthly interest rate (rD / 12)"; wl["B7"] = f"={N('Debt cost rD')}/12"; wl["B7"].number_format = '0.0000%'
    wl["A8"] = "First repayment month"; wl["B8"] = f"={N('First principal repayment month')}"; wl["B8"].number_format = DATEF; wl["B8"].font = GREEN
    wl["A9"] = "Unit ledger: committed loan at d = 100% [kEUR]"; wl["B9"] = "=B4"; wl["B9"].number_format = KEUR
    wl["A10"] = "Unit monthly principal"; wl["B10"] = f"=B9/{N('Number of principal instalments (Apr 2028 - Dec 2037)')}"; wl["B10"].number_format = KEUR1
    # annual DSCR table rows 19..29 (2028..2037) -> at row 20+ (y-2028)
    wl["A18"] = "ANNUAL POST-GRACE DSCR TEST"; wl["A18"].font = H2
    hdr = ["Year", "CFADS (selected)", "CFADS [kEUR]", "Interest (selected)", "Principal (selected)", "Debt service (selected)", "DSCR (selected d)", "Unit debt service (d=100%)", "Unit final balance", "Covenant", "Pass?"]
    for j, h in enumerate(hdr, 1): wl.cell(19, j, h).font = BOLD; wl.cell(19, j).fill = GREY
    LFIRST = 33
    hdr2 = ["Month", "m", "CAPEX [kEUR]", "Draw (d x CAPEX)", "Balance after draw", "Interest", "Principal", "Balance end", "Debt service", "CFADS", "Repayment month? (1/0)", "",
            "UNIT (d=100%): Draw", "Balance after draw", "Interest", "Principal", "Balance end", "Debt service"]
    for j, h in enumerate(hdr2, 1): wl.cell(LFIRST - 1, j, h).font = BOLD; wl.cell(LFIRST - 1, j).fill = GREY
    for m in range(NMONTHS):
        rr = LFIRST + m; mc = L(FIRSTC + m)
        wl.cell(rr, 1, f"=Monthly!{mc}3").number_format = 'mmm-yy'; wl.cell(rr, 1).font = GREEN
        wl.cell(rr, 2, f"=Monthly!{mc}5").font = GREEN
        wl.cell(rr, 3, f"=Monthly!{mc}13").number_format = KEUR; wl.cell(rr, 3).font = GREEN
        wl.cell(rr, 4, f"=$B$3*C{rr}").number_format = KEUR
        wl.cell(rr, 5, (f"=D{rr}" if m == 0 else f"=H{rr-1}+D{rr}")).number_format = KEUR
        wl.cell(rr, 6, f"=E{rr}*$B$7").number_format = KEUR
        wl.cell(rr, 11, f"=IF(A{rr}>=$B$8,1,0)")
        wl.cell(rr, 7, f"=K{rr}*MIN($B$6,E{rr})").number_format = KEUR
        wl.cell(rr, 8, f"=E{rr}-G{rr}").number_format = KEUR
        wl.cell(rr, 9, f"=F{rr}+G{rr}").number_format = KEUR
        wl.cell(rr, 10, f"=Monthly!{mc}23").number_format = KEUR; wl.cell(rr, 10).font = GREEN
        wl.cell(rr, 13, f"=C{rr}").number_format = KEUR
        wl.cell(rr, 14, (f"=M{rr}" if m == 0 else f"=Q{rr-1}+M{rr}")).number_format = KEUR
        wl.cell(rr, 15, f"=N{rr}*$B$7").number_format = KEUR
        wl.cell(rr, 16, f"=K{rr}*MIN($B$10,N{rr})").number_format = KEUR
        wl.cell(rr, 17, f"=N{rr}-P{rr}").number_format = KEUR
        wl.cell(rr, 18, f"=O{rr}+P{rr}").number_format = KEUR
    LLAST = LFIRST + NMONTHS - 1
    for k, y in enumerate(years):
        rr = 20 + k
        wl.cell(rr, 1, y).font = BLUE
        ycond = f"\">=\"&DATE({y},IF({y}=2028,4,1),1),A${LFIRST}:A${LLAST},\"<=\"&DATE({y},12,31)"
        wl.cell(rr, 2, f"=SUMIFS(J${LFIRST}:J${LLAST},A${LFIRST}:A${LLAST},{ycond})").number_format = KEUR
        wl.cell(rr, 3, f"=B{rr}").number_format = KEUR
        wl.cell(rr, 4, f"=SUMIFS(F${LFIRST}:F${LLAST},A${LFIRST}:A${LLAST},{ycond})").number_format = KEUR
        wl.cell(rr, 5, f"=SUMIFS(G${LFIRST}:G${LLAST},A${LFIRST}:A${LLAST},{ycond})").number_format = KEUR
        wl.cell(rr, 6, f"=D{rr}+E{rr}").number_format = KEUR
        wl.cell(rr, 7, f"=IF(F{rr}>0,C{rr}/F{rr},\"n/a\")").number_format = '0.00'
        wl.cell(rr, 8, f"=SUMIFS(R${LFIRST}:R${LLAST},A${LFIRST}:A${LLAST},{ycond})").number_format = KEUR
        wl.cell(rr, 10, f"={N('Minimum DSCR (covenant)')}").number_format = '0.00'; wl.cell(rr, 10).font = GREEN
        wl.cell(rr, 11, f'=IF(F{rr}=0,"n/a",IF(G{rr}>=J{rr},"PASS","FAIL"))')
    wl.cell(17, 1, "Unit final balance at maturity (must be 0)"); wl.cell(17, 9, f"=Q{LLAST}").number_format = KEUR
    wl.cell(12, 1, "Final drawn balance at maturity, selected d (must be 0)"); wl.cell(12, 2, f"=H{LLAST}").number_format = KEUR
    wl.cell(13, 1, "Total interest paid [kEUR]"); wl.cell(13, 2, f"=SUM(F{LFIRST}:F{LLAST})").number_format = KEUR
    wl.cell(14, 1, "Total principal repaid [kEUR] (= committed loan)"); wl.cell(14, 2, f"=SUM(G{LFIRST}:G{LLAST})").number_format = KEUR
    wl.cell(15, 1, "Interest during grace (Jan 2027 - Mar 2028), funded by sponsor, not in FCFF"); wl.cell(15, 2, f"=SUMIFS(F{LFIRST}:F{LLAST},K{LFIRST}:K{LLAST},0)").number_format = KEUR
    wl.cell(16, 1, "Minimum annual DSCR (selected d)"); wl.cell(16, 2, "=MIN(G20:G29)").number_format = '0.00'; wl.cell(16, 3, '=IF(B16>=J20,"COVENANT MET","COVENANT BREACHED")')
    wl.column_dimensions["A"].width = 46
    for j in range(2, 19): wl.column_dimensions[L(j)].width = 15
    wl.freeze_panes = f"B{LFIRST}"

    # ------------------------------------------------------------------ Resources (verification)
    wr = wb.create_sheet("Resources")
    wr["A1"] = "RESOURCE-FEASIBILITY VERIFICATION - peak concurrent units on any working day vs Max. Units (values from the schedule engine; cross-check in MS Project Resource Usage)"; wr["A1"].font = H1
    hdr = ["Resource", "Max. Units", "Peak units used", "Peak date", "Compliant?", "Std rate [EUR/h]"]
    for j, h in enumerate(hdr, 1): wr.cell(3, j, h).font = BOLD; wr.cell(3, j).fill = GREY
    rr = 4
    for name, rs in RES.items():
        if rs["type"] != "Work": continue
        arr = r["usage"][name]
        peak = max(arr); pd = arr.index(peak)
        wr.cell(rr, 1, name); wr.cell(rr, 2, rs["max"]).font = BLUE; wr.cell(rr, 3, round(peak, 2)).font = BLUE
        wr.cell(rr, 4, wd_date(pd) if peak > 0 else None).number_format = DATEF
        wr.cell(rr, 5, f'=IF(C{rr}<=B{rr}+0.0001,"OK","OVER")'); wr.cell(rr, 6, rs["rate"]).font = BLUE
        rr += 1
    wr.cell(rr, 1, "Heavy-equipment corridor (team assumption: max 2 concurrent heavy-equipment installation fronts)"); wr.cell(rr, 2, CORRIDOR_MAX).font = BLUE
    wr.cell(rr, 3, round(max(r["usage"]["__corridor__"]), 2)).font = BLUE; wr.cell(rr, 5, f'=IF(C{rr}<=B{rr},"OK","OVER")')
    wr.column_dimensions["A"].width = 60
    for j in range(2, 7): wr.column_dimensions[L(j)].width = 16

    # ------------------------------------------------------------------ Summary
    wsum = wb.create_sheet("Summary", 1)
    wsum["A1"] = "SUMMARY OF REQUIRED FINANCIAL OUTPUTS"; wsum["A1"].font = H1
    items = [
        ("Highest feasible program NPV [kEUR] (at WACC of the selected D/E mix)", "=Financing!C214", KEUR),
        ("Selected debt share d = D / CAPEX", "=Financing!C5", PCT),
        ("WACC", f"={N('WACC at selected d')}", PCT),
        ("Total program CAPEX [kEUR] (activities, equipment, lots, expediting, transport, overheads)", "=Monthly!B31", KEUR),
        ("Debt D [kEUR]", "=Financing!B7", KEUR),
        ("Equity E [kEUR]", "=Financing!B8", KEUR),
        ("Hall A commercial-service date", f"={N('Hall A commercial-service date used in the model')}", DATEF),
        ("Hall B commercial-service date", f"={N('Hall B commercial-service date used in the model')}", DATEF),
        ("Final Program Acceptance date", f"={N('Final Program Acceptance used in the model')}", DATEF),
        ("Final Program Acceptance deadline (Month 32 after NTP)", f"={N('Final Program Acceptance deadline (Month 32)')}", DATEF),
        ("Completion verification", f"={N('Deadline check (Final Program Acceptance <= Month 32)')}", '@'),
        ("Minimum annual post-grace DSCR (selected d)", "=Debt_Ledger!B16", '0.00"x"'),
        ("DSCR covenant verification", "=Debt_Ledger!C16", '@'),
        ("Full repayment by maturity (final balance, kEUR)", "=Debt_Ledger!B12", KEUR),
        ("Undiscounted cumulative FCFF at horizon [kEUR]", "=Monthly!B33", KEUR),
        ("Payback month (cumulative FCFF turns positive)", "=Monthly!B32", 'mmm-yy'),
        ("Total activity costs incl. equipment and lots [kEUR]", f"=Monthly!{L(tc)}7", KEUR),
        ("Total expedite premiums [kEUR]", f"=Monthly!{L(tc)}8", KEUR),
        ("Total transport costs [kEUR]", f"=Monthly!{L(tc)}9", KEUR),
        ("Total overheads [kEUR]", f"=Monthly!{L(tc)}10+Monthly!{L(tc)}11+Monthly!{L(tc)}12", KEUR),
        ("Total revenue 2027-2037 [kEUR]", f"=Monthly!{L(tc)}21", KEUR),
        ("Total electricity 2027-2037 [kEUR]", f"=Monthly!{L(tc)}22", KEUR),
        ("NPV with 1.50x covenant (re-optimised d) [kEUR]", "=Financing!D218", KEUR),
        ("NPV with 1.80x covenant (re-optimised d) [kEUR]", "=Financing!D220", KEUR),
    ]
    for i, (lab, f, fmt) in enumerate(items, 3):
        wsum.cell(i, 1, lab); c = wsum.cell(i, 2, f); c.number_format = fmt; c.font = GREEN
    wsum.cell(3, 2).font = Font(name="Arial", size=11, bold=True, color="008000")
    # acceleration decisions
    rr = 3 + len(items) + 2
    wsum.cell(rr, 1, "ACCELERATION (CRASHING) DECISIONS").font = H2; rr += 1
    hdr = ["Activity", "Normal [wd]", "Chosen [wd]", "Minimum [wd]", "Crash resources and units at chosen duration", "Extra cost vs normal [kEUR]"]
    for j, h in enumerate(hdr, 1): wsum.cell(rr, j, h).font = BOLD; wsum.cell(rr, j).fill = GREY
    rr += 1
    for aid, a in ACTS.items():
        if a["mindur"] and dec["crash"].get(aid, a["dur"]) < a["dur"]:
            d = dec["crash"][aid]
            units = crash_units(a, d)
            normal_cost = sum(u * 8 * RES[rn]["rate"] for rn, u in a["work"]) * a["dur"]
            new_cost = sum(u * 8 * RES[rn]["rate"] for rn, u in units) * d
            wsum.cell(rr, 1, f"{aid} {a['name']}"); wsum.cell(rr, 2, a["dur"]); wsum.cell(rr, 3, d); wsum.cell(rr, 4, a["mindur"])
            wsum.cell(rr, 5, ", ".join(f"{rn} {u*100:.0f}%" for rn, u in units if rn in a["crash"]))
            wsum.cell(rr, 6, round((new_cost - normal_cost) / 1000, 1)).number_format = KEUR1
            rr += 1
    rr += 1
    wsum.cell(rr, 1, "PROCUREMENT DECISIONS - see sheet Procurement (PO dates, expediting, transport mode, delivery vs required-on-site)").font = H2
    wsum.column_dimensions["A"].width = 85; wsum.column_dimensions["B"].width = 18; wsum.column_dimensions["E"].width = 60; wsum.column_dimensions["F"].width = 18

    # ------------------------------------------------------------------ Sensitivity (simple one-way on NPV via formulas)
    wsn = wb.create_sheet("Sensitivity")
    wsn["A1"] = "ONE-WAY SENSITIVITIES ON NPV (selected WACC held constant; FCFF rebuilt analytically from Monthly totals)"; wsn["A1"].font = H1
    wsn["A3"] = "Base NPV [kEUR]"; wsn["B3"] = "=Financing!C214"; wsn["B3"].number_format = KEUR; wsn["B3"].font = GREEN
    wsn["A4"] = "PV of revenue at selected WACC [kEUR]"; wsn["B4"] = f"=SUMPRODUCT(Monthly!B21:{LASTC}21,Monthly!B26:{LASTC}26)"; wsn["B4"].number_format = KEUR
    wsn["A5"] = "PV of electricity at selected WACC [kEUR]"; wsn["B5"] = f"=SUMPRODUCT(Monthly!B22:{LASTC}22,Monthly!B26:{LASTC}26)"; wsn["B5"].number_format = KEUR
    wsn["A6"] = "PV of CAPEX at selected WACC [kEUR]"; wsn["B6"] = f"=SUMPRODUCT(Monthly!B13:{LASTC}13,Monthly!B26:{LASTC}26)"; wsn["B6"].number_format = KEUR
    wsn["A8"] = "Selling rate change"; wsn["B8"] = "NPV [kEUR]"; wsn["A8"].font = BOLD; wsn["B8"].font = BOLD
    for k, chg in enumerate([-0.3, -0.2, -0.1, 0, 0.1, 0.2]):
        wsn.cell(9 + k, 1, chg).number_format = '+0%;-0%;0%'; wsn.cell(9 + k, 1).font = BLUE
        wsn.cell(9 + k, 2, f"=B4*(1+A{9+k})-B5-B6").number_format = KEUR
    wsn["D8"] = "Electricity price change"; wsn["E8"] = "NPV [kEUR]"; wsn["D8"].font = BOLD; wsn["E8"].font = BOLD
    for k, chg in enumerate([-0.3, 0, 0.3, 0.6, 1.0]):
        wsn.cell(9 + k, 4, chg).number_format = '+0%;-0%;0%'; wsn.cell(9 + k, 4).font = BLUE
        wsn.cell(9 + k, 5, f"=B4-B5*(1+D{9+k})-B6").number_format = KEUR
    wsn["G8"] = "WACC (absolute)"; wsn["H8"] = "NPV [kEUR]"; wsn["G8"].font = BOLD; wsn["H8"].font = BOLD
    for k, w in enumerate([0.08, 0.10, 0.12, 0.14, 0.16]):
        wsn.cell(9 + k, 7, w).number_format = PCT; wsn.cell(9 + k, 7).font = BLUE
        wsn.cell(9 + k, 8, f"=SUMPRODUCT(Monthly!$B$24:${LASTC}$24,(1+G{9+k})^(-(Monthly!$B$5:${LASTC}$5+0.5)/12))").number_format = KEUR
    wsn["A16"] = "Commercial-service delay: set Inputs!B24 (months) and read Summary; each month of delay of both halls removes about one month of net revenue at the end of the horizon and shifts the DSCR profile."
    wsn["A17"] = "Note: the selling-rate and electricity sensitivities hold the financing mix (WACC) constant; a lower CFADS would also reduce the feasible debt share under the 1.65x covenant."
    wsn.column_dimensions["A"].width = 44; wsn.column_dimensions["D"].width = 26; wsn.column_dimensions["G"].width = 18
    for w_ in wb.worksheets:
        style_ws(w_)
    wb.save(out_path)
    return r

if __name__ == "__main__":
    r = build(sys.argv[1] if len(sys.argv) > 1 else "solution.json", sys.argv[2] if len(sys.argv) > 2 else "../deliverables/Aurelia_Financial_Model.xlsx")
    print("engine NPV kEUR:", r["npv"] / 1000, "d:", r["d"], "capex kEUR:", sum(r["cf"]["capex"]) / 1000)
