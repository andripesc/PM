"""Generate the Program Charter, the Program Risk Management Plan (docx) and the Risk Register (xlsx)."""
import json, copy, sys, datetime as dt
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from engine import *
from links import derive_links
from optimize import unserial
from risks import RISKS, counts

def fmt_d(d): return d.strftime("%d %b %Y")
def k(x): return f"{x/1000:,.0f}"
def m(x): return f"{x/1e6:,.1f}"

# ------------------------------------------------------------------ docx helpers
def new_doc(landscape=False):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "Arial"; st.font.size = Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for s in ("Heading 1", "Heading 2", "Heading 3", "Title"):
        doc.styles[s].font.name = "Arial"; doc.styles[s].font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    doc.styles["Heading 1"].font.size = Pt(15); doc.styles["Heading 2"].font.size = Pt(12); doc.styles["Heading 3"].font.size = Pt(11)
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2); sec.top_margin = sec.bottom_margin = Cm(2)
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = sec.page_height, sec.page_width
    return doc

def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor); tcPr.append(shd)

def set_col_widths(t, widths_cm):
    t.autofit = False
    tblPr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout"); layout.set(qn("w:type"), "fixed"); tblPr.append(layout)
    grid = t._tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if i < len(widths_cm): gc.set(qn("w:w"), str(int(widths_cm[i] * 567)))
    for row in t.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)

from docx.enum.section import WD_SECTION
def landscape(doc):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    if sec.orientation != WD_ORIENT.LANDSCAPE:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin = sec.right_margin = Cm(1.5); sec.top_margin = sec.bottom_margin = Cm(1.5)
def portrait(doc):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    if sec.orientation != WD_ORIENT.PORTRAIT:
        sec.orientation = WD_ORIENT.PORTRAIT; sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin = sec.right_margin = Cm(2); sec.top_margin = sec.bottom_margin = Cm(2)

def table(doc, header, rows, widths=None, size=8, header_fill="1F3A5F"):
    t = doc.add_table(rows=1, cols=len(header)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; p = c.paragraphs[0]; r = p.add_run(str(h)); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); shade(c, header_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""; p = cells[i].paragraphs[0]; r = p.add_run("" if v is None else str(v)); r.font.size = Pt(size)
    if not widths:
        widths = [17.0 / len(header)] * len(header)
    set_col_widths(t, widths)
    doc.add_paragraph()
    return t

def para(doc, text, bold=False, italic=False, size=None, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    return p

def bullets(doc, items):
    for it in items:
        doc.add_paragraph(it, style="List Bullet")

def kv(doc, pairs, widths=(6, 11)):
    t = doc.add_table(rows=0, cols=2); t.style = "Table Grid"
    for a, b in pairs:
        cells = t.add_row().cells; cells[0].text = ""; cells[1].text = ""
        r = cells[0].paragraphs[0].add_run(a); r.bold = True; r.font.size = Pt(9); shade(cells[0], "E8EEF5")
        r = cells[1].paragraphs[0].add_run(str(b)); r.font.size = Pt(9)
    set_col_widths(t, list(widths))
    doc.add_paragraph()

# ------------------------------------------------------------------ analysis helpers
def whatif(dec, label, mutate):
    n = copy.deepcopy(dec); mutate(n)
    try:
        r = evaluate(n)
    except AssertionError:
        return None
    return dict(label=label, npv=r["npv"], hall_a=r["cf"]["hall_a"], hall_b=r["cf"]["hall_b"], fpa=r["cf"]["fpa"], d=r["d"], wacc=r["wacc"], capex=sum(r["cf"]["capex"]), feasible=r["feasible"])

def build_all(solution_path, out_dir):
    sol = json.load(open(solution_path)); dec = unserial(sol["decisions"])
    r = evaluate(dec); extra, _ = derive_links(dec, r)
    tasks, sched, dl, cf = r["tasks"], r["sched"], r["deliveries"], r["cf"]
    L = ledger(cf, r["d"])
    capex_tot = sum(cf["capex"])
    def act_cost(prefix): return sum(t["cost"] for tid, t in tasks.items() if tid.startswith(prefix) and t["base"] not in ("A05", "A06"))
    lots = sum(t["mat_cost"] for t in tasks.values() if t["base"] in ("A05", "A06"))
    exp_tot = sum(i["premium"] for i in dl.values()); tr_tot = sum(i["transport_cost"] for i in dl.values()); oh_tot = sum(cf["capex_oh"])
    crash_premium = 0.0
    crash_rows = []
    for aid, a in ACTS.items():
        if a["mindur"] and dec["crash"].get(aid, a["dur"]) < a["dur"]:
            d = dec["crash"][aid]; units = crash_units(a, d)
            nc = sum(u * 8 * RES[rn]["rate"] for rn, u in a["work"]) * a["dur"]; cc = sum(u * 8 * RES[rn]["rate"] for rn, u in units) * d
            crash_premium += cc - nc
            crash_rows.append([aid, a["name"], a["dur"], d, a["mindur"], ", ".join(f"{rn} {u*100:.0f}%" for rn, u in units if rn in a["crash"]), f"{(cc-nc)/1000:,.0f}",
                               {"F23": "Drives Hall A: first activity after transformer delivery", "F24": "Must finish before F23 monopolises the E&I crew (196%)", "F29": "Hall A fit-out before IT installation (A08)",
                                "F31": "Drives Hall B: first activity after transformer lot 2", "A08": "Hall A chain (compute installation)", "A10": "Hall B chain (compute installation)"}.get(aid, "")])
    # what-ifs for the decision log
    W = []
    W.append(whatif(dec, "Base case (selected plan)", lambda n: None))
    W.append(whatif(dec, "No expediting of power transformers (both lots)", lambda n: [po.update(expedite=False) for po in n["pos"]["Power transformers"]]))
    W.append(whatif(dec, "Transformers lot 2 not expedited (lot 1 only)", lambda n: n["pos"]["Power transformers"][1].update(expedite=False)))
    W.append(whatif(dec, "Generators lot 1 not expedited (sea transport)", lambda n: n["pos"]["Generators"][0].update(expedite=False, mode="Sea (short-sea ship)")))
    W.append(whatif(dec, "Expedite AI compute lot 1 (5% premium)", lambda n: n["pos"]["AI compute capacity"][0].update(expedite=True)))
    W.append(whatif(dec, "AI compute lot 1 by sea instead of air", lambda n: n["pos"]["AI compute capacity"][0].update(mode="Sea (container ship)")))
    W.append(whatif(dec, "No crashing at all (normal durations)", lambda n: n["crash"].clear()))
    W.append(whatif(dec, "F23 at normal duration (35 wd)", lambda n: n["crash"].pop("F23", None)))
    W.append(whatif(dec, "All crashable activities at minimum duration", lambda n: n["crash"].update({a: ACTS[a]["mindur"] for a in ACTS if ACTS[a]["mindur"]})))
    W.append(whatif(dec, "Generators shipped disassembled (lot 1, air)", lambda n: (n["gen_disassembled"].update({1: True}), n["pos"]["Generators"][0].update(mode="Air (freighter)"))))
    W.append(whatif(dec, "Compute / network lots ordered immediately after A04 (no JIT)", lambda n: n.update(jit=False)))
    W = [w for w in W if w]
    base = W[0]
    fin_rows = []
    for dshare in (0.0, 0.25, 0.40, r["d"], 0.55, 0.70):
        Lx = ledger(cf, dshare); mn = min((a["dscr"] for a in Lx["years"].values()), default=float("inf"))
        fin_rows.append([f"{dshare:.1%}", f"{wacc(dshare):.2%}", f"{npv(cf['fcff'], wacc(dshare))/1e6:,.1f}", ("n/a" if dshare == 0 else f"{mn:.2f}x"), "yes" if (dshare == 0 or mn >= FIN["dscr_min"]) else "NO (covenant)"])
    d15 = max_debt_share(cf, FIN["dscr_lo"]); d18 = max_debt_share(cf, FIN["dscr_hi"])

    # ============================================================== PROGRAM CHARTER
    doc = new_doc()
    doc.add_paragraph().add_run("Aurelia AI Infrastructure SpA").bold = True
    t = doc.add_paragraph(); rr = t.add_run("European AI Datacenter Conversion Program"); rr.bold = True; rr.font.size = Pt(20); rr.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    t = doc.add_paragraph(); rr = t.add_run("PROGRAM CHARTER"); rr.bold = True; rr.font.size = Pt(16)
    para(doc, "including the Project Delivery, Release and Contracting Strategy and the Procurement and Logistics Plan", italic=True)
    para(doc, "Project Management Course - Capstone Project - Fall Term, A.Y. 2026-27 - Politecnico di Torino", size=9)
    para(doc, "Team: [TEAM LEADER SURNAME] - [member names and IDs] - Version 1.0 - [date]", size=9)
    kv(doc, [("Program sponsor", "Aurelia AI Infrastructure SpA (Program Director: [name])"),
             ("Program", "Conversion of an existing ~18,000 m2 industrial building into a 20 MW, two-hall, liquid-cooled AI datacenter (P1 Facility Conversion, P2 AI Platform, P3 Operations Readiness)"),
             ("Notice to Proceed", fmt_d(NTP)), ("Hall A commercial service", fmt_d(cf["hall_a"])), ("Hall B commercial service", fmt_d(cf["hall_b"])),
             ("Final Program Acceptance", f"{fmt_d(cf['fpa'])} (deadline {fmt_d(FIN['deadline'])}, Month 32 - {(FIN['deadline']-cf['fpa']).days} days of program float)"),
             ("Total program CAPEX", f"{m(capex_tot)} MEUR (activities and equipment {m(capex_tot-exp_tot-tr_tot-oh_tot)} MEUR, expediting {m(exp_tot)} MEUR, transport {m(tr_tot)} MEUR, overheads {m(oh_tot)} MEUR)"),
             ("Financing", f"Debt {r['d']:.1%} / Equity {1-r['d']:.1%} of CAPEX (D = {m(L['D'])} MEUR), WACC {r['wacc']:.2%}, minimum annual DSCR {min(a['dscr'] for a in L['years'].values()):.2f}x (covenant 1.65x)"),
             ("Program NPV (base case)", f"{m(r['npv'])} MEUR at WACC {r['wacc']:.2%} (monthly mid-period discounting, horizon 31 Dec 2037)")])

    doc.add_heading("1. Purpose, background and objectives", 1)
    para(doc, "This Charter authorises the European AI Datacenter Conversion Program and records the integrated execution strategy selected by the team: how the three projects are delivered and contracted, how long-lead equipment is procured and shipped, how the schedule is accelerated, when each data hall enters commercial service and how the program is financed. All quantitative inputs are taken unchanged from the instructor workbook (Datacenter_Data_Updated.xlsx); every assumption not prescribed by the case is recorded in Section 10.")
    para(doc, "Aurelia, a proposed spin-off of a large Italian corporation, intends to enter the fast-growing European AI-infrastructure market (CBRE: ~13 GW of European datacenter capacity by end-2026; EU AI Gigafactories; IEA: datacenter electricity demand doubling to ~950 TWh by 2030). Converting an existing industrial building secures a site with an available grid connection faster than a greenfield development, but introduces brownfield uncertainty (structure, utilities, a single heavy-equipment corridor) that must be managed.")
    para(doc, "Program objectives (from the case):", bold=True)
    bullets(doc, ["Maximise reuse of the existing facility where technically feasible and economically justified.",
                  "Bring AI computing capacity into commercial service in the sequence and at the time that maximises value: Hall A first, Hall B second.",
                  "Coordinate engineering, procurement, construction, platform deployment and operational preparation across P1, P2 and P3 through one Integrated Master Schedule and one resource pool.",
                  "Select delivery and contracting strategies that allocate responsibilities and risks to the party best able to manage them.",
                  f"Maximise feasible program NPV subject to all technical, physical, resource, schedule (Final Acceptance by Month 32) and financial (DSCR >= 1.65x) constraints: result {m(r['npv'])} MEUR."])

    doc.add_heading("2. Strategy summary and key decisions", 1)
    para(doc, "The economics of the case are dominated by the time value of commercial service: each hall earns 4,096 GPU-eq x 8,760 h x 2.50 EUR = 89.7 MEUR of revenue per year (about 7.5 MEUR per month) against 8.9 MEUR of electricity; the first partial DSCR test year (Apr-Dec 2028) also determines how much debt the program can carry, hence the WACC. The strategy therefore (i) orders the critical long-lead power equipment at the earliest technically possible date and expedites the single binding package (power transformers), (ii) crashes only the activities on the Hall A and Hall B driving chains, (iii) orders the 160 MEUR compute lots just in time because they are paid at order placement, and (iv) sizes the debt share at the maximum that satisfies the covenant in every year.")
    table(doc, ["Decision area", "Decision", "Effect"], [
        ["Release strategy", "Two commercial releases: Hall A (Train 1) then Hall B (Train 2), each with its own functional tests; Whole-Plant IST and operations handover close the program.", f"Hall A {fmt_d(cf['hall_a'])}; Hall B {fmt_d(cf['hall_b'])}; Final Acceptance {fmt_d(cf['fpa'])}"],
        ["Power equipment ordering", "One ordering event (F08) on " + fmt_d(dl[('Power transformers', 1)]['po_date']) + " placing separate POs per lot for transformers, generators and switchgear/UPS, so that lot 1 ships as soon as it is ready (lot 2 follows on the same supplier line).", "Transformer lot 1 on site " + fmt_d(dl[('Power transformers', 1)]['delivery']) + "; lot 2 " + fmt_d(dl[('Power transformers', 2)]['delivery'])],
        ["Expediting", "Power transformers lots 1 and 2 (x0.84, 12%) and generators lot 1 (x0.85, 10%). All other packages at normal manufacturing time.", f"Expedite premiums {m(exp_tot)} MEUR; Hall A ~2 months earlier than without expediting (see decision log)"],
        ["Transport", "Road for out-of-gauge transformers and generators lot 1 (3 cd); short-sea for non-critical European lots (cheapest); air for compute lot 1 and rail for network lot 1 (time value of 160 MEUR / 25 MEUR lots); road for Hall B East-Asia lots.", f"Transport {m(tr_tot)} MEUR in total"],
        ["Generators", "Shipped assembled (out-of-gauge by road); disassembled shipping rejected (reassembly activity F46/F47 and extra cost without schedule benefit).", "F46/F47 not required"],
        ["Crashing", f"{len(crash_rows)} activities crashed on the driving chains (F23, F24, F29, F31, A08, A10).", f"Crash premium {m(crash_premium)} MEUR; Hall A about one month earlier"],
        ["AI compute / network POs", "Just-in-time: each PO placed so that the lot arrives shortly before the installing/testing activity (lots are paid at order placement).", "No idle capital; ~2 weeks of delivery float per lot"],
        ["Financing", f"Debt share {r['d']:.1%} of CAPEX (highest share with DSCR >= 1.65x in 2028 Apr-Dec and all later years); equal monthly principal Apr 2028 - Dec 2037.", f"WACC {r['wacc']:.2%}; NPV {m(r['npv'])} MEUR"],
    ], widths=[3.0, 9.6, 4.4])

    doc.add_heading("3. Scope and work breakdown structure", 1)
    para(doc, "The program is one integrated submission with three interdependent projects. The WBS below maps the case activity catalogue to projects and work packages; the same structure is used in the MS Project master plan, the cost model and the risk plan.")
    table(doc, ["WBS", "Work package", "Activities", "Main deliverable / handover"], [
        ["1", "P1 - Facility Conversion", "F01-F47", "Converted facility: shared infrastructure, two electrical trains, two cooling trains, two data halls, tested and handed over for IT installation"],
        ["1.1", "Engineering and design", "F01-F07, F09-F11, F13-F16", "Surveys, Basis of Design, 30% design and logistics plan, permits, tender designs, IFC packages"],
        ["1.2", "Facility equipment procurement and logistics", "F08, F12 (+ team-dated POs for busway sets and shared controls)", "POs, manufacturing, FAT, transport, delivery on site"],
        ["1.3", "Building conversion", "F17-F21", "Temporary facilities, isolation, demolition, heavy-equipment access route, structural reinforcement"],
        ["1.4-1.6", "Train 1 and Data Hall A", "F22-F29, F39-F41", "Power Train 1, Cooling Train 1, Hall A envelope / MEP / fit-out, functional tests"],
        ["1.7-1.9", "Train 2 and Data Hall B", "F30-F37, F42-F44", "Power Train 2, Cooling Train 2, Hall B, functional tests"],
        ["1.10", "Shared systems and IST", "F38, F45", "Controls, life-safety, security, network backbone; Whole-Plant Integrated Systems Test"],
        ["2", "P2 - AI Platform", "A01-A12", "Tested and integrated 8,192 GPU-eq platform ready for operational acceptance"],
        ["2.1-2.2", "Architecture, specification and procurement", "A01-A06", "Capacity confirmation, architecture, specifications, RFPs, compute and network/storage POs"],
        ["2.3-2.6", "Platform software, hall deployments, integration", "A07-A12", "Cluster management platform, Hall A/B installation and test, cross-hall integration"],
        ["3", "P3 - Operations Readiness", "O01-O15", "Operating organisation ready to assume responsibility for each hall"],
        ["3.1-3.2", "Organisation, procedures, recruitment", "O01-O08, O12", "Shift strategy, staffing plan, O&M procedures, core and incremental teams"],
        ["3.3-3.5", "Hall readiness and handover", "O09-O11, O13-O15", "Vendor training, readiness drills, operations handover and service acceptance"],
        ["0", "Program management (overheads)", "-", "PMO, general site services and operations-readiness management (time-dependent overheads, not resources)"],
    ], widths=[1.3, 4.4, 4.0, 7.3])
    para(doc, "Out of scope: site selection, upstream grid reinforcement, steady-state operations after Final Program Acceptance, taxation and terminal value (per the case).")

    landscape(doc)
    doc.add_heading("4. Governance and organisation (OBS / RACI)", 1)
    para(doc, "Aurelia retains program integration, acceptance and the core operating organisation. The Program Director chairs a weekly Program Integration Meeting with the three project managers, the Owner's Engineer, procurement and the commissioning manager; the PMO maintains the Integrated Master Schedule, the single resource pool, the procurement tracker and the risk register.")
    table(doc, ["Deliverable / decision", "Program Director (Aurelia)", "PMO / Master scheduler", "P1 PM (Facility)", "P2 PM (AI Platform)", "P3 PM (Operations)", "Owner's Engineer", "Procurement Manager", "Design-Build contractor", "Systems Integrator", "Commissioning Manager"], [
        ["Program Charter and strategy", "A", "R", "C", "C", "C", "C", "C", "I", "I", "I"],
        ["Basis of Design, 30% design, permits", "A", "I", "R", "C", "I", "R", "I", "C", "I", "C"],
        ["Long-lead equipment POs (F08, F12)", "A", "I", "C", "I", "I", "C", "R", "I", "I", "I"],
        ["Construction and installation (F17-F38)", "I", "C", "A", "I", "I", "C", "I", "R", "I", "C"],
        ["Facility tests and IST (F39-F45)", "A", "I", "R", "C", "C", "C", "I", "C", "C", "R"],
        ["Compute / network RFPs and POs (A04-A06)", "A", "I", "I", "R", "I", "I", "R", "I", "C", "I"],
        ["Platform installation, test, integration (A07-A12)", "I", "C", "C", "A", "C", "I", "I", "I", "R", "C"],
        ["Recruitment, procedures, training, drills (O01-O14)", "A", "I", "C", "C", "R", "I", "I", "I", "C", "C"],
        ["Commercial-service declaration (Hall A / B)", "A", "C", "R", "R", "R", "C", "I", "I", "I", "R"],
        ["Financing structure and covenant compliance", "A (with CFO)", "R", "I", "I", "I", "I", "I", "I", "I", "I"],
    ], widths=[5.2, 2.2, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0, 2.2, 2.0, 2.1], size=7)
    para(doc, "R = responsible, A = accountable, C = consulted, I = informed.", italic=True, size=8)
    portrait(doc)

    doc.add_heading("5. Project delivery, release and contracting strategy", 1)
    para(doc, "The 'Delivery Methods' sheet referred to in the briefing is not present in the workbook version received by the team; the delivery methods below are selected from the standard spectrum named in the briefing (Design-Bid-Build to Build-Operate-Transfer) and justified by the risk profile of each project. The choice must be confirmed against the instructor's sheet during the Q&A sessions (see Section 10).")
    doc.add_heading("5.1 P1 - Facility Conversion: Design-Build with Early Contractor Involvement and owner-procured long-lead equipment", 2)
    table(doc, ["Element", "Decision", "Rationale"], [
        ["Delivery method", "Two-stage Design-Build (ECI): stage 1 pre-construction services during F04-F06; stage 2 lump-sum / target-cost conversion contract after the 30% design. Long-lead power, cooling, busway and controls packages are procured by Aurelia (through the Owner's Engineer tender designs F07/F11) and free-issued to the Design-Build contractor for installation.", "Brownfield uncertainty and the single corridor require builder input in the logistics plan and design; owner procurement protects the critical lead times (transformers 50 weeks) from the contractor award cycle; a single installer keeps the P1 interface simple."],
        ["Design responsibility", "Aurelia's Owner's Engineer develops the Basis of Design, 30% design, tender designs and reviews vendor drawings; the Design-Build contractor completes IFC packages (F15/F16) with the owner's architect/BIM team and bears buildability risk; design changes through a formal change-control board chaired by the P1 PM.", "Keeps performance specification and system boundaries with the owner, buildability with the builder."],
        ["Contracting structure", "Aurelia holds: Owner's Engineer (multi-discipline), Design-Build contractor (civil, general, E&I, MEP, controls, network, security crews), equipment supply contracts (6 packages), commissioning agent (independent). Heavy-lift crew subcontracted to the Design-Build contractor.", "Independent commissioning avoids self-certification; owner-held supply contracts allow expediting and transport decisions to be taken by Aurelia."],
        ["Procurement and award strategy", "Tender designs F07/F11 released after the Basis of Design; POs placed at the end of F08 (" + fmt_d(dl[('Power transformers', 1)]['po_date']) + ") and F12 (" + fmt_d(dl[('Cooling train equipment', 1)]['po_date']) + "); busway sets ordered at the end of each hall's detailed design; shared controls after the 30% design. Design-Build stage 2 award after permits are filed.", "Earliest technically possible PO dates; award when design maturity supports a firm price."],
        ["Interface management", "Hall / train handover certificates (envelope, MEP, fit-out, tests) gate the start of IT installation; BIM clash detection before each IFC issue; corridor booking calendar in the logistics plan.", "Translates the Physical-Constraints logic into contractual gates."],
        ["Risk allocation", "Owner: ground/structural conditions above the surveyed baseline (provisional sums), equipment lead times, permits. Contractor: workmanship, buildability, crash-resource mobilisation, site safety, schedule with LDs on Train 1 / Hall A milestones. Vendors: FAT performance, delivery dates (LDs), warranty.", "Each risk sits with the party that controls it; see Risk Management Plan."],
    ], widths=[2.8, 8.4, 5.8])
    doc.add_heading("5.2 P2 - AI Platform: owner-procured technology with a Systems Integrator (EPCM-type) contract", 2)
    table(doc, ["Element", "Decision", "Rationale"], [
        ["Delivery method", "Aurelia procures the compute lots and the network/storage lots directly from the technology vendors (RFP A04, POs A05/A06); a Systems Integrator contract covers platform architecture support, cluster-management configuration (A07), installation (A08/A10 with the IT installation team), testing (A09/A11) and cross-hall integration (A12).", "Direct purchase gives Aurelia control of allocation, lead time, substitution and delivery flexibility for the 320 MEUR compute spend; one integrator carries integration risk."],
        ["Design responsibility", "Aurelia's HPC Architect and Network Engineer define the architecture and specifications (A02/A03); vendors are responsible for compliance; the integrator for the integrated design and commissioning.", "Keeps the service model with the owner."],
        ["Procurement and award strategy", "Single RFP round for both packages after A03; separate POs per hall placed just in time (A05-1/A06-1 for Hall A; A05-2/A06-2 for Hall B) with lead-time guarantees and LDs; no expediting (5-6% premium not justified).", "Lots are paid at order placement: ordering later saves cost of capital; a 2-week float protects the installation dates."],
        ["Interface management", "Hall readiness gate (F29 fit-out, F39/F40 train tests) before A08; hall-systems test (F41) before A09; platform training (O10) with the installed system; IST (F45) after A12.", "Prevents installation into an unsafe or untested hall."],
        ["Risk allocation", "Vendors: delivery date, performance, warranty, transit insurance. Integrator: integration, software configuration, test completion. Aurelia: technology selection and commercial (GPU-hour price) risk.", "Market and technology risk cannot be transferred economically."],
    ], widths=[2.8, 8.4, 5.8])
    doc.add_heading("5.3 P3 - Operations Readiness: in-house organisation with outsourced recruitment and vendor training", 2)
    table(doc, ["Element", "Decision", "Rationale"], [
        ["Delivery method", "Internal project led by the Operations Program Manager; recruitment agencies (third-party services in O04-O07, O12) under a recruitment-process-outsourcing agreement; vendor training (O09, O10, O13) contracted as part of the equipment and platform POs; managed-service option for platform operations kept as fallback.", "The core operating organisation must remain Aurelia's; outsourcing is limited to recruitment and training, clearly separated from construction contracts."],
        ["Design responsibility", "Aurelia defines organisation, staffing plan and competency matrix (O01/O02); procedures and runbooks by the Procedure & Training Lead with the facilities and platform leads (O03).", "Procedures are validated by the readiness drills before each service date."],
        ["Interface management", "Readiness drills (O11/O14) scheduled in the last week before each hall's commercial service; shift and vendor-escalation arrangements (O08) active before Hall A; handover (O15) after IST.", "Commercial-service conditions of the case."],
        ["Risk allocation", "Aurelia: recruitment lead times and skill shortages (mitigated by early start and agencies); vendors: training content and trainer availability (contractual dates).", "See risks R14-R16."],
    ], widths=[2.8, 8.4, 5.8])
    doc.add_heading("5.4 Release strategy", 2)
    para(doc, f"Hall A is released first because the single supplier production line delivers lot 1 of every package one manufacturing cycle before lot 2; Train 1 and Hall A activities are therefore prioritised in engineering (vendor-drawing reviews and detailed design of Train 1/Hall A are sequenced before Train 2/Hall B), in construction crews and in the single commissioning team. Hall A enters commercial service on {fmt_d(cf['hall_a'])}, Hall B on {fmt_d(cf['hall_b'])}; the Whole-Plant IST ({fmt_d(wd_date(sched['F45']['start']))} - {fmt_d(wd_date(sched['F45']['finish']))}) and the operations handover complete the program on {fmt_d(cf['fpa'])}, {(FIN['deadline']-cf['fpa']).days} days before the Month-32 deadline.")

    doc.add_heading("6. Integrated master schedule", 1)
    para(doc, "The Program Master Plan (MS Project file Aurelia_Program_Master_Plan) contains the 74 catalogue activities (order-placement activities instantiated per ordering event / PO), 15 procurement chains (manufacturing - transport - delivery), the program milestones, one Mon-Fri 8-hour calendar, the single resource pool with Max. Units, all logical links and the explicit resource-sequencing links; no automatic leveling and no task splitting are used. Activity costs are computed by MS Project from the assignments (work: units x 8 h x days x rate; material: quantity x unit price; all prorated).")
    table(doc, ["Milestone", "Date", "Last condition met"], [
        ["Notice to Proceed", fmt_d(NTP), "-"],
        ["Power equipment ordering event (end of F08)", fmt_d(dl[('Power transformers', 1)]['po_date']), "F07 power tender design"],
        ["Transformer lot 1 delivered on site", fmt_d(dl[('Power transformers', 1)]['delivery']), "Expedited manufacturing + road"],
        ["Hall A ready for IT installation (F29 fit-out, F39/F40 tests)", fmt_d(wd_date(max(sched['F29']['finish'], sched['F39']['finish'], sched['F40']['finish']))), "Hall A handover gate"],
        ["Hall A commercial service", fmt_d(cf["hall_a"]), f"O11 readiness drill ({fmt_d(wd_date(sched['O11']['finish']))}); F41 and A09 complete"],
        ["Transformer lot 2 delivered on site", fmt_d(dl[('Power transformers', 2)]['delivery']), "Second lot on the supplier line"],
        ["Hall B commercial service", fmt_d(cf["hall_b"]), f"O14 readiness drill ({fmt_d(wd_date(sched['O14']['finish']))}); F44 and A11 complete"],
        ["Whole-Plant Integrated Systems Test complete", fmt_d(wd_date(sched["F45"]["finish"])), "A12 cross-hall integration"],
        ["Final Program Acceptance", fmt_d(cf["fpa"]), f"O15 operations handover; deadline {fmt_d(FIN['deadline'])}"],
    ], widths=[6.5, 3.0, 7.5])
    doc.add_heading("6.1 Driving chains", 2)
    para(doc, f"Hall A: F08 ordering event -> transformer lot 1 manufacturing (42 weeks expedited) and road transport -> F23 Train 1 electrical installation (crashed to {tasks['F23']['dur']} wd) -> F39 Train 1 test -> A08 Hall A compute installation (crashed to {tasks['A08']['dur']} wd) -> O10 platform training -> O11 readiness drill -> commercial service. The hall civil chain (F21 reinforcement -> F27 -> F28 -> F29) and the cooling chain (F25 -> F26 -> F40) are completed before the transformer arrives, so they are not crashed beyond what the E&I/MEP crew limits require.")
    para(doc, f"Hall B: transformer lot 2 (second cycle on the supplier line, delivered {fmt_d(dl[('Power transformers', 2)]['delivery'])}) -> F31 (crashed to {tasks['F31']['dur']} wd) -> F42 -> A10 -> O13 -> O14 -> commercial service; then A11/A12 -> F45 IST -> O15 handover -> Final Program Acceptance.")
    doc.add_heading("6.2 Acceleration (crashing) decisions", 2)
    table(doc, ["ID", "Activity", "Normal [wd]", "Chosen [wd]", "Min [wd]", "Crash resources at chosen duration", "Extra cost [kEUR]", "Why"], [[c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7]] for c in crash_rows], widths=[1.0, 4.6, 1.2, 1.2, 1.0, 3.8, 1.3, 2.9], size=7)
    para(doc, f"Total crash premium {m(crash_premium)} MEUR. Crash units follow workbook note 4 (u x (1 + (Normal/d - 1)/p), rounded up to the whole percent) and are entered as fixed-duration tasks in MS Project. Activities with fixed durations (permits, PO placement, heavy-lift setting, tests, IST, training, drills) are not crashed. One-day crashes suggested by the search for F09/F10/F13/F14 and O05/O07/O08 were rejected: their +0.26 MEUR NPV arose only from shifting a compute-lot payment across a month boundary, not from an earlier service date.")
    doc.add_heading("6.3 Resource feasibility and sequencing", 2)
    para(doc, "Peak concurrent units of every resource respect the Max. Units of the pool (verified in the Financial Model, sheet Resources, and visible in MS Project Resource Usage). The binding resources are the E&I and MEP contractor crews (2 units each) around the Train 1 installation, the single Commissioning Team (which serialises F39, F40, F41, A09 and later the Train 2 tests), and the engineering teams (Owner's Engineer, BIM, Electrical/Mechanical Engineers, 2 units each) during the vendor-drawing reviews. Conflicts are resolved with explicit finish-to-start sequencing links (listed in the task notes of the MS Project file), giving priority to the Hall A chain:")
    bullets(doc, [f"{p} -> {s}" + (f" (+{lag} wd)" if lag else "") for p, s, lag in extra])
    para(doc, "Team assumption on the heavy-equipment corridor: at most two heavy-equipment installation fronts (one electrical, one cooling) use the corridor at the same time; demolition, access-route works and structural reinforcement are complete before any heavy equipment is delivered.")

    landscape(doc)
    doc.add_heading("7. Procurement and logistics plan", 1)
    para(doc, "Purchase-order quantities and dates, expediting and transport modes were chosen package by package against the supplier lead-time rule of the case (one production line per package; lot ready = max(PO date, previous lot ready) + manufacturing time; all lots of a PO ship together; delivery = ready + transport duration). Facility equipment is paid as installed through the material resources of the installing activity; compute and network lots are paid with their order-placement activity; expedite premiums at order placement; transport at delivery.")
    rows = []
    for (pkg, lot), info in sorted(dl.items(), key=lambda kv_: (kv_[1]["po_date"], kv_[0])):
        acts = NEEDS[pkg][lot]
        if pkg == "Generators" and info["disassembled"]: acts = ["F46" if lot == 1 else "F47"]
        req = wd_date(min(sched[a]["start"] for a in acts))
        rows.append([pkg, lot, info["po_event"], fmt_d(info["po_date"]), "yes" if info["expedite"] else "no", f"{info['mfg_cd']}", fmt_d(info["ready"]), info["mode"], f"{info['transport_cd']}", fmt_d(info["delivery"]), "/".join(acts) + " " + fmt_d(req), f"{(req-info['delivery']).days-1}", f"{info['teu']:.0f}", f"{info['transport_cost']/1000:,.1f}", f"{info['premium']/1000:,.0f}"])
    table(doc, ["Package", "Lot", "PO event", "PO date", "Exp.", "Mfg [cd]", "Ready", "Mode", "Tr. [cd]", "Delivered", "Required on site (activity, start)", "Float [cd]", "TEU", "Transport [kEUR]", "Premium [kEUR]"], rows,
          widths=[3.2, 0.8, 1.5, 2.0, 0.9, 1.1, 2.0, 2.4, 1.0, 2.0, 3.2, 1.0, 0.9, 1.5, 1.5], size=7)
    para(doc, "Notes:", bold=True)
    bullets(doc, [
        "Power packages (transformers, generators, switchgear/UPS): one ordering event F08 (10 wd after the power tender design) with six POs (one per package and lot). Separate POs let lot 1 pass FAT and ship alone; a combined two-lot PO would hold lot 1 at the factory for a second manufacturing cycle.",
        "Transformers are the binding package (50 weeks per lot): both POs expedited (x0.84 = 294 cd per lot, premium 12% = 816 kEUR per lot) and moved by road (out-of-gauge pieces: 2 x 8 TEU x 1,000 EUR; 3 cd). Generators lot 1 is expedited (620 kEUR) so that F24 (generators and UPS) is completed before the crashed F23 needs 196% of the E&I crew; generators lot 2 and all other packages arrive with float at normal lead time and by the cheapest mode (short-sea).",
        "Busway sets are ordered at the end of each hall's detailed design (F15/F16) and the shared controls package at the end of F05, i.e. at the earliest permitted date (paid as installed, no holding cost in the model); they arrive with large float.",
        "AI compute lots (160 MEUR each) and network/storage lots (25 MEUR each) are ordered just in time (team-chosen start-no-earlier-than dates on A05/A06) because they are paid at order placement: lot 1 of compute travels by air (19 TEU x 25 kEUR = 475 kEUR) and lot 1 of network by rail, which costs less than the cost of capital of ordering 35 or 20 days earlier; Hall B lots travel by road (17 cd), the cheapest mode consistent with the JIT date.",
        "Staging: outdoor equipment (transformers, generators, chillers/dry coolers) is staged in the laydown area and set by the heavy-lift crew through the corridor; indoor equipment (switchgear, UPS, CDUs, busway, racks) is delivered to the respective electrical room or hall after the envelope is closed. Deliveries are booked in the corridor calendar of the Construction Logistics Plan (F05).",
        "Generators disassembled option rejected: it removes the out-of-gauge charge (32 kEUR) and allows air freight, but adds the reassembly activity F46/F47 (5 wd, about 185 kEUR) with no schedule benefit because the generators are not on the driving chain.",
    ])

    portrait(doc)
    doc.add_heading("8. CAPEX and financial outcome", 1)
    table(doc, ["Item", "MEUR", "Basis"], [
        ["P1 activities incl. facility equipment and materials", m(act_cost("F")), "Resource assignments at scheduled durations (incl. crash units)"],
        ["P2 activities (excluding lots)", m(act_cost("A")), "Idem"],
        ["P3 activities", m(act_cost("O")), "Idem"],
        ["AI compute and network/storage lots", m(lots), "Paid with A05/A06 order-placement activities"],
        ["Expedite premiums", m(exp_tot), "Paid at PO"],
        ["Transport", m(tr_tot), "Paid at delivery"],
        ["Time-dependent overheads", m(oh_tot), f"630 kEUR/month prorated by calendar day to {fmt_d(cf['fpa'])}"],
        ["Total program CAPEX", m(capex_tot), "D + E"],
    ], widths=[7, 2.5, 7.5])
    para(doc, f"Unlevered FCFF: CAPEX from NTP, revenue and electricity from each commercial-service date to 31 Dec 2037, no tax, inflation or terminal value. Cumulative undiscounted FCFF at the horizon: {m(sum(cf['fcff']))} MEUR; payback (cumulative FCFF turns positive) in mid-2032.")
    doc.add_heading("8.1 Debt/equity optimisation and DSCR", 2)
    para(doc, "FCFF is identical for every financing mix; each mix is discounted at its own WACC = d x 6.0% + (1 - d) x 16.0%. Because debt is cheaper than equity, NPV rises with the debt share until the annual DSCR covenant binds; the first test year (Apr-Dec 2028) is the binding one because only Hall A is in service.")
    table(doc, ["Debt share d", "WACC", "NPV [MEUR]", "Min annual DSCR", "Feasible"], fin_rows, widths=[3, 3, 3, 3, 4])
    table(doc, ["Year", "CFADS [MEUR]", "Interest [MEUR]", "Principal [MEUR]", "Debt service [MEUR]", "DSCR"], [[y, m(a["cfads"]), m(a["interest"]), m(a["principal"]), m(a["ds"]), f"{a['dscr']:.2f}x"] for y, a in L["years"].items()], widths=[2, 3, 3, 3, 3, 2])
    para(doc, f"Selected structure: D = {r['d']:.1%} ({m(L['D'])} MEUR), E = {1-r['d']:.1%} ({m(capex_tot-L['D'])} MEUR); WACC {r['wacc']:.2%}; monthly draws pro rata to CAPEX; interest on the drawn balance ({m(sum(x['interest'] for x in L['rows'] if x['m'] < FIRST_REPAY_M))} MEUR of interest during the grace period funded by the sponsor, outside FCFF); equal monthly principal of {m(L['P'])} MEUR from Apr 2028 to Dec 2037 (117 instalments, fully repaid at maturity). Covenant sensitivities: at 1.50x the feasible debt share rises to {d15:.1%} (NPV {m(npv(cf['fcff'], wacc(d15)))} MEUR); at 1.80x it falls to {d18:.1%} (NPV {m(npv(cf['fcff'], wacc(d18)))} MEUR). Full monthly ledger and annual tests are in the Financial Model.")

    landscape(doc)
    doc.add_heading("9. Decision log", 1)
    para(doc, "Each line re-runs the integrated model with one decision changed; the deltas therefore include the knock-on effects on resource sequencing, service dates, feasible debt share and WACC.")
    table(doc, ["Alternative evaluated", "Hall A", "Hall B", "Final Acceptance", "Debt share", "WACC", "NPV [MEUR]", "Delta vs base [MEUR]", "Conclusion"],
          [[w["label"], fmt_d(w["hall_a"]), fmt_d(w["hall_b"]), fmt_d(w["fpa"]) + ("" if w["feasible"] else " (LATE)"), f"{w['d']:.1%}", f"{w['wacc']:.2%}", m(w["npv"]), ("-" if w is base else f"{(w['npv']-base['npv'])/1e6:+,.1f}"),
            ("Selected" if w is base else ("Rejected: lower NPV" if w["npv"] < base["npv"] - 1 else "Equivalent"))] for w in W], widths=[6.6, 2.4, 2.4, 2.8, 2.0, 2.0, 2.2, 2.4, 2.9], size=8)
    para(doc, "Search method: the team evaluated the plan with every crashable activity at minimum duration, every PO expedited and the fastest transport, then relaxed every decision that did not pay (un-crash, un-expedite, cheaper mode) and finally tested single-decision improvements, keeping the best feasible NPV; the model is a serial resource-constrained scheduler with the case lead-time rules and a monthly FCFF/DSCR engine reproduced formula-by-formula in the Financial Model. Key lessons: (1) the transformer lead time, not construction, sets the Hall A date; (2) the 2028 partial-year DSCR test links the Hall A date to the debt capacity (about 4 pp of debt share per month); (3) crashing is worth it only on the two transformer-driven chains; (4) lots paid at order must be ordered just in time, and fast transport is cheaper than early ordering for the 160 MEUR compute lot.")

    doc.add_heading("10. Assumptions register", 1)
    table(doc, ["#", "Assumption", "Rationale / source", "Impact if different"], [
        ["A1", "The workbook received ('Datacenter_Data_Updated.xlsx') does not contain the 'Physical Constraints' and 'Delivery Methods' sheets referred to in the briefing. Precedence, safety, test and commercial-release conditions were therefore derived from the activity notes, the briefing text and engineering practice (Annex A).", "Workbook contents; briefing sections 'Brownfield Construction Constraints', 'Installation and Integration Constraints', 'Commissioning, Commercial Service and Final Acceptance'.", "To be verified with the instructors; any prescribed condition not reproduced in Annex A must be added to the schedule (the model and plan are built so that links can be changed and the result recomputed)."],
        ["A2", "Hall A commercial service requires: F41 hall systems test, A09 compute test, O08 shift arrangements, O09/O10 training and the O11 readiness drill (same logic for Hall B with F44, A11, O13, O14). Revenue starts the calendar day after the last condition.", "Case: commercial-service conditions and revenue rule.", "-"],
        ["A3", "IT installation (A08/A10) may start when the hall fit-out (F29/F37) and both train tests (F39/F40, F42/F43) are complete; compute testing (A09/A11) requires the hall systems test (F41/F44), the cluster platform (A07) and the network/storage lot.", "Briefing: 'installation cannot begin until the required physical and safety conditions are satisfied'.", "If IT installation may start before the train tests, Hall A could be ~3 weeks earlier."],
        ["A4", "Shared systems (F38) require Train 1 electrical rooms (F22), the Hall A envelope (F27) and the controls/fibre lot; train and hall tests require F38.", "Briefing: halls operate independently once the shared infrastructure is installed.", "-"],
        ["A5", "Demolition (F19) requires the hazardous-material survey (F03), permits (F06) and isolation (F18); temporary facilities (F17) may start after the 30% design and logistics plan (F05), before permits.", "Activity notes; non-intrusive temporary works do not need the renovation permit.", "If F17 needs the permit, the hall civil chain shifts ~2 months but remains off the Hall A driving chain."],
        ["A6", "Tender designs (F07, F11) start after the Basis of Design (F04); vendor-drawing reviews after the PO; detailed hall design after F05 and the AI specifications (A03).", "Activity notes ('specifies the power equipment', 'vendor data define foundations').", "Earlier tender design would advance the transformer PO and Hall A."],
        ["A7", "Heavy-equipment corridor: at most two concurrent heavy-equipment installation fronts; demolition, access route and reinforcement complete before deliveries.", "Briefing (single corridor) and Heavy Lift Crew Max. Units (1 unit = four 25% assignments).", "A stricter one-at-a-time rule would delay Train 2 installation, not Hall A."],
        ["A8", "P3 sequence: organisation -> staffing plan -> procedures / recruitment of leaders -> core teams -> shift arrangements; training on installed systems; drill in the last week before service; incremental staff (O12) after the core teams.", "Activity notes.", "Large float; no effect on service dates."],
        ["A9", "Lead-time rounding: manufacturing days = round(weeks x 7 x expedite factor); lots served in lot order; delivery day = ready + transport days; installation starts the next working day.", "Case lead-time and expediting rules.", "-"],
        ["A10", "CFADS for the DSCR test = revenue - electricity (operating cash flow); CAPEX after April 2028 is financed by the continuing D/E draws; overheads are part of CAPEX.", "Financing Inputs: 'payments start in Apr 2028 even if Hall B CAPEX continues; account for simultaneous borrowing and repayment'.", "If CAPEX were deducted from CFADS, no debt would pass the 2028 test (CFADS negative) - the instructor's intent is clearly the operating definition."],
        ["A11", "Interest accrues monthly at rD/12 on the balance after the month's draw; discounting (1+WACC)^(-(m+0.5)/12) with m = 0 for Jan 2027.", "Financing Inputs (drawn-balance basis, monthly mid-period).", "Nominal-monthly discounting available as a switch in the model (NPV slightly lower)."],
        ["A12", "Delivery methods selected from the spectrum named in the briefing (Design-Build with ECI for P1, owner-procured technology with Systems Integrator for P2, in-house with RPO for P3).", "Briefing section 'Project Delivery and Contracting Strategy'; sheet not available.", "To be aligned with the Delivery Methods sheet conditions."],
        ["A13", "No cost is modelled for early-delivery storage, warranty erosion or working capital of facility equipment (paid as installed); these are treated qualitatively in the risk plan.", "Case: no quantitative data.", "Would favour later busway/controls orders; no effect on NPV ranking."],
    ], widths=[1.0, 9.0, 8.2, 7.5], size=7.5)

    portrait(doc)
    doc.add_heading("Annex A - Activity dependency logic (team assumptions)", 1)
    rows = []
    for p, s in r["links"]:
        if s in tasks and p in tasks:
            rows.append([p, s, "FS", ""])
    rationale = {"F02->F04": "BoD needs structural capacity", "F04->F07": "tender design after BoD", "F05->F17": "logistics plan before site establishment", "F06->F19": "permit before intrusive demolition", "F03->F19": "haz-mat survey before demolition", "F20->F21": "access route before reinforcement", "F09->F22": "vendor data define train 1 civils", "F22->F23": "rooms before equipment", "F29->A08": "fit-out before IT installation", "F39->A08": "tested power before IT installation", "F40->A08": "tested cooling before IT installation", "F41->A09": "hall systems tested before compute test", "A07->A09": "cluster platform before compute test", "F24->O09": "training on installed systems", "A08->O10": "training on installed platform", "A12->F45": "IST needs integrated platform", "F45->O15": "handover after IST", "F09->F10": "sequencing policy: Train 1 first", "F13->F14": "sequencing policy: Train 1 first", "F15->F16": "sequencing policy: Hall A first"}
    for row in rows:
        row[3] = rationale.get(f"{row[0]}->{row[1]}", "")
    table(doc, ["Predecessor", "Successor", "Type", "Rationale (where not self-evident)"], rows, widths=[2.5, 2.5, 1.2, 10.8], size=7.5)
    para(doc, "Delivery links: each 'Delivered on site' milestone precedes its installing activity (transformers -> F23/F31; generators -> F24/F32; switchgear/UPS -> F23, F24 / F31, F32; cooling -> F26/F34; busway -> F28/F36; controls -> F38; compute -> A08/A10; network/storage -> A09/A11).", size=8)
    doc.save(f"{out_dir}/Aurelia_Program_Charter.docx")

    # ============================================================== RISK MANAGEMENT PLAN
    tot, hi, rhi = counts()
    doc = new_doc()
    doc.add_paragraph().add_run("Aurelia AI Infrastructure SpA - European AI Datacenter Conversion Program").bold = True
    t = doc.add_paragraph(); rr = t.add_run("PROGRAM RISK MANAGEMENT PLAN"); rr.bold = True; rr.font.size = Pt(18); rr.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    para(doc, "Summary document - details in Aurelia_Risk_Register.xlsx", italic=True)
    para(doc, "Team: [TEAM LEADER SURNAME] - Version 1.0 - [date]", size=9)
    doc.add_heading("1. Purpose and approach", 1)
    para(doc, "This plan identifies, assesses and treats the risks of the three projects and of their interdependencies, consistently with the Program Charter, the Integrated Master Schedule, the contracting strategy and the Financial Model. Risks were identified by work package and interface (brownfield conditions, procurement and logistics, installation prerequisites, commissioning, recruitment and training, resources shared across projects, financing), using the case description, the activity notes and the sensitivities of the integrated model.")
    para(doc, "Assessment scale: probability 1 (rare, <10%) to 5 (almost certain, >70%); impact 1 (negligible: <1 week / <0.5 MEUR) to 5 (severe: >2 months on a commercial-service date, >20 MEUR of NPV, or a safety fatality). Exposure = probability x impact: Low 1-4, Medium 5-9, High 10-15, Critical 16-25. Each risk has an owner, a response strategy (avoid, mitigate, transfer, accept), mitigation measures and a residual assessment after mitigation. Schedule and cost assumptions associated with material exposures are stated per risk.")
    doc.add_heading("2. Summary of results", 1)
    kv(doc, [("Total number of risks identified", str(tot)), ("Risks initially High or Critical", f"{hi} ({', '.join(x['id'] for x in RISKS if x['rating'] in ('High','Critical'))})"),
             ("Risks remaining High or Critical after mitigation", f"{rhi} ({', '.join(x['id'] for x in RISKS if x['rrating'] in ('High','Critical')) or 'none'})"),
             ("Program schedule buffer", f"Final Program Acceptance {fmt_d(cf['fpa'])} vs deadline {fmt_d(FIN['deadline'])}: {(FIN['deadline']-cf['fpa']).days} calendar days of float at program level (the Hall A chain has none: every week of delay costs about 1.7 MEUR of revenue and reduces the feasible debt share)."),
             ("Cost assumptions", "Provisional sums and contingencies recommended in the register (not included in the base-case NPV, which uses case data unchanged): structural strengthening 1.3 MEUR, hazardous materials 0.3 MEUR, re-testing 0.25 MEUR, integration 7.4 MEUR, temporary test power 0.2 MEUR; expected monetary value of the register about " + f"{sum(x['p']/5*x['i']*2 for x in RISKS):.0f}" + " MEUR using a 2 MEUR-per-impact-point convention."),
             ("Financial risk exposure", f"Selling-rate -10%: NPV about -51 MEUR; Hall A one month later: NPV about -9 MEUR and debt share -4 pp; covenant 1.80x instead of 1.65x: NPV -{(r['npv']-npv(cf['fcff'], wacc(d18)))/1e6:,.1f} MEUR (see Financial Model, Sensitivity and Financing).")])
    landscape(doc)
    doc.add_heading("3. Risk profile by project", 1)
    for proj, title in (("P1", "P1 - Facility Conversion"), ("P2", "P2 - AI Platform"), ("P3", "P3 - Operations Readiness"), ("Program", "Program-level and interdependency risks")):
        doc.add_heading(title, 2)
        table(doc, ["ID", "Risk", "Cause / consequence", "P", "I", "Exposure", "Owner / strategy", "Mitigation measures", "Residual"], [[x["id"], x["title"], x["cause"] + " -> " + x["consequence"], x["p"], x["i"], f"{x['score']} {x['rating']}", x["owner"] + " / " + x["strategy"], x["mitigation"], f"{x['rscore']} {x['rrating']}"] for x in RISKS if x["project"] == proj],
              widths=[1.0, 3.4, 6.2, 0.6, 0.6, 1.6, 3.0, 7.6, 1.7], size=7)
    portrait(doc)
    doc.add_heading("4. Top risks and their treatment", 1)
    for x in sorted(RISKS, key=lambda z: -z["score"])[:6]:
        doc.add_heading(f"{x['id']} - {x['title']} ({x['rating']} -> {x['rrating']})", 3)
        kv(doc, [("Cause", x["cause"]), ("Consequence", x["consequence"]), ("Owner / strategy", f"{x['owner']} / {x['strategy']}"), ("Mitigation", x["mitigation"]), ("Schedule assumption", x["sched"]), ("Cost assumption", x["cost"])], widths=(3.5, 13.5))
    doc.add_heading("5. Consistency with schedule, contracts and financial model", 1)
    bullets(doc, [
        "Schedule: the gates assumed in the master schedule (hall handover before IT installation, tests before commercial service, drills in the last week) are the controls for R06, R13 and R15; the explicit sequencing links and the Train 1 priority rule address R18.",
        "Contracts: owner-held equipment contracts with expediting, LDs and witnessed FAT (R02, R09, R11); Design-Build contractor with provisional sums for brownfield conditions and crash-resource mobilisation clauses (R01, R07, R08); systems-integrator single-point responsibility (R10); recruitment-process outsourcing and training dates in the POs (R14, R15).",
        "Financial model: debt share selected with the covenant exactly met in 2028; the register quantifies how Hall A delay, selling rate and covenant level change the debt capacity and NPV (R12, R17, R20, R23); contingencies listed above are recommended on top of the base-case CAPEX.",
        "Monitoring: the register is reviewed weekly by the PMO; triggers (factory progress reports, recruitment pipeline, test results) are tracked against the dates in the Procurement plan and the schedule."])
    doc.save(f"{out_dir}/Aurelia_Program_Risk_Management_Plan.docx")

    # ============================================================== RISK REGISTER XLSX
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook(); ws = wb.active; ws.title = "Risk Register"
    ws["A1"] = "PROGRAM RISK REGISTER - Aurelia European AI Datacenter Conversion Program"; ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = "Scale: probability 1-5, impact 1-5; exposure = P x I; Low 1-4, Medium 5-9, High 10-15, Critical 16-25 (thresholds on sheet Scales). Exposure, rating and summary counts are formulas."
    hdr = ["ID", "Project", "Category", "Risk", "Cause", "Consequence", "Probability (1-5)", "Impact (1-5)", "Exposure", "Rating", "Owner", "Response strategy", "Mitigation measures", "Residual probability", "Residual impact", "Residual exposure", "Residual rating", "Schedule assumption", "Cost assumption"]
    ws.append([]); ws.append(hdr)
    for c in ws[4]: c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3A5F"); c.alignment = Alignment(wrap_text=True, vertical="top")
    rr = 5
    for x in RISKS:
        ws.append([x["id"], x["project"], x["category"], x["title"], x["cause"], x["consequence"], x["p"], x["i"], f"=G{rr}*H{rr}", f'=IF(I{rr}>=Scales!$B$5,"Critical",IF(I{rr}>=Scales!$B$4,"High",IF(I{rr}>=Scales!$B$3,"Medium","Low")))',
                   x["owner"], x["strategy"], x["mitigation"], x["rp"], x["ri"], f"=N{rr}*O{rr}", f'=IF(P{rr}>=Scales!$B$5,"Critical",IF(P{rr}>=Scales!$B$4,"High",IF(P{rr}>=Scales!$B$3,"Medium","Low")))', x["sched"], x["cost"]])
        rr += 1
    last = rr - 1
    for row in ws.iter_rows(min_row=5, max_row=last):
        for c in row:
            c.font = Font(name="Arial", size=9); c.alignment = Alignment(wrap_text=True, vertical="top")
        for ci in (6, 7, 13, 14): row[ci].font = Font(name="Arial", size=9, color="0000FF")
    widths = [6, 8, 18, 34, 40, 40, 10, 8, 9, 9, 24, 14, 60, 10, 10, 10, 10, 36, 36]
    for i, w in enumerate(widths, 1): ws.column_dimensions[L_(i)].width = w
    ws.freeze_panes = "E5"
    rr = last + 2
    ws.cell(rr, 1, "SUMMARY").font = Font(name="Arial", bold=True)
    ws.cell(rr + 1, 1, "Total number of risks identified"); ws.cell(rr + 1, 4, f"=COUNTA(A5:A{last})")
    ws.cell(rr + 2, 1, "Risks initially High or Critical"); ws.cell(rr + 2, 4, f'=COUNTIF(J5:J{last},"High")+COUNTIF(J5:J{last},"Critical")')
    ws.cell(rr + 3, 1, "Risks remaining High or Critical after mitigation"); ws.cell(rr + 3, 4, f'=COUNTIF(Q5:Q{last},"High")+COUNTIF(Q5:Q{last},"Critical")')
    ws.cell(rr + 4, 1, "Initial exposure by rating: Critical / High / Medium / Low"); ws.cell(rr + 4, 4, f'=COUNTIF(J5:J{last},"Critical")'); ws.cell(rr + 4, 5, f'=COUNTIF(J5:J{last},"High")'); ws.cell(rr + 4, 6, f'=COUNTIF(J5:J{last},"Medium")'); ws.cell(rr + 4, 7, f'=COUNTIF(J5:J{last},"Low")')
    ws.cell(rr + 5, 1, "Residual exposure by rating: Critical / High / Medium / Low"); ws.cell(rr + 5, 4, f'=COUNTIF(Q5:Q{last},"Critical")'); ws.cell(rr + 5, 5, f'=COUNTIF(Q5:Q{last},"High")'); ws.cell(rr + 5, 6, f'=COUNTIF(Q5:Q{last},"Medium")'); ws.cell(rr + 5, 7, f'=COUNTIF(Q5:Q{last},"Low")')
    ws.cell(rr + 6, 1, "Sum of initial exposure scores / residual"); ws.cell(rr + 6, 4, f"=SUM(I5:I{last})"); ws.cell(rr + 6, 5, f"=SUM(P5:P{last})")
    for r_ in range(rr, rr + 7):
        for c_ in range(1, 8): ws.cell(r_, c_).font = Font(name="Arial", size=10, bold=(r_ == rr))
    w2 = wb.create_sheet("Scales")
    w2.append(["Rating", "Minimum exposure"]); w2.append(["Low", 1]); w2.append(["Medium", 5]); w2.append(["High", 10]); w2.append(["Critical", 16])
    w2.append([]); w2.append(["Probability", "Meaning"]); w2.append([1, "Rare (<10%)"]); w2.append([2, "Unlikely (10-30%)"]); w2.append([3, "Possible (30-50%)"]); w2.append([4, "Likely (50-70%)"]); w2.append([5, "Almost certain (>70%)"])
    w2.append([]); w2.append(["Impact", "Meaning (schedule / cost)"]); w2.append([1, "< 1 week / < 0.5 MEUR"]); w2.append([2, "1-2 weeks / 0.5-2 MEUR"]); w2.append([3, "2-4 weeks on a service date / 2-5 MEUR"]); w2.append([4, "1-2 months / 5-20 MEUR NPV"]); w2.append([5, "> 2 months / > 20 MEUR NPV or a fatality"])
    for row in w2.iter_rows():
        for c in row: c.font = Font(name="Arial", size=10)
    for c in (w2["B2"], w2["B3"], w2["B4"], w2["B5"]): c.font = Font(name="Arial", size=10, color="0000FF")
    w2.column_dimensions["A"].width = 14; w2.column_dimensions["B"].width = 44
    w3 = wb.create_sheet("Matrix")
    w3["A1"] = "Risk matrix (count of risks, initial -> residual)"; w3["A1"].font = Font(name="Arial", bold=True)
    w3["A2"] = "Impact \\ Probability"
    for p_ in range(1, 6): w3.cell(2, 1 + p_, p_).font = Font(name="Arial", bold=True)
    for i_ in range(5, 0, -1):
        rrow = 3 + (5 - i_); w3.cell(rrow, 1, i_).font = Font(name="Arial", bold=True)
        for p_ in range(1, 6):
            w3.cell(rrow, 1 + p_, f'=COUNTIFS(\'Risk Register\'!$G$5:$G${last},{p_},\'Risk Register\'!$H$5:$H${last},{i_})&" -> "&COUNTIFS(\'Risk Register\'!$N$5:$N${last},{p_},\'Risk Register\'!$O$5:$O${last},{i_})')
            sc = p_ * i_
            w3.cell(rrow, 1 + p_).fill = PatternFill("solid", fgColor=("FF9999" if sc >= 16 else "FFCC99" if sc >= 10 else "FFFF99" if sc >= 5 else "CCFFCC"))
    for row in w3.iter_rows():
        for c in row:
            if not c.font.bold: c.font = Font(name="Arial", size=10)
    wb.save(f"{out_dir}/Aurelia_Risk_Register.xlsx")
    return r

def L_(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)

if __name__ == "__main__":
    build_all(sys.argv[1] if len(sys.argv) > 1 else "solution.json", sys.argv[2] if len(sys.argv) > 2 else "../deliverables")
    print("docs written")
