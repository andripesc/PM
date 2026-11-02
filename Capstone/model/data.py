"""Load the case data from the instructor workbook (authoritative source)."""
import re, datetime as dt
import openpyxl

WB_PATH = __file__.rsplit("/", 1)[0] + "/Datacenter_Data_Updated.xlsx"

def _parse_assign(s):
    """'Architect, BIM Coordinator [25%]' -> [('Architect',1.0),('BIM Coordinator',0.25)]"""
    out = []
    if not s or s.strip() == "-":
        return out
    for part in s.split(","):
        part = part.strip()
        m = re.match(r"(.+?)\s*\[(\d+)%\]$", part)
        if m:
            out.append((m.group(1).strip(), int(m.group(2)) / 100.0))
        else:
            out.append((part, 1.0))
    return out

def _parse_mat(s):
    """'Power transformer [2 each], Installation materials [2,620 k€]' -> [('Power transformer',2),('Installation materials',2620)]
    (commas inside the brackets are thousands separators, so the string is parsed with a regex rather than split on commas)"""
    out = []
    if not s or s.strip() == "-":
        return out
    for name, qty in re.findall(r"([^,\[\]]+?)\s*\[([\d,\.]+)\s*[^\]]*\]", s):
        out.append((name.strip(), float(qty.replace(",", ""))))
    return out

def load():
    wb = openpyxl.load_workbook(WB_PATH, data_only=False)
    # ---- Activities
    ws = wb["Activity Dictionary"]
    acts = {}
    for r in range(5, 79):
        aid = ws.cell(r, 1).value
        if not aid:
            continue
        mind = ws.cell(r, 6).value
        prod = ws.cell(r, 10).value
        acts[aid] = dict(
            id=aid, area=ws.cell(r, 2).value, name=ws.cell(r, 3).value, loc=ws.cell(r, 4).value,
            dur=int(ws.cell(r, 5).value), mindur=(int(mind) if isinstance(mind, (int, float)) else None),
            work=_parse_assign(ws.cell(r, 7).value), mat=_parse_mat(ws.cell(r, 8).value),
            crash=[x[0] for x in _parse_assign(ws.cell(r, 9).value)],
            p=(float(prod) if isinstance(prod, (int, float)) else None), note=ws.cell(r, 11).value or "")
    # ---- Resources
    ws = wb["Resources"]
    res = {}
    for r in range(5, 53):
        name = ws.cell(r, 2).value
        if not name:
            continue
        mx = ws.cell(r, 3).value
        res[name] = dict(type=ws.cell(r, 1).value, name=name, max=(float(mx) if isinstance(mx, (int, float)) else None),
                         rate=float(ws.cell(r, 4).value), per=ws.cell(r, 5).value, label=ws.cell(r, 6).value,
                         accrue=ws.cell(r, 7).value, note=ws.cell(r, 8).value)
    # ---- Procurement packages
    ws = wb["Procurement and Logistics"]
    pk = {}
    for r in range(5, 13):
        name = ws.cell(r, 1).value
        pk[name] = dict(name=name, contents=ws.cell(r, 2).value, lots=int(ws.cell(r, 3).value),
                        value_total=float(ws.cell(r, 5).value), mfg_wk=float(ws.cell(r, 6).value),
                        exp_factor=float(ws.cell(r, 9).value), exp_premium=float(ws.cell(r, 10).value),
                        origin=ws.cell(r, 13).value, order_act=ws.cell(r, 14).value, spec_by=ws.cell(r, 15).value,
                        required_before=ws.cell(r, 16).value)
        pk[name]["lot_value"] = pk[name]["value_total"] / pk[name]["lots"]
    # shipping volumes
    items = {}
    for r in range(44, 54):
        pkg = ws.cell(r, 1).value
        items.setdefault(pkg, []).append(dict(item=ws.cell(r, 2).value, units=float(ws.cell(r, 3).value),
                                             vol=float(ws.cell(r, 4).value), wt=float(ws.cell(r, 5).value),
                                             oog=(ws.cell(r, 6).value == "Yes")))
    for pkg, lst in items.items():
        pk[pkg]["vol_lot"] = sum(i["units"] * i["vol"] for i in lst if not i["oog"])
        pk[pkg]["wt_lot"] = sum(i["units"] * i["wt"] for i in lst)
        pk[pkg]["oog_pieces"] = sum(int(i["units"]) for i in lst if i["oog"])
    # generators disassembled per lot
    gen_dis_vol = 4 * sum(float(ws.cell(r, 4).value) for r in range(58, 61))
    gen_dis_wt = 4 * sum(float(ws.cell(r, 5).value) for r in range(58, 61))
    transport = {}
    for r in range(26, 34):
        region, mode = ws.cell(r, 1).value, ws.cell(r, 2).value
        oog = ws.cell(r, 6).value
        transport[(region, mode)] = dict(region=region, mode=mode, cd=int(ws.cell(r, 3).value), rate=float(ws.cell(r, 4).value),
                                         oog_teu=(float(oog) if isinstance(oog, (int, float)) else None),
                                         air=("Air" in mode))
    # ---- Finance
    ws = wb["Revenue and Finance"]
    fin = dict(ntp=ws["B5"].value.date(), appraisal_end=ws["B6"].value.date(), hall_mw=float(ws["B8"].value),
               gpu_per_hall=float(ws["B9"].value), rate_gpuh=float(ws["B10"].value), hours=float(ws["B11"].value),
               elec_price=float(ws["B13"].value), pue=float(ws["B14"].value),
               oh_pmo=float(ws["B20"].value), oh_site=float(ws["B21"].value), oh_ops=float(ws["B22"].value))
    ws = wb["Financing Inputs"]
    fin.update(rd=float(ws["B21"].value) + float(ws["B22"].value), re=float(ws["B6"].value), dscr_min=float(ws["B7"].value),
               dscr_lo=float(ws["B8"].value), dscr_hi=float(ws["B9"].value), grace_end=ws["B10"].value.date(),
               maturity=ws["B11"].value.date())
    fin["deadline"] = dt.date(2029, 9, 11)  # EDATE(NTP, 32)
    return dict(acts=acts, res=res, pk=pk, transport=transport, fin=fin, gen_dis=dict(vol=gen_dis_vol, wt=gen_dis_wt))

if __name__ == "__main__":
    d = load()
    for a in d["acts"].values():
        print(a["id"], a["dur"], a["mindur"], a["work"], a["mat"], a["crash"], a["p"])
    for p in d["pk"].values():
        print(p)
    print(d["fin"]); print(d["gen_dis"]); print(d["transport"])
