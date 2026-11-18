import csv, sys, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

SRC, OUT, SRC2, SRC3, SBCF = sys.argv[1:6]
SNAPSHOT = datetime.datetime(2026, 11, 18)
FY_START, FY_END = 2025, 2029
LOOKBACKS = [0, -7, -14, -30, -60, -90, -120, -180, -365]

COMPANIES = [  # VA code, company name, currency
    ("KLAC", "KLA Corporation", "USD"), ("LRCX", "Lam Research", "USD"),
    ("ASML", "ASML Holding", "EUR"), ("CSCO", "Cisco Systems", "USD"),
    ("AMAT", "Applied Materials", "USD"), ("000660_KR", "SK Hynix", "KRW"),
    ("MU", "Micron Technology", "USD"), ("AVGO", "Broadcom", "USD"),
    ("2330_TW", "TSMC", "TWD"), ("PLTR_US", "Palantir Technologies", "USD"),
    ("COST", "Costco Wholesale", "USD"), ("WMT", "Walmart", "USD"),
    ("SPCX_US", "Space Exploration Technologies", "USD"), ("TSLA", "Tesla", "USD"),
    ("SHOP_US", "Shopify", "USD"), ("ARM_GB", "Arm Holdings", "USD"),
    ("CRWD_US", "CrowdStrike", "USD"), ("PANW", "Palo Alto Networks", "USD"),
    ("INTC", "Intel", "USD"), ("MNST", "Monster Beverage", "USD"),
    ("WELL_US", "Welltower", "USD"), ("BE_US", "Bloom Energy", "USD"),
    ("NET_US", "Cloudflare", "USD"),
    ("NVDA", "NVIDIA", "USD"),
]
# Other EPS / revenue lines seen in the VA screenshots, for review
NEARBY = {
    "ASML": "EPS - Basic | EPS - Diluted (Analyst published) | Sales - System sales",
    "000660_KR": "EPS - Basic | EPS excl. treasury | EPS incl. treasury | EPS excl. treasury - Applicable to common stockholders",
    "AVGO": "Revenue - Operating | EPS - Diluted incl. SBC and excl. others | EPS - Diluted-Applicable to common stockholders",
    "2330_TW": "Sales revenue, net ($) | EPS - Diluted - ADR",
    "COST": "Net sales | EPS Basic",
    "SPCX_US": "EPS - Diluted applicable to common stockholders",
    "ARM_GB": "EPS - Diluted - Continuing operations | EPS - Diluted - Continuing operations - Operating",
    "INTC": "Net revenue - Operating | EPS attributable to common shareholders - Basic",
    "MNST": "Gross revenue",
    "WELL_US": "Resident fees and services | FFO per share-Diluted | Normalized FAD per share-Diluted",
    "KLAC": "EPS - Diluted - Pre Split | EPS - Diluted - Operating - Pre Split",
    "MU": "EPS to common stock - Diluted incl. Sbc and excl. others",
}

def plain(code):
    return code.split("_")[0]

tick_id = {c[0]: i + 1 for i, c in enumerate(COMPANIES)}
name = {c[0]: c[1] for c in COMPANIES}
ccy = {c[0]: c[2] for c in COMPANIES}
def _rows(p):
    rs = [r for r in csv.DictReader(open(p))]
    rs.sort(key=lambda r: tick_id[r["Ticker"]])
    return rs
_base, _marg, _cf = _rows(SRC), _rows(SRC2), _rows(SRC3)
# keep existing Item_IDs: original lines first, then NVIDIA, then cash flow
items = ([r for r in _base if r["Ticker"] != "NVDA"] + [r for r in _marg if r["Ticker"] != "NVDA"] +
         [r for r in _base if r["Ticker"] == "NVDA"] + [r for r in _marg if r["Ticker"] == "NVDA"] + _cf)
SBC_MANUAL = [r for r in csv.DictReader(open(SBCF))]
assert {r["Ticker"] for r in items} == set(tick_id)

PERIODS = []
for fy in range(FY_START, FY_END + 1):
    for q in range(1, 5):
        PERIODS.append((f"{q}QFY-{fy}", "Quarterly", fy, f"Q{q}"))
    PERIODS.append((f"FY-{fy}", "Annual", fy, "FY"))
NP = len(PERIODS)

HFILL = PatternFill("solid", fgColor="1F4E78")
HFONT = Font(bold=True, color="FFFFFF", size=10)
BFONT = Font(bold=True, color="1F1F1F", size=10)
INFILL = PatternFill("solid", fgColor="FFF2CC")
DATE = r"dd\-mmm\-yyyy"
EPSF = r"#,##0.0000;[Red]\(#,##0.0000\);\-"
REVF = r"#,##0.0;[Red]\(#,##0.0\);\-"
PCT = r"0.0%;[Red]\(0.0%\);\-"

def sheet(wb, title, headers, widths, freeze):
    ws = wb.create_sheet(title)
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h); c.fill, c.font = HFILL, HFONT
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = freeze
    return ws

wb = Workbook()
ctl = wb.active; ctl.title = "Control"
wb.remove(ctl)

# ---------- build row maps first ----------
n_items = len(items)
n_act = n_items * NP
n_con = n_act * len(LOOKBACKS)

# Control
ctl = wb.create_sheet("Control")
ctl.sheet_view.showGridLines = False
ctl.column_dimensions["A"].width = 31; ctl.column_dimensions["B"].width = 98
ctl["A2"] = "Visible Alpha revenue and EPS line-item pull"; ctl["A2"].font = Font(bold=True, size=16, color="1F1F1F")
for c in ("A4", "B4", "A19", "B19"):
    ctl[c].fill, ctl[c].font = HFILL, HFONT
ctl["A4"], ctl["B4"] = "Setting", "Value"
settings = [
    ("Purpose", "Refresh revenue, GAAP and operating diluted EPS, gross profit, operating income, operating cash flow, capex and SBC line items across actuals, dated consensus and earnings beat/miss; margins calculated on the Margins sheet."),
    ("Snapshot date", SNAPSHOT),
    ("Requested line items", n_items),
    ("Companies", len(COMPANIES)),
    ("Exact DB matches", "Not checked"),
    ("Not in DuckDB", "Not checked"),
    ("Fiscal periods per line", NP),
    ("Period range", f"1QFY-{FY_START} through FY-{FY_END}, all quarters and annual periods"),
    ("Consensus lookbacks", "0, 7, 14, 30, 60, 90, 120, 180, 365 days before anchor"),
    ("Actual rows", n_act),
    ("Consensus rows", n_con),
    ("Beat/miss rows", n_act),
]
for i, (k, v) in enumerate(settings, 5):
    ctl.cell(i, 1, k).font = BFONT
    c = ctl.cell(i, 2, v); c.font = BFONT; c.alignment = Alignment(horizontal="left")
ctl["B6"].fill = INFILL; ctl["B6"].number_format = DATE
ctl["A19"], ctl["B19"] = "Method", "Treatment"
method = [
    ("Refresh", "Open in Excel with the Visible Alpha add-in, recalculate fully, and save."),
    ("Actuals", "Pulled for review even when the VA earnings date is scheduled or stale. Confirm release dates before any import."),
    ("Historical consensus", "Anchored to the earnings date, then reduced by the selected lookback days."),
    ("Future consensus", "Anchored to the snapshot date in B6, then reduced by the selected lookback days."),
    ("Beat/miss", "Provisional until the actual earnings release date is confirmed; actual less consensus seven days before that date."),
    ("Date review", "VA earnings dates may be estimates. Enter the confirmed date in Date_Override and its evidence source before import. No result here is import-approved."),
    ("Report dates", "One row per company and period. Every line item for that company uses the same date."),
    ("Units", "Native VA scale and currency: ASML in EUR m, TSMC in TWD m, SK Hynix in KRW bn, all others USD m. EPS per share in the same currency. Welltower operating line is normalized FFO per share."),
    ("Margins", "Margins sheet = gross profit and operating income from the Actuals sheet divided by revenue for the same company and period. Reported where Value_Date is filled; otherwise the VA pull is consensus. Operating versions use VA '- Operating' (adjusted) lines where they exist. Welltower has no gross profit or operating income line: its gross margin column is net operating income / property revenues (NOI margin)."),
    ("Cash flow", "Cash_Flow sheet: FCF = operating cash flow - ABS(capex), so it is right whatever sign VA shows capex with (Lam shows it negative). FCF after SBC also subtracts ABS(stock-based compensation). FCF incl. software/intangibles also subtracts capitalised software or intangibles purchases where VA has that line (ASML, Arm, CrowdStrike, Shopify, Cloudflare). SBC comes from the VA cash flow line where one exists; otherwise from the Manual_SBC sheet (company filings for reported years, our revenue-share projection for forecast years, labelled). Monster Beverage cash flow lines not yet supplied."),
    ("Fiscal calendars", "Period labels follow each company's own fiscal year, so the same label is not the same calendar quarter across companies."),
]
for i, (k, v) in enumerate(method, 20):
    ctl.cell(i, 1, k).font = BFONT
    c = ctl.cell(i, 2, v); c.font = BFONT; c.alignment = Alignment(wrap_text=True, vertical="top")

# DB_Review
dbr = sheet(wb, "DB_Review", ["Item_ID", "Ticker", "VA_Ticker_Code", "Company_Name", "Line_Item", "DB_Status",
            "DB_Actual_Periods", "DB_Consensus_Periods", "Actual_Years_in_DB", "Consensus_Years_in_DB",
            "DB_Metric_ID", "Nearby_EPS_Lines", "Dividend_Related_Names"],
            {"A": 10, "B": 18, "D": 27, "E": 62, "F": 22, "I": 38, "K": 37, "L": 60}, "F2")
for i, it in enumerate(items, 1):
    t = it["Ticker"]
    for j, v in enumerate([i, plain(t), t, name[t], it["Line Item"], "Not checked", None, None, None, None, None,
                           NEARBY.get(t)], 1):
        if v is not None:
            dbr.cell(i + 1, j, v)
dbr.conditional_formatting.add(f"F2:F{n_items+1}", Rule(type="containsText", operator="containsText", text="Not checked",
    formula=[f'NOT(ISERROR(SEARCH("Not checked",F2)))'],
    dxf=DifferentialStyle(fill=PatternFill("solid", bgColor="FFF2CC", fgColor="FFF2CC"))))

# Report_Dates (per company x period)
rd = sheet(wb, "Report_Dates", ["Ticker", "VA_Ticker_Code", "Period_Label", "DB_Reported_Date", "VA_Earnings_Date",
           "Date_Override", "Effective_Reported_Date", "Date_Status", "Date_Evidence_Source"],
           {"A": 19, "D": 23, "I": 70}, "D2")
rd_row = {}
r = 2
for code, _, _ in COMPANIES:
    for pi, (p, *_rest) in enumerate(PERIODS):
        rd_row[(code, pi)] = r
        rd.cell(r, 1, plain(code)); rd.cell(r, 2, code); rd.cell(r, 3, p)
        rd.cell(r, 5, f'=_xll.VADetail("EarningsDate",$B{r},$C{r})')
        rd.cell(r, 7, f'=IF(ISNUMBER($F{r}),$F{r},IF(IFERROR(ISNUMBER($E{r}),FALSE),$E{r},IF(ISNUMBER($D{r}),$D{r},"")))')
        rd.cell(r, 8, f'=IF(AND(ISNUMBER($F{r}),$I{r}<>""),"Confirmed override",IF(ISNUMBER($G{r}),IF($G{r}<=Control!$B$6,"Past date / unverified","Scheduled / future"),"Date unavailable"))')
        for col in (4, 5, 6, 7):
            rd.cell(r, col).number_format = DATE
        for col in (6, 9):
            rd.cell(r, col).fill = INFILL
        r += 1

COMMON = ["Observation_Key", "Ticker", "VA_Ticker_Code", "Company_Name", "Line_Item", "Period_Label",
          "Period_Type", "Fiscal_Year", "Quarter", "VA_Data_Type", "Currency"]
WID = {"A": 28, "B": 18, "D": 28, "E": 66, "F": 18}
act = sheet(wb, "Actuals", COMMON + ["Reported_Date", "Actual_Value", "Value_Date"], WID, "F2")
con = sheet(wb, "Consensus", COMMON + ["Snapshot_Date", "Reported_Date", "Anchor_Date", "Lookback_Days", "As_Of_Date", "Consensus_Value"],
            dict(WID, A=30), "L2")
bm = sheet(wb, "Beat_Miss", COMMON + ["Reported_Date", "Actual_Value", "Consensus_7D_Pre_Report", "Surprise", "Surprise_Pct", "Value_Date"],
           dict(WID, A=30), "F2")

ar = cr = 2
for iid, it in enumerate(items, 1):
    t, li = it["Ticker"], it["Line Item"]
    valf = EPSF if "/share" in it["Unit"] else REVF
    for pi, (p, ptype, fy, q) in enumerate(PERIODS):
        pid = pi + 1
        base = [plain(t), t, name[t], li, p, ptype, fy, q, "CD", ccy[t]]
        rdr = rd_row[(t, pi)]
        # Actuals
        act.cell(ar, 1, f"{plain(t)}|{iid}|{pid}|A")
        for j, v in enumerate(base, 2):
            act.cell(ar, j, v)
        act.cell(ar, 12, f"=Report_Dates!$G${rdr}")
        act.cell(ar, 13, f'=_xll.VAData($C{ar},$F{ar},$E{ar},"consensus.vaactuals",$J{ar},,,,,$K{ar})')
        act.cell(ar, 14, f'=IF(AND(ISNUMBER(M{ar}),ISNUMBER(L{ar}),L{ar}<=Control!$B$6),L{ar},"")')
        act.cell(ar, 12).number_format = DATE; act.cell(ar, 13).number_format = valf; act.cell(ar, 14).number_format = DATE
        # Beat_Miss (same row numbers as Actuals)
        bm.cell(ar, 1, f"{plain(t)}|{iid}|{pid}|BM")
        for j, v in enumerate(base, 2):
            bm.cell(ar, j, v)
        bm.cell(ar, 12, f"=Actuals!L{ar}")
        bm.cell(ar, 13, f"=Actuals!M{ar}")
        bm.cell(ar, 14, f'=IF(IFERROR(AND(ISNUMBER(L{ar}),ISNUMBER(M{ar}),L{ar}<=Control!$B$6),FALSE),_xll.VAData($C{ar},$F{ar},$E{ar},"consensus.vaactuals",$J{ar},,L{ar}-7,,,$K{ar}),"")')
        bm.cell(ar, 15, f'=IF(COUNT(M{ar}:N{ar})=2,M{ar}-N{ar},"")')
        bm.cell(ar, 16, f'=IFERROR(IF(COUNT(M{ar}:N{ar})=2,M{ar}/N{ar}-1,""),"")')
        bm.cell(ar, 17, f'=IF(ISNUMBER(M{ar}),L{ar},"")')
        bm.cell(ar, 12).number_format = DATE
        for col in (13, 14, 15):
            bm.cell(ar, col).number_format = valf
        bm.cell(ar, 16).number_format = PCT; bm.cell(ar, 17).number_format = DATE
        ar += 1
        # Consensus
        for lb in LOOKBACKS:
            con.cell(cr, 1, f"{plain(t)}|{iid}|{pid}|{lb}")
            for j, v in enumerate(base, 2):
                con.cell(cr, j, v)
            con.cell(cr, 12, "=Control!$B$6")
            con.cell(cr, 13, f"=Report_Dates!$G${rdr}")
            con.cell(cr, 14, f"=IF(IFERROR(AND(ISNUMBER(M{cr}),M{cr}<=L{cr}),FALSE),M{cr},L{cr})")
            con.cell(cr, 15, lb)
            con.cell(cr, 16, f"=N{cr}+O{cr}")
            con.cell(cr, 17, f'=_xll.VAData($C{cr},$F{cr},$E{cr},"consensus.vaactuals",$J{cr},,P{cr},,,$K{cr})')
            for col in (12, 13, 14, 16):
                con.cell(cr, col).number_format = DATE
            con.cell(cr, 17).number_format = valf
            cr += 1

# Line_Items
lit = sheet(wb, "Line_Items", ["Item_ID", "Ticker", "VA_Ticker_Code", "Company_Name", "Line_Item", "VA_Data_Type", "Currency", "DB_Status"],
            {"A": 18, "E": 65, "F": 18}, "A2")
for i, it in enumerate(items, 1):
    t = it["Ticker"]
    for j, v in enumerate([i, plain(t), t, name[t], it["Line Item"], "CD", ccy[t], "Not checked"], 1):
        lit.cell(i + 1, j, v)

# Tickers
tk = sheet(wb, "Tickers", ["Ticker_ID", "Ticker", "VA_Ticker_Code", "Company_Name", "Currency"], {"A": 18, "D": 30, "E": 16}, "A2")
for i, (code, nm, cy) in enumerate(COMPANIES, 1):
    for j, v in enumerate([i, plain(code), code, nm, cy], 1):
        tk.cell(i + 1, j, v)

# Periods
pe = sheet(wb, "Periods", ["Period_ID", "Period_Label", "Period_Type", "Fiscal_Year", "Quarter", "Period_Order"], {"A": 18}, "A2")
for i, (p, ptype, fy, q) in enumerate(PERIODS, 1):
    for j, v in enumerate([i, p, ptype, fy, q, i], 1):
        pe.cell(i + 1, j, v)

# Lookbacks
lk = sheet(wb, "Lookbacks", ["Lookback_ID", "Lookback_Days", "Is_Active", "Meaning"], {"A": 18, "D": 32}, None)
for i, lb in enumerate(LOOKBACKS, 1):
    meaning = "At anchor date" if lb == 0 else f"{-lb} days before anchor"
    for j, v in enumerate([i, lb, 1, meaning], 1):
        lk.cell(i + 1, j, v)

# ---------- Margins sheet ----------
key_of = {}
for iid, it in enumerate(items, 1):
    key_of[(it["Ticker"], it["Metric"])] = iid
def arow(iid, pi):
    return 2 + (iid - 1) * NP + pi
mg = sheet(wb, "Margins", ["Ticker", "VA_Ticker_Code", "Company_Name", "Period_Label", "Period_Type", "Fiscal_Year",
           "Currency", "Status", "Revenue", "Gross_Profit", "Gross_Margin", "Gross_Profit_Operating",
           "Gross_Margin_Operating", "Operating_Income", "Operating_Margin", "Operating_Income_Operating",
           "Operating_Margin_Operating"], {"A": 12, "B": 14, "C": 28, "D": 13, "I": 14, "J": 14, "L": 16, "N": 15, "P": 17}, "E2")
PCTF = r"0.0%;[Red]\-0.0%;\-"
mr = 2
for code, _, cy in COMPANIES:
    rev_id = next(key_of[(code, m)] for m in ("Revenue",) if (code, m) in key_of)
    for pi, (p, ptype, fy, q) in enumerate(PERIODS):
        rr = arow(rev_id, pi)
        mg.cell(mr, 1, plain(code)); mg.cell(mr, 2, code); mg.cell(mr, 3, name[code])
        mg.cell(mr, 4, p); mg.cell(mr, 5, ptype); mg.cell(mr, 6, fy); mg.cell(mr, 7, cy)
        mg.cell(mr, 8, f'=IF(ISNUMBER(Actuals!$N${rr}),"Reported",IF(ISNUMBER(Actuals!$M${rr}),"Consensus",""))')
        mg.cell(mr, 9, f'=IF(ISNUMBER(Actuals!$M${rr}),Actuals!$M${rr},"")')
        for vcol, mcol, metric in ((10, 11, "Gross profit"), (12, 13, "Gross profit (operating)"),
                                   (14, 15, "Operating income"), (16, 17, "Operating income (operating)")):
            iid = key_of.get((code, metric)) or (key_of.get((code, "Gross profit (NOI proxy)")) if metric == "Gross profit" else None)
            if iid:
                r2 = arow(iid, pi)
                mg.cell(mr, vcol, f'=IF(ISNUMBER(Actuals!$M${r2}),Actuals!$M${r2},"")')
                mg.cell(mr, mcol, f'=IF(AND(ISNUMBER({mg.cell(mr, vcol).coordinate}),ISNUMBER($I{mr}),N($I{mr})<>0),{mg.cell(mr, vcol).coordinate}/$I{mr},"")')
            else:
                mg.cell(mr, vcol, "n/a"); mg.cell(mr, mcol, "n/a")
            mg.cell(mr, vcol).number_format = REVF; mg.cell(mr, mcol).number_format = PCTF
        mg.cell(mr, 9).number_format = REVF
        mr += 1

# ---------- Manual_SBC sheet ----------
ms = sheet(wb, "Manual_SBC", ["Key", "VA_Ticker_Code", "Period_Label", "SBC", "Type", "Source"],
           {"A": 22, "B": 14, "C": 13, "D": 12, "E": 26, "F": 110}, "A2")
for i, mrow in enumerate(SBC_MANUAL, 2):
    proj = mrow["Filing"].startswith("Our assumption")
    ms.cell(i, 1, f'=B{i}&"|"&C{i}'); ms.cell(i, 2, mrow["Ticker"]); ms.cell(i, 3, mrow["Fiscal_Period"])
    c = ms.cell(i, 4, float(mrow["SBC_USD_m"])); c.fill = INFILL; c.font = Font(color="0000FF"); c.number_format = REVF
    ms.cell(i, 5, "Projection (our assumption)" if proj else "Reported (filing)")
    ms.cell(i, 6, mrow["Source line"] + " | " + mrow["Filing"])
MS_LAST = len(SBC_MANUAL) + 1

# ---------- Cash_Flow sheet ----------
cf = sheet(wb, "Cash_Flow", ["Ticker", "VA_Ticker_Code", "Company_Name", "Period_Label", "Period_Type", "Fiscal_Year",
           "Currency", "Status", "Revenue", "Operating_Cash_Flow", "Capex", "FCF", "FCF_Margin", "SBC", "SBC_Source",
           "FCF_After_SBC", "FCF_After_SBC_Margin", "Software_Intangibles_Capex", "FCF_Incl_Software_Intangibles"],
           {"A": 12, "B": 14, "C": 28, "D": 13, "I": 14, "J": 18, "K": 13, "L": 14, "N": 12, "O": 26, "P": 15, "R": 16, "S": 18}, "E2")
r_ = 2
for code, _, cy in COMPANIES:
    rev_id = key_of.get((code, "Revenue"))
    ocf_id, cap_id = key_of.get((code, "Operating cash flow")), key_of.get((code, "Capex"))
    sbc_id, oth_id = key_of.get((code, "Stock-based compensation")), key_of.get((code, "Other capex (software/intangibles)"))
    has_manual = any(m["Ticker"] == code for m in SBC_MANUAL)
    for pi, (p, ptype, fy, q) in enumerate(PERIODS):
        A = lambda iid: f"Actuals!$M${arow(iid, pi)}"
        cf.cell(r_, 1, plain(code)); cf.cell(r_, 2, code); cf.cell(r_, 3, name[code]); cf.cell(r_, 4, p)
        cf.cell(r_, 5, ptype); cf.cell(r_, 6, fy); cf.cell(r_, 7, cy)
        if not ocf_id:
            for col in range(8, 20): cf.cell(r_, col, "n/a")
            cf.cell(r_, 8, "Cash flow lines not supplied"); r_ += 1; continue
        cf.cell(r_, 8, f'=IF(ISNUMBER(Actuals!$N${arow(ocf_id, pi)}),"Reported",IF(ISNUMBER({A(ocf_id)}),"Consensus",""))')
        cf.cell(r_, 9, f'=IF(ISNUMBER({A(rev_id)}),{A(rev_id)},"")')
        cf.cell(r_, 10, f'=IF(ISNUMBER({A(ocf_id)}),{A(ocf_id)},"")')
        cf.cell(r_, 11, f'=IF(ISNUMBER({A(cap_id)}),ABS({A(cap_id)}),"")')
        cf.cell(r_, 12, f'=IF(COUNT(J{r_},K{r_})=2,J{r_}-K{r_},"")')
        cf.cell(r_, 13, f'=IF(AND(ISNUMBER(L{r_}),N(I{r_})<>0),L{r_}/I{r_},"")')
        if sbc_id:
            cf.cell(r_, 14, f'=IF(ISNUMBER({A(sbc_id)}),ABS({A(sbc_id)}),"")')
            cf.cell(r_, 15, "VA line")
        elif has_manual:
            look = f'INDEX(Manual_SBC!$D$2:$D${MS_LAST},MATCH($B{r_}&"|"&$D{r_},Manual_SBC!$A$2:$A${MS_LAST},0))'
            cf.cell(r_, 14, f'=IFERROR({look},"")')
            cf.cell(r_, 15, f'=IFERROR(INDEX(Manual_SBC!$E$2:$E${MS_LAST},MATCH($B{r_}&"|"&$D{r_},Manual_SBC!$A$2:$A${MS_LAST},0)),"No SBC figure (quarter)")')
        else:
            cf.cell(r_, 14, ""); cf.cell(r_, 15, "No SBC line")
        cf.cell(r_, 16, f'=IF(COUNT(L{r_},N{r_})=2,L{r_}-N{r_},"")')
        cf.cell(r_, 17, f'=IF(AND(ISNUMBER(P{r_}),N(I{r_})<>0),P{r_}/I{r_},"")')
        if oth_id:
            cf.cell(r_, 18, f'=IF(ISNUMBER({A(oth_id)}),ABS({A(oth_id)}),"")')
            cf.cell(r_, 19, f'=IF(ISNUMBER(L{r_}),L{r_}-N(R{r_}),"")')
        else:
            cf.cell(r_, 18, "none"); cf.cell(r_, 19, f'=L{r_}')
        for col in (9, 10, 11, 12, 14, 16, 18, 19): cf.cell(r_, col).number_format = REVF
        for col in (13, 17): cf.cell(r_, col).number_format = PCTF
        r_ += 1

ORDER = ["Control", "Margins", "Cash_Flow", "Manual_SBC", "DB_Review", "Actuals", "Consensus", "Beat_Miss", "Report_Dates", "Line_Items", "Tickers", "Periods", "Lookbacks"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
wb.save(OUT)
print("items", n_items, "periods", NP, "actuals", n_act, "consensus", n_con, "report_dates", r - 2)
