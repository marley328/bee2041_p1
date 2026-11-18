"""One clean, values-only workbook with everything collected so far."""
import csv, json, sys, datetime as dt
import numpy as np, pandas as pd, openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

S = "/tmp/claude-0/-home-user-bee2041-p1/d82c9cc5-7263-5271-8116-79a86939a805/scratchpad"
VA = "/root/.claude/uploads/d82c9cc5-7263-5271-8116-79a86939a805/ceea9a74-VA_Revenue_EPS_Margins_CashFlow_Actuals_Consensus_Beat_Miss_2025_2029_AsOf_2026-11-18.xlsx"
PRICES = f"{S}/Adjusted_Close_2026-01-01_to_2026-11-16.xlsx"
QUOTES = f"{S}/charts/fmp_quotes.json"
OUT = sys.argv[1]
SNAP = pd.Timestamp("2026-11-18")

CO = [  # VA code, ticker key, name, sector, currency, VA revenue unit, FYE month, FMP symbol
    ("NVDA", "NVDA", "NVIDIA", "Tech", "USD", "USD m", 1, "NVDA"),
    ("KLAC", "KLAC", "KLA Corporation", "Tech", "USD", "USD m", 6, "KLAC"),
    ("LRCX", "LRCX", "Lam Research", "Tech", "USD", "USD m", 6, "LRCX"),
    ("ASML", "ASML", "ASML Holding", "Tech", "EUR", "EUR m", 12, "ASML"),
    ("AMAT", "AMAT", "Applied Materials", "Tech", "USD", "USD m", 10, "AMAT"),
    ("000660_KR", "000660", "SK Hynix", "Tech", "KRW", "KRW bn", 12, "000660.KS"),
    ("MU", "MU", "Micron Technology", "Tech", "USD", "USD m", 8, "MU"),
    ("2330_TW", "2330", "TSMC", "Tech", "TWD", "TWD m", 12, "2330.TW"),
    ("AVGO", "AVGO", "Broadcom", "Tech", "USD", "USD m", 10, "AVGO"),
    ("INTC", "INTC", "Intel", "Tech", "USD", "USD m", 12, "INTC"),
    ("ARM_GB", "ARM", "Arm Holdings", "Tech", "USD", "USD m", 3, "ARM"),
    ("CSCO", "CSCO", "Cisco Systems", "Tech", "USD", "USD m", 7, "CSCO"),
    ("PLTR_US", "PLTR", "Palantir Technologies", "Tech", "USD", "USD m", 12, "PLTR"),
    ("SHOP_US", "SHOP", "Shopify", "Tech", "USD", "USD m", 12, "SHOP"),
    ("SPCX_US", "SPCX", "Space Exploration Technologies", "Tech", "USD", "USD m", 12, "SPCX"),
    ("CRWD_US", "CRWD", "CrowdStrike", "Cybersecurity", "USD", "USD m", 1, "CRWD"),
    ("PANW", "PANW", "Palo Alto Networks", "Cybersecurity", "USD", "USD m", 7, "PANW"),
    ("NET_US", "NET", "Cloudflare", "Cybersecurity", "USD", "USD m", 12, "NET"),
    ("COST", "COST", "Costco Wholesale", "Everything else", "USD", "USD m", 8, "COST"),
    ("WMT", "WMT", "Walmart", "Everything else", "USD", "USD m", 1, "WMT"),
    ("TSLA", "TSLA", "Tesla", "Everything else", "USD", "USD m", 12, "TSLA"),
    ("MNST", "MNST", "Monster Beverage", "Everything else", "USD", "USD m", 12, "MNST"),
    ("WELL_US", "WELL", "Welltower", "Everything else", "USD", "USD m", 12, "WELL"),
    ("BE_US", "BE", "Bloom Energy", "Everything else", "USD", "USD m", 12, "BE"),
]
C = {c[0]: c for c in CO}
SCALE = {"KRW bn": 1e9}  # revenue unit -> currency units; default 1e6
MONTH = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

METRIC = {"Revenue": "revenue", "GAAP diluted EPS": "eps_gaap", "GAAP EPS (no diluted line in VA)": "eps_gaap",
          "Operating diluted EPS": "eps_op", "Operating diluted EPS (normalized FFO/share)": "ffo",
          "Gross profit": "gp", "Gross profit (NOI proxy)": "noi", "Gross profit (operating)": "gp_adj",
          "Operating income": "oi", "Operating income (operating)": "oi_adj", "Operating cash flow": "ocf",
          "Capex": "capex", "Other capex (software/intangibles)": "capex_other", "Stock-based compensation": "sbc"}
LABEL = {"revenue": "Revenue", "eps_gaap": "GAAP diluted EPS", "eps_op": "Non-GAAP (operating) diluted EPS",
         "ffo": "Normalized FFO per diluted share", "gp": "Gross profit (GAAP)", "gp_adj": "Gross profit (adjusted)",
         "noi": "Net operating income", "oi": "Operating income (GAAP)", "oi_adj": "Operating income (adjusted)",
         "ocf": "Operating cash flow", "capex": "Capex (PP&E)", "capex_other": "Software / intangibles capex",
         "sbc": "Stock-based compensation"}
linemap = {}
for f in ("va_line_items", "va_margin_line_items", "va_cashflow_line_items"):
    for r in csv.DictReader(open(f"{S}/{f}.csv")):
        linemap[(r["Ticker"], r["Line Item"])] = METRIC[r["Metric"]]

# ------------------------------------------------------------------ read VA workbook values
wb = openpyxl.load_workbook(VA, data_only=True, read_only=True)
def frame(name):
    rows = list(wb[name].iter_rows(values_only=True))
    return pd.DataFrame(rows[1:], columns=rows[0])
num = lambda s: pd.to_numeric(s, errors="coerce")
def xl(s):
    def one(v):
        if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)): return pd.Timestamp(v)
        if isinstance(v, (int, float)) and not pd.isna(v) and 0 < v < 100000: return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
        return pd.NaT
    return pd.Series([one(v) for v in s], index=s.index, dtype="datetime64[ns]")
act, con, bm = frame("Actuals"), frame("Consensus"), frame("Beat_Miss")
for d in (act, con, bm):
    d["metric"] = [linemap.get((c, l)) for c, l in zip(d.VA_Ticker_Code, d.Line_Item)]
assert act.metric.notna().all(), act[act.metric.isna()][["VA_Ticker_Code", "Line_Item"]].drop_duplicates()
act["val"] = num(act.Actual_Value)
act["reported"] = xl(act.Value_Date).notna()
act["rdate"] = xl(act.Reported_Date)
con["val"] = num(con.Consensus_Value); con["asof"] = xl(con.As_Of_Date)
c0 = con[con.Lookback_Days == 0][["VA_Ticker_Code", "metric", "Period_Label", "val"]].rename(columns={"val": "cons"})
L = act.merge(c0, on=["VA_Ticker_Code", "metric", "Period_Label"], how="left")
# SpaceX 2025 periods: VA actuals exist but VA lists no report date (pre-IPO)
spx = (L.VA_Ticker_Code == "SPCX_US") & (L.Fiscal_Year == 2025) & L.val.notna()
L.loc[spx, "reported"] = True
L["value"] = np.where(L.reported, L.val, L.cons)
# period status from revenue line
st = L[L.metric == "revenue"].set_index(["VA_Ticker_Code", "Period_Label"])
L["Status"] = [("Reported" if st.loc[(c, p), "reported"] else "Consensus") for c, p in zip(L.VA_Ticker_Code, L.Period_Label)]
L.loc[spx, "Status"] = "Reported (VA actual, no report date)"

# manual SBC (filings + our projections)
man = pd.read_csv(f"{S}/sbc_from_10k.csv")
man["Type"] = np.where(man.Filing.str.startswith("Our assumption"), "Our estimate", "Company disclosure")

# ------------------------------------------------------------------ wide period table
W = L.pivot_table(index=["VA_Ticker_Code", "Period_Label", "Period_Type", "Fiscal_Year", "Quarter"], columns="metric",
                  values="value", aggfunc="first").reset_index()
stat = L.groupby(["VA_Ticker_Code", "Period_Label"]).Status.first()
rdate = L[L.metric == "revenue"].set_index(["VA_Ticker_Code", "Period_Label"]).rdate
W["Status"] = [stat[(c, p)] for c, p in zip(W.VA_Ticker_Code, W.Period_Label)]
W["Report date"] = [rdate.get((c, p)) for c, p in zip(W.VA_Ticker_Code, W.Period_Label)]
for m in LABEL:
    if m not in W: W[m] = np.nan
welli = W.VA_Ticker_Code == "WELL_US"
W.loc[welli, "gp"] = W.loc[welli, "noi"]   # Welltower: gross profit = net operating income
W["capex"] = W.capex.abs(); W["capex_other"] = W.capex_other.abs(); W["sbc"] = W.sbc.abs()
W["sbc_source"] = np.where(W.sbc.notna(), "Visible Alpha", "")
mk = {(r.Ticker, r.Fiscal_Period): (r.SBC_USD_m, r.Type) for r in man.itertuples()}
for i, r in W.iterrows():
    if pd.isna(r.sbc) and (r.VA_Ticker_Code, r.Period_Label) in mk:
        v, t = mk[(r.VA_Ticker_Code, r.Period_Label)]
        W.at[i, "sbc"] = v; W.at[i, "sbc_source"] = t + " (see Estimates_Disclosures)"
for i, r in W[(W.Period_Type == "Quarterly") & W.sbc.isna()].iterrows():
    k = (r.VA_Ticker_Code, f"FY-{int(r.Fiscal_Year)}")
    if k in mk:
        v, t = mk[k]
        W.at[i, "sbc"] = v / 4
        W.at[i, "sbc_source"] = f"Our estimate: {t.lower()} annual SBC ÷ 4"
W["fcf"] = W.ocf - W.capex
W["fcf_after_sbc"] = W.fcf - W.sbc
W["fcf_incl_other"] = W.fcf - W.capex_other.fillna(0)
div = lambda a, b: np.where((b.notna()) & (b != 0), a / b, np.nan)
W["gm"] = div(W.gp, W.revenue); W["gm_adj"] = div(W.gp_adj, W.revenue); W["noi_m"] = div(W.noi, W.revenue)
W["om"] = div(W.oi, W.revenue); W["om_adj"] = div(W.oi_adj, W.revenue)
W["fcf_m"] = div(W.fcf, W.revenue); W["fcf_sbc_m"] = div(W.fcf_after_sbc, W.revenue)
W["Ticker"] = W.VA_Ticker_Code.map(lambda c: C[c][1]); W["Company"] = W.VA_Ticker_Code.map(lambda c: C[c][2])
W["Sector"] = W.VA_Ticker_Code.map(lambda c: C[c][3]); W["Currency"] = W.VA_Ticker_Code.map(lambda c: C[c][4])
W["Unit (money)"] = W.VA_Ticker_Code.map(lambda c: C[c][5])
order = {c[0]: i for i, c in enumerate(CO)}
W["_o"] = W.VA_Ticker_Code.map(order)
pord = {f"{q}QFY-{y}": (y, q) for y in range(2025, 2030) for q in range(1, 5)} | {f"FY-{y}": (y, 5) for y in range(2025, 2030)}
W["_p"] = W.Period_Label.map(pord)
W = W.sort_values(["_o", "_p"])
COLS = [("Ticker", "Ticker"), ("VA_Ticker_Code", "VA code"), ("Company", "Company"), ("Sector", "Sector"),
        ("Period_Label", "Period"), ("Fiscal_Year", "Fiscal year"), ("Quarter", "Quarter"), ("Status", "Status"),
        ("Report date", "Report date"), ("Currency", "Currency"), ("Unit (money)", "Unit (money)"),
        ("revenue", "Revenue"), ("gp", "Gross profit (GAAP)"), ("gp_adj", "Gross profit (adjusted)"),
        ("noi", "Net operating income (WELL)"), ("oi", "Operating income (GAAP)"), ("oi_adj", "Operating income (adjusted)"),
        ("eps_gaap", "GAAP diluted EPS"), ("eps_op", "Non-GAAP diluted EPS"), ("ffo", "Normalized FFO/share (WELL)"),
        ("ocf", "Operating cash flow"), ("capex", "Capex (PP&E)"), ("capex_other", "Software/intangibles capex"),
        ("sbc", "Stock-based comp"), ("sbc_source", "SBC source"), ("fcf", "FCF (OCF − capex)"),
        ("fcf_after_sbc", "FCF after SBC"), ("fcf_incl_other", "FCF incl. software/intangibles"),
        ("gm", "Gross margin (GAAP)"), ("gm_adj", "Gross margin (adjusted)"), ("noi_m", "NOI margin (WELL)"),
        ("om", "Operating margin (GAAP)"), ("om_adj", "Operating margin (adjusted)"), ("fcf_m", "FCF margin"),
        ("fcf_sbc_m", "FCF after SBC margin")]
def tidy(df):
    out = df[[c for c, _ in COLS]].rename(columns=dict(COLS)).copy()
    out["Report date"] = pd.to_datetime(out["Report date"], errors="coerce").dt.date
    return out
annual = tidy(W[W.Period_Type == "Annual"]).drop(columns=["Quarter"])
quarterly = tidy(W[W.Period_Type == "Quarterly"])

# ------------------------------------------------------------------ prices, valuation
pw = list(openpyxl.load_workbook(PRICES, data_only=True, read_only=True)["Prices"].iter_rows(values_only=True))
prices = pd.DataFrame(pw[1:], columns=pw[0]).set_index("date")
keep = [c[7] for c in CO if c[7] in prices.columns] + [x for x in ("EURUSD", "TWDUSD", "KRWUSD") if x in prices.columns]
prices = prices[keep].dropna(how="all")
pdate = prices.dropna(how="all").index.max()
last = {c: prices[c].dropna().iloc[-1] for c in prices.columns if prices[c].notna().any()}
quotes = json.load(open(QUOTES))
def fwd_fy(fye):
    for y in range(2025, 2030):
        if pd.Timestamp(year=y, month=fye, day=1) + pd.offsets.MonthEnd(0) > SNAP + pd.DateOffset(months=3):
            return y
Ann = W[W.Period_Type == "Annual"].set_index(["VA_Ticker_Code", "Fiscal_Year"])
vrows = []
for code, tk, nm, sec, cy, unit, fye, fmp in CO:
    q = quotes.get(tk) or quotes.get(code)
    price = last.get(fmp)
    if code == "ASML": price = price / last["EURUSD"]  # US ADR in USD -> EUR to match VA
    shares = q["shares"] if q else np.nan
    mc = price * shares if price is not None else np.nan
    fy = fwd_fy(fye)
    rep = [y for y in range(2025, 2030) if (code, y) in Ann.index and str(Ann.loc[(code, y), "Status"]).startswith("Reported")]
    ty = rep[-1] if rep else np.nan
    g = lambda m, y: Ann.loc[(code, y), m] if (code, y) in Ann.index else np.nan
    sc = SCALE.get(unit, 1e6)
    pos = lambda v: v if (pd.notna(v) and v > 0) else np.nan
    gr = lambda m: (g(m, fy) / g(m, fy - 1) - 1) if pd.notna(g(m, fy - 1)) and g(m, fy - 1) > 0 else np.nan
    vrows.append({
        "Ticker": tk, "Company": nm, "Sector": sec, "Currency": cy, "Price (local)": price, "Price date": pdate.date(),
        "Shares (FMP)": shares, "Market cap (local)": mc, "Trailing FY": ty, "Forward FY": fy,
        "P/S trailing": mc / (pos(g("revenue", ty)) * sc) if pd.notna(ty) else np.nan,
        "P/S forward": mc / (pos(g("revenue", fy)) * sc),
        "P/E GAAP trailing": price / pos(g("eps_gaap", ty)) if pd.notna(ty) else np.nan,
        "P/E GAAP forward": price / pos(g("eps_gaap", fy)),
        "P/E non-GAAP trailing": price / pos(g("eps_op", ty)) if pd.notna(ty) else np.nan,
        "P/E non-GAAP forward": price / pos(g("eps_op", fy)),
        "FCF yield trailing": g("fcf", ty) * sc / mc if pd.notna(ty) else np.nan,
        "FCF yield forward": g("fcf", fy) * sc / mc,
        "FCF yield after SBC trailing": g("fcf_after_sbc", ty) * sc / mc if pd.notna(ty) else np.nan,
        "FCF yield after SBC forward": g("fcf_after_sbc", fy) * sc / mc,
        "Revenue growth (fwd FY)": gr("revenue"), "GAAP EPS growth (fwd FY)": gr("eps_gaap"),
        "Non-GAAP EPS growth (fwd FY)": gr("eps_op"),
    })
val = pd.DataFrame(vrows)

# ------------------------------------------------------------------ consensus history & revisions (annual)
ch = con[(con.Period_Type == "Annual")].copy()
ch["Ticker"] = ch.VA_Ticker_Code.map(lambda c: C[c][1])
ch["Metric"] = ch.metric.map(LABEL)
ch = ch[["Ticker", "VA_Ticker_Code", "Metric", "Line_Item", "Period_Label", "Fiscal_Year", "Lookback_Days", "asof", "val"]]
ch.columns = ["Ticker", "VA code", "Metric", "VA line item", "Period", "Fiscal year", "Lookback days", "As-of date", "Consensus"]
ch["As-of date"] = ch["As-of date"].dt.date
ch["_o"] = ch["VA code"].map(order); ch = ch.sort_values(["_o", "Metric", "Fiscal year", "Lookback days"]).drop(columns="_o")
rv = []
for code, tk, nm, sec, *_r in CO:
    fy = fwd_fy(C[code][6])
    for m in ("revenue", "eps_gaap", "eps_op", "ocf", "capex", "gp", "oi"):
        s = con[(con.VA_Ticker_Code == code) & (con.metric == m) & (con.Fiscal_Year == fy) & (con.Period_Type == "Annual")]
        if s.empty: continue
        s = s.set_index("Lookback_Days").val
        now = s.get(0)
        r = {"Ticker": tk, "Company": nm, "Sector": sec, "Metric": LABEL[m], "Forward FY": fy, "Consensus now": now}
        for lb, w in ((-30, "1M"), (-90, "3M"), (-180, "6M"), (-365, "12M")):
            then = s.get(lb)
            r[f"Change {w}"] = (now / then - 1) if pd.notna(now) and pd.notna(then) and then > 0 else np.nan
        rv.append(r)
rev = pd.DataFrame(rv)

# ------------------------------------------------------------------ beat / miss
bm["Surprise_Pct"] = num(bm.Surprise_Pct)
b = bm[bm.metric.isin(["revenue", "eps_gaap", "eps_op"]) & num(bm.Surprise).notna()].copy()
b["Ticker"] = b.VA_Ticker_Code.map(lambda c: C[c][1])
b["Metric"] = b.metric.map(LABEL)
b["Result"] = np.where(num(b.Surprise) > 0, "Beat", np.where(num(b.Surprise) < 0, "Miss", "In line"))
b["Reported date"] = xl(b.Reported_Date).dt.date
b = b[["Ticker", "VA_Ticker_Code", "Metric", "Period_Label", "Reported date", "Actual_Value",
       "Consensus_7D_Pre_Report", "Surprise", "Surprise_Pct", "Result"]]
b.columns = ["Ticker", "VA code", "Metric", "Period", "Reported date", "Actual", "Consensus 7 days before", "Surprise",
             "Surprise %", "Result"]
for c in ("Actual", "Consensus 7 days before", "Surprise"): b[c] = num(b[c])
b["_o"] = b["VA code"].map(order); b = b.sort_values(["_o", "Metric", "Period"]).drop(columns="_o")

# ------------------------------------------------------------------ line map
lm = []
for f in ("va_line_items", "va_margin_line_items", "va_cashflow_line_items"):
    for r in csv.DictReader(open(f"{S}/{f}.csv")):
        lm.append({"Ticker": C[r["Ticker"]][1], "VA code": r["Ticker"], "Metric": LABEL[METRIC[r["Metric"]]],
                   "Exact VA line item": r["Line Item"], "Unit": r["Unit"]})
lmap = pd.DataFrame(lm); lmap["_o"] = lmap["VA code"].map(order)
lmap = lmap.sort_values(["_o"]).drop(columns="_o")

# ------------------------------------------------------------------ estimates & disclosures
est = man.copy()
est["Ticker"] = est.Ticker.map(lambda c: C[c][1])
est["Unit"] = est.Ticker.map({c[1]: c[5] for c in CO})
est = est[["Ticker", "Fiscal_Period", "Type", "SBC_USD_m", "Unit", "Source line", "Filing"]]
est.columns = ["Ticker", "Period", "Type", "Stock-based compensation", "Unit", "Detail", "Source"]
other_est = pd.DataFrame([
    ("All", "Forward FY", "Our convention", "First fiscal year ending more than 3 months after 18 Nov 2026 (FY28 for January year-ends, FY27 for the rest)."),
    ("All", "Sector groups", "Our convention", "Tech / Cybersecurity / Everything else, as listed on the Companies tab."),
    ("All", "Market cap", "Our calculation", "Last adjusted close in the price file x FMP shares outstanding."),
    ("ASML", "Price", "Our calculation", "US-listed ADR price (USD) converted to EUR at the EURUSD rate on the same date, to match VA's EUR figures."),
    ("SPCX", "FY2025 status", "Our treatment", "VA holds FY2025 actuals but lists no report date (pre-IPO); treated as reported."),
    ("WMT", "SBC base for projections", "Our adjustment", "FY26 SBC 3,603 less ~700 one-off PhonePe modification charge = 2,903 base."),
    ("WELL", "SBC base for projections", "Our adjustment", "FY25 SBC 1,556.7 less 1,408.7 one-off Executive & Key Employee LTIP awards = 148.1 base."),
    ("WELL", "Gross profit", "Definition (Kolytics)", "Welltower gross profit = Net operating income (VA line 'Net operating income'); gross margin = NOI / property revenues."),
    ("WMT/COST/WELL/TSMC/SKH/MNST", "Quarterly SBC", "Our estimate", "Quarterly SBC = that fiscal year's annual SBC ÷ 4 (filing or projection). Welltower FY25 quarters include ¼ of the one-off LTIP award."),
    ("All", "FCF", "Our calculation", "FCF = operating cash flow − ABS(PP&E capex); after SBC also subtracts ABS(SBC)."),
], columns=["Ticker", "Item", "Type", "Detail"])

# ------------------------------------------------------------------ missing data audit
miss = []
need = {"revenue": "Revenue", "gp": "Gross profit (GAAP)", "gp_adj": "Gross profit (adjusted)", "oi": "Operating income (GAAP)",
        "oi_adj": "Operating income (adjusted)", "eps_gaap": "GAAP EPS", "eps_op": "Non-GAAP EPS", "ocf": "Operating cash flow",
        "capex": "Capex", "sbc": "SBC"}
A2 = W[W.Period_Type == "Annual"]
for code, tk, nm, *_r in CO:
    d = A2[A2.VA_Ticker_Code == code]
    for m, lab in need.items():
        have = d[m].notna() if m in d else pd.Series(False, index=d.index)
        missing_years = d.loc[~have, "Fiscal_Year"].tolist()
        if not missing_years: continue
        if not have.any():
            why = {"gp_adj": "No adjusted gross profit line in VA", "oi_adj": "No adjusted operating income line in VA",
                   "eps_op": "No non-GAAP EPS line in VA", "sbc": "No SBC line in VA or filings",
                   "ocf": "Cash flow screenshot not supplied", "capex": "Cash flow screenshot not supplied",
                   "gp": "No gross profit line in VA", "oi": "No operating income line in VA"}.get(m, "No data")
            if tk == "WELL" and m in ("oi", "oi_adj"): why = "REIT: definition being agreed (see chat); VA lines not yet pulled"
            if tk == "WELL" and m == "gp_adj": why = "REIT: no adjusted NOI used"
            if tk == "WELL" and m == "eps_op": why = "REIT: normalized FFO/share used instead"
            if tk == "MNST" and m in ("ocf", "capex"): continue
            miss.append((tk, nm, lab, "All years", why))
        else:
            miss.append((tk, nm, lab, ", ".join(f"FY{int(y)}" for y in missing_years), "Consensus not available for these years"))
miss.append(("MNST", "Monster Beverage", "Operating cash flow / capex", "All periods", "Monster-only VA workbook sent for refresh; not yet returned"))
for r in rev.itertuples():
    if pd.isna(r._10) and r.Metric in ("Revenue", "GAAP diluted EPS"):
        miss.append((r.Ticker, r.Company, f"12M consensus revision ({r.Metric})", f"FY{r._5}", "No consensus 12 months ago (recent listing / coverage)"))
missing = pd.DataFrame(miss, columns=["Ticker", "Company", "Data", "Periods", "Reason"])

# ------------------------------------------------------------------ companies
comp = pd.DataFrame([{"Ticker": c[1], "VA code": c[0], "FMP symbol": c[7], "Company": c[2], "Sector": c[3],
                      "Reporting currency": c[4], "VA money unit": c[5], "Fiscal year end": MONTH[c[6]],
                      "Forward FY": fwd_fy(c[6])} for c in CO])

# ------------------------------------------------------------------ write
HF, HFont = PatternFill("solid", fgColor="1F4E78"), Font(bold=True, color="FFFFFF")
EST = PatternFill("solid", fgColor="FFF2CC")
MONEY, EPS, PCT, MULT = "#,##0.0;[Red]-#,##0.0", "#,##0.00;[Red]-#,##0.00", "0.0%;[Red]-0.0%", "0.0x"
wbo = Workbook(); wbo.remove(wbo.active)

def put(name, df, fmts=None, widths=None, freeze="B2", highlight=None):
    ws = wbo.create_sheet(name)
    for j, c in enumerate(df.columns, 1):
        cell = ws.cell(1, j, c); cell.fill, cell.font = HF, HFont
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and (np.isnan(v) or np.isinf(v)): v = None
            if isinstance(v, (np.floating,)): v = None if np.isnan(v) else float(v)
            if isinstance(v, (np.integer,)): v = int(v)
            ws.cell(i, j, v)
    for j, c in enumerate(df.columns, 1):
        f = (fmts or {}).get(c)
        if f:
            for i in range(2, len(df) + 2): ws.cell(i, j).number_format = f
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(c, max(10, min(40, len(str(c)) + 2)))
    if highlight:
        col, test = highlight
        j = list(df.columns).index(col) + 1
        for i in range(2, len(df) + 2):
            if test(ws.cell(i, j).value):
                for jj in range(1, len(df.columns) + 1): ws.cell(i, jj).fill = EST
    ws.freeze_panes = freeze; ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32
    return ws

readme = wbo.create_sheet("README")
lines = [
    ("Visible Alpha consensus pack: 24 companies, FY2025–FY2029", True),
    ("", False),
    (f"VA snapshot: {SNAP:%d %b %Y}.  Prices: adjusted close to {pdate:%d %b %Y}.  Values only (no VA formulas).", False),
    ("", False),
    ("Tabs", True),
    ("Companies — tickers, sector grouping, currency, VA unit, fiscal year end, forward FY.", False),
    ("Annual — one row per company and fiscal year: revenue, gross profit, operating income, EPS, cash flow, FCF, margins.", False),
    ("Quarterly — the same for each quarter.", False),
    ("Valuation — price, market cap, P/S, P/E (GAAP and non-GAAP kept separate), FCF yield, forward growth.", False),
    ("Revisions — change in forward-FY consensus over 1, 3, 6 and 12 months.", False),
    ("Consensus_History — every annual consensus value at each lookback date (0 to −365 days).", False),
    ("Beat_Miss — reported quarters: actual vs consensus 7 days before the report.", False),
    ("Prices — daily adjusted close and FX.", False),
    ("Estimates_Disclosures — data not from Visible Alpha: company filings and our own estimates/conventions (yellow rows = our estimates).", False),
    ("VA_Line_Map — exact Visible Alpha line item behind every metric.", False),
    ("Missing_Data — what is not available and why.", False),
    ("", False),
    ("Definitions", True),
    ("Status: Reported = company has reported the period; Consensus = latest Visible Alpha consensus at the snapshot.", False),
    ("GAAP and non-GAAP are never mixed: a company without a non-GAAP line is left blank, not filled with GAAP.", False),
    ("FCF = operating cash flow − PP&E capex. FCF after SBC also subtracts stock-based compensation.", False),
    ("Money in each company's VA unit (USD m, EUR m for ASML, TWD m for TSMC, KRW bn for SK Hynix). EPS per share in local currency.", False),
    ("SBC source column shows where SBC came from when it is not a Visible Alpha line.", False),
]
for i, (t, bold) in enumerate(lines, 1):
    c = readme.cell(i, 1, t); c.font = Font(bold=bold, size=14 if i == 1 else 11)
readme.column_dimensions["A"].width = 130

put("Companies", comp)
money_cols = ["Revenue", "Gross profit (GAAP)", "Gross profit (adjusted)", "Net operating income (WELL)", "Operating income (GAAP)",
              "Operating income (adjusted)", "Operating cash flow", "Capex (PP&E)", "Software/intangibles capex", "Stock-based comp",
              "FCF (OCF − capex)", "FCF after SBC", "FCF incl. software/intangibles"]
eps_cols = ["GAAP diluted EPS", "Non-GAAP diluted EPS", "Normalized FFO/share (WELL)"]
pct_cols = ["Gross margin (GAAP)", "Gross margin (adjusted)", "NOI margin (WELL)", "Operating margin (GAAP)",
            "Operating margin (adjusted)", "FCF margin", "FCF after SBC margin"]
F = {c: MONEY for c in money_cols} | {c: EPS for c in eps_cols} | {c: PCT for c in pct_cols} | {"Report date": "dd-mmm-yyyy"}
Wd = {"Company": 26, "SBC source": 34, "Status": 16}
put("Annual", annual, F, Wd, "F2")
put("Quarterly", quarterly, F, Wd, "F2")
VF = {c: MULT for c in val.columns if c.startswith("P/")} | {c: PCT for c in val.columns if "yield" in c or "growth" in c} | \
     {"Price (local)": "#,##0.00", "Shares (FMP)": "#,##0", "Market cap (local)": "#,##0", "Price date": "dd-mmm-yyyy"}
put("Valuation", val, VF, {"Company": 26}, "C2")
put("Revisions", rev, {c: PCT for c in rev.columns if c.startswith("Change")} | {"Consensus now": "#,##0.00"}, {"Company": 26}, "E2")
put("Consensus_History", ch, {"Consensus": "#,##0.00", "As-of date": "dd-mmm-yyyy"}, {"VA line item": 50}, "D2")
put("Beat_Miss", b, {"Actual": "#,##0.00", "Consensus 7 days before": "#,##0.00", "Surprise": "#,##0.00", "Surprise %": PCT,
                     "Reported date": "dd-mmm-yyyy"}, None, "D2")
pr = prices.reset_index(); pr["date"] = pd.to_datetime(pr["date"]).dt.date
put("Prices", pr, {"date": "dd-mmm-yyyy"} | {c: "#,##0.00" for c in pr.columns if c != "date"})
ws = put("Estimates_Disclosures", est, {"Stock-based compensation": "#,##0.0"}, {"Detail": 80, "Source": 45, "Type": 20},
         "C2", ("Type", lambda v: v == "Our estimate"))
r0 = len(est) + 4
ws.cell(r0, 1, "Other estimates, adjustments and conventions").font = Font(bold=True, size=12)
for j, c in enumerate(other_est.columns, 1):
    cell = ws.cell(r0 + 1, j, c); cell.fill, cell.font = HF, HFont
for i, row in enumerate(other_est.itertuples(index=False), r0 + 2):
    for j, v in enumerate(row, 1): ws.cell(i, j, v)
put("VA_Line_Map", lmap, None, {"Exact VA line item": 70, "Metric": 32})
put("Missing_Data", missing, None, {"Data": 34, "Periods": 30, "Reason": 60, "Company": 26})
wbo.save(OUT)
print("annual", len(annual), "quarterly", len(quarterly), "history", len(ch), "beat/miss", len(b), "missing", len(missing))
print(missing.to_string(index=False))
