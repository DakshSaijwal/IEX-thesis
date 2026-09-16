"""
build_model.py
---------------
Indian Energy Exchange (NSE: IEX) -- short thesis scenario model.

The catalyst: CERC's "market coupling" mandate centralises electricity price
discovery across all power exchanges, eroding IEX's liquidity-driven pricing
power. The legal and regulatory fight is live and unresolved as of Sept 2026
(a Supreme Court challenge and Grid India's own technical objections to the
draft rules are both outstanding).

This model does NOT use any non-public information. Every figure is built
from IEX's own public disclosures and press coverage of the regulatory
process, current as of the build date. Treat every yellow cell as your own
assumption, not a fact -- verify against a live quote and IEX's own filings
before presenting this anywhere.

Run:  python scripts/build_model.py
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

BLUE = Font(name="Arial", size=10, color="0000FF")
BLACK = Font(name="Arial", size=10, color="000000")
GREEN = Font(name="Arial", size=10, color="008000")
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=13, bold=True)
SUB = Font(name="Arial", size=10, italic=True, color="595959")
HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")

YELLOW = PatternFill("solid", fgColor="FFFF00")
NAVY = PatternFill("solid", fgColor="1F3864")
GREYFILL = PatternFill("solid", fgColor="F2F2F2")

TOPBORDER = Border(top=Side(style="thin"))
TOPDOUBLE = Border(top=Side(style="thin"), bottom=Side(style="double"))

NUM = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%'
MULT = '0.0x'
RS = '#,##0.00'
UNITRATE = '0.0000'


def put(ws, cell, value, font=BLACK, fmt=None, fill=None, border=None, align=None):
    c = ws[cell]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def label(ws, row, text, bold=False, indent=0, col=1):
    c = ws.cell(row=row, column=col)
    c.value = text
    c.font = BOLD if bold else BLACK
    if indent:
        c.alignment = Alignment(indent=indent)


def widths(ws, a=50, rest=14):
    ws.column_dimensions["A"].width = a
    for col in ["B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = rest


wb = Workbook()

# =============================================================== COVER =====
ws = wb.active
ws.title = "Cover"
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 84

put(ws, "A1", "INDIAN ENERGY EXCHANGE (NSE: IEX)", TITLE)
put(ws, "A2", "Short Thesis: Market Coupling and the End of the Liquidity Moat", SUB)

put(ws, "A4", "THE PITCH IN THREE SENTENCES", TITLE)
pitch = ("CERC's market coupling mandate will centralise electricity price discovery across all power "
         "exchanges, and the direction of that policy is no longer in question, only its timing. IEX's "
         "85-90% Day-Ahead Market share is a function of liquidity, which coupling makes irrelevant, so its "
         "premium pricing power should compress even if a residual customer relationship survives. At INR "
         "120, the market is pricing in something close to the bull case (moat mostly intact); this model "
         "argues the base case (partial share and fee erosion) is more likely and implies roughly 45-50% "
         "downside, with a further, less likely bear case implying over 70% downside.")
put(ws, "A5", pitch, BLACK)
ws["A5"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A5:B5")
ws.row_dimensions[5].height = 90

rows = [
    ("Recommendation", "SHORT / SELL (illustrative, educational exercise -- not investment advice)"),
    ("Catalyst", "Supreme Court ruling on IEX's challenge to market coupling; CERC's final notification following the Grid India technical review. Both outstanding as of Sept 2026, no confirmed date."),
    ("Current price (confirmed live)", "INR 113.90, confirmed 13-Sep-2026 (kotakneo.com live quote). No longer an approximation."),
    ("Base case target", "INR 58 (-49% vs current price)"),
    ("Bear case target", "INR 31 (-73% vs current price)"),
    ("Bull case target (the risk to being short)", "INR 147 (+29% vs current price): if API-based customer lock-in fully offsets uniform pricing"),
    ("Sources", "Public news coverage of the CERC market coupling proceeding, IEX's own investor disclosures (Screener.in, exchange filings), and analyst commentary reported in the financial press. No non-public information of any kind was used."),
    ("Disclaimer", "Educational exercise built for interview preparation. Not investment research, not a recommendation to buy or sell any security, and not based on any professional research mandate."),
]
r = 13
for k, v in rows:
    put(ws, f"A{r}", k, BOLD)
    put(ws, f"B{r}", v, BLACK)
    ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 44 if len(v) > 90 else 18
    r += 2

# =========================================================== ASSUMPTIONS ===
aw = wb.create_sheet("Assumptions")
widths(aw, a=52)
put(aw, "A1", "CURRENT STATE AND SCENARIO ASSUMPTIONS", TITLE)
put(aw, "A2", "Every yellow cell is an input you own and must justify before presenting this model.", SUB)

rows = [
    (4, "CURRENT STATE (approx, Sept 2026 -- verify against a live quote)", None, None, None),
    (5, "Current share price (INR)", 113.90, RS, "CONFIRMED 13-Sep-2026: kotakneo.com live quote, Rs 113.90 (day range Rs 113.00-115.91). This IS the live price on the memo date."),
    (6, "Market capitalisation (INR Cr)", 10517.52, NUM, "CONFIRMED 13-Sep-2026: indmoney.com, Rs 10,517.52 Cr (as of 10-Sep-2026 intraday)."),
    (7, "Implied diluted shares (Cr)", "=B6/B5", NUM, "NOTE: back-derived from two different vendor snapshots taken at slightly different times; will not perfectly match IEX's own reported diluted share count in its filings."),
    (8, "TTM revenue (INR Cr)", 624, NUM, "Screener.in, TTM as of approx Sept 2026."),
    (9, "TTM net profit (INR Cr)", 487, NUM, "Screener.in, TTM as of approx Sept 2026."),
    (10, "TTM PAT margin", "=B9/B8", PCT, None),
    (11, "TTM EPS (INR)", "=B9/B7", RS, None),
    (12, "Implied current P/E", "=B5/B11", MULT, None),
    (13, "IEX's own FY26 traded volume (billion units, BU)", 145, NUM,
         "Derived from FY25's reported 121 BU and ~19-20% YoY volume growth seen in recent monthly prints."),
    (14, "IEX's current share of total exchange-traded market", 0.87, PCT,
         "Widely reported 85-90% Day-Ahead Market share; midpoint used."),
    (15, "Implied total exchange-traded market, FY26 (BU)", "=B13/B14", NUM, None),
    (16, "IEX's current take rate (INR Cr revenue per BU traded)", "=B8/B13", UNITRATE,
         "This is IEX's own blended fee per unit of electricity traded on its platform."),
    (18, "FORWARD MARKET ASSUMPTION", None, None, None),
    (19, "Total exchange-traded market CAGR (2-year forward)", 0.16, PCT,
         "Underlying secular growth: rising renewable capacity, gradual shift from long-term PPAs to short-term exchange trading. This is a growing-pie assumption independent of who captures share."),
    (20, "Forward horizon (years)", 2, NUM, "To FY28, matching the timeframe most analyst share-loss estimates use."),
    (21, "Implied total exchange-traded market, FY28E (BU)", "=B15*(1+B19)^B20", NUM, None),
    (23, "SCENARIO DEFINITIONS (FY28E)", None, None, None),
]
for r, name, v, fmt, note in rows:
    if v is None:
        label(aw, r, name, bold=True)
        aw.cell(row=r, column=1).fill = GREYFILL
        continue
    is_formula = isinstance(v, str) and v.startswith("=")
    is_input = not is_formula
    label(aw, r, name, indent=1)
    put(aw, f"B{r}", v, BLUE if is_input else BOLD, fmt,
        fill=YELLOW if is_input else None, border=TOPBORDER if is_formula else None)
    if note:
        put(aw, f"D{r}", note, SUB)
        aw[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        aw.row_dimensions[r].height = 44 if len(note) > 140 else (30 if len(note) > 70 else 16)
        aw.column_dimensions["D"].width = 66

SCEN_R0 = 25
put(aw, f"A{SCEN_R0-1}", "", None)
scen_hdr_row = SCEN_R0
put(aw, f"A{scen_hdr_row}", "Scenario", HDR, fill=NAVY)
for col, hdr in zip(["B", "C", "D"], ["BULL", "BASE", "BEAR"]):
    put(aw, f"{col}{scen_hdr_row}", hdr, HDR, fill=NAVY, align="center")

scen_rows = [
    ("FY28E market share", [0.70, 0.57, 0.50],
     "BULL: API-based 'tight coupling' with 70%+ of members preserves most switching friction. "
     "BASE: meaningful but partial erosion. BEAR: matches published analyst projections of ~50% by FY28."),
    ("Take rate change vs today", [-0.08, -0.25, -0.35],
     "How much IEX's blended fee per unit compresses as uniform pricing removes its differentiation."),
    ("FY28E PAT margin", [0.78, 0.72, 0.65],
     "High fixed-cost, asset-light exchange economics mean margin holds best when share/pricing hold; "
     "competitive response (marketing, fee wars) compresses margin further in weaker scenarios."),
    ("Exit P/E multiple", [28, 18, 14],
     "BULL: re-rates as a proven monopoly. BASE: 'show me' multiple, still uncertain. "
     "BEAR: structurally impaired, low-growth multiple."),
]
for i, (name, vals, note) in enumerate(scen_rows):
    r = SCEN_R0 + 1 + i
    label(aw, r, name, indent=1)
    for col, v in zip(["B", "C", "D"], vals):
        fmt = PCT if abs(v) < 1 else MULT
        put(aw, f"{col}{r}", v, BLUE, fmt, fill=YELLOW)
    put(aw, f"F{r}", note, SUB)
    aw[f"F{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    aw.row_dimensions[r].height = 50
    aw.column_dimensions["F"].width = 60
SHARE_R, TAKE_R, MARGIN_R, MULT_R = SCEN_R0 + 1, SCEN_R0 + 2, SCEN_R0 + 3, SCEN_R0 + 4

A = "Assumptions"

# ============================================================ SCENARIOS ====
sw = wb.create_sheet("Scenario Model")
widths(sw, a=44)
put(sw, "A1", "SCENARIO MODEL -- FY28E", TITLE)
put(sw, "A2", "Each column is fully independent and formula-driven off the Assumptions tab.", SUB)

put(sw, "A4", "Scenario", HDR, fill=NAVY)
for col, hdr in zip(["B", "C", "D"], ["BULL", "BASE", "BEAR"]):
    put(sw, f"{col}4", hdr, HDR, fill=NAVY, align="center")

rows = [
    (5, "FY28E total market volume (BU)", f"={A}!$B$21"),
    (6, "x IEX market share", f"={A}!{{c}}${SHARE_R}"),
    (7, "IEX FY28E traded volume (BU)", "={c}5*{c}6"),
    (9, "Current take rate (INR Cr / BU)", f"={A}!$B$16"),
    (10, "x (1 + take rate change)", f"=1+{A}!{{c}}${TAKE_R}"),
    (11, "FY28E take rate (INR Cr / BU)", "={c}9*{c}10"),
    (13, "FY28E revenue (INR Cr)", "={c}7*{c}11"),
    (14, "x PAT margin", f"={A}!{{c}}${MARGIN_R}"),
    (15, "FY28E net profit (INR Cr)", "={c}13*{c}14"),
    (17, "Diluted shares (Cr)", f"={A}!$B$7"),
    (18, "FY28E EPS (INR)", "={c}15/{c}17"),
    (19, "x Exit P/E multiple", f"={A}!{{c}}${MULT_R}"),
    (20, "FY28E TARGET PRICE (INR)", "={c}18*{c}19"),
    (22, "Current price (INR)", f"={A}!$B$5"),
    (23, "IMPLIED RETURN", "={c}20/{c}22-1"),
]
for r, name, f in rows:
    bold = r in (7, 13, 15, 18, 20, 23)
    label(sw, r, name, bold=bold)
    for col in ["B", "C", "D"]:
        formula = f.replace("{c}", col)
        fmt = RS if r in (18, 20, 22) else (PCT if r == 23 else NUM)
        if r in (9, 11):
            fmt = UNITRATE
        put(sw, f"{col}{r}", formula, BOLD if bold else BLACK, fmt,
            border=TOPBORDER if bold else None)

# ============================================================ DOWNSIDE BRIDGE
bw = wb.create_sheet("Downside Bridge")
widths(bw, a=54)
put(bw, "A1", "DOWNSIDE BRIDGE: WHERE THE RETURN COMES FROM", TITLE)
put(bw, "A2", "Decomposes the BASE and BEAR case returns into the EPS-decline effect and the "
              "multiple-compression effect, holding the other constant -- same technique as the LBO "
              "returns bridge in the companion Havells repo, applied here to a short thesis.", SUB)

for col, scen in zip(["B", "C"], ["BASE", "BEAR"]):
    put(bw, f"{col}4", scen, HDR, fill=NAVY, align="center")
put(bw, "A4", "", HDR, fill=NAVY)

scen_col_map = {"BASE": "C", "BEAR": "D"}
rows = [
    (5, "Current EPS (INR)", f"='Scenario Model'!$B$18" if False else None),
    (6, "Scenario EPS (INR)", None),
    (7, "Current P/E (as implied today)", None),
    (8, "Price if EPS falls, multiple UNCHANGED (INR)", None),
    (9, "  return from EPS decline alone", None),
    (10, "Scenario exit multiple", None),
    (11, "Price if multiple ALSO compresses (INR) = full target", None),
    (12, "  additional return from multiple compression", None),
    (13, "TOTAL IMPLIED RETURN (check vs Scenario Model tab)", None),
]
for r, name, _ in rows:
    bold = r in (8, 11, 13)
    label(bw, r, name, bold=bold)

for out_col, scen in zip(["B", "C"], ["BASE", "BEAR"]):
    src_col = scen_col_map[scen]
    put(bw, f"{out_col}5", f"='Assumptions'!$B$11", GREEN, RS)
    put(bw, f"{out_col}6", f"='Scenario Model'!{src_col}18", GREEN, RS)
    put(bw, f"{out_col}7", f"='Assumptions'!$B$12", GREEN, MULT)
    put(bw, f"{out_col}8", f"={out_col}6*{out_col}7", BOLD, RS, border=TOPBORDER)
    put(bw, f"{out_col}9", f"={out_col}8/'Assumptions'!$B$5-1", BLACK, PCT)
    put(bw, f"{out_col}10", f"='Scenario Model'!{src_col}19", GREEN, MULT)
    put(bw, f"{out_col}11", f"={out_col}6*{out_col}10", BOLD, RS, border=TOPBORDER)
    put(bw, f"{out_col}12", f"={out_col}11/{out_col}8-1", BLACK, PCT)
    put(bw, f"{out_col}13", f"={out_col}11/'Assumptions'!$B$5-1", BOLD, PCT, border=TOPDOUBLE)

put(bw, "A16", "READING THIS", BOLD)
put(bw, "A17", "In both scenarios, most of the downside comes from the EPS decline itself (IEX simply "
               "earning less), not from the market additionally punishing the stock with a lower multiple. "
               "That matters for the pitch: you do not need to argue the market will panic and de-rate the "
               "stock unfairly. You only need the earnings decline to happen, and the price follows "
               "mechanically even at an unchanged multiple. The multiple compression is upside to the thesis, "
               "not the load-bearing assumption.", SUB)
bw["A17"].alignment = Alignment(wrap_text=True, vertical="top")
bw.merge_cells("A17:D17")
bw.row_dimensions[17].height = 90

# ============================================================ SENSITIVITY ==
vw = wb.create_sheet("Sensitivity")
widths(vw, a=34, rest=12)
put(vw, "A1", "SENSITIVITY: MARKET SHARE x TAKE-RATE COMPRESSION", TITLE)
put(vw, "A2", "Implied target price (INR) at a fixed 18x exit multiple (the BASE case multiple), "
              "so the grid isolates the operating assumptions from the re-rating assumption.", SUB)
vw["A2"].alignment = Alignment(wrap_text=True)
vw.merge_cells("A2:H2")
vw.row_dimensions[2].height = 30

share_grid = [0.85, 0.75, 0.65, 0.57, 0.50, 0.40]
take_grid = [0.00, -0.10, -0.20, -0.25, -0.30, -0.40]

put(vw, "B5", "Share \\ Take-rate chg", BOLD, fill=GREYFILL, align="center")
for j, t in enumerate(take_grid):
    col = get_column_letter(3 + j)
    put(vw, f"{col}5", f"{t:.0%}", BOLD, align="center", fill=GREYFILL)

for i, sh in enumerate(share_grid):
    r = 6 + i
    put(vw, f"B{r}", f"{sh:.0%}", BOLD, fill=GREYFILL)
    for j, t in enumerate(take_grid):
        col = get_column_letter(3 + j)
        mkt = f"'Assumptions'!$B$21"
        vol = f"({mkt}*{sh})"
        take_rate = f"('Assumptions'!$B$16*(1+{t}))"
        rev = f"({vol}*{take_rate})"
        margin = "0.72"  # held at BASE margin for a clean two-variable grid
        pat = f"({rev}*{margin})"
        eps = f"({pat}/'Assumptions'!$B$7)"
        target = f"({eps}*18)"
        put(vw, f"{col}{r}", f"={target}", BLACK, RS, align="center")

put(vw, "A14", "Current price for reference (INR)", BOLD)
put(vw, "B14", "='Assumptions'!$B$5", GREEN, RS)
put(vw, "A15", "NOTE", BOLD)
put(vw, "A16", "PAT margin held fixed at 72% (the BASE case) throughout this grid so it isolates only "
               "share and take-rate, the two variables the market coupling debate is actually about. "
               "Every cell below the current price row is a scenario in which the short thesis is right.", SUB)
vw["A16"].alignment = Alignment(wrap_text=True)
vw.merge_cells("A16:H16")
vw.row_dimensions[16].height = 40

# =========================================================== CATALYST TIMELINE
cw = wb.create_sheet("Catalyst Timeline")
widths(cw, a=28, rest=20)
put(cw, "A1", "CATALYST TIMELINE", TITLE)
put(cw, "A2", "Dated or approximately-dated events that could force a re-rating in either direction. "
              "None of these dates are confirmed by IEX or CERC as of the build date -- verify each before citing it.", SUB)
cw.column_dimensions["C"].width = 70

put(cw, "A4", "Date / Status", HDR, fill=NAVY)
put(cw, "B4", "Event", HDR, fill=NAVY)
put(cw, "C4", "Why it matters", HDR, fill=NAVY)

events = [
    ("Jul 2025", "CERC's original market coupling order", "The regulatory direction was set here; the stock fell ~30% in a single session."),
    ("Feb 2026", "APTEL (appellate tribunal) dismisses IEX's initial legal challenge", "First sign the legal path favours implementation, not IEX."),
    ("Apr 2026", "CERC issues draft Second Amendment Regulations; Grid India raises its own technical objections", "A regulator's own implementation partner questioning execution readiness is a stronger delay signal than IEX's objections alone."),
    ("Ongoing, 2026", "Supreme Court agrees to examine IEX's plea", "The highest possible legal escalation; a ruling either way is the clearest re-rating catalyst on the table."),
    ("Quarterly", "IEX volume and revenue prints", "Watch for early signs of share loss or fee pressure showing up in the numbers ahead of coupling's actual launch."),
    ("Unconfirmed", "Final CERC notification and Power Market Coupling Procedure (PMCP) publication by Grid India", "The actual implementation date; every delay here is a reason the BULL case can persist longer than the BEAR case assumes."),
]
for i, (date, event, why) in enumerate(events):
    r = 5 + i
    put(cw, f"A{r}", date, BLACK)
    put(cw, f"B{r}", event, BLACK)
    put(cw, f"C{r}", why, BLACK)
    cw[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    cw.row_dimensions[r].height = 34

# ================================================================ COMPS ====
cw2 = wb.create_sheet("Comps")
widths(cw2)
put(cw2, "A1", "COMPARABLE COMPANY CHECK: WHAT MOAT-INTACT INDIAN MARKET INFRASTRUCTURE TRADES AT", TITLE)
put(cw2, "A2", "Sourced 13-Sep-2026. These are the businesses IEX most resembles structurally (near-monopoly, "
               "asset-light, high-margin market infrastructure) BEFORE the market coupling threat. Used to "
               "sanity-check, not set, the exit multiples on the Scenario Model tab.", SUB)
cw2["A2"].alignment = Alignment(wrap_text=True)
cw2.merge_cells("A2:F2")
cw2.row_dimensions[2].height = 40

put(cw2, "A4", "Company", HDR, fill=NAVY)
put(cw2, "B4", "Business", HDR, fill=NAVY)
put(cw2, "C4", "P/E", HDR, fill=NAVY, align="center")
put(cw2, "D4", "Source, 13-Sep-2026", HDR, fill=NAVY)

comps_data = [
    ("MCX (Multi Commodity Exchange)", "India's dominant commodity derivatives exchange, near-monopoly", 56.6,
     "valueresearchonline.com, 29-May-2026 (P/E 56.60x, peer median 60.51x)"),
    ("CDSL (Central Depository Services)", "Depository duopoly, 120M+ demat accounts, recurring AMC/fee income", 60.5,
     "Derived: Screener.in market cap Rs 28,424 Cr / TTM profit Rs 470 Cr"),
    ("CAMS (Computer Age Management Services)", "Mutual fund registrar/transfer agent, dominant market share", 36.1,
     "dhan.co live quote, 11-Sep-2026"),
]
for i, (name, biz, pe, src) in enumerate(comps_data):
    r = 5 + i
    put(cw2, f"A{r}", name, BLACK)
    put(cw2, f"B{r}", biz, BLACK)
    cw2[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    put(cw2, f"C{r}", pe, BLUE, MULT, align="center")
    put(cw2, f"D{r}", src, SUB)
    cw2[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    cw2.row_dimensions[r].height = 32

put(cw2, "A9", "Peer median P/E (moat-intact market infrastructure)", BOLD)
put(cw2, "C9", "=MEDIAN(C5:C7)", BOLD, MULT, border=TOPBORDER)

put(cw2, "A11", "IEX current TTM P/E (at today's price, reflecting the regulatory discount)", BOLD)
put(cw2, "C11", "=Assumptions!B5/Assumptions!B11", BOLD, MULT, border=TOPBORDER)

put(cw2, "A13", "Implied discount to peers already priced in", BOLD)
put(cw2, "C13", "=1-C11/C9", BOLD, PCT, border=TOPDOUBLE)

put(cw2, "A16", "READING THIS AGAINST THE SCENARIO MODEL'S EXIT MULTIPLES", BOLD)
comps_note = ("IEX already trades at a large discount to peers with an intact moat, which is exactly what the "
              "market coupling threat would predict. This is a genuine cross-check on the three exit multiples "
              "used on the Scenario Model tab, not just an assertion: the BULL case (28x) sits well BELOW the "
              "peer median above, meaning that even in the scenario where IEX's moat mostly survives, the model "
              "does not assume IEX re-rates all the way back to where MCX or CDSL trade, a conservative choice, "
              "if anything. The BASE (18x) and BEAR (14x) cases sit further below still, consistent with a "
              "business that has lost the structural characteristic (near-monopoly liquidity) that lets these "
              "peers command 36-60x. If a bull case wanted to argue IEX fully re-rates to the peer median once "
              "coupling is defeated or delayed indefinitely, that would imply a target price meaningfully ABOVE "
              "this model's own bull case: worth stating explicitly as the true upside risk to being short, "
              "rather than treating 28x as the ceiling.")
put(cw2, "A17", comps_note, SUB)
cw2["A17"].alignment = Alignment(wrap_text=True, vertical="top")
cw2.merge_cells("A17:F17")
cw2.row_dimensions[17].height = 130

# --------------------------------------------------------------- save ------
out = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "model", "IEX_Short_Thesis_Model.xlsx"))
os.makedirs(os.path.dirname(out), exist_ok=True)
wb.save(out)
print("Written:", out)
