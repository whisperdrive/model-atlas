"""Write a synthetic three-statement infrastructure model and four altered variants to tests/sample_models/
(git-ignored; re-run to regenerate). They exist to test the statement detector and the identity registry.

A fictional water utility ("Harbourline Water Utility"), annual, financial year ending 30 June, FY2026..FY2040.
Inputs, Revenue, Costs, Capex (with the fixed-assets corkscrew), Debt (sources and uses, then the senior-debt
corkscrew), PnL, Equity (retained earnings), CashFlow, BalanceSheet, Checks and an equity DCF of the distributions.
Every timeline sheet has labels in column A, units in B, the first period in D, and the P&L and cash flow sheets
carry a Total column after the last period. Revenue and cash inflows are positive, costs and outflows negative.

  threeway_model.xlsx       the model; every identity holds (the balance sheet balances, cash ties, sources = uses,
                            debt ties) and Checks says OK
  threeway_unbalanced.xlsx  Closing retained earnings leaves out Distributions, cached values and all, so the balance
                            sheet really does not balance from FY2026 and Checks says ERROR: a genuine model error
  threeway_moved.xlsx       the same numbers with the furniture changed: three memo rows above the P&L block, the P&L
                            sheet renamed, three lines relabelled, no Balance check row, no Checks sheet. Only
                            structure is left to go on
  threeway_stale.xlsx       the Inputs tariff shows 2.40 but every saved result was computed at 2.10: inputs changed,
                            workbook never recalculated
  threeway_assets.xlsx      Revenue split into RevenueNorth (60%) and RevenueSouth (40%) with a Revenue sheet that
                            consolidates them; totals equal the base

Every formula is written with its cached value computed here, so build_map (which reads saved values) sees the
numbers Excel would have saved. The opening balance sheet is pro forma for the financial close: cash 25 plus fixed
assets 650 equals senior debt 400 plus share capital 200 plus the 100 contribution plus retained earnings of -25 (the
40 of transaction costs is taken through retained earnings; the 460 refinanced facility is repaid out of the 400 debt
and 100 equity raised, with the 40 of costs). That is why Opening retained earnings is -25.
    uv run python tests/make_threeway_model.py
"""
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as COL

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from modelatlas import dcf  # noqa: E402

OUT = ROOT / "tests" / "sample_models"
K = 15            # FY2026 .. FY2040
FIRST_COL = 3     # column D
FY0 = 2026
TOTAL_COL = FIRST_COL + K   # column S
START_ROW = 7     # first data row on a timeline sheet (1-based)
VALUE_COL = 2     # column C, where Inputs and Checks keep their numbers
SNAP = 1e-9

# (key, label, value, unit, number format)
INPUTS = [
    ("val_date", "Valuation date", date(2025, 6, 30), "date", "date"),
    ("open_vol", "Opening volume", 42000.0, "ML", "num"),
    ("vol_growth", "Volume growth", 0.015, "%", "pct"),
    ("open_tariff", "Opening tariff", 2.10, "$/kL", "tariff"),
    ("cpi", "CPI", 0.025, "%", "pct"),
    ("res_share", "Residential share", 0.6, "%", "pct"),
    ("fixed_opex", "Fixed opex", 18.0, "$m p.a.", "num"),
    ("var_opex", "Variable opex (% of revenue)", 0.22, "%", "pct"),
    ("growth_capex", "Growth capex", 12.0, "$m p.a.", "num"),
    ("refurb", "Major refurbishment", 40.0, "$m", "num"),
    ("refurb_y1", "Refurbishment year 1", 2031, "year", "year"),
    ("refurb_y2", "Refurbishment year 2", 2036, "year", "year"),
    ("life", "Depreciation life", 25.0, "years", "num"),
    ("debt_drawn", "Senior debt drawn at financial close", 400.0, "$m", "num"),
    ("rate", "Interest rate", 0.055, "%", "pct"),
    ("tenor", "Debt tenor", 12, "years", "year"),
    ("first_repay", "First repayment year", 2027, "year", "year"),
    ("tax", "Tax rate", 0.30, "%", "pct"),
    ("payout", "Distribution payout", 0.80, "%", "pct"),
    ("open_cash", "Opening cash", 25.0, "$m", "num"),
    ("open_fa", "Opening fixed assets", 650.0, "$m", "num"),
    ("share_cap", "Share capital", 200.0, "$m", "num"),
    ("open_re", "Opening retained earnings", -25.0, "$m", "num"),
    ("eq_contrib", "Equity contribution", 100.0, "$m", "num"),
    ("refi", "Refinanced facility", 460.0, "$m", "num"),
    ("costs", "Transaction costs", 40.0, "$m", "num"),
    ("disc", "Discount rate", 0.085, "%", "pct"),
    ("tg", "Terminal growth", 0.02, "%", "pct"),
]
ASSET_INPUTS = [("north_share", "North volume share", 0.6, "%", "pct"),
                ("south_share", "South volume share", 0.4, "%", "pct")]
NAMES = {"val_date": "Val_date", "disc": "Disc_rate", "tax": "Tax_rate", "cpi": "CPI"}
RELABEL = {"Revenue": "Turnover", "NPAT": "Net profit after tax", "Total revenue": "Revenue total"}


def snap(x: float) -> float:
    return 0.0 if abs(x) < SNAP else x


def q(name: str) -> str:
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) else "'" + name.replace("'", "''") + "'"


# ---- the numbers -----------------------------------------------------------------------------------------

def compute(P: dict, unbalanced: bool = False, assets: bool = False) -> dict:
    """Every row's values, by (sheet, key). P maps input key to value."""
    V: dict = {}
    ends = [date(FY0 + k, 6, 30) for k in range(K)]

    def segment(sheet: str, share: float | None) -> None:
        vol, tar = [], []
        for k in range(K):
            v0 = P["open_vol"] * (share if share is not None else 1.0)
            vol.append(v0 if k == 0 else vol[-1] * (1 + P["vol_growth"]))
            tar.append(P["open_tariff"] if k == 0 else tar[-1] * (1 + P["cpi"]))
        res = [v * t * P["res_share"] / 1000 for v, t in zip(vol, tar)]
        non = [v * t * (1 - P["res_share"]) / 1000 for v, t in zip(vol, tar)]
        V.update({(sheet, "vol"): vol, (sheet, "tariff"): tar, (sheet, "res"): res, (sheet, "non"): non,
                  (sheet, "total"): [a + b for a, b in zip(res, non)]})

    if assets:
        segment("RevenueNorth", P["north_share"])
        segment("RevenueSouth", P["south_share"])
        N, S = "RevenueNorth", "RevenueSouth"
        for key in ("vol", "res", "non"):
            V[("Revenue", key)] = [a + b for a, b in zip(V[(N, key)], V[(S, key)])]
        V[("Revenue", "total")] = [a + b for a, b in zip(V[("Revenue", "res")], V[("Revenue", "non")])]
    else:
        segment("Revenue", None)
    rev = V[("Revenue", "total")]

    fixed, var = [], []
    for k in range(K):
        fixed.append(-P["fixed_opex"] if k == 0 else fixed[-1] * (1 + P["cpi"]))
        var.append(-rev[k] * P["var_opex"])
    V.update({("Costs", "fixed"): fixed, ("Costs", "var"): var, ("Costs", "total"): [a + b for a, b in zip(fixed, var)]})

    growth, refurb = [], []
    for k in range(K):
        growth.append(-P["growth_capex"] if k == 0 else growth[-1] * (1 + P["cpi"]))
        refurb.append(-P["refurb"] if FY0 + k in (P["refurb_y1"], P["refurb_y2"]) else 0.0)
    tcap = [a + b for a, b in zip(growth, refurb)]
    fa_open, fa_add, fa_dep, fa_close = [], [], [], []
    for k in range(K):
        fa_open.append(P["open_fa"] if k == 0 else fa_close[-1])
        fa_add.append(-tcap[k])
        fa_dep.append(-fa_open[k] / P["life"])
        fa_close.append(fa_open[k] + fa_add[k] + fa_dep[k])
    V.update({("Capex", "growth"): growth, ("Capex", "refurb"): refurb, ("Capex", "total"): tcap,
              ("Capex", "open"): fa_open, ("Capex", "add"): fa_add, ("Capex", "dep"): fa_dep,
              ("Capex", "close"): fa_close})

    d_open, d_draw, d_rep, d_close, d_int = [], [], [], [], []
    for k in range(K):
        d_open.append(0.0 if k == 0 else d_close[-1])
        d_draw.append(P["debt_drawn"] if k == 0 else 0.0)
        y = FY0 + k
        d_rep.append(-P["debt_drawn"] / P["tenor"] if P["first_repay"] <= y < P["first_repay"] + P["tenor"] else 0.0)
        d_close.append(d_open[k] + d_draw[k] + d_rep[k])
        d_int.append(-d_open[k] * P["rate"])
    sources = P["debt_drawn"] + P["eq_contrib"]
    uses = P["refi"] + P["costs"]
    V.update({("Debt", "open"): d_open, ("Debt", "draw"): d_draw, ("Debt", "rep"): d_rep, ("Debt", "close"): d_close,
              ("Debt", "int"): d_int, ("Debt", "su_debt"): [P["debt_drawn"]], ("Debt", "su_eq"): [P["eq_contrib"]],
              ("Debt", "su_src"): [sources], ("Debt", "su_refi"): [P["refi"]], ("Debt", "su_costs"): [P["costs"]],
              ("Debt", "su_uses"): [uses], ("Debt", "su_chk"): [snap(sources - uses)]})

    opc = V[("Costs", "total")]
    ebitda = [r + o for r, o in zip(rev, opc)]
    ebit = [e + d for e, d in zip(ebitda, fa_dep)]
    pbt = [e + i for e, i in zip(ebit, d_int)]
    tax = [-max(0.0, p) * P["tax"] for p in pbt]
    npat = [p + t for p, t in zip(pbt, tax)]
    V.update({("PnL", "rev"): rev, ("PnL", "opc"): opc, ("PnL", "ebitda"): ebitda, ("PnL", "dep"): fa_dep,
              ("PnL", "ebit"): ebit, ("PnL", "int"): d_int, ("PnL", "pbt"): pbt, ("PnL", "tax"): tax,
              ("PnL", "npat"): npat})

    r_open, dist, r_close = [], [], []
    for k in range(K):
        r_open.append(P["open_re"] if k == 0 else r_close[-1])
        dist.append(-P["payout"] * max(0.0, npat[k]))
        r_close.append(r_open[k] + npat[k] + (0.0 if unbalanced else dist[k]))
    V.update({("Equity", "open"): r_open, ("Equity", "npat"): npat, ("Equity", "dist"): dist,
              ("Equity", "close"): r_close})

    c_open, c_close, cfo, cfi, cff, net = [], [], [], [], [], []
    eq_in = [P["eq_contrib"] if k == 0 else 0.0 for k in range(K)]
    refi = [-uses if k == 0 else 0.0 for k in range(K)]
    for k in range(K):
        c_open.append(P["open_cash"] if k == 0 else c_close[-1])
        cfo.append(ebitda[k] + tax[k])
        cfi.append(tcap[k])
        cff.append(d_draw[k] + d_rep[k] + d_int[k] + dist[k] + eq_in[k] + refi[k])
        net.append(cfo[k] + cfi[k] + cff[k])
        c_close.append(c_open[k] + net[k])
    V.update({("CashFlow", "open"): c_open, ("CashFlow", "ebitda"): ebitda, ("CashFlow", "tax"): tax,
              ("CashFlow", "cfo"): cfo, ("CashFlow", "capex"): tcap, ("CashFlow", "cfi"): cfi,
              ("CashFlow", "draw"): d_draw, ("CashFlow", "rep"): d_rep, ("CashFlow", "int"): d_int,
              ("CashFlow", "dist"): dist, ("CashFlow", "eq"): eq_in, ("CashFlow", "refi"): refi,
              ("CashFlow", "cff"): cff, ("CashFlow", "net"): net, ("CashFlow", "close"): c_close})

    sc = [P["share_cap"] + P["eq_contrib"]] * K
    ta = [c + f for c, f in zip(c_close, fa_close)]
    te = [s + r for s, r in zip(sc, r_close)]
    tle = [d + e for d, e in zip(d_close, te)]
    V.update({("BalanceSheet", "cash"): c_close, ("BalanceSheet", "fa"): fa_close, ("BalanceSheet", "ta"): ta,
              ("BalanceSheet", "debt"): d_close, ("BalanceSheet", "tl"): d_close, ("BalanceSheet", "sc"): sc,
              ("BalanceSheet", "re"): r_close, ("BalanceSheet", "te"): te, ("BalanceSheet", "tle"): tle,
              ("BalanceSheet", "chk"): [snap(a - b) for a, b in zip(ta, tle)]})

    chk = [sum(abs(x) for x in V[("BalanceSheet", "chk")]),
           sum(abs(snap(a - b)) for a, b in zip(c_close, c_close)),
           abs(V[("Debt", "su_chk")][0]),
           sum(abs(snap(a - b)) for a, b in zip(d_close, d_close))]
    allc = sum(chk)
    V.update({("Checks", "bs"): [chk[0]], ("Checks", "cash"): [chk[1]], ("Checks", "su"): [chk[2]],
              ("Checks", "debt"): [chk[3]], ("Checks", "all"): [allc],
              ("Checks", "ok"): ["OK" if snap(allc) == 0 else "ERROR"]})

    vd = P["val_date"]
    pos = [-x for x in dist]
    fac = [1 / (1 + P["disc"]) ** dcf.yearfrac(vd, e, "actual/actual") for e in ends]
    pv = [a * b for a, b in zip(pos, fac)]
    V.update({("DCF", "dist"): pos, ("DCF", "df"): fac, ("DCF", "pv"): pv,
              ("DCF", "equity"): [sum(a * b for a, b in zip(pos, fac))]})
    return V


# ---- the layout ------------------------------------------------------------------------------------------

@dataclass
class Row:
    key: str = ""
    label: str = ""
    unit: str = ""
    fn: Callable | None = None     # fn(ctx) -> formula text (str) or a constant number
    kind: str = "row"              # row | section | blank | memo
    bold: bool = False
    fmt: str = "num"
    total: str | None = "sum"      # sum | first | last | None: what the Total column holds
    only0: bool = False            # a single cell (Debt sources and uses, DCF equity value, Checks)
    skip: bool = False


@dataclass
class Sheet:
    key: str
    title: str
    rows: list = field(default_factory=list)
    timeline: bool = True
    total_col: bool = False
    cell_col: int = FIRST_COL
    layout_row: dict = field(default_factory=dict)


class Model:
    def __init__(self, sheets: list[Sheet], names: dict[str, str], inputs: list):
        self.sheets = {s.key: s for s in sheets}
        self.order = [s.key for s in sheets]
        self.names = names
        self.inputs = inputs
        self.in_row = {k: 2 + i + 1 for i, (k, *_r) in enumerate(inputs)}   # Inputs data from row 3 (1-based)
        self.rows: dict[tuple[str, str], int] = {}
        for s in sheets:
            r = START_ROW
            for row in s.rows:
                if row.skip:
                    continue
                if row.key:
                    self.rows[(s.key, row.key)] = r
                r += 1


class Ctx:
    """Formula-building view of one column of one sheet."""

    def __init__(self, M: Model, sheet: str, k: int, col: int):
        self.M, self.sheet, self.k = M, sheet, k
        self.c, self.p = COL(col), COL(col - 1)

    def row(self, sheet: str, key: str) -> int:
        return self.M.rows[(sheet, key)]

    def loc(self, key: str) -> str:
        return f"{self.c}{self.row(self.sheet, key)}"

    def prev(self, key: str) -> str:
        return f"{self.p}{self.row(self.sheet, key)}"

    def x(self, sheet: str, key: str) -> str:
        return f"{q(self.M.names[sheet])}!{self.c}{self.row(sheet, key)}"

    def span(self, sheet: str, key: str, local: bool = False) -> str:
        r = self.row(sheet, key)
        rng = f"{COL(FIRST_COL)}{r}:{COL(FIRST_COL + K - 1)}{r}"
        return rng if local else f"{q(self.M.names[sheet])}!{rng}"

    def cell(self, sheet: str, key: str, local: bool = False) -> str:
        a = f"${COL(self.M.sheets[sheet].cell_col)}${self.row(sheet, key)}"
        return a if local else f"{q(self.M.names[sheet])}!{a}"

    def inp(self, key: str) -> str:
        return NAMES.get(key) or f"Inputs!$C${self.M.in_row[key]}"

    def year(self) -> str:
        return f"YEAR({self.c}$3)"


def build_sheets(variant: str) -> tuple[list[Sheet], dict[str, str], list]:
    moved, assets = variant == "moved", variant == "assets"
    R = Row

    def sec(label):
        return R(label=label, kind="section")

    def segment_rows(sheet: str, share_key: str | None) -> list[Row]:
        def vol(x):
            base = x.inp("open_vol") + (f"*{x.inp(share_key)}" if share_key else "")
            return f"={base}" if x.k == 0 else f"={x.prev('vol')}*(1+{x.inp('vol_growth')})"
        return [
            R("vol", "Volume", "ML", vol, total=None),
            R("tariff", "Tariff", "$/kL",
              lambda x: f"={x.inp('open_tariff')}" if x.k == 0 else f"={x.prev('tariff')}*(1+CPI)",
              fmt="tariff", total=None),
            R("res", "Residential revenue", "$m",
              lambda x: f"={x.loc('vol')}*{x.loc('tariff')}*{x.inp('res_share')}/1000", total=None),
            R("non", "Non-residential revenue", "$m",
              lambda x: f"={x.loc('vol')}*{x.loc('tariff')}*(1-{x.inp('res_share')})/1000", total=None),
            R("total", "Total revenue", "$m", lambda x: f"=SUM({x.loc('res')}:{x.loc('non')})", bold=True, total=None),
        ]

    sheets: list[Sheet] = []
    if assets:
        sheets.append(Sheet("RevenueNorth", "Revenue - North", segment_rows("RevenueNorth", "north_share")))
        sheets.append(Sheet("RevenueSouth", "Revenue - South", segment_rows("RevenueSouth", "south_share")))
        both = lambda key: (lambda x: f"={x.x('RevenueNorth', key)}+{x.x('RevenueSouth', key)}")  # noqa: E731
        sheets.append(Sheet("Revenue", "Revenue (consolidated)", [
            R("vol", "Volume", "ML", both("vol"), total=None),
            R("res", "Residential revenue", "$m", both("res"), total=None),
            R("non", "Non-residential revenue", "$m", both("non"), total=None),
            R("total", "Total revenue", "$m", lambda x: f"=SUM({x.loc('res')}:{x.loc('non')})", bold=True,
              total=None)]))
    else:
        sheets.append(Sheet("Revenue", "Revenue", segment_rows("Revenue", None)))

    sheets.append(Sheet("Costs", "Operating costs", [
        R("fixed", "Fixed opex", "$m",
          lambda x: f"=-{x.inp('fixed_opex')}" if x.k == 0 else f"={x.prev('fixed')}*(1+CPI)", total=None),
        R("var", "Variable opex", "$m",
          lambda x: f"=-{x.x('Revenue', 'total')}*{x.inp('var_opex')}", total=None),
        R("total", "Total operating costs", "$m", lambda x: f"=SUM({x.loc('fixed')}:{x.loc('var')})", bold=True,
          total=None)]))

    sheets.append(Sheet("Capex", "Capital expenditure and fixed assets", [
        R("growth", "Growth capex", "$m",
          lambda x: f"=-{x.inp('growth_capex')}" if x.k == 0 else f"={x.prev('growth')}*(1+CPI)", total=None),
        R("refurb", "Major refurbishment", "$m",
          lambda x: (f"=-IF(OR({x.year()}={x.inp('refurb_y1')},{x.year()}={x.inp('refurb_y2')}),"
                     f"{x.inp('refurb')},0)"), total=None),
        R("total", "Total capex", "$m", lambda x: f"=SUM({x.loc('growth')}:{x.loc('refurb')})", bold=True,
          total=None),
        R(kind="blank"),
        sec("Fixed assets"),
        R("open", "Opening fixed assets", "$m",
          lambda x: f"={x.inp('open_fa')}" if x.k == 0 else f"={x.prev('close')}", total=None),
        R("add", "Additions", "$m", lambda x: f"=-{x.loc('total')}", total=None),
        R("dep", "Depreciation", "$m", lambda x: f"=-{x.loc('open')}/{x.inp('life')}", total=None),
        R("close", "Closing fixed assets", "$m", lambda x: f"=SUM({x.loc('open')}:{x.loc('dep')})", bold=True,
          total=None)]))

    one = lambda fn: dict(fn=fn, only0=True, total=None)  # noqa: E731
    sheets.append(Sheet("Debt", "Debt", [
        sec("Sources and uses"),
        sec("Sources"),
        R("su_debt", "Senior debt", "$m", **one(lambda x: f"={x.inp('debt_drawn')}")),
        R("su_eq", "Equity contribution", "$m", **one(lambda x: f"={x.inp('eq_contrib')}")),
        R("su_src", "Total sources", "$m", bold=True,
          **one(lambda x: f"=SUM({x.loc('su_debt')}:{x.loc('su_eq')})")),
        sec("Uses"),
        R("su_refi", "Refinanced facility", "$m", **one(lambda x: f"={x.inp('refi')}")),
        R("su_costs", "Transaction costs", "$m", **one(lambda x: f"={x.inp('costs')}")),
        R("su_uses", "Total uses", "$m", bold=True,
          **one(lambda x: f"=SUM({x.loc('su_refi')}:{x.loc('su_costs')})")),
        R("su_chk", "S&U check", "$m", **one(lambda x: f"={x.loc('su_src')}-{x.loc('su_uses')}")),
        R(kind="blank"),
        sec("Senior debt"),
        R("open", "Opening balance", "$m",
          lambda x: 0 if x.k == 0 else f"={x.prev('close')}", total=None),
        R("draw", "Drawdown", "$m",
          lambda x: f"=${COL(FIRST_COL)}${x.row('Debt', 'su_debt')}" if x.k == 0 else 0, total=None),
        R("rep", "Repayment", "$m",
          lambda x: (f"=-IF(AND({x.year()}>={x.inp('first_repay')},{x.year()}<{x.inp('first_repay')}+"
                     f"{x.inp('tenor')}),{x.inp('debt_drawn')}/{x.inp('tenor')},0)"), total=None),
        R("close", "Closing balance", "$m", lambda x: f"=SUM({x.loc('open')}:{x.loc('rep')})", bold=True,
          total=None),
        R("int", "Interest", "$m", lambda x: f"=-{x.loc('open')}*{x.inp('rate')}", total=None)]))

    pnl_rows = [
        R("rev", "Revenue", "$m", lambda x: f"={x.x('Revenue', 'total')}"),
        R("opc", "Operating costs", "$m", lambda x: f"={x.x('Costs', 'total')}"),
        R("ebitda", "EBITDA", "$m", lambda x: f"={x.loc('rev')}+{x.loc('opc')}", bold=True),
        R("dep", "Depreciation", "$m", lambda x: f"={x.x('Capex', 'dep')}"),
        R("ebit", "EBIT", "$m", lambda x: f"={x.loc('ebitda')}+{x.loc('dep')}", bold=True),
        R("int", "Interest", "$m", lambda x: f"={x.x('Debt', 'int')}"),
        R("pbt", "PBT", "$m", lambda x: f"={x.loc('ebit')}+{x.loc('int')}", bold=True),
        R("tax", "Tax", "$m", lambda x: f"=-MAX(0,{x.loc('pbt')})*Tax_rate"),
        R("npat", "NPAT", "$m", lambda x: f"={x.loc('pbt')}+{x.loc('tax')}", bold=True)]
    if moved:
        pnl_rows = [R(label=t, kind="memo") for t in
                    ("Draft - management case, not for distribution", "Memo: costs shown as negatives",
                     "Memo: figures in $m unless stated")] + pnl_rows
    sheets.append(Sheet("PnL", "Profit and loss", pnl_rows, total_col=True))

    sheets.append(Sheet("Equity", "Equity", [
        sec("Retained earnings"),
        R("open", "Opening retained earnings", "$m",
          lambda x: f"={x.inp('open_re')}" if x.k == 0 else f"={x.prev('close')}", total=None),
        R("npat", "NPAT", "$m", lambda x: f"={x.x('PnL', 'npat')}", total=None),
        R("dist", "Distributions", "$m",
          lambda x: f"=-{x.inp('payout')}*MAX(0,{x.loc('npat')})", total=None),
        R("close", "Closing retained earnings", "$m",
          (lambda x: f"={x.loc('open')}+{x.loc('npat')}") if variant == "unbalanced"
          else (lambda x: f"=SUM({x.loc('open')}:{x.loc('dist')})"), bold=True, total=None)]))

    sheets.append(Sheet("CashFlow", "Cash flow", [
        R("open", "Opening cash", "$m",
          lambda x: f"={x.inp('open_cash')}" if x.k == 0 else f"={x.prev('close')}", total="first"),
        sec("Operating"),
        R("ebitda", "EBITDA", "$m", lambda x: f"={x.x('PnL', 'ebitda')}"),
        R("tax", "Tax paid", "$m", lambda x: f"={x.x('PnL', 'tax')}"),
        R("cfo", "Cash flow from operations", "$m", lambda x: f"=SUM({x.loc('ebitda')}:{x.loc('tax')})", bold=True),
        sec("Investing"),
        R("capex", "Capital expenditure", "$m", lambda x: f"={x.x('Capex', 'total')}"),
        R("cfi", "Cash flow from investing", "$m", lambda x: f"={x.loc('capex')}", bold=True),
        sec("Financing"),
        R("draw", "Drawdowns", "$m", lambda x: f"={x.x('Debt', 'draw')}"),
        R("rep", "Repayments", "$m", lambda x: f"={x.x('Debt', 'rep')}"),
        R("int", "Interest paid", "$m", lambda x: f"={x.x('Debt', 'int')}"),
        R("dist", "Distributions paid", "$m", lambda x: f"={x.x('Equity', 'dist')}"),
        R("eq", "Equity contribution", "$m",
          lambda x: f"={x.cell('Debt', 'su_eq')}" if x.k == 0 else 0),
        R("refi", "Refinancing and transaction costs", "$m",
          lambda x: f"=-{x.cell('Debt', 'su_uses')}" if x.k == 0 else 0),
        R("cff", "Cash flow from financing", "$m", lambda x: f"=SUM({x.loc('draw')}:{x.loc('refi')})", bold=True),
        R("net", "Net cash flow", "$m",
          lambda x: f"={x.loc('cfo')}+{x.loc('cfi')}+{x.loc('cff')}", bold=True),
        R("close", "Closing cash", "$m", lambda x: f"={x.loc('open')}+{x.loc('net')}", bold=True, total="last")],
        total_col=True))

    bs = [
        sec("Assets"),
        R("cash", "Cash", "$m", lambda x: f"={x.x('CashFlow', 'close')}", total=None),
        R("fa", "Fixed assets", "$m", lambda x: f"={x.x('Capex', 'close')}", total=None),
        R("ta", "Total assets", "$m", lambda x: f"=SUM({x.loc('cash')}:{x.loc('fa')})", bold=True, total=None),
        sec("Liabilities"),
        R("debt", "Senior debt", "$m", lambda x: f"={x.x('Debt', 'close')}", total=None),
        R("tl", "Total liabilities", "$m", lambda x: f"=SUM({x.loc('debt')}:{x.loc('debt')})", bold=True, total=None),
        sec("Equity"),
        R("sc", "Share capital", "$m", lambda x: f"={x.inp('share_cap')}+{x.inp('eq_contrib')}", total=None),
        R("re", "Retained earnings", "$m", lambda x: f"={x.x('Equity', 'close')}", total=None),
        R("te", "Total equity", "$m", lambda x: f"=SUM({x.loc('sc')}:{x.loc('re')})", bold=True, total=None),
        R("tle", "Total liabilities and equity", "$m", lambda x: f"={x.loc('tl')}+{x.loc('te')}", bold=True,
          total=None),
        R("chk", "Balance check", "$m", lambda x: f"={x.loc('ta')}-{x.loc('tle')}", total=None, skip=moved)]
    sheets.append(Sheet("BalanceSheet", "Balance sheet", bs))

    if not moved:
        def abs_sum(expr):
            return lambda x: f"=SUMPRODUCT(ABS({expr(x)}))"
        z = lambda fn: dict(fn=fn, only0=True, total=None)  # noqa: E731
        sheets.append(Sheet("Checks", "Model checks", [
            R("bs", "Balance sheet balances", "$m",
              **z(abs_sum(lambda x: x.span("BalanceSheet", "chk")))),
            R("cash", "Cash ties to balance sheet", "$m",
              **z(abs_sum(lambda x: f"{x.span('CashFlow', 'close')}-{x.span('BalanceSheet', 'cash')}"))),
            R("su", "Sources equal uses", "$m", **z(lambda x: f"=ABS({x.cell('Debt', 'su_chk')})")),
            R("debt", "Debt ties to balance sheet", "$m",
              **z(abs_sum(lambda x: f"{x.span('Debt', 'close')}-{x.span('BalanceSheet', 'debt')}"))),
            R("all", "All checks", "$m", bold=True, **z(lambda x: f"=SUM({x.loc('bs')}:{x.loc('debt')})")),
            R("ok", "Model OK", "", **z(lambda x: f'=IF({x.loc("all")}=0,"OK","ERROR")'))],
            timeline=False, cell_col=VALUE_COL))

    sheets.append(Sheet("DCF", "Equity DCF of distributions", [
        R("dist", "Distributions", "$m", lambda x: f"=-{x.x('Equity', 'dist')}", total=None),
        R("df", "Discount factor", "x", lambda x: f"=1/(1+Disc_rate)^YEARFRAC(Val_date,{x.c}$3,1)", fmt="factor",
          total=None),
        R("pv", "Present value", "$m", lambda x: f"={x.loc('dist')}*{x.loc('df')}", total=None),
        R(kind="blank"),
        R("equity", "Equity value", "$m", bold=True,
          **one(lambda x: f"=SUMPRODUCT({x.span('DCF', 'dist', True)},{x.span('DCF', 'df', True)})"))]))

    names = {s.key: s.key for s in sheets}
    if moved:
        names["PnL"] = "Income statement"
    inputs = INPUTS + (ASSET_INPUTS if assets else [])
    return sheets, names, inputs


# ---- the workbook ----------------------------------------------------------------------------------------

def write(path: Path, variant: str) -> tuple[Model, dict]:
    P = {k: v for k, _l, v, _u, _f in INPUTS + ASSET_INPUTS}
    shown_tariff = P["open_tariff"]
    if variant == "stale":
        shown_tariff = 2.40   # the Inputs cell moved on; the saved results did not
    V = compute(P, unbalanced=variant == "unbalanced", assets=variant == "assets")
    sheets, names, inputs = build_sheets(variant)
    M = Model(sheets, names, inputs)

    wb = xlsxwriter.Workbook(path)
    fmts: dict = {}

    def fm(kind: str, bold: bool = False):
        key = (kind, bold)
        if key not in fmts:
            nf = {"num": "#,##0.0;(#,##0.0);-", "pct": "0.00%", "tariff": "0.000", "factor": "0.0000",
                  "date": "dd-mmm-yy", "year": "0", "text": "General"}[kind]
            fmts[key] = wb.add_format({"num_format": nf, "bold": bold})
        return fmts[key]

    title, head = wb.add_format({"bold": True}), wb.add_format({"italic": True, "font_color": "#747480"})
    bold = wb.add_format({"bold": True})
    wi = wb.add_worksheet("Inputs")
    ws = {key: wb.add_worksheet(names[key]) for key in M.order}
    wi.write(0, 0, "Harbourline Water Utility - inputs", title)
    for i, (key, label, value, unit, nf) in enumerate(inputs):
        r = 2 + i
        wi.write(r, 0, label)
        wi.write(r, 1, unit)
        if isinstance(value, date):
            wi.write_datetime(r, VALUE_COL, value, fm("date"))
        else:
            wi.write_number(r, VALUE_COL, shown_tariff if key == "open_tariff" else value, fm(nf))
        if NAMES.get(key):
            wb.define_name(NAMES[key], f"=Inputs!$C${r + 1}")
    wi.write(M.in_row["open_re"] - 1, 3, "Pro forma for the transaction costs taken at financial close")

    ends = [date(FY0 + k, 6, 30) for k in range(K)]
    for key in M.order:
        s, w = M.sheets[key], ws[key]
        w.write(0, 0, f"{s.title} - Harbourline Water Utility", title)
        w.set_column(0, 0, 34)
        if s.timeline:
            w.write(2, 0, "Period ending", bold)
            w.write(3, 0, "Financial year")
            w.write(4, 0, "Period flag")
            w.write(4, 1, "flag")
            for k in range(K):
                w.write_datetime(2, FIRST_COL + k, ends[k], fm("date"))
                w.write(3, FIRST_COL + k, f"FY{FY0 + k}")
                w.write_number(4, FIRST_COL + k, 1)
            if s.total_col:
                w.write(3, TOTAL_COL, "Total", bold)
        r = START_ROW
        for row in s.rows:
            if row.skip:
                continue
            if row.kind == "blank":
                r += 1
                continue
            label = RELABEL.get(row.label, row.label) if variant == "moved" else row.label
            if row.kind in ("section", "memo"):
                w.write(r - 1, 0, label, head)
                r += 1
                continue
            w.write(r - 1, 0, label, fm("text", row.bold))
            w.write(r - 1, 1, row.unit)
            vals = V[(key, row.key)]
            cols = [s.cell_col] if row.only0 else [FIRST_COL + k for k in range(K)]
            for k, col in enumerate(cols):
                got = row.fn(Ctx(M, key, k, col))
                nf = "text" if isinstance(vals[k], str) else row.fmt
                if isinstance(got, str):
                    w.write_formula(r - 1, col, got, fm(nf, row.bold), vals[k])
                else:
                    w.write_number(r - 1, col, got, fm(nf, row.bold))
                    assert abs(got - vals[k]) < 1e-12, (key, row.key, k)
            if s.total_col and row.total and not row.only0:
                a, b = COL(FIRST_COL) + str(r), COL(FIRST_COL + K - 1) + str(r)
                f, v = {"sum": (f"=SUM({a}:{b})", sum(vals)), "first": (f"={a}", vals[0]),
                        "last": (f"={b}", vals[-1])}[row.total]
                w.write_formula(r - 1, TOTAL_COL, f, fm(row.fmt, row.bold), v)
            r += 1
    wb.close()
    return M, V


def build(out_dir: Path) -> dict[str, Path]:
    """Write the base model and its four variants into out_dir. Returns {"base", "unbalanced", "moved", "stale",
    "assets"} -> path. Asserts the base model's identities hold exactly and the unbalanced one's do not."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    files = {"base": "threeway_model.xlsx", "unbalanced": "threeway_unbalanced.xlsx",
             "moved": "threeway_moved.xlsx", "stale": "threeway_stale.xlsx", "assets": "threeway_assets.xlsx"}
    paths, models = {}, {}
    for variant, name in files.items():
        paths[variant] = out_dir / name
        models[variant] = write(paths[variant], "base" if variant == "base" else variant)

    P = {k: v for k, _l, v, _u, _f in INPUTS}
    # the opening balance sheet, pro forma for the close
    assert P["open_cash"] + P["open_fa"] == (P["debt_drawn"] + P["share_cap"] + P["eq_contrib"] + P["open_re"])
    V = models["base"][1]
    assert all(v == 0 for v in V[("BalanceSheet", "chk")]), "balance sheet must balance every period"
    assert V[("Checks", "all")] == [0.0] and V[("Checks", "ok")] == ["OK"]
    assert V[("Debt", "su_chk")] == [0.0]
    assert all(abs(a - b) < SNAP for a, b in zip(V[("CashFlow", "close")], V[("BalanceSheet", "cash")]))
    assert all(abs(a - b) < SNAP for a, b in zip(V[("Debt", "close")], V[("BalanceSheet", "debt")]))
    assert abs(V[("Debt", "close")][-1]) < SNAP, "the senior debt is fully repaid by FY2040"
    assert all(n > 0 for n in V[("PnL", "npat")][:3]), "distributions are paid, so dropping them unbalances"
    U = models["unbalanced"][1]
    assert U[("BalanceSheet", "chk")][0] != 0 and U[("Checks", "ok")] == ["ERROR"]
    assert all(c != 0 for c in U[("BalanceSheet", "chk")][:3])
    A = models["assets"][1]
    for key in ("vol", "res", "non", "total"):
        assert all(abs(a - b) < 1e-9 for a, b in zip(A[("Revenue", key)], V[("Revenue", key)])), key
    S = models["stale"][1]
    assert S[("PnL", "npat")] == V[("PnL", "npat")], "stale results are the base results"
    return paths


def main() -> None:
    paths = build(OUT)
    for role, p in paths.items():
        print(f"{role:12s} {p.relative_to(ROOT)}")
    M, V = write(OUT / "threeway_model.xlsx", "base")
    print("\nbase: bold total rows and key cells (1-based rows)")
    for key in M.order:
        s = M.sheets[key]
        bolds = [f"{row.label} r{M.rows[(key, row.key)]}" for row in s.rows if row.bold and not row.skip]
        print(f"  {M.names[key]:14s} " + "; ".join(bolds))
    print(f"  DCF equity value  DCF!D{M.rows[('DCF', 'equity')]} = {V[('DCF', 'equity')][0]:,.2f}")
    print(f"  Checks All checks Checks!C{M.rows[('Checks', 'all')]}; Model OK Checks!C{M.rows[('Checks', 'ok')]}")
    print(f"  Balance check     BalanceSheet!D{M.rows[('BalanceSheet', 'chk')]}:R{M.rows[('BalanceSheet', 'chk')]}")


if __name__ == "__main__":
    main()
