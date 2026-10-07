"""The three statements (modelatlas/threeway.py) on the three-way fixture, a schedules-only variant, the Riverbend client
model, a quarterly workbook and a workbook with nothing in it.

  base            all three statements `extracted`, every check holds, NPAT is the PnL sheet's NPAT row, closing cash is
                  the balance sheet's cash, the financial-year view equals the period view for an annual model
  unbalanced      the balance sheet check fails with the residual the fixture wrote, and the residual analysis names the
                  distributions: the cash flow pays them, retained earnings never move by them
  moved           the same statements `extracted` from the renamed sheet
  schedules_only  no CashFlow, BalanceSheet or Checks sheet: P&L `extracted`, cash flow and balance sheet `derived` from the
                  schedules. The opening cash typed on Inputs was read only by the deleted CashFlow sheet, so derived cash
                  is cumulative from the first period and differs from the base model's by a constant 375: the +100
                  equity and -500 refinancing live in the single-column sources and uses, which are on no timeline (+400),
                  less the 25 opening cash; the balance sheet's MOVEMENT is zero in every period
  derived P&L     the P&L roles taken away from the same result: derived from the revenue build, the costs total, the
                  fixed-assets schedule and the debt schedule; no tax row, so the balance sheet movement is no test
  riverbend       P&L only: partial P&L, derived cash flow and balance sheet with missing lines, no crash
  small workbook  built here (`three`): costs, depreciation, interest and tax stored positive (P&L, derived cash flow and
                  balance sheet right, the tax sign without NPAT to fit it); 'Repayment of borrowings' and 'Interest paid
                  on borrowings' read as what they are; the debt schedule's interest charge, not its rate; liabilities
                  and equity stored negative; two faults at once (distributions left out of retained earnings and a capex
                  gap: both named, their sum is the movement) and an equity contribution the cross-checks leave
                  unexplained; no equity roll-forward (distributions against retained earnings is no test); a check on
                  a derived line says it is no test; a label that would start a formula is text in the CSV; a line with
                  no number in any year is `missing` in the year view
  quarterly       ten quarters to FY2028: flows sum, stocks take the last period, a two-quarter first year is `partial`
  empty           nothing bound: no periods, every statement `none`, text and CSV still work

    uv run python tests/check_threeway.py
"""
import contextlib
import copy
import csv
import io
import json
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import xlsxwriter  # noqa: E402

from modelatlas import build_map, statements, threeway  # noqa: E402
import check_statements  # noqa: E402
import make_sample_models  # noqa: E402
import make_threeway_model  # noqa: E402

PNL_KEYS = ["revenue", "revenue_part_0", "revenue_part_1", "opex", "ebitda", "depreciation", "ebit", "interest", "pbt", "tax", "npat"]
CF_KEYS = ["cf_open", "ebitda", "tax_paid", "cfo", "capex", "cfi", "draw", "rep", "int_paid", "dist", "other_fin", "cff", "net", "cf_close"]
BS_KEYS = ["cash", "fa", "other_assets", "ta", "debt", "other_liab", "tl", "share_cap", "re", "te", "tle", "bs_check"]
TOL = 1e-6


def build_db(xlsx: Path, tmp: Path, name: str) -> str:
    return build_map.main(str(xlsx), str(tmp / f"{name}__0a1b2c3d"))["db"]


def vals(tw, st, key):
    return next(l for l in tw["statements"][st]["lines"] if l["key"] == key)["values"]


def line(tw, st, key):
    return next(l for l in tw["statements"][st]["lines"] if l["key"] == key)


def arr(v):
    return np.array([np.nan if x is None else x for x in v], dtype=float)


def close(a, b, tol=TOL):
    a, b = arr(a), arr(b)
    return a.shape == b.shape and bool(np.all(np.isfinite(a) == np.isfinite(b))) and \
        bool(np.all(np.abs(a[np.isfinite(a)] - b[np.isfinite(b)]) <= tol))


def checks(tw):
    return {c["key"]: c for s in tw["statements"].values() for c in s["checks"]}


def residual(tw, key):
    return next(r for r in tw["residuals"] if r["key"] == key)


def plain(tw):
    assert json.loads(json.dumps(tw)) == tw, "the result is not plain JSON"
    for st in tw["statements"].values():
        for l in st["lines"]:
            assert len(l["values"]) == len(tw["periods"]), (st["title"], l["key"])
            assert l["source"] in ("extracted", "derived", "missing") and l["kind"] in ("flow", "stock", "subtotal", "check")
            assert (l["source"] == "missing") == all(v is None for v in l["values"]), l["key"]
            assert l["source"] == "missing" or l["from"] or l["kind"] == "check", l["key"]


# ---- a small three-statement workbook, built here -------------------------------------------------------------------

def three(path: Path, pos_costs=False, cf=True, bs=True, omit_dist_in_re=False, capex_gap=0.0, labels=None, rate_row=False,
          neg_le=False, seg_label=None, n=6):
    """Sheets PnL, Assets (fixed-assets corkscrew), Debt (debt corkscrew), Equity (retained earnings corkscrew), and
    optionally CashFlow (with a cash corkscrew) and BalanceSheet, six annual periods to June. Returns {key: values}.
      pos_costs        costs, depreciation, interest and tax stored as positive amounts (subtotals subtract them)
      omit_dist_in_re  retained earnings never move by the distributions the cash flow pays
      capex_gap        cash flow capex is this much more than the fixed-asset additions
      labels           {row key: label} to rename rows; rate_row adds an 'Interest rate' row (%) to the debt schedule
      neg_le           liabilities, equity and their total stored as negatives (assets + liabilities + equity = 0)
      seg_label        a Revenue sheet of two segments, the first with this label, and their total
    The cash flow has a 300 equity contribution in the first period; the balance sheet carries it as share capital."""
    sg = -1 if pos_costs else 1                 # cost rows are stored as sg * (a negative amount)
    rows = []                                   # (sheet, key, label, units, [formula or None per period], values)
    V = {}

    def row(sheet, key, label, values, formula=None, units="$m"):
        V[key] = [float(x) for x in values]
        rows.append((sheet, key, (labels or {}).get(key, label), units, formula, V[key]))

    rng = range(n)
    if seg_label:
        row("Revenue", "seg1", seg_label, [60.0 + 6 * i for i in rng])
        row("Revenue", "seg2", "South revenue", [40.0 + 4 * i for i in rng])
        row("Revenue", "rev_total", "Total revenue", [100.0 + 10 * i for i in rng], "{seg1}+{seg2}")
        row("PnL", "rev", "Revenue", V["rev_total"], "{rev_total}")
    else:
        row("PnL", "rev", "Revenue", [100.0 + 10 * i for i in rng])
    op = "-" if pos_costs else "+"
    row("PnL", "opc", "Operating costs", [sg * -(40.0 + 3 * i) for i in rng])
    row("PnL", "ebitda", "EBITDA", [V["rev"][i] - 40.0 - 3 * i for i in rng], "{rev}" + op + "{opc}")
    row("Assets", "fa_add", "Additions", [200.0 if i == 0 else 10.0 for i in rng])
    row("PnL", "dep", "Depreciation", [sg * -20.0] * n)
    row("PnL", "ebit", "EBIT", [V["ebitda"][i] - 20.0 for i in rng], "{ebitda}" + op + "{dep}")
    row("Debt", "d_draw", "Drawdown", [150.0 if i == 0 else 0.0 for i in rng])
    row("Debt", "d_rep", "Repayment", [0.0 if i == 0 else -10.0 for i in rng])
    dop, dcl = [], []
    for i in rng:
        dop.append(0.0 if i == 0 else dcl[-1])
        dcl.append(dop[-1] + V["d_draw"][i] + V["d_rep"][i])
    row("Debt", "d_open", "Opening debt", dop, "{d_close@-1}")
    row("Debt", "d_close", "Closing debt", dcl, "{d_open}+{d_draw}+{d_rep}")
    if rate_row:
        row("Debt", "d_rate", "Interest rate", [0.05] * n, "0.05+0*{d_draw}", units="%")
    row("Debt", "d_int", "Interest expense", [sg * -0.05 * d for d in dop], ("" if pos_costs else "-") + "0.05*{d_open}")
    row("PnL", "int", "Interest", V["d_int"], "{d_int}")
    row("PnL", "pbt", "Profit before tax", [V["ebit"][i] - 0.05 * dop[i] for i in rng], "{ebit}" + op + "{int}")
    row("PnL", "tax", "Tax", [sg * -0.3 * x for x in V["pbt"]], ("" if pos_costs else "-") + "0.3*{pbt}")
    row("PnL", "npat", "NPAT", [0.7 * x for x in V["pbt"]], "{pbt}" + op + "{tax}")
    fop, fcl = [], []
    for i in rng:
        fop.append(0.0 if i == 0 else fcl[-1])
        fcl.append(fop[-1] + V["fa_add"][i] - 20.0)
    row("Assets", "fa_open", "Opening fixed assets", fop, "{fa_close@-1}")
    row("Assets", "fa_dep", "Depreciation charge", [-20.0] * n, "-20+0*{fa_add}")
    row("Assets", "fa_close", "Closing fixed assets", fcl, "{fa_open}+{fa_add}+{fa_dep}")
    row("Equity", "e_dist", "Distributions", [-0.5 * max(0.0, x) for x in V["npat"]], "-0.5*MAX(0,{npat})")
    rop, rcl = [], []
    for i in rng:
        rop.append(0.0 if i == 0 else rcl[-1])
        rcl.append(rop[-1] + V["npat"][i] + (0.0 if omit_dist_in_re else V["e_dist"][i]))
    row("Equity", "re_open", "Opening retained earnings", rop, "{re_close@-1}")
    row("Equity", "re_npat", "NPAT", V["npat"], "{npat}")
    row("Equity", "re_close", "Closing retained earnings", rcl, "{re_open}+{re_npat}" + ("" if omit_dist_in_re else "+{e_dist}"))
    if cf:
        neg = "-" if pos_costs else ""
        row("CashFlow", "c_ebitda", "EBITDA", V["ebitda"], "{ebitda}")
        row("CashFlow", "c_tax", "Tax paid", [-0.3 * x for x in V["pbt"]], neg + "{tax}")
        row("CashFlow", "cfo", "Cash flow from operations", [a + b for a, b in zip(V["c_ebitda"], V["c_tax"])], "{c_ebitda}+{c_tax}")
        row("CashFlow", "c_capex", "Capital expenditure", [-a - capex_gap for a in V["fa_add"]], f"-{{fa_add}}-{capex_gap}")
        row("CashFlow", "cfi", "Cash flow from investing", V["c_capex"], "{c_capex}")
        row("CashFlow", "c_draw", "Drawdowns", V["d_draw"], "{d_draw}")
        row("CashFlow", "c_rep", "Repayments", V["d_rep"], "{d_rep}")
        row("CashFlow", "c_int", "Interest paid", [-0.05 * d for d in dop], neg + "{int}")
        row("CashFlow", "c_dist", "Distributions paid", V["e_dist"], "{e_dist}")
        row("CashFlow", "c_eq", "Equity contributions", [300.0 if i == 0 else 0.0 for i in rng])
        row("CashFlow", "cff", "Cash flow from financing",
            [sum(V[k][i] for k in ("c_draw", "c_rep", "c_int", "c_dist", "c_eq")) for i in rng],
            "{c_draw}+{c_rep}+{c_int}+{c_dist}+{c_eq}")
        row("CashFlow", "net", "Net cash flow", [V["cfo"][i] + V["cfi"][i] + V["cff"][i] for i in rng], "{cfo}+{cfi}+{cff}")
        cop, ccl = [], []
        for i in rng:
            cop.append(10.0 if i == 0 else ccl[-1])
            ccl.append(cop[-1] + V["net"][i])
        row("CashFlow", "c_open", "Opening cash", cop, "{c_close@-1}")
        row("CashFlow", "c_close", "Closing cash", ccl, "{c_open}+{net}")
    if bs and cf:
        ls, m = (-1, "-") if neg_le else (1, "")
        row("BalanceSheet", "b_cash", "Cash", V["c_close"], "{c_close}")
        row("BalanceSheet", "b_fa", "Fixed assets", fcl, "{fa_close}")
        row("BalanceSheet", "b_ta", "Total assets", [a + b for a, b in zip(V["c_close"], fcl)], "{b_cash}+{b_fa}")
        row("BalanceSheet", "b_debt", "Debt", [ls * d for d in dcl], m + "{d_close}")
        row("BalanceSheet", "b_tl", "Total liabilities", V["b_debt"], "{b_debt}")
        row("BalanceSheet", "b_sc", "Share capital", [ls * 310.0] * n)
        row("BalanceSheet", "b_re", "Retained earnings", [ls * r for r in rcl], m + "{re_close}")
        row("BalanceSheet", "b_te", "Total equity", [a + b for a, b in zip(V["b_sc"], V["b_re"])], "{b_sc}+{b_re}")
        row("BalanceSheet", "b_tle", "Total liabilities and equity", [a + b for a, b in zip(V["b_tl"], V["b_te"])], "{b_tl}+{b_te}")
    # write: one sheet per name, dates in row 3 from column D, a row's formula in every period (an opening row is typed
    # in the first period and the previous closing after)
    at, per = {}, {}
    for sheet, key, *_ in rows:
        per[sheet] = per.get(sheet, 4) + 1
        at[key] = (sheet, per[sheet])
    col = [xlsxwriter.utility.xl_col_to_name(3 + i) for i in rng]

    def cell(here, key, i):
        sh, r = at[key]
        return f"{col[i]}{r}" if sh == here else f"'{sh}'!{col[i]}{r}"

    wb = xlsxwriter.Workbook(str(path))
    fmt = wb.add_format({"num_format": "dd-mmm-yy"})
    sheets = {}
    for sheet, key, label, units, formula, values in rows:
        if sheet not in sheets:
            ws = sheets[sheet] = wb.add_worksheet(sheet)
            ws.write(2, 0, "Period ending")
            for i in rng:
                ws.write_datetime(2, 3 + i, __import__("datetime").datetime(2026 + i, 6, 30), fmt)
        ws, r = sheets[sheet], at[key][1]
        ws.write_string(r - 1, 0, label)
        ws.write(r - 1, 1, units)
        for i in rng:
            if formula is None or ("@-1}" in formula and i == 0):
                ws.write_number(r - 1, 3 + i, values[i])
                continue
            f = formula
            for k in {x.strip("{}").split("@")[0] for x in __import__("re").findall(r"\{[a-z_0-9@-]+\}", formula)}:
                f = f.replace("{" + k + "@-1}", cell(sheet, k, i - 1) if i else "0").replace("{" + k + "}", cell(sheet, k, i))
            ws.write_formula(r - 1, 3 + i, "=" + f, None, values[i])
    wb.close()
    return V


def _strip(res, blocks=(), roles=(), cs_kind=None):
    """A copy of a detect result with blocks unbound, roles removed and corkscrews of a kind taken away."""
    r2 = copy.deepcopy(res)
    for b in r2["blocks"]:
        if b["type"] in blocks:
            b["rows"], b["bound"] = {}, False
        for ro in roles:
            b["rows"].pop(ro, None)
    if cs_kind:
        r2["corkscrews"] = [c for c in r2["corkscrews"] if c["kind"] != cs_kind]
    return r2


# ---- base --------------------------------------------------------------------------------------------------------------

def base(db, V):
    tw = threeway.build(db)
    plain(tw)
    assert tw["by"] == "period" and tw["units"] == "$m" and tw["periodicity"] == "annual" and tw["fy_end_month"] == 6
    assert [p["label"] for p in tw["periods"]] == [f"FY{y}" for y in range(2026, 2041)]
    assert tw["periods"][0] == {"label": "FY2026", "end": "2026-06-30", "n": 1, "partial": False}
    assert {k: s["mode"] for k, s in tw["statements"].items()} == {"pnl": "extracted", "cf": "extracted", "bs": "extracted"}
    assert [l["key"] for l in tw["statements"]["pnl"]["lines"]] == PNL_KEYS
    assert [l["key"] for l in tw["statements"]["cf"]["lines"]] == CF_KEYS
    assert [l["key"] for l in tw["statements"]["bs"]["lines"]] == BS_KEYS
    assert [l["level"] for l in tw["statements"]["pnl"]["lines"]][:3] == [0, 1, 1]
    # every check holds; nothing is unbound in a complete model
    ck = checks(tw)
    assert {c["status"] for c in ck.values()} == {"holds"}, {k: (c["status"], c["why"]) for k, c in ck.items() if c["status"] != "holds"}
    assert set(ck) >= {"pnl_ebitda", "pnl_ebit", "pnl_pbt", "pnl_npat", "pnl_rev_parts", "cf_net", "cf_roll", "cf_cash_tie", "bs_balance"}
    assert all(r["status"] == "holds" for r in tw["residuals"]), [(r["key"], r["status"]) for r in tw["residuals"]]
    # the model's own rows
    assert close(vals(tw, "pnl", "npat"), V[("PnL", "npat")]) and close(vals(tw, "pnl", "ebitda"), V[("PnL", "ebitda")])
    assert close(vals(tw, "pnl", "revenue"), V[("PnL", "rev")]) and close(vals(tw, "pnl", "opex"), V[("PnL", "opc")])
    assert close(vals(tw, "cf", "cf_close"), V[("BalanceSheet", "cash")]) and close(vals(tw, "cf", "cf_close"), V[("CashFlow", "close")])
    assert close(vals(tw, "bs", "cash"), V[("BalanceSheet", "cash")]) and close(vals(tw, "bs", "ta"), V[("BalanceSheet", "ta")])
    assert close(vals(tw, "bs", "tle"), V[("BalanceSheet", "tle")]) and close(vals(tw, "bs", "re"), V[("BalanceSheet", "re")])
    assert close(vals(tw, "bs", "bs_check"), [0.0] * 15, 1e-6)
    assert close(vals(tw, "cf", "cf_open"), V[("CashFlow", "open")]) and close(vals(tw, "cf", "dist"), V[("CashFlow", "dist")])
    assert close(vals(tw, "cf", "draw"), V[("CashFlow", "draw")]) and close(vals(tw, "cf", "capex"), V[("CashFlow", "capex")])
    # costs and outflows negative; the financing plug is the equity contribution less the refinancing
    assert max(vals(tw, "pnl", "opex")) < 0 and max(vals(tw, "cf", "capex")) <= 0 and max(vals(tw, "cf", "dist")) <= 0
    assert abs(vals(tw, "cf", "other_fin")[0] - (100.0 - 500.0)) < 1e-9 and line(tw, "cf", "other_fin")["source"] == "derived"
    assert line(tw, "bs", "other_assets")["source"] == "derived" and max(map(abs, vals(tw, "bs", "other_assets"))) < 1e-9
    assert close(vals(tw, "bs", "share_cap"), [300.0] * 15)
    # every line points at its source rows
    rev = line(tw, "pnl", "revenue")
    assert [f["sheet"] for f in rev["from"]] == ["PnL"] and rev["from"][0]["label"] == "Revenue" and rev["from"][0]["role"] == "revenue"
    assert {f["sheet"] for f in line(tw, "pnl", "revenue_part_0")["from"]} == {"Revenue"}
    assert line(tw, "cf", "tax_paid")["from"][0]["sheet"] == "CashFlow" and line(tw, "bs", "debt")["from"][0]["sheet"] == "BalanceSheet"
    assert line(tw, "bs", "other_liab")["formula"] == "other_liab = tl - debt"
    # the financial-year view of an annual model is the period view
    fy = threeway.build(db, by="fy")
    plain(fy)
    assert fy["by"] == "fy" and fy["periods"] == tw["periods"]
    for k in ("by", "stats"):
        fy.pop(k)
    ref = {k: v for k, v in tw.items() if k not in ("by", "stats")}
    assert _same(fy, ref), "financial-year view differs from the period view of an annual model"
    # the same result passed in is the same answer; a different tolerance is accepted; bad arguments are refused
    res = statements.detect(db)
    again = threeway.build(db, result=res)
    assert {k: v for k, v in again.items() if k != "stats"} == {k: v for k, v in tw.items() if k != "stats"}
    assert threeway.build(db, result=res, tolerance=(1e-3, 1e-3))["statements"]["bs"]["mode"] == "extracted"
    for bad in ({"by": "month"}, {"fy_end_month": 13}):
        try:
            threeway.build(db, **bad)
            raise AssertionError(bad)
        except ValueError:
            pass
    return tw


def _same(a, b):
    """Equal up to float noise (the year view sums one period; the check lines are recomputed)."""
    if isinstance(a, dict):
        return a.keys() == b.keys() and all(_same(a[k], b[k]) for k in a)
    if isinstance(a, list):
        return len(a) == len(b) and all(_same(x, y) for x, y in zip(a, b))
    if isinstance(a, float) or isinstance(b, float):
        return (a is None and b is None) or (a is not None and b is not None and abs(a - b) <= 1e-9)
    return a == b


def sign_normalised(tmp):
    """Costs stored as positive amounts are shown negative, with a note saying so."""
    p = tmp / "poscost.xlsx"
    wb = xlsxwriter.Workbook(str(p))
    ws = wb.add_worksheet("P")
    fmt = wb.add_format({"num_format": "dd-mmm-yy"})
    ws.write(2, 0, "Period ending")
    for i in range(5):
        ws.write_datetime(2, 3 + i, __import__("datetime").datetime(2026 + i, 6, 30), fmt)
    rows = [("Revenue", [100.0 + 10 * i for i in range(5)], None), ("Costs", [40.0 + 3 * i for i in range(5)], None)]
    for r, (label, v, _f) in enumerate(rows, start=5):
        ws.write(r - 1, 0, label)
        for i, x in enumerate(v):
            ws.write_number(r - 1, 3 + i, x)
    ws.write(6, 0, "EBITDA")
    for i in range(5):
        c = chr(ord("D") + i)
        ws.write_formula(f"{c}7", f"={c}5-{c}6", None, (100.0 + 10 * i) - (40.0 + 3 * i))
    wb.close()
    db = build_db(p, tmp, "poscost")
    tw = threeway.build(db)
    assert statements.detect(db)["blocks"][0]["bound"]
    assert line(tw, "pnl", "opex")["source"] == "extracted" and max(vals(tw, "pnl", "opex")) < 0, vals(tw, "pnl", "opex")
    assert "sign flipped" in line(tw, "pnl", "opex")["note"]
    assert close(vals(tw, "pnl", "ebitda"), [60.0 + 7 * i for i in range(5)]) and checks(tw)["pnl_ebitda"]["status"] == "holds"


# ---- signs, labels and readings on the small workbook ------------------------------------------------------------------

def positive_costs(tmp):
    """Costs, depreciation, interest and tax stored as positive amounts, schedules only (no cash flow, no balance sheet):
    the P&L comes out in the standard convention, NPAT is right, and the derived balance sheet moves in step."""
    V = three(tmp / "pos.xlsx", pos_costs=True, cf=False, bs=False)
    db = build_db(tmp / "pos.xlsx", tmp, "pos")
    tw = threeway.build(db)
    plain(tw)
    assert tw["statements"]["pnl"]["mode"] == "extracted" and tw["statements"]["bs"]["mode"] == "derived"
    for k in ("opex", "depreciation", "interest", "tax"):
        assert max(vals(tw, "pnl", k)) <= 0 and "sign flipped" in line(tw, "pnl", k)["note"], (k, vals(tw, "pnl", k))
    assert close(vals(tw, "pnl", "npat"), V["npat"]) and close(vals(tw, "pnl", "tax"), [-x for x in V["tax"]])
    assert close(vals(tw, "cf", "tax_paid"), [-x for x in V["tax"]]) and close(vals(tw, "cf", "capex"), [-x for x in V["fa_add"]])
    assert close(vals(tw, "bs", "re"), V["re_close"]) and close(vals(tw, "bs", "debt"), V["d_close"])
    assert all(abs(x) < 1e-9 for x in vals(tw, "bs", "bs_move")) and checks(tw)["bs_movement"]["status"] == "holds"
    # the tax row with no NPAT to fit its sign against: still taken as a cost, and the note says how
    tw = threeway.build(db, result=_strip(statements.detect(db), roles=("npat",)))
    assert close(vals(tw, "pnl", "npat"), V["npat"]), vals(tw, "pnl", "npat")
    assert "taken from the row's total" in line(tw, "pnl", "tax")["note"]


def financing_labels(tmp):
    """'Repayment of borrowings' is a repayment and 'Interest paid on borrowings' is interest, not drawdowns."""
    three(tmp / "lab.xlsx", labels={"c_draw": "Proceeds from borrowings", "c_rep": "Repayment of borrowings",
                                    "c_int": "Interest paid on borrowings"})
    tw = threeway.build(build_db(tmp / "lab.xlsx", tmp, "lab"))
    got = {k: (line(tw, "cf", k)["source"], [f["label"] for f in line(tw, "cf", k)["from"]]) for k in ("draw", "rep", "int_paid")}
    assert got == {"draw": ("extracted", ["Proceeds from borrowings"]), "rep": ("extracted", ["Repayment of borrowings"]),
                   "int_paid": ("extracted", ["Interest paid on borrowings"])}, got
    assert residual(tw, "debt_flows")["status"] == "holds" and abs(vals(tw, "cf", "other_fin")[1]) < 1e-9


def interest_row(tmp):
    """With no P&L, interest comes from the debt schedule's interest charge, not from its interest rate row."""
    three(tmp / "rate.xlsx", rate_row=True, cf=False, bs=False)
    db = build_db(tmp / "rate.xlsx", tmp, "rate")
    tw = threeway.build(db, result=_strip(statements.detect(db), blocks=("income_statement",)))
    i = line(tw, "pnl", "interest")
    assert i["source"] == "derived" and [f["label"] for f in i["from"]] == ["Interest expense"], i["from"]


def negative_le(tmp):
    """Liabilities, equity and their total stored as negatives: shown positive, and the balance sheet balances."""
    three(tmp / "negle.xlsx", neg_le=True)
    tw = threeway.build(build_db(tmp / "negle.xlsx", tmp, "negle"))
    assert min(vals(tw, "bs", "tle")) > 0 and "sign flipped" in line(tw, "bs", "tle")["note"]
    assert checks(tw)["bs_balance"]["status"] == "holds" and residual(tw, "bs_move")["status"] == "holds"


def two_faults(tmp):
    """Distributions left out of retained earnings AND cash flow capex 5 above the additions: both comparisons fail,
    their sum is the balance sheet movement, nothing is left over, and the reading names both."""
    V = three(tmp / "two.xlsx", omit_dist_in_re=True, capex_gap=5.0)
    tw = threeway.build(build_db(tmp / "two.xlsx", tmp, "two"))
    mv, dist, cap = (arr(residual(tw, k)["values"]) for k in ("bs_move", "dist_vs_re", "capex_additions"))
    assert residual(tw, "dist_vs_re")["status"] == "fails" and residual(tw, "capex_additions")["status"] == "fails"
    assert np.allclose(dist, V["e_dist"]) and np.allclose(cap, -5.0) and np.allclose(mv, dist + cap)
    assert residual(tw, "unexplained")["status"] == "holds"
    r = residual(tw, "bs_move")["reading"]
    assert "distributions" in r and "capex" in r and "unexplained" not in r and "expenses something" not in r, r
    assert "not complete" not in r
    # one cross-check that cannot be built: the reading names what it found and says the split is not complete
    db = build_db(tmp / "two.xlsx", tmp, "two_b")
    tw = threeway.build(db, result=_strip(statements.detect(db), roles=("fa_depreciation",)))
    r = residual(tw, "bs_move")["reading"]
    assert residual(tw, "dep_schedule")["status"] == "unbound" and "distributions" in r and "not complete" in r, r
    # the same without a balance sheet: the 300 equity contribution reaches no derived line, and the reading says the
    # cross-checks leave it unexplained instead of letting the two faults stand for the whole movement
    three(tmp / "two_nobs.xlsx", omit_dist_in_re=True, capex_gap=5.0, bs=False)
    tw = threeway.build(build_db(tmp / "two_nobs.xlsx", tmp, "two_nobs"))
    un = arr(residual(tw, "unexplained")["values"])
    assert abs(un[0] - 300.0) < 1e-9 and np.allclose(un[1:], 0.0)
    assert "300.0 is left unexplained" in residual(tw, "bs_move")["reading"], residual(tw, "bs_move")["reading"]


def no_equity_rollforward(tmp):
    """No retained earnings roll-forward: the movement in retained earnings is built from the cash flow's distributions,
    so distributions against retained earnings agrees by construction and is no test."""
    three(tmp / "nore.xlsx", bs=False)
    db = build_db(tmp / "nore.xlsx", tmp, "nore")
    tw = threeway.build(db, result=_strip(statements.detect(db), blocks=("equity",), cs_kind="retained"))
    r = residual(tw, "dist_vs_re")
    assert r["status"] == "unbound" and "by construction" in r["reading"] and "no equity roll-forward" in r["reading"], r


def derived_not_tautology(db):
    """A check whose line is derived from elsewhere (not from the check's own lines) is no test, and says why truly."""
    tw = threeway.build(db, result=_strip(statements.detect(db), roles=("cfo",)))
    c = checks(tw)["cf_net"]
    assert line(tw, "cf", "cfo")["source"] == "derived" and c["status"] == "unbound"
    assert "not the model's own row" in c["why"] and "cannot differ" not in c["why"], c["why"]


def csv_injection(tmp):
    """A model label that would start a formula reaches the CSV as text."""
    three(tmp / "inj.xlsx", seg_label="=1+1 north revenue")
    tw = threeway.build(build_db(tmp / "inj.xlsx", tmp, "inj"))
    assert any(l["label"].startswith("=1+1") for l in tw["statements"]["pnl"]["lines"])
    rows = list(csv.reader(io.StringIO(threeway.csv(tw, "pnl"))))
    cell = next(r[0] for r in rows if "1+1" in r[0])
    assert cell.startswith("'") and cell.lstrip("' ").startswith("=1+1"), cell
    for x in ("=A1", "+A1", "-A1", "@SUM(A1)", "  =A1", "\t=A1"):
        assert threeway._text_cell(x).startswith("'"), x
    assert threeway._text_cell("Revenue") == "Revenue" and threeway._text_cell("  EBITDA") == "  EBITDA"
    assert all(not r[3].startswith("'") for r in rows[1:] if len(r) > 3 and r[3])   # numbers are left as numbers


def fy_missing():
    """A line with no number in any financial year is `missing` there, as plain() requires."""
    S = threeway.Stmt("pnl", 4)
    S.lines["revenue"].set([1.0, np.nan, 2.0, np.nan], "extracted", [{"sheet": "x", "row": 1, "label": "r", "role": "revenue"}])
    S.lines["opex"].set([-1.0, -1.0, -1.0, -1.0], "extracted", [{"sheet": "x", "row": 2, "label": "c", "role": "opex"}])
    threeway._set_from_calc(S, "ebitda", [(1, "revenue"), (1, "opex")], True)
    assert S.lines["ebitda"].source == "derived"
    T = threeway._to_fy({"pnl": S}, [("FY1", [0, 1], False), ("FY2", [2, 3], False)])["pnl"]
    assert T.lines["revenue"].source == "missing" and T.lines["opex"].source == "extracted"
    assert T.lines["ebitda"].source == "missing" and not np.isfinite(T.lines["ebitda"].v).any(), T.lines["ebitda"].v


# ---- unbalanced --------------------------------------------------------------------------------------------------------

def unbalanced(db, U):
    tw = threeway.build(db)
    plain(tw)
    assert {k: s["mode"] for k, s in tw["statements"].items()} == {"pnl": "extracted", "cf": "extracted", "bs": "extracted"}
    ck = checks(tw)
    chk = U[("BalanceSheet", "chk")]
    assert ck["bs_balance"]["status"] == "fails" and close(ck["bs_balance"]["residuals"], chk)
    assert abs(ck["bs_balance"]["max_residual"] - max(map(abs, chk))) < 1e-9
    assert close(vals(tw, "bs", "bs_check"), chk)
    assert {k for k, c in ck.items() if c["status"] == "fails"} == {"bs_balance"}, {k: c["status"] for k, c in ck.items()}
    # the residual analysis: the balance sheet moves out of balance by the distributions, and nothing else differs
    lvl, mv = residual(tw, "bs_level"), residual(tw, "bs_move")
    assert lvl["status"] == "fails" and close(lvl["values"], chk)
    assert close(mv["values"], [chk[0]] + [b - a for a, b in zip(chk, chk[1:])])
    dist = residual(tw, "dist_vs_re")
    assert dist["status"] == "fails" and close(dist["values"], U[("Equity", "dist")])
    assert np.allclose(np.cumsum(arr(dist["values"])), chk, atol=1e-9)    # their running total is the balance sheet's residual
    for key in ("cash_vs_profit", "dep_schedule", "capex_additions", "debt_flows", "unexplained"):
        assert residual(tw, key)["status"] == "holds", key
        assert max(map(abs, residual(tw, key)["values"])) < 1e-9, key
    assert "distributions" in mv["reading"] and "retained earnings" in mv["reading"] and "FY2026" in mv["reading"]
    assert "distributions" in dist["reading"] and "retained earnings" in dist["reading"]
    # by year: the same residual (annual model)
    fy = threeway.build(db, by="fy")
    assert close(residual(fy, "dist_vs_re")["values"], dist["values"]) and checks(fy)["bs_balance"]["status"] == "fails"


def moved(db, V):
    tw = threeway.build(db)
    plain(tw)
    assert {k: s["mode"] for k, s in tw["statements"].items()} == {"pnl": "extracted", "cf": "extracted", "bs": "extracted"}
    assert line(tw, "pnl", "revenue")["from"][0]["sheet"] == "Income statement" and line(tw, "pnl", "revenue")["from"][0]["label"] == "Turnover"
    assert close(vals(tw, "pnl", "npat"), V[("PnL", "npat")]) and close(vals(tw, "bs", "cash"), V[("BalanceSheet", "cash")])
    assert {c["status"] for c in checks(tw).values()} == {"holds"}


# ---- schedules only ----------------------------------------------------------------------------------------------------

def schedules_only(db, V):
    tw = threeway.build(db)
    plain(tw)
    assert {k: s["mode"] for k, s in tw["statements"].items()} == {"pnl": "extracted", "cf": "derived", "bs": "derived"}, \
        {k: (s["mode"], s["why"]) for k, s in tw["statements"].items()}
    assert close(vals(tw, "pnl", "npat"), V[("PnL", "npat")])
    # derived cash flow: from the P&L and the schedules; what the schedules cannot give is missing
    for k, src in (("ebitda", "derived"), ("tax_paid", "derived"), ("capex", "derived"), ("draw", "derived"), ("rep", "derived"),
                   ("int_paid", "derived"), ("dist", "derived"), ("other_fin", "missing"), ("net", "derived"), ("cf_close", "derived")):
        assert line(tw, "cf", k)["source"] == src, (k, line(tw, "cf", k)["source"])
    assert close(vals(tw, "cf", "capex"), V[("CashFlow", "capex")]) and close(vals(tw, "cf", "draw"), V[("CashFlow", "draw")])
    assert close(vals(tw, "cf", "rep"), V[("CashFlow", "rep")]) and close(vals(tw, "cf", "dist"), V[("CashFlow", "dist")])
    assert close(vals(tw, "cf", "int_paid"), V[("CashFlow", "int")]) and close(vals(tw, "cf", "cfo"), V[("CashFlow", "cfo")])
    # cash is cumulative from the first period: opening cash was typed on Inputs and read by nothing left
    assert vals(tw, "cf", "cf_open")[0] is None and "opening cash not found" in line(tw, "cf", "cf_close")["note"]
    base_cash = arr(V[("BalanceSheet", "cash")])
    off = arr(vals(tw, "cf", "cf_close")) - base_cash
    # +100 equity and -500 refinancing (the one-column sources and uses, on no timeline) are not in the derived cash
    # flow, so it is 400 higher; less the 25 opening cash typed on Inputs: a constant 375. The drawdown is in both.
    assert np.allclose(off, off[0], atol=1e-9) and abs(off[0] - 375.0) < 1e-9, off
    assert "excludes equity contributions / other financing" in line(tw, "cf", "cf_close")["note"]
    assert close(vals(tw, "bs", "cash"), vals(tw, "cf", "cf_close"))
    # derived balance sheet: the schedules' closing balances
    assert close(vals(tw, "bs", "fa"), V[("BalanceSheet", "fa")]) and close(vals(tw, "bs", "debt"), V[("BalanceSheet", "debt")])
    assert close(vals(tw, "bs", "re"), V[("BalanceSheet", "re")])
    assert line(tw, "bs", "share_cap")["source"] == "missing" and line(tw, "bs", "bs_check")["source"] == "missing"
    assert "excludes" in line(tw, "bs", "te")["note"] and "share capital" in line(tw, "bs", "te")["note"]
    ck = checks(tw)
    assert ck["bs_balance"]["status"] == "unbound" and "moves" in ck["bs_balance"]["why"]
    mv = line(tw, "bs", "bs_move")
    assert mv["kind"] == "check" and mv["label"] == "Balance movement check"
    assert all(v is not None and abs(v) < 1e-9 for v in mv["values"]), mv["values"]       # zero in EVERY period, the first too
    assert ck["bs_movement"]["status"] == "holds"
    # by construction, said so: what is taken from the schedule that is also the other side
    r = {x["key"]: x for x in tw["residuals"]}
    assert r["bs_move"]["status"] == "holds" and "bs_level" not in r
    assert r["dep_schedule"]["status"] == "holds"
    for k in ("capex_additions", "debt_flows", "cash_vs_profit", "dist_vs_re"):
        assert r[k]["status"] == "unbound" and "by construction" in r[k]["reading"], (k, r[k])
    # the same, by year
    fy = threeway.build(db, by="fy")
    assert close(vals(fy, "bs", "bs_move"), mv["values"]) and close(vals(fy, "cf", "cf_close"), vals(tw, "cf", "cf_close"))
    return tw


def derived_pnl(db, V):
    res = statements.detect(db)
    r2 = copy.deepcopy(res)
    for b in r2["blocks"]:
        if b["type"] == "income_statement":
            b["rows"], b["bound"] = {}, False
    tw = threeway.build(db, result=r2)
    plain(tw)
    assert tw["statements"]["pnl"]["mode"] == "derived", tw["statements"]["pnl"]["why"]
    src = {k: line(tw, "pnl", k)["source"] for k in ("revenue", "opex", "ebitda", "depreciation", "ebit", "interest", "pbt", "tax", "npat")}
    assert src == {"revenue": "derived", "opex": "derived", "ebitda": "derived", "depreciation": "derived", "ebit": "derived",
                   "interest": "derived", "pbt": "derived", "tax": "missing", "npat": "missing"}, src
    assert close(vals(tw, "pnl", "revenue"), V[("PnL", "rev")]) and close(vals(tw, "pnl", "opex"), V[("PnL", "opc")])
    assert close(vals(tw, "pnl", "depreciation"), V[("PnL", "dep")]) and close(vals(tw, "pnl", "pbt"), V[("PnL", "pbt")])
    assert close(vals(tw, "pnl", "interest"), V[("PnL", "int")])
    assert "costs total found by its label" in line(tw, "pnl", "opex")["note"]
    assert checks(tw)["pnl_ebitda"]["status"] == "unbound" and "computed from the other lines" in checks(tw)["pnl_ebitda"]["why"]
    # without a tax row the cash flow has no tax paid, so the balance sheet movement is the tax: no test, and it says why
    assert line(tw, "cf", "tax_paid")["source"] == "missing"
    mv = residual(tw, "bs_move")
    assert mv["status"] == "unbound" and mv["reading"].startswith("Not a test") and "tax paid in the cash flow" in mv["reading"]
    assert checks(tw)["bs_movement"]["status"] == "unbound"
    assert abs(vals(tw, "bs", "bs_move")[1] - V[("PnL", "tax")][1]) < 1e-9 or abs(vals(tw, "bs", "bs_move")[1] + V[("PnL", "tax")][1]) < 1e-9


# ---- Riverbend, quarterly, empty -----------------------------------------------------------------------------------------

def riverbend(tmp):
    src = ROOT / "tests" / "sample_models" / "Riverbend_BP25_client_model.xlsx"
    if not src.exists():
        make_sample_models.main()
    db = build_db(src, tmp, "riverbend")
    for by in ("period", "fy"):
        tw = threeway.build(db, by=by)
        plain(tw)
        st = tw["statements"]
        assert st["pnl"]["mode"] == "partial" and st["cf"]["mode"] == "derived" and st["bs"]["mode"] == "derived", \
            {k: v["mode"] for k, v in st.items()}
        assert [line(tw, "pnl", k)["source"] for k in ("revenue", "opex", "ebitda", "depreciation", "npat")] == \
            ["extracted", "extracted", "extracted", "missing", "missing"]
        assert checks(tw)["pnl_ebitda"]["status"] == "holds" and checks(tw)["pnl_ebit"]["status"] == "unbound"
        assert line(tw, "cf", "tax_paid")["source"] == "missing" and line(tw, "cf", "capex")["source"] == "missing"
        assert line(tw, "cf", "ebitda")["source"] == "derived"
        # cash is cumulative EBITDA, and the movement is no test: the lines that would make it one are missing
        assert close(vals(tw, "cf", "net"), vals(tw, "pnl", "ebitda"))
        assert checks(tw)["bs_movement"]["status"] == "unbound" and checks(tw)["bs_balance"]["status"] == "unbound"
        mv = residual(tw, "bs_move")
        assert mv["status"] == "unbound" and mv["reading"].startswith("Not a test, because these were not found")
        assert residual(tw, "unexplained")["status"] == "unbound" and "not available" in residual(tw, "unexplained")["reading"]
        assert threeway.text(tw) and all(threeway.csv(tw, s) for s in ("pnl", "cf", "bs"))
    # the same model with the overlay sheets: still no crash
    ov = ROOT / "tests" / "sample_models" / "Riverbend_BP25_with_overlay.xlsx"
    if ov.exists():
        plain(threeway.build(build_db(ov, tmp, "overlay")))


def quarterly(tmp):
    qs = [date(2026, 3, 31), date(2026, 6, 30), date(2026, 9, 30), date(2026, 12, 31), date(2027, 3, 31), date(2027, 6, 30),
          date(2027, 9, 30), date(2027, 12, 31), date(2028, 3, 31), date(2028, 6, 30)]
    p = tmp / "quarterly.xlsx"
    check_statements.minimal(p, dates=qs)
    db = build_db(p, tmp, "quarterly")
    per = threeway.build(db)
    plain(per)
    assert per["periodicity"] == "quarterly" and per["fy_end_month"] == 6
    assert [p_["label"] for p_ in per["periods"]][:3] == ["Mar-26", "Jun-26", "Sep-26"]
    assert per["statements"]["bs"]["mode"] != "none" and line(per, "bs", "cash")["source"] == "extracted"
    fy = threeway.build(db, by="fy")
    plain(fy)
    assert [(x["label"], x["n"], x["partial"]) for x in fy["periods"]] == [("FY2026", 2, True), ("FY2027", 4, False), ("FY2028", 4, False)]
    assert [x["end"] for x in fy["periods"]] == ["2026-06-30", "2027-06-30", "2028-06-30"]
    groups = [(0, 2), (2, 6), (6, 10)]
    for key in ("revenue", "opex", "ebitda", "tax", "npat"):                       # flows sum
        p_, f_ = arr(vals(per, "pnl", key)), arr(vals(fy, "pnl", key))
        assert np.allclose([p_[a:b].sum() for a, b in groups], f_), key
    for key in ("cf_close", "net"):
        p_, f_ = arr(vals(per, "cf", key)), arr(vals(fy, "cf", key))
        want = [p_[b - 1] for a, b in groups] if key == "cf_close" else [p_[a:b].sum() for a, b in groups]
        assert np.allclose(want, f_), key
    p_, f_ = arr(vals(per, "cf", "cf_open")), arr(vals(fy, "cf", "cf_open"))     # opening cash: the first period of the year
    assert np.allclose([p_[a] for a, b in groups], f_)
    for key in ("cash", "ta", "tle", "re"):                                         # stocks: the last period of the year
        p_, f_ = arr(vals(per, "bs", key)), arr(vals(fy, "bs", key))
        assert np.allclose([p_[b - 1] for a, b in groups], f_), key
    assert checks(fy)["bs_balance"]["status"] == "holds" and checks(fy)["cf_roll"]["status"] == "holds"
    assert checks(fy)["pnl_ebitda"]["status"] == "holds"
    # a year ending in December: 4, 4, 2
    dec = threeway.build(db, by="fy", fy_end_month=12)
    assert [(x["label"], x["n"], x["partial"]) for x in dec["periods"]] == [("FY2026", 4, False), ("FY2027", 4, False), ("FY2028", 2, True)]
    assert dec["fy_end_month"] == 12
    # CSV: a column per year, a star on a partial one
    rows = list(csv.reader(io.StringIO(threeway.csv(fy, "pnl"))))
    assert rows[0][3:] == ["FY2026*", "FY2027", "FY2028"] and rows[0][1] == "Source", rows[0]
    assert float(rows[1][3]) == fy["statements"]["pnl"]["lines"][0]["values"][0]


def empty(db):
    res = statements.detect(db)
    r2 = copy.deepcopy(res)
    for b in r2["blocks"]:
        b["rows"], b["bound"] = {}, False
    r2["corkscrews"] = []
    for by in ("period", "fy"):
        tw = threeway.build(db, result=r2, by=by)
        plain(tw)
        assert tw["periods"] == [] and {s["mode"] for s in tw["statements"].values()} == {"none"}
        assert all(l["source"] == "missing" for s in tw["statements"].values() for l in s["lines"])
        assert threeway.text(tw) and threeway.csv(tw, "bs").startswith("Balance sheet")
    assert tw["stats"]["lines"]["missing"] > 0


def cli_and_csv(db, tmp):
    tw = threeway.build(db)
    rows = list(csv.reader(io.StringIO(threeway.csv(tw, "cf"))))
    assert rows[0][:3] == ["Cash flow ($m)", "Source", "Note"] and len(rows[0]) == 3 + 15
    assert rows[1][0] == "Opening cash" and rows[2][0] == "  EBITDA"
    assert any(r[0].startswith("check: ") for r in rows)
    try:
        threeway.csv(tw, "equity")
        raise AssertionError("an unknown statement")
    except ValueError:
        pass
    out = tmp / "csv_out"
    js = tmp / "tw.json"
    with contextlib.redirect_stdout(io.StringIO()) as printed:
        threeway.main([db, "--fy", "--json", str(js), "--csv", str(out)])
    assert "by financial year" in printed.getvalue()
    assert json.loads(js.read_text())["by"] == "fy" and {f.name for f in out.iterdir()} == {"pnl.csv", "cf.csv", "bs.csv"}
    text = threeway.text(tw)
    assert "Profit and loss  [extracted]" in text and "Residual analysis" in text and "Experimental" in text


def main():
    tmp = Path(tempfile.mkdtemp(prefix="threeway_"))
    try:
        paths = make_threeway_model.build(tmp)
        paths["schedules_only"] = make_threeway_model.build_schedules_only(tmp)
        dbs = {name: build_db(Path(p), tmp, name) for name, p in paths.items()}
        _M, V = make_threeway_model.write(tmp / "v_base.xlsx", "base")
        _M, U = make_threeway_model.write(tmp / "v_unbalanced.xlsx", "unbalanced")
        base(dbs["base"], V)
        sign_normalised(tmp)
        unbalanced(dbs["unbalanced"], U)
        moved(dbs["moved"], V)
        print("ok: base, unbalanced, moved, a sign flipped")
        schedules_only(dbs["schedules_only"], V)
        derived_pnl(dbs["schedules_only"], V)
        print("ok: schedules only (derived cash flow and balance sheet), derived P&L")
        positive_costs(tmp)
        financing_labels(tmp)
        interest_row(tmp)
        negative_le(tmp)
        two_faults(tmp)
        no_equity_rollforward(tmp)
        derived_not_tautology(dbs["base"])
        csv_injection(tmp)
        fy_missing()
        print("ok: positive costs, financing labels, interest row, negative L&E, two faults, no equity roll-forward, "
              "derived lines, CSV text cells, year view")
        riverbend(tmp)
        quarterly(tmp)
        empty(dbs["base"])
        cli_and_csv(dbs["base"], tmp)
        print("ok: Riverbend, quarterly to financial years, nothing bound, CSV and CLI")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
