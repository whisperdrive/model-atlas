"""The three financial statements of a model, laid out from the blocks modelatlas/statements.py bound.

For each of the P&L, the cash flow and the balance sheet the lines are in a fixed order, the same for every model, so two
models' statements can be set side by side. A line the model has is EXTRACTED (the model's own row, signs normalised to
the standard convention: costs and outflows negative, inflows positive; assets, liabilities and equity positive). A line
the model does not have is DERIVED from the schedules (the revenue build, the fixed-assets, debt and equity corkscrews,
the capex block, the P&L) when they are enough, and is `missing` (None in every period) when they are not. Nothing is
invented: every line says where it came from (`from`: the source rows, `formula`: how a derived line was computed,
`note`: what was assumed), and a statement says how much of it was found (`mode`: extracted, derived, partial or none).

  extracted  every line is the model's own row (a plug such as 'Other assets' is the difference of its totals)
  derived    none is: the lines come from schedules, other statements, or a cumulative roll of the flows
  partial    some are the model's own rows, others derived or not found
  none       nothing could be built

The checks are identities between lines (EBITDA is revenue plus costs, net cash flow is operating plus investing plus
financing, closing cash is opening plus net flow, the balance sheet balances), each `holds`, `fails`, or `unbound` when a
line it needs is missing or when it is true by construction because the line was computed from the others.

The point of the derived case is the residual analysis. A balance sheet rolled forward from schedules does not know its
opening balances, so its level cannot be compared; what can be compared is how it MOVES: the change in assets less the
change in liabilities and equity, each change built from the roll-forward movements, which needs no opening. When that
movement is not zero the analysis splits it into what could explain it: profit against operating cash (tax timing,
working capital), depreciation in the schedule against the P&L, capex in the cash flow against fixed-asset additions,
debt flows against the movement in the debt balance, distributions paid against distributions reaching retained earnings.
Each has a one-sentence reading. The same split runs on an extracted balance sheet that does not balance.

By financial year (`by="fy"`): flows sum, stocks take the last period of the year (opening cash the first), subtotals
the model did not give are recomputed from the aggregated lines, checks and residuals are recomputed; a year with fewer
periods than a full year is marked `partial`.

No model (LLM) calls. The result is plain JSON values.

    uv run python -m modelatlas.threeway out/<dir>/model.db [--fy] [--json path] [--csv dir]
"""
import argparse
import csv as _csv
import io
import json
import re
import time
from collections import OrderedDict
from datetime import date
from pathlib import Path

import numpy as np

from . import NOTICE, __version__, version_line
from . import depgraph
from . import statements

STATEMENTS = ("pnl", "cf", "bs")
TITLES = {"pnl": "Profit and loss", "cf": "Cash flow", "bs": "Balance sheet"}
BLOCK_OF = {"pnl": "income_statement", "cf": "cash_flow", "bs": "balance_sheet"}
PPY = {"monthly": 12, "quarterly": 4, "half-yearly": 2, "annual": 1}   # periods in a full year

# (key, label, level, kind, agg, counts toward the mode). kind: flow | stock | subtotal | check. agg: how a financial year
# takes the periods: sum | last | first.
PNL = [("revenue", "Revenue", 0, "flow", "sum", True), ("opex", "Operating costs", 0, "flow", "sum", True),
       ("ebitda", "EBITDA", 0, "subtotal", "sum", True), ("depreciation", "Depreciation", 0, "flow", "sum", True),
       ("ebit", "EBIT", 0, "subtotal", "sum", True), ("interest", "Interest", 0, "flow", "sum", True),
       ("pbt", "Profit before tax", 0, "subtotal", "sum", True), ("tax", "Tax", 0, "flow", "sum", True),
       ("npat", "NPAT", 0, "subtotal", "sum", True)]
CF = [("cf_open", "Opening cash", 0, "stock", "first", True), ("ebitda", "EBITDA", 1, "flow", "sum", True),
      ("tax_paid", "Tax paid", 1, "flow", "sum", True), ("cfo", "Cash flow from operations", 0, "subtotal", "sum", True),
      ("capex", "Capital expenditure", 1, "flow", "sum", True), ("cfi", "Cash flow from investing", 0, "subtotal", "sum", True),
      ("draw", "Drawdowns", 1, "flow", "sum", True), ("rep", "Repayments", 1, "flow", "sum", True),
      ("int_paid", "Interest paid", 1, "flow", "sum", True), ("dist", "Distributions", 1, "flow", "sum", True),
      ("other_fin", "Equity contributions / other financing", 1, "flow", "sum", False),
      ("cff", "Cash flow from financing", 0, "subtotal", "sum", True), ("net", "Net cash flow", 0, "subtotal", "sum", True),
      ("cf_close", "Closing cash", 0, "stock", "last", True)]
BS = [("cash", "Cash", 1, "stock", "last", True), ("fa", "Fixed assets", 1, "stock", "last", True),
      ("other_assets", "Other assets", 1, "stock", "last", False), ("ta", "Total assets", 0, "subtotal", "last", True),
      ("debt", "Debt", 1, "stock", "last", True), ("other_liab", "Other liabilities", 1, "stock", "last", False),
      ("tl", "Total liabilities", 0, "subtotal", "last", True),
      ("share_cap", "Share capital / other equity", 1, "stock", "last", False),
      ("re", "Retained earnings", 1, "stock", "last", True), ("te", "Total equity", 0, "subtotal", "last", True),
      ("tle", "Total liabilities and equity", 0, "subtotal", "last", True)]
ORDER = {"pnl": PNL, "cf": CF, "bs": BS}

COSTS_TOTAL = re.compile(r"^\s*(total\s+)?(operating\s+)?(costs?|expenses?|expenditure|opex)(\s+total)?\s*$|"
                         r"^\s*total\s+(operating\s+)?(costs?|expenses?|opex)\b", re.I)
INTEREST = re.compile(r"interest|finance (costs?|charges?)", re.I)
CFO_EBITDA = re.compile(r"\bebitda\b|operating (profit|result)|earnings before", re.I)
CFO_TAX = re.compile(r"\btax", re.I)
FIN_DRAW = re.compile(r"draw|borrow|advance|proceeds|new (debt|loan)|debt (raised|issue)|issue of (debt|notes)", re.I)
FIN_REPAY = re.compile(r"repay|redemption|principal|amorti[sz]ation of (debt|loan)", re.I)
FIN_INT = re.compile(r"interest|finance (costs?|charges?)|coupon", re.I)
FIN_DIST = re.compile(r"distribution|dividend", re.I)
# a financing line's label is read for one kind only: 'Repayment of borrowings' is a repayment, not a drawdown, and
# 'Interest paid on borrowings' is interest
FIN_NOT = {"draw": (FIN_REPAY, FIN_INT, FIN_DIST), "rep": (FIN_INT, FIN_DIST), "int_paid": (FIN_DIST,), "dist": ()}
# rows of a debt schedule whose label says interest but that are not the interest charge
NOT_INTEREST = re.compile(r"\brates?\b|%|\bcover|margin|\bbasis\b|\bdays\b|capitali[sz]|\bswap\b|\bhedg|\bindex", re.I)


# ---- small helpers ---------------------------------------------------------------------------------------------------

def _nan(n):
    return np.full(n, np.nan)


def _known(a) -> bool:
    return a is not None and bool(np.isfinite(a).any())


def _fit_sign(a, x, b):
    """+1 or -1: whether a + x or a - x is b in the periods all three have; None when they share none."""
    if a is None or x is None or b is None:
        return None
    m = np.isfinite(a) & np.isfinite(x) & np.isfinite(b)
    if not m.any():
        return None
    return 1 if np.abs(a[m] + x[m] - b[m]).max() <= np.abs(a[m] - x[m] - b[m]).max() else -1


def _outflow_sign(x):
    """A row expected to be an outflow: -1 when it is stored as positive amounts."""
    t = np.nansum(x) if _known(x) else 0.0
    return -1 if t > 0 else 1


def _sum_known(arrs):
    """Sum of the arrays that have numbers; NaN where none of them does. Missing (None) arrays are skipped."""
    arrs = [a for a in arrs if a is not None]
    if not arrs:
        return None
    m = np.vstack(arrs)
    out = np.nansum(m, axis=0)
    out[~np.isfinite(m).any(axis=0)] = np.nan
    return out


def _fmt(x) -> str:
    x = abs(float(x))
    return f"{x:,.1f}" if x >= 100 else f"{x:,.2f}" if x >= 1 else f"{x:.3g}"


def _short(text, n=110) -> str:
    text = str(text or "")
    return text if len(text) <= n else text[:n].rsplit(" ", 1)[0].rstrip(",;: ") + "..."


def _list(items) -> str:
    items = list(items)
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " and " + items[-1]


# ---- a statement being built ---------------------------------------------------------------------------------------

class Line:
    def __init__(self, key, label, level, kind, agg, core, n, hidden=False):
        self.key, self.label, self.level, self.kind, self.agg, self.core, self.hidden = key, label, level, kind, agg, core, hidden
        self.v = _nan(n)
        self.source = "missing"          # extracted | derived | missing
        self.frm: list = []              # [{"sheet","row","label","role"}]
        self.formula = None
        self.note = []
        self.calc = None                 # (terms [(coef, key)], strict) for a line computed from its siblings

    def set(self, v, source, frm=(), formula=None, note=None, calc=None):
        self.v = np.array(v, dtype=float)
        self.source = source if _known(self.v) else "missing"
        self.frm, self.formula, self.calc = list(frm), formula, calc
        if note:
            self.note.append(note)
        return self


class Stmt:
    def __init__(self, key, n):
        self.key, self.title, self.n = key, TITLES[key], n
        self.lines: "OrderedDict[str, Line]" = OrderedDict()
        for k, label, level, kind, agg, core in ORDER[key]:
            self.lines[k] = Line(k, label, level, kind, agg, core, n)
        self.mode, self.why = "none", ""
        self.checks: list = []
        self.level_ok = False     # (balance sheet) the opening balances are known, so the level check means something
        self.extra: dict = {}

    def hidden(self, key, label, v, frm=(), formula=None):
        ln = Line(key, label, 0, "flow", "sum", False, self.n, hidden=True)
        self.lines[key] = ln
        return ln.set(v, "derived", frm, formula)

    def arr(self, key):
        ln = self.lines.get(key)
        return ln.v if ln is not None and ln.source != "missing" else None

    def line(self, key):
        return self.lines[key]


def _merge_from(*groups):
    seen, out = set(), []
    for g in groups:
        for f in g:
            k = (f["sheet"], f["row"])
            if k not in seen:
                seen.add(k)
                out.append(f)
    return out


class Ctx:
    """The workbook, the statements result and the period axis the statements are laid out on."""

    def __init__(self, book, res):
        self.b, self.res = book, res
        self.roles: dict = {}
        for blk in res["blocks"]:
            for role, v in blk["rows"].items():
                self.roles[role] = v
        self.bound = {b["type"]: b["bound"] for b in res["blocks"]}
        self.unbound_why = {b["type"]: next(iter(b["unbound"].values()), "") for b in res["blocks"]}
        self.cs = {(c["closing"]["sheet"], c["closing"]["row"]): c for c in res["corkscrews"]}
        # the period axis: the dates of the sheets that hold the bound rows, in the periodicity most of them have
        sheets = [r["sheet"] for v in self.roles.values() for r in (v if isinstance(v, list) else [v])]
        per = {}
        for s in set(sheets):
            if s in book.dates:
                per.setdefault(book.periodicity(s), []).append(s)
        self.periodicity = None
        dates: set = set()
        if per:
            self.periodicity = max(per, key=lambda p: (sum(sheets.count(s) for s in per[p]), p or ""))
            for s in per[self.periodicity]:
                dates |= {d for _, d in book.dates[s]}
        self.dates = sorted(dates)
        self.idx = np.array([book.axis_ix[d] for d in self.dates], dtype=int)
        self.n = len(self.dates)

    # -- rows
    def key(self, ref):
        return (ref["sheet"], ref["row"])

    def has(self, role):
        return role in self.roles

    def ref(self, role):
        return self.roles.get(role)

    def fr(self, key, role=None):
        return {"sheet": key[0], "row": key[1], "label": self.b.label(key), "role": role}

    def v(self, key, sign=1):
        return sign * self.b.vec(key)[self.idx] if self.n else _nan(0)

    def role_v(self, role, sign=1, shift=0):
        """(values, [from]) of a single-row role, or (None, [])."""
        ref = self.roles.get(role)
        if ref is None or isinstance(ref, list):
            return None, []
        a = self.v(self.key(ref), sign)
        if shift:
            a = np.concatenate([_nan(shift), a[:-shift]]) if self.n else a
        return a, [self.fr(self.key(ref), role)]

    def movements(self, closing_role):
        """The corkscrew of a closing row: [(key, sign)] of its movements, its opening key (or None)."""
        ref = self.roles.get(closing_role)
        cs = self.cs.get(self.key(ref)) if ref and not isinstance(ref, list) else None
        if cs is None:
            return None, None
        return [((m["sheet"], m["row"]), m["sign"]) for m in cs["movements"]], cs

    def leaves(self, key, sign=1, seen=None):
        """The rows a statement subtotal is made of, on its own sheet: [(key, sign)]. Links and sums are opened."""
        seen = seen or set()
        if key in seen:
            return [(key, sign)]
        seen = seen | {key}
        b = self.b
        lk = b.link(key)
        if lk and lk[1][0] == key[0]:
            return self.leaves(lk[1], sign * lk[0], seen)
        if b.is_total(key) and all(t[1] == key[0] and t[3] == 0 for t in b.terms(key)):
            out = []
            for sg, sh, r, _d in b.terms(key):
                out += self.leaves((sh, r), sign * sg, seen)
            return out
        return [(key, sign)]

    def find_rows(self, sheets, pattern, need_formula=True):
        """Rows on the given sheets with a timeline whose label matches, that have numbers: [(key)]."""
        out = []
        for k, m in self.b.meta.items():
            if k[0] in sheets and k[0] in self.b.dates and pattern.search(m.get("label") or "") \
                    and (m["n_formula"] or not need_formula) and self.b.nonzero(k):
                out.append(k)
        return sorted(out)


# ---- statement builders ----------------------------------------------------------------------------------------------

def _set_from_calc(S, key, terms, strict, note_missing=True):
    """Compute a subtotal from its siblings (terms [(coef, key)]): strict needs every component, else the sum of the
    components that have numbers (and says which are missing)."""
    ln = S.lines[key]
    comps = [(c, S.lines[k]) for c, k in terms]
    have = [(c, l) for c, l in comps if l.source != "missing"]
    miss = [l for _, l in comps if l.source == "missing"]
    if not have or (strict and miss):
        return False
    arr = [c * l.v for c, l in have]
    if strict:
        v = np.sum(arr, axis=0)
    else:
        v = _sum_known(arr)
    sym = " ".join((("+ " if c > 0 else "- ") if i else ("" if c > 0 else "- ")) + l.key for i, (c, l) in enumerate(comps))
    note = f"excludes {_list([l.label.lower() for l in miss])} (not found)" if miss and note_missing else None
    ln.set(v, "derived", _merge_from(*[l.frm for _, l in have]), f"{key} = {sym}", note, calc=(terms, strict))
    return True


def _pnl(C):
    S = Stmt("pnl", C.n)
    n = C.n
    ext, extfrom = {}, {}
    for r in ("revenue", "opex", "ebitda", "depreciation", "ebit", "interest", "pbt", "tax", "npat"):
        a, f = C.role_v(r)
        if a is not None:
            ext[r], extfrom[r] = a, f
    # the sign each cost row was stored in, from the subtotal chain; else costs are negative
    flips, guessed = {}, {}
    chain = {"opex": ("revenue", "ebitda"), "depreciation": ("ebitda", "ebit"), "interest": ("ebit", "pbt"), "tax": ("pbt", "npat")}
    for r, (a, b) in chain.items():
        if r in ext:
            s = _fit_sign(ext.get(a), ext[r], ext.get(b))
            if s is None:   # no subtotal pair to fit it against: a cost row is taken as an outflow (tax too)
                s, guessed[r] = _outflow_sign(ext[r]), True
            if s == -1:
                ext[r] = -ext[r]
                flips[r] = True
    fl = "sign flipped: the model stores this row as a positive amount; shown as a negative (costs are negative)"
    gu = "its sign is not fixed by the subtotals around it, so it is taken from the row's total (a cost is negative overall)"
    L = S.lines

    def extracted(key):
        if key in ext:
            note = "; ".join(x for x in (fl if flips.get(key) else None, gu if guessed.get(key) else None) if x) or None
            L[key].set(ext[key], "extracted", extfrom[key], note=note)
            return True
        return False

    # revenue
    if not extracted("revenue"):
        a, f = C.role_v("rev_total")
        if a is not None:
            L["revenue"].set(a, "derived", f, note="from the revenue build: the model's P&L revenue row was not found")
    # segments, as level-1 parts of revenue
    segs = C.roles.get("rev_segments*")
    if segs and "rev_total" in C.roles and L["revenue"].source != "missing":
        new = OrderedDict()
        for k, ln in S.lines.items():
            new[k] = ln
            if k == "revenue":
                for i, sg in enumerate(segs):
                    key = (sg["sheet"], sg["row"])
                    p = Line(f"revenue_part_{i}", sg["label"] or f"{sg['sheet']}!r{sg['row']}", 1, "flow", "sum", False, n)
                    p.set(C.v(key, sg.get("sign", 1)), "derived", [C.fr(key, "rev_segments*")],
                          note="a segment of the revenue build")
                    new[p.key] = p
        S.lines = new
        L = S.lines
    # the model's own subtotals first (so the leaves can be filled from them)
    for k in ("ebitda", "ebit", "pbt", "npat"):
        extracted(k)
    # leaves
    if not extracted("opex"):
        if L["ebitda"].source != "missing" and L["revenue"].source != "missing":
            L["opex"].set(L["ebitda"].v - L["revenue"].v, "derived", _merge_from(L["ebitda"].frm, L["revenue"].frm), "opex = ebitda - revenue",
                          calc=([(1, "ebitda"), (-1, "revenue")], True))
        else:
            rows = C.find_rows({s for s in C.b.dates}, COSTS_TOTAL)
            stmt_sheets = {C.roles[r]["sheet"] for r in ("cfo", "assets") if r in C.roles and not isinstance(C.roles[r], list)}
            rows = [k for k in rows if k[0] not in stmt_sheets]
            if rows:
                v = C.v(rows[0])
                s = _outflow_sign(v)
                L["opex"].set(s * v, "derived", [C.fr(rows[0], "opex")],
                              note="a costs total found by its label" + (f" ({len(rows)} candidates, the first taken)" if len(rows) > 1 else "")
                              + ("; sign flipped (costs are negative)" if s == -1 else ""))
    if not extracted("depreciation"):
        mv, cs = C.movements("fa_closing")
        dep = C.roles.get("fa_depreciation")
        if mv and dep and not isinstance(dep, list):
            dk = C.key(dep)
            sg = next((s for k, s in mv if k == dk), 1)
            L["depreciation"].set(C.v(dk, sg), "derived", [C.fr(dk, "fa_depreciation")],
                                  note="the depreciation movement of the fixed-assets schedule")
        elif L["ebit"].source != "missing" and L["ebitda"].source != "missing":
            L["depreciation"].set(L["ebit"].v - L["ebitda"].v, "derived", _merge_from(L["ebit"].frm, L["ebitda"].frm), "depreciation = ebit - ebitda",
                                  calc=([(1, "ebit"), (-1, "ebitda")], True))
    if not extracted("interest"):
        sheets = {C.roles[r]["sheet"] for r in ("debt_closing",) if r in C.roles and not isinstance(C.roles[r], list)}
        rows = C.find_rows(sheets, INTEREST) if sheets else []
        rows = [k for k in rows if not NOT_INTEREST.search(C.b.label(k) or "") and "%" not in (C.b.meta[k].get("units") or "")]
        if rows:
            v = C.v(rows[0])
            s = _outflow_sign(v)
            L["interest"].set(s * v, "derived", [C.fr(rows[0], "interest")],
                              note="the interest row of the debt schedule, found by its label" + ("; sign flipped (costs are negative)" if s == -1 else ""))
        elif L["pbt"].source != "missing" and L["ebit"].source != "missing":
            L["interest"].set(L["pbt"].v - L["ebit"].v, "derived", _merge_from(L["pbt"].frm, L["ebit"].frm), "interest = pbt - ebit",
                              calc=([(1, "pbt"), (-1, "ebit")], True))
    if not extracted("tax"):
        if L["npat"].source != "missing" and L["pbt"].source != "missing":
            L["tax"].set(L["npat"].v - L["pbt"].v, "derived", _merge_from(L["npat"].frm, L["pbt"].frm), "tax = npat - pbt", calc=([(1, "npat"), (-1, "pbt")], True))
    # subtotals the model did not give, in order
    for k, terms in (("ebitda", [(1, "revenue"), (1, "opex")]), ("ebit", [(1, "ebitda"), (1, "depreciation")]),
                     ("pbt", [(1, "ebit"), (1, "interest")]), ("npat", [(1, "pbt"), (1, "tax")])):
        if L[k].source == "missing":
            _set_from_calc(S, k, terms, True)
    _mode(S, C, "pnl")
    return S


def _cf(C, P):
    S = Stmt("cf", C.n)
    n, L = C.n, S.lines
    roles = C.roles

    def role_line(line, role, note=None, sign=1):
        a, f = C.role_v(role, sign)
        if a is not None:
            L[line].set(a, "extracted", f, note=note)
            return True
        return False

    cfo_k = C.key(roles["cfo"]) if "cfo" in roles and not isinstance(roles["cfo"], list) else None
    cfi_k = C.key(roles["cfi"]) if "cfi" in roles and not isinstance(roles["cfi"], list) else None
    cff_k = C.key(roles["cff"]) if "cff" in roles and not isinstance(roles["cff"], list) else None

    def from_leaves(line, parent, pattern=None, role=None):
        """Extract a line as the leaves of the parent subtotal that match the label (or are the role's row)."""
        if parent is None:
            return False
        rk = C.key(roles[role]) if role and role in roles and not isinstance(roles[role], list) else None
        excl = FIN_NOT.get(line, ())
        hit = [(k, s) for k, s in C.leaves(parent) if (rk is not None and k == rk) or (rk is None and pattern is not None
               and pattern.search(C.b.label(k) or "") and not any(x.search(C.b.label(k) or "") for x in excl)
               and not (k in used))]
        if not hit:
            return False
        for k, _ in hit:
            used.add(k)
        v = np.sum([C.v(k, s) for k, s in hit], axis=0)
        L[line].set(v, "extracted", [C.fr(k, role or line) for k, _ in hit])
        return True

    used: set = set()
    # operating
    from_leaves("ebitda", cfo_k, CFO_EBITDA)
    from_leaves("tax_paid", cfo_k, CFO_TAX)
    if L["ebitda"].source == "missing" and P is not None and P.lines["ebitda"].source != "missing":
        pl = P.lines["ebitda"]
        L["ebitda"].set(pl.v, "derived", pl.frm, note="from the P&L: not found among the operating cash flows")
    if L["tax_paid"].source == "missing" and P is not None and P.lines["tax"].source != "missing":
        pl = P.lines["tax"]
        L["tax_paid"].set(pl.v, "derived", pl.frm, note="the P&L tax (no cash tax row found): tax paid is taken as tax charged")
    if not role_line("cfo", "cfo"):
        _set_from_calc(S, "cfo", [(1, "ebitda"), (1, "tax_paid")], False)
    # investing
    if cfi_k is not None and "cf_capex" in roles and not isinstance(roles["cf_capex"], list):
        ck = C.key(roles["cf_capex"])
        sg = next((s for k, s in C.leaves(cfi_k) if k == ck), None)
        a = C.v(ck, sg if sg is not None else 1)
        note = None
        if sg is None:
            s2 = _outflow_sign(a)
            a, note = s2 * a, ("sign flipped: capex is stored as a positive amount" if s2 == -1 else None)
        L["capex"].set(a, "extracted", [C.fr(ck, "cf_capex")], note=note)
    else:
        add = roles.get("fa_additions")
        mv, _cs = C.movements("fa_closing")
        if add and not isinstance(add, list) and mv:
            ak = C.key(add)
            sg = next((s for k, s in mv if k == ak), 1)
            L["capex"].set(-C.v(ak, sg), "derived", [C.fr(ak, "fa_additions")],
                           note="minus the additions of the fixed-assets schedule (capex is an outflow)")
        else:
            a, f = C.role_v("capex_total")
            if a is not None:
                s2 = _outflow_sign(a)
                L["capex"].set(s2 * a, "derived", f, note="the capex block's total" + ("; sign flipped (outflows are negative)" if s2 == -1 else ""))
    if not role_line("cfi", "cfi"):
        _set_from_calc(S, "cfi", [(1, "capex")], False)
    # financing
    from_leaves("draw", cff_k, FIN_DRAW)
    from_leaves("rep", cff_k, FIN_REPAY)
    from_leaves("int_paid", cff_k, FIN_INT)
    if not from_leaves("dist", cff_k, None, "cf_distributions") and not from_leaves("dist", cff_k, FIN_DIST):
        pass
    mv, cs = C.movements("debt_closing")
    if mv and (L["draw"].source == "missing" or L["rep"].source == "missing"):
        dc = C.roles["debt_closing"]
        dsign = -1 if _known(C.v(C.key(dc))) and np.nansum(C.v(C.key(dc))) < 0 else 1
        draws, reps, frm_d, frm_r = [], [], [], []
        for k, s in mv:
            v = dsign * C.v(k, s)
            lab = C.b.label(k) or ""
            if FIN_REPAY.search(lab) and not FIN_DRAW.search(lab):
                reps.append(v)
                frm_r.append(C.fr(k, "debt_movements*"))
            elif FIN_DRAW.search(lab) and not FIN_REPAY.search(lab):
                draws.append(v)
                frm_d.append(C.fr(k, "debt_movements*"))
            else:   # a movement the label does not place: positive periods are drawdowns, negative are repayments
                draws.append(np.where(np.isfinite(v), np.maximum(v, 0), np.nan))
                reps.append(np.where(np.isfinite(v), np.minimum(v, 0), np.nan))
                frm_d.append(C.fr(k, "debt_movements*"))
                frm_r.append(C.fr(k, "debt_movements*"))
        if L["draw"].source == "missing" and draws:
            L["draw"].set(np.sum(draws, axis=0), "derived", frm_d, note="the drawdowns of the debt schedule (movements that raise the balance)")
        if L["rep"].source == "missing" and reps:
            L["rep"].set(np.sum(reps, axis=0), "derived", frm_r, note="the repayments of the debt schedule (movements that lower the balance)")
    if L["int_paid"].source == "missing" and P is not None and P.lines["interest"].source != "missing":
        pl = P.lines["interest"]
        L["int_paid"].set(pl.v, "derived", pl.frm, note="the P&L interest: interest paid is taken as interest charged")
    if L["dist"].source == "missing":
        rd = roles.get("re_distributions")
        mv2, _ = C.movements("re_closing")
        if rd and not isinstance(rd, list) and mv2 and any(k == C.key(rd) for k, _ in mv2):
            rk = C.key(rd)
            sg = next(s for k, s in mv2 if k == rk)
            L["dist"].set(C.v(rk, sg), "derived", [C.fr(rk, "re_distributions")],
                          note="the distributions movement of the equity roll-forward (an outflow)")
        else:
            dd = roles.get("dcf_distributions") or rd
            if dd and not isinstance(dd, list):
                a = C.v(C.key(dd))
                s2 = _outflow_sign(a)
                L["dist"].set(s2 * a, "derived", [C.fr(C.key(dd), "dcf_distributions")],
                              note="the distributions row the DCF discounts" + ("; sign flipped (outflows are negative)" if s2 == -1 else ""))
    have_cff = role_line("cff", "cff")
    if have_cff:
        known = _sum_known([L[k].v for k in ("draw", "rep", "int_paid", "dist") if L[k].source != "missing"])
        if known is not None:
            rest = L["cff"].v - np.nan_to_num(known)
            missing = [L[k].label.lower() for k in ("draw", "rep", "int_paid", "dist") if L[k].source == "missing"]
            L["other_fin"].set(rest, "derived", L["cff"].frm, "other_fin = cff - draw - rep - int_paid - dist",
                               note="financing flows not listed above" + (f" (and {_list(missing)}, which were not found)" if missing else ""))
    else:
        _set_from_calc(S, "cff", [(1, "draw"), (1, "rep"), (1, "int_paid"), (1, "dist"), (1, "other_fin")], False)
    # net, closing, opening
    if not role_line("net", "cf_net"):
        _set_from_calc(S, "net", [(1, "cfo"), (1, "cfi"), (1, "cff")], False)
    ref_close = roles.get("cf_closing")
    close_k = C.key(ref_close) if ref_close and not isinstance(ref_close, list) else None
    if close_k is not None:
        L["cf_close"].set(C.v(close_k), "extracted", [C.fr(close_k, "cf_closing")])
        ref_open = roles.get("cf_opening")
        if ref_open and not isinstance(ref_open, list):
            ok = C.key(ref_open)
            sh = 1 if ok == close_k else 0
            a, f = C.role_v("cf_opening", shift=sh)
            L["cf_open"].set(a, "extracted", f, note="the closing cash one period back (no opening row)" if sh else None)
    elif L["net"].source != "missing":
        gaps = [f"{L[k].label.lower()} {x}" for k in ("cfo", "cfi", "cff", "net") if L[k].source == "derived"
                for x in L[k].note if x.startswith("excludes")]
        L["cf_close"].set(np.cumsum(L["net"].v), "derived", L["net"].frm, "cf_close = cumulative net",
                          note="opening cash not found; cumulative from the first period (the level is the cash moved since the start)"
                          + ("; " + "; ".join(gaps) + ", so the level also leaves out what those flows moved" if gaps else ""))
        if n:
            L["cf_open"].set(np.concatenate([[np.nan], L["cf_close"].v[:-1]]), "derived", L["net"].frm,
                             "cf_open = previous cf_close", note="opening cash not found; the first period has none")
    _mode(S, C, "cf")
    return S


def _bs(C, P, F):
    S = Stmt("bs", C.n)
    n, L, roles = C.n, S.lines, C.roles
    A, Li, E = (C.role_v(r)[0] for r in ("assets", "liabilities", "equity"))
    sL = sE = 1
    if A is not None and Li is not None and E is not None:
        best = None
        for s1, s2 in ((1, 1), (-1, 1), (1, -1), (-1, -1)):
            m = np.isfinite(A) & np.isfinite(Li) & np.isfinite(E)
            if m.any():
                r = np.abs(A[m] - s1 * Li[m] - s2 * E[m]).max()
                if best is None or r < best[0] - 1e-12:
                    best = (r, s1, s2)
        if best:
            sL, sE = best[1], best[2]
    TLE = C.role_v("total_le")[0]
    sT = (-1 if sL == sE == -1 else 1) if A is None or TLE is None else (_fit_sign(np.zeros(n), TLE, A) or 1)
    fl = "sign flipped: the model stores this on the other side of the balance (liabilities and equity are shown positive)"

    def role_line(line, role, sign=1, note=None):
        a, f = C.role_v(role, sign)
        if a is not None:
            L[line].set(a, "extracted", f, note=(fl if sign == -1 else None) or note)
            return True
        return False

    role_line("ta", "assets")
    role_line("tl", "liabilities", sL)
    role_line("te", "equity", sE)
    role_line("tle", "total_le", sT)
    role_line("cash", "bs_cash")
    role_line("fa", "bs_fixed_assets")
    role_line("debt", "bs_debt", sL)
    role_line("re", "bs_retained", sE)
    # derived fill-ins, from the other statements and the schedules
    if L["cash"].source == "missing" and F is not None and F.lines["cf_close"].source != "missing":
        c = F.lines["cf_close"]
        L["cash"].set(c.v, "derived", c.frm, note="the closing cash of the cash flow statement" +
                      ("; " + "; ".join(c.note) if c.note and c.source == "derived" else ""))
    if L["fa"].source == "missing":
        a, f = C.role_v("fa_closing")
        if a is not None:
            L["fa"].set(a, "derived", f, note="the closing balance of the fixed-assets schedule")
        elif F is not None and F.lines["capex"].source != "missing" and P is not None and P.lines["depreciation"].source != "missing":
            L["fa"].set(np.cumsum(-F.lines["capex"].v + P.lines["depreciation"].v), "derived",
                        _merge_from(F.lines["capex"].frm, P.lines["depreciation"].frm), "fa = cumulative(-capex + depreciation)",
                        note="opening fixed assets not found; cumulative from the first period")
    if L["debt"].source == "missing":
        a, f = C.role_v("debt_closing")
        if a is not None:
            s = -1 if np.nansum(a) < 0 else 1
            L["debt"].set(s * a, "derived", f, note="the closing balance of the debt schedule" + ("; sign flipped (liabilities are positive)" if s == -1 else ""))
    if L["re"].source == "missing":
        a, f = C.role_v("re_closing")
        if a is not None:
            L["re"].set(a, "derived", f, note="the closing balance of the equity roll-forward")
        elif P is not None and P.lines["npat"].source != "missing":
            flow = P.lines["npat"].v + (F.lines["dist"].v if F is not None and F.lines["dist"].source != "missing" else 0.0)
            L["re"].set(np.cumsum(flow), "derived", _merge_from(P.lines["npat"].frm, F.lines["dist"].frm if F is not None else []),
                        "re = cumulative(npat + dist)", note="opening retained earnings not found; cumulative from the first period")
    # totals
    if L["ta"].source == "extracted":
        known = _sum_known([L[k].v for k in ("cash", "fa") if L[k].source != "missing"])
        rest = L["ta"].v - (np.nan_to_num(known) if known is not None else 0.0)
        L["other_assets"].set(rest, "derived", L["ta"].frm, "other_assets = ta - cash - fa",
                              note="total assets less the lines found" + ("" if known is not None else " (no line was found)"))
    else:
        _set_from_calc(S, "ta", [(1, "cash"), (1, "fa"), (1, "other_assets")], False)
    if L["tl"].source == "extracted":
        known = L["debt"].v if L["debt"].source != "missing" else None
        rest = L["tl"].v - (np.nan_to_num(known) if known is not None else 0.0)
        L["other_liab"].set(rest, "derived", L["tl"].frm, "other_liab = tl - debt", note="total liabilities less the debt line")
    else:
        _set_from_calc(S, "tl", [(1, "debt"), (1, "other_liab")], False)
    if L["te"].source == "extracted":
        known = L["re"].v if L["re"].source != "missing" else None
        rest = L["te"].v - (np.nan_to_num(known) if known is not None else 0.0)
        L["share_cap"].set(rest, "derived", L["te"].frm, "share_cap = te - re",
                           note="total equity less retained earnings (share capital and other reserves)" if known is not None else
                           "total equity: no retained earnings line was found")
    else:
        _set_from_calc(S, "te", [(1, "share_cap"), (1, "re")], False)
    if L["tle"].source == "missing":
        _set_from_calc(S, "tle", [(1, "tl"), (1, "te")], False)
    S.level_ok = (L["ta"].source == "extracted" and L["tle"].source in ("extracted", "derived")
                  and (L["tle"].source == "extracted" or (L["tl"].source == "extracted" and L["te"].source == "extracted")))
    # the movements the roll-forwards give, which need no opening balance
    mv, _ = C.movements("fa_closing")
    if mv:
        fa_add = roles.get("fa_additions")
        fa_dep = roles.get("fa_depreciation")
        sg = {k: s for k, s in mv}
        d = np.sum([C.v(k, s) for k, s in mv], axis=0)
        S.hidden("d_fa", "Movement in fixed assets", d, [C.fr(k, "fa_movements*") for k, _ in mv], "d_fa = sum of the schedule's movements")
        if fa_add and not isinstance(fa_add, list) and C.key(fa_add) in sg:
            S.hidden("fa_additions", "Fixed-asset additions", C.v(C.key(fa_add), sg[C.key(fa_add)]), [C.fr(C.key(fa_add), "fa_additions")])
        if fa_dep and not isinstance(fa_dep, list) and C.key(fa_dep) in sg:
            S.hidden("fa_dep", "Depreciation in the schedule", C.v(C.key(fa_dep), sg[C.key(fa_dep)]), [C.fr(C.key(fa_dep), "fa_depreciation")])
    elif F is not None and F.lines["capex"].source != "missing" and P is not None and P.lines["depreciation"].source != "missing":
        S.hidden("d_fa", "Movement in fixed assets", -F.lines["capex"].v + P.lines["depreciation"].v,
                 _merge_from(F.lines["capex"].frm, P.lines["depreciation"].frm), "d_fa = -capex + depreciation")
    mv, _ = C.movements("debt_closing")
    if mv:
        dc = C.v(C.key(roles["debt_closing"]))
        ds = -1 if _known(dc) and np.nansum(dc) < 0 else 1
        S.hidden("d_debt", "Movement in debt", ds * np.sum([C.v(k, s) for k, s in mv], axis=0),
                 [C.fr(k, "debt_movements*") for k, _ in mv], "d_debt = sum of the schedule's movements")
    mv, _ = C.movements("re_closing")
    if mv:
        S.hidden("d_re", "Movement in retained earnings", np.sum([C.v(k, s) for k, s in mv], axis=0),
                 [C.fr(k, "re_movements*") for k, _ in mv], "d_re = sum of the roll-forward's movements")
    elif F is not None and P is not None and P.lines["npat"].source != "missing":
        d = P.lines["npat"].v + (F.lines["dist"].v if F.lines["dist"].source != "missing" else 0.0)
        S.hidden("d_re", "Movement in retained earnings", d, P.lines["npat"].frm, "d_re = npat + dist")
    if F is not None and F.lines["net"].source != "missing":
        S.hidden("d_cash", "Movement in cash", F.lines["net"].v, F.lines["net"].frm, "d_cash = net cash flow")
    _mode(S, C, "bs")
    return S


def _mode(S, C, which):
    core = [l for l in S.lines.values() if l.core and not l.hidden]
    ext = [l for l in core if l.source == "extracted"]
    der = [l for l in core if l.source == "derived"]
    mis = [l for l in core if l.source == "missing"]
    if not ext and not der:
        S.mode = "none"
        S.why = f"nothing to build it from: {C.unbound_why.get(BLOCK_OF[which]) or 'the model has no such rows'}"
    elif not der and not mis:
        S.mode = "extracted"
        S.why = f"every line is the model's own row ({len(ext)} lines), signs normalised"
    elif not ext:
        S.mode = "derived"
        S.why = ("the model's statements were not found (" + _short(C.unbound_why.get(BLOCK_OF[which]) or "no block bound") + "); "
                 f"{len(der)} lines derived from schedules and the other statements" + (f", {len(mis)} not found" if mis else ""))
    else:
        S.mode = "partial"
        S.why = (f"{len(ext)} lines are the model's own rows" + (f", {len(der)} derived" if der else "") +
                 (f", {len(mis)} not found ({_list([l.label.lower() for l in mis])})" if mis else ""))


# ---- period view -> financial-year view --------------------------------------------------------------------------------

def _groups(dates, fy_end_month, periodicity):
    groups: "OrderedDict[int, list]" = OrderedDict()
    for i, d in enumerate(dates):
        groups.setdefault(d.year + (1 if d.month > fy_end_month else 0), []).append(i)
    full = PPY.get(periodicity)
    return [(f"FY{y}", ix, bool(full and len(ix) < full)) for y, ix in groups.items()]


def _agg(v, ixs, how):
    out = []
    for ix in ixs:
        seg = v[ix]
        if how == "sum":
            out.append(float(seg.sum()) if np.isfinite(seg).all() else np.nan)
        elif how == "first":
            out.append(float(seg[0]))
        else:
            out.append(float(seg[-1]))
    return np.array(out) if out else _nan(0)


def _to_fy(stmts, groups):
    ixs = [g[1] for g in groups]
    out = {}
    for k, S in stmts.items():
        T = Stmt.__new__(Stmt)
        T.key, T.title, T.n, T.mode, T.why, T.checks, T.level_ok, T.extra = S.key, S.title, len(groups), S.mode, S.why, [], S.level_ok, dict(S.extra)
        T.lines = OrderedDict()
        for key, ln in S.lines.items():
            m = Line(ln.key, ln.label, ln.level, ln.kind, ln.agg, ln.core, len(groups), ln.hidden)
            m.source, m.frm, m.formula, m.note, m.calc = ln.source, ln.frm, ln.formula, list(ln.note), ln.calc
            m.v = _agg(ln.v, ixs, ln.agg)
            if not _known(m.v):
                m.source = "missing"
            T.lines[key] = m
        # a subtotal the model did not give is recomputed from the aggregated lines
        for key, ln in T.lines.items():
            if ln.calc:
                terms, strict = ln.calc
                comps = [(c, T.lines[kk]) for c, kk in terms]
                have = [(c, l) for c, l in comps if l.source != "missing"]
                if strict and len(have) < len(comps):   # a component has no number in any year: neither has the subtotal
                    ln.v = _nan(T.n)
                elif have:
                    ln.v = np.sum([c * l.v for c, l in have], axis=0) if strict else _sum_known([c * l.v for c, l in have])
                ln.source = ln.source if _known(ln.v) else "missing"
        out[k] = T
    return out


# ---- checks, movements, residuals -------------------------------------------------------------------------------------

def _tol_of(stmts, tolerance):
    scale = 0.0
    for S in stmts.values():
        for ln in S.lines.values():
            if ln.source != "missing" and np.isfinite(ln.v).any():
                scale = max(scale, float(np.nanmax(np.abs(ln.v))))
    return statements._tol(scale, tolerance)


def _clean(v):
    return [None if not np.isfinite(x) else float(x) for x in v]


def _status(resid, tol):
    f = resid[np.isfinite(resid)]
    if not f.size:
        return "unbound"
    return "holds" if float(np.abs(f).max()) <= tol else "fails"


def _mkcheck(S, key, title, terms, tol, result_key=None, need=None):
    """terms [(coef, line key)] summing to zero. unbound when a line is missing, or when the result line was computed
    from the others (true by construction)."""
    lines = [S.lines.get(k) for _, k in terms]
    names = [S.lines[k].label if k in S.lines else k for _, k in terms]
    miss = [l.label for l in lines if l is None or l.source == "missing"]
    if miss:
        ch = {"key": key, "title": title, "status": "unbound", "residuals": [None] * S.n, "max_residual": None,
              "why": f"{_list([m.lower() for m in miss])} not found"}
    else:
        resid = np.sum([c * l.v for (c, _), l in zip(terms, lines)], axis=0)
        st = _status(resid, tol)
        rk = result_key or terms[-1][1]
        why = ""
        by_construction = [l for l in [S.lines[rk]] + lines if l.source == "derived" and l.calc] + \
            ([S.lines[rk]] if S.lines[rk].source == "derived" else [])
        if by_construction:
            l0 = by_construction[0]
            if l0.calc and {kk for _, kk in l0.calc[0]} <= {kk for _, kk in terms}:
                why = f"{l0.label.lower()} is computed from the other lines, so it cannot differ"
            else:   # derived from elsewhere: not the model's own row, so the identity says nothing about the model
                how = l0.formula or (l0.note[0] if l0.note else "from the schedules")
                why = f"{l0.label.lower()} is not the model's own row ({how}), so this is not a test of the model"
            st = "unbound"
        elif st == "unbound":
            why = "no period has every line"
        mx = float(np.nanmax(np.abs(resid))) if np.isfinite(resid).any() else None
        ch = {"key": key, "title": title, "status": st, "residuals": _clean(resid), "max_residual": mx, "why": why}
    S.checks.append(ch)
    return ch


def _finalize(stmts, tol):
    """Check lines, the balance sheet movement, and the checks, on whichever view (period or year) the lines are in."""
    P, F, B = stmts["pnl"], stmts["cf"], stmts["bs"]
    # P&L
    for k, t, terms in (("pnl_ebitda", "Revenue plus operating costs is EBITDA", [(1, "revenue"), (1, "opex"), (-1, "ebitda")]),
                        ("pnl_ebit", "EBITDA plus depreciation is EBIT", [(1, "ebitda"), (1, "depreciation"), (-1, "ebit")]),
                        ("pnl_pbt", "EBIT plus interest is profit before tax", [(1, "ebit"), (1, "interest"), (-1, "pbt")]),
                        ("pnl_npat", "Profit before tax plus tax is NPAT", [(1, "pbt"), (1, "tax"), (-1, "npat")])):
        _mkcheck(P, k, t, terms, tol)
    parts = [k for k in P.lines if k.startswith("revenue_part_")]
    if parts:
        ch = _mkcheck(P, "pnl_rev_parts", "The revenue segments sum to revenue",
                      [(1, k) for k in parts] + [(-1, "revenue")], tol)
    # cash flow
    _mkcheck(F, "cf_net", "Operating, investing and financing flows sum to net cash flow",
             [(1, "cfo"), (1, "cfi"), (1, "cff"), (-1, "net")], tol)
    _mkcheck(F, "cf_roll", "Opening cash plus net cash flow is closing cash", [(1, "cf_open"), (1, "net"), (-1, "cf_close")], tol)
    cash, close = B.lines["cash"], F.lines["cf_close"]
    if cash.source == "extracted" and close.source == "extracted":
        resid = close.v - cash.v
        F.checks.append({"key": "cf_cash_tie", "title": "Closing cash on the cash flow is cash on the balance sheet",
                         "status": _status(resid, tol), "residuals": _clean(resid),
                         "max_residual": float(np.nanmax(np.abs(resid))) if np.isfinite(resid).any() else None, "why": ""})
    else:
        F.checks.append({"key": "cf_cash_tie", "title": "Closing cash on the cash flow is cash on the balance sheet", "status": "unbound",
                         "residuals": [None] * F.n, "max_residual": None,
                         "why": "cash on the balance sheet is not the model's own row" if cash.source != "extracted" else "closing cash is not the model's own row"})
    # balance sheet: the level check and the movement
    ta, tle = B.lines["ta"], B.lines["tle"]
    chk = B.lines["bs_check"] = Line("bs_check", "Check (assets - liabilities and equity)", 0, "check", "last", False, B.n)
    if B.level_ok and ta.source != "missing" and tle.source != "missing":
        chk.set(ta.v - tle.v, "derived", _merge_from(ta.frm, tle.frm), "bs_check = ta - tle")
        resid = chk.v
        B.checks.append({"key": "bs_balance", "title": "Assets equal liabilities plus equity", "status": _status(resid, tol),
                         "residuals": _clean(resid), "max_residual": float(np.nanmax(np.abs(resid))) if np.isfinite(resid).any() else None,
                         "why": ""})
        prev = np.concatenate([[0.0], chk.v[:-1]]) if B.n else chk.v
        move = chk.v - prev                  # the first period assumes the opening balance sheet balanced
        move_note = "the first period assumes the opening balance sheet balanced"
    else:
        why = ("the balance sheet is not the model's own, or an opening balance is not known" if B.mode != "none" else "no balance sheet")
        chk.source, chk.v = "missing", _nan(B.n)
        B.checks.append({"key": "bs_balance", "title": "Assets equal liabilities plus equity", "status": "unbound",
                         "residuals": [None] * B.n, "max_residual": None,
                         "why": why + ": a rolled-forward balance sheet is compared by how it moves, below"})
        parts = [(1, "d_cash"), (1, "d_fa"), (-1, "d_debt"), (-1, "d_re")]
        have = [(c, B.lines[k]) for c, k in parts if k in B.lines and B.lines[k].source != "missing"]
        move = _sum_known([c * l.v for c, l in have]) if have else _nan(B.n)
        move_note = "share capital, other assets and other liabilities are assumed not to move"
    mv = Line("bs_move", "Balance movement check", 0, "check", "sum", False, B.n)
    if np.isfinite(move).any():
        mv.set(move, "derived", [], "bs_move = change in assets - change in liabilities and equity", move_note)
    if not B.level_ok:
        B.lines["bs_move"] = mv
        lacking = [f"the movement in {name}" for name, key in (("cash", "d_cash"), ("fixed assets", "d_fa"), ("debt", "d_debt"),
                                                               ("retained earnings", "d_re"))
                   if key not in B.lines or B.lines[key].source == "missing"]
        need = ["ebitda", "tax_paid", "capex"] + (["draw", "rep", "int_paid"] if "d_debt" in B.lines else []) + \
            (["dist"] if "d_re" in B.lines else [])
        lacking += [f"{F.lines[k].label.lower()} in the cash flow" for k in need if F.lines[k].source == "missing"]
        st = _status(move, tol) if np.isfinite(move).any() else "unbound"
        why = ""
        if lacking:
            st, why = "unbound", f"not found: {_list(lacking)}, so this is not a test"
            B.extra["lacking"] = _list(lacking)
        B.checks.append({"key": "bs_movement", "title": "Assets and liabilities plus equity move together", "status": st,
                         "residuals": _clean(move), "max_residual": float(np.nanmax(np.abs(move))) if np.isfinite(move).any() else None,
                         "why": why})
    B.extra["move"] = move


COMPONENTS = [
    ("cash_vs_profit", "Operating cash against profit",
     "cash flow from operations (after interest paid) and profit before depreciation",
     "tax paid differs from tax charged, or working capital or another non-cash item moves"),
    ("dep_schedule", "Depreciation in the schedule against the P&L",
     "depreciation in the fixed-assets schedule and in the P&L", "the schedule and the P&L charge different amounts"),
    ("capex_additions", "Capex against fixed-asset additions",
     "the capex in the cash flow and the fixed-asset additions",
     "a timing difference, a disposal, capex paid outside the schedule, or spend capitalised in one and expensed in the other"),
    ("debt_flows", "Debt flows against the debt balance",
     "the debt drawn and repaid in the cash flow and the movement in the debt balance",
     "a non-cash movement, or a flow that reaches one side only"),
    ("dist_vs_re", "Distributions paid against retained earnings",
     "the distributions paid in the cash flow and the distributions that reach retained earnings",
     "the retained earnings roll-forward leaves distributions out, or the cash flow pays what the equity schedule does not"),
]


def _residuals(stmts, labels, tol):
    P, F, B = stmts["pnl"], stmts["cf"], stmts["bs"]
    n = B.n
    move = B.extra.get("move", _nan(n))
    got = lambda S, k: S.lines[k].v if k in S.lines and S.lines[k].source != "missing" else None  # noqa: E731
    npat, dep = got(P, "npat"), got(P, "depreciation")
    cfo, intp, capex, draw, rep, dist = (got(F, k) for k in ("cfo", "int_paid", "capex", "draw", "rep", "dist"))
    add, depsch, ddebt, dre = (got(B, k) for k in ("fa_additions", "fa_dep", "d_debt", "d_re"))
    comp = {k: None for k, *_ in COMPONENTS}
    if all(x is not None for x in (cfo, intp, npat, dep)):
        comp["cash_vs_profit"] = (cfo + intp) - (npat - dep)
    if depsch is not None and dep is not None:
        comp["dep_schedule"] = depsch - dep
    if add is not None and capex is not None:
        comp["capex_additions"] = capex + add
    if ddebt is not None and draw is not None and rep is not None:
        comp["debt_flows"] = (draw + rep) - ddebt
    elif ddebt is not None and (draw is not None or rep is not None):
        comp["debt_flows"] = (draw if draw is not None else 0.0) + (rep if rep is not None else 0.0) - ddebt
    if dist is not None and dre is not None and npat is not None:
        comp["dist_vs_re"] = dist - (dre - npat)
    allk = all(comp[k] is not None for k, *_ in COMPONENTS)
    out = []
    # the balance sheet: the level check (when the opening balances are known), and always how it moves
    lvl = B.level_ok
    chk = got(B, "bs_check")
    if lvl and chk is not None:
        out.append(_residual_line("bs_level", "Balance sheet check (assets less liabilities and equity)", chk, labels, tol, "level", comp, B))
    title = "Balance sheet movement (change in assets less change in liabilities and equity)" if not lvl else \
        "Balance sheet movement (change in the check)"
    line = _residual_line("bs_move", title, move, labels, tol, "movement", comp, B)
    mvck = next((c for c in B.checks if c["key"] == "bs_movement"), None)
    if mvck and mvck["status"] == "unbound" and np.isfinite(move).any():
        first = int(np.nonzero(np.isfinite(move))[0][0])
        line.update(status="unbound", reading=f"Not a test, because these were not found: {B.extra.get('lacking', 'some lines')}. "
                                              f"What was found moves by {_fmt(move[first])} in {labels[first]}.")
    out.append(line)
    # a comparison is no test when one side was taken from the other's own schedule
    roles_of = lambda S, k: {f.get("role") for f in S.lines[k].frm} if k in S.lines and S.lines[k].source == "derived" else set()  # noqa: E731
    triv = {}
    if roles_of(F, "draw") & {"debt_movements*"} or roles_of(F, "rep") & {"debt_movements*"}:
        triv["debt_flows"] = "the drawdowns and repayments in the cash flow are the debt schedule's own movements"
    if roles_of(F, "capex") & {"fa_additions"}:
        triv["capex_additions"] = "the capex in the cash flow is the fixed-asset additions"
    if roles_of(P, "depreciation") & {"fa_depreciation"}:
        triv["dep_schedule"] = "the P&L depreciation is the fixed-assets schedule's own"
    if F.lines["cfo"].source == "derived" and F.lines["int_paid"].source == "derived":
        triv["cash_vs_profit"] = "the cash flow's operating lines and interest are the P&L's own"
    if roles_of(F, "dist") & {"re_distributions"}:
        triv["dist_vs_re"] = "the distributions in the cash flow are the equity roll-forward's own movement"
    elif "d_re" in B.lines and B.lines["d_re"].formula == "d_re = npat + dist":
        triv["dist_vs_re"] = ("no equity roll-forward was found, so the movement in retained earnings is built from the "
                              "cash flow's distributions")
    for k, name, what, why in COMPONENTS:
        ln = _residual_line(k, name, comp[k], labels, tol, "comp", comp, B, what, why)
        if k in triv and ln["status"] == "holds":
            ln.update(status="unbound", reading=f"Agrees by construction, so it tests nothing: {triv[k]}.")
        out.append(ln)
    if allk and np.isfinite(move).any():
        rest = move - np.sum([comp[k] for k, *_ in COMPONENTS], axis=0)
        out.append(_residual_line("unexplained", "Not explained by the lines above", rest, labels, tol, "rest", comp, B))
    else:
        out.append({"key": "unexplained", "title": "Not explained by the lines above", "values": [None] * n, "status": "unbound",
                    "reading": "The split into causes is not available: " + (
                        _list([name[0].lower() + name[1:] for k, name, what, _w in COMPONENTS if comp[k] is None]) + " could not be built, "
                        "because a line each needs was not found" if any(comp[k] is None for k, *_ in COMPONENTS)
                        else "the balance sheet movement is not known") + "."})
    return out


def _residual_line(key, title, v, labels, tol, kind, comp, B, what=None, why=None):
    n = len(labels)
    if v is None or not np.isfinite(v).any():
        reading = {"bs": "The balance sheet cannot be tested: its lines or their movements were not found.",
                   "rest": "Not known."}.get(kind, f"Not testable: a line it needs was not found ({what}).")
        return {"key": key, "title": title, "values": [None] * n, "status": "unbound", "reading": reading}
    bad = np.nonzero(np.isfinite(v) & (np.abs(v) > tol))[0]
    fin = int(np.isfinite(v).sum())
    status = "holds" if not len(bad) else "fails"
    if not len(bad):
        reading = {"level": "The balance sheet balances in every period.",
                   "movement": ("The balance sheet's check does not change in any period." if B.level_ok else
                                "The derived balance sheet moves in step in every period: assets and liabilities plus equity change by the same amount."),
                   "rest": "Nothing is left over: the lines above explain the whole movement."}.get(kind) or \
            f"{what[0].upper() + what[1:]} agree in every period."
        return {"key": key, "title": title, "values": _clean(v), "status": status, "reading": reading}
    first = int(bad[0])
    big = int(bad[np.argmax(np.abs(v[bad]))])
    tail = "" if len(bad) == 1 else f"; {len(bad)} of {fin} periods differ, the most by {_fmt(v[big])} in {labels[big]}" if big != first else \
        f"; {len(bad)} of {fin} periods differ, the most by {_fmt(v[big])}"
    if kind == "level":
        end = int(np.nonzero(np.isfinite(v))[0][-1])
        head = f"The balance sheet is out of balance by {_fmt(v[first])} in {labels[first]}"
        if big != first:
            head += f", by {_fmt(v[big])} in {labels[big]} at most"
        if end not in (first, big) and abs(v[end]) > tol:
            head += f" and by {_fmt(v[end])} in {labels[end]}"
        return {"key": key, "title": title, "values": _clean(v), "status": status,
                "reading": head + f"; {len(bad)} of {fin} periods are out. The movement below says what changed it."}
    if kind == "movement":
        head = (f"The balance sheet moves out of balance by {_fmt(v[first])} in {labels[first]}" if B.level_ok else
                f"The derived balance sheet moves out of balance by {_fmt(v[first])} in {labels[first]}")
        clauses = []
        for k, name, w, why_ in COMPONENTS:
            c = comp.get(k)
            if c is not None and np.isfinite(c[first]) and abs(c[first]) > tol:
                clauses.append((abs(c[first]), f"{w} differ by {_fmt(c[first])}: {why_}"))
        left = ""
        if clauses and all(comp.get(k) is not None for k, *_ in COMPONENTS):
            rest = v[first] - sum(float(comp[k][first]) for k, *_ in COMPONENTS if np.isfinite(comp[k][first]))
            if abs(rest) > tol:   # the cross-checks explain part of it: say how much they leave
                left = (f"; after them {_fmt(rest)} is left unexplained (other financing flows or other balance sheet "
                        f"lines)")
        clauses.sort(key=lambda x: -x[0])
        if clauses:
            head += ", the year " if labels[first].startswith("FY") else ", the period "
            head += _list([c for _, c in clauses[:3]]) + left
            if any(comp.get(k) is None for k, *_ in COMPONENTS):
                head += "; the split into causes is not complete, because some lines it needs were not found"
        elif any(comp.get(k) is None for k, *_ in COMPONENTS):
            head += "; the split into causes is not complete, because some lines it needs were not found"
        else:
            head += "; none of the cross-checks below differs, so the difference is in other balance sheet lines or other financing flows"
        return {"key": key, "title": title, "values": _clean(v), "status": status, "reading": head + tail + "."}
    if kind == "rest":
        return {"key": key, "title": title, "values": _clean(v), "status": status,
                "reading": f"Left over after the lines above: {_fmt(v[first])} in {labels[first]}{tail}: other financing flows or movements in other balance sheet lines."}
    why = dict((k, w) for k, _n, _a, w in COMPONENTS)[key]
    return {"key": key, "title": title, "values": _clean(v), "status": status,
            "reading": f"{what[0].upper() + what[1:]} differ by {_fmt(v[first])} in {labels[first]}{tail}: {why}."}


# ---- the public API ---------------------------------------------------------------------------------------------------

def _parse_tol(tolerance):
    if tolerance is None:
        return (statements.REL_TOL, statements.ABS_TOL)
    if isinstance(tolerance, (tuple, list)):
        return (float(tolerance[0]), float(tolerance[1]))
    if isinstance(tolerance, dict):
        return (float(tolerance.get("rel", statements.REL_TOL)), float(tolerance.get("abs", statements.ABS_TOL)))
    return (float(tolerance), statements.ABS_TOL)


def build(db_path, result=None, by="period", fy_end_month=None, tolerance=None) -> dict:
    """The three statements of a model.db laid out from its bound blocks (see the module docstring). `result` is a
    statements.detect result (detected here when None); by is "period" or "fy"."""
    if by not in ("period", "fy"):
        raise ValueError("by must be 'period' or 'fy'")
    if fy_end_month is not None and not 1 <= int(fy_end_month) <= 12:
        raise ValueError("fy_end_month must be 1..12")
    t0 = time.time()
    tol_cfg = _parse_tol(tolerance)
    if result is None:
        result = statements.detect(db_path, tolerance)
    book = statements.Book(db_path)
    try:
        C = Ctx(book, result)
        P = _pnl(C)
        F = _cf(C, P)
        B = _bs(C, P, F)
        stmts = {"pnl": P, "cf": F, "bs": B}
        units = _units(C, stmts)
        end_month = int(fy_end_month) if fy_end_month else (C.dates[-1].month if C.dates else 12)
        if by == "fy" and C.n:
            groups = _groups(C.dates, end_month, C.periodicity)
            stmts = _to_fy(stmts, groups)
            periods = [{"label": lab, "end": str(C.dates[ix[-1]]), "n": len(ix), "partial": part} for lab, ix, part in groups]
        else:
            periods = [{"label": _plabel(d, C.periodicity), "end": str(d), "n": 1, "partial": False} for d in C.dates]
        tol = _tol_of(stmts, tol_cfg)
        _finalize(stmts, tol)
        labels = [p["label"] for p in periods]
        residuals = _residuals(stmts, labels, tol)
        out = {"version": __version__, "workbook": result.get("workbook") or depgraph._workbook(db_path), "by": by,
               "periodicity": C.periodicity, "fy_end_month": end_month, "periods": periods, "units": units,
               "tolerance": tol, "statements": {k: _stmt_json(S) for k, S in stmts.items()}, "residuals": residuals}
        cnt = {"extracted": 0, "derived": 0, "missing": 0}
        for S in stmts.values():
            for ln in S.lines.values():
                if not ln.hidden and ln.core:
                    cnt[ln.source] += 1
        out["stats"] = {"secs": round(time.time() - t0, 2), "periods": len(periods), "lines": cnt,
                        "modes": {k: S.mode for k, S in stmts.items()}}
        return depgraph._clean(out)
    finally:
        book.db.close()


def _plabel(d, periodicity):
    return f"FY{d.year}" if periodicity == "annual" else f"{d:%b-%y}"


def _units(C, stmts):
    seen = {}
    for S in stmts.values():
        for ln in S.lines.values():
            if ln.source != "missing" and not ln.hidden:
                for f in ln.frm[:1]:
                    r = C.b.db.execute("SELECT units FROM rows WHERE sheet=? AND row=?", (f["sheet"], f["row"])).fetchone()
                    if r and r[0] and str(r[0]).strip():
                        seen[str(r[0]).strip()] = seen.get(str(r[0]).strip(), 0) + 1
    return max(seen, key=lambda k: seen[k]) if seen else ""


def _stmt_json(S):
    lines = []
    for ln in S.lines.values():
        if ln.hidden:
            continue
        lines.append({"key": ln.key, "label": ln.label, "level": ln.level, "kind": ln.kind, "values": _clean(ln.v),
                      "source": ln.source, "from": ln.frm, "formula": ln.formula, "note": "; ".join(dict.fromkeys(ln.note)) or None})
    return {"title": S.title, "mode": S.mode, "why": S.why, "lines": lines, "checks": S.checks}


# ---- text and CSV -------------------------------------------------------------------------------------------------------

def _num(x):
    return "" if x is None else f"{(0.0 if abs(x) < 0.05 else x):,.1f}"


def text(tw: dict, max_cols: int | None = None) -> str:
    labels = [p["label"] for p in tw["periods"]]
    cols = list(range(len(labels)))[:max_cols] if max_cols else list(range(len(labels)))
    out = [f"{tw['workbook']}: three statements by {'financial year' if tw['by'] == 'fy' else 'period'} "
           f"({tw['stats']['secs']} s), Model Atlas {tw.get('version') or __version__}" + (f", {tw['units']}" if tw["units"] else "")]
    for k in STATEMENTS:
        s = tw["statements"][k]
        out += ["", f"{s['title']}  [{s['mode']}]  {s['why']}"]
        w = max([len(l["label"]) + 2 * l["level"] for l in s["lines"]] + [10])
        out.append(" " * (w + 11) + "".join(f"{labels[i]:>11}" for i in cols))
        for l in s["lines"]:
            tag = {"extracted": "", "derived": "(derived)", "missing": "(not found)"}[l["source"]]
            vals = "".join(f"{_num(l['values'][i]):>11}" for i in cols)
            out.append(f"{'  ' * l['level'] + l['label']:<{w}} {tag:<10}{vals}")
        for c in s["checks"]:
            out.append(f"  check {c['status']:<8} {c['title']}" + (f", max residual {c['max_residual']:.3g}" if c["max_residual"] is not None else "")
                       + (f" ({c['why']})" if c["why"] else ""))
    out += ["", "Residual analysis"]
    for r in tw["residuals"]:
        out.append(f"  [{r['status']}] {r['title']}: {r['reading']}")
    out += ["", NOTICE]
    return "\n".join(out)


_FORMULA_START = ("=", "+", "-", "@", "\t", "\r", "\n")


def _text_cell(x) -> str:
    """A text cell for a spreadsheet: one that would start a formula (=, +, -, @, or a control character, after any
    leading spaces) is prefixed with an apostrophe, so a model's label is never run as a formula when the CSV is opened."""
    x = "" if x is None else str(x)
    return "'" + x if x.lstrip(" ").startswith(_FORMULA_START) else x


def csv(tw: dict, statement: str) -> str:
    """One statement as CSV: Line, Source, Note, then a column per period. The checks follow the lines."""
    if statement not in STATEMENTS:
        raise ValueError("statement must be one of pnl, cf, bs")
    s = tw["statements"][statement]
    buf = io.StringIO()
    w = _csv.writer(buf, lineterminator="\n")
    unit = f" ({tw['units']})" if tw.get("units") else ""
    t = _text_cell
    w.writerow([t(f"{s['title']}{unit}"), "Source", "Note"] + [t(p["label"] + ("*" if p.get("partial") else "")) for p in tw["periods"]])
    for l in s["lines"]:
        w.writerow([t(("  " * l["level"]) + l["label"]), t(l["source"]), t(l.get("note") or l.get("formula") or "")] +
                   ["" if x is None else repr(x) for x in l["values"]])
    for c in s["checks"]:
        w.writerow([t(f"check: {c['title']}"), t(c["status"]), t(c.get("why") or "")] + ["" if x is None else repr(x) for x in c["residuals"]])
    return buf.getvalue()


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("db")
    ap.add_argument("--fy", action="store_true", help="by financial year instead of by model period")
    ap.add_argument("--fy-end-month", type=int, help="the month the financial year ends in (default: the month the timeline ends in)")
    ap.add_argument("--json", help="write the full result here")
    ap.add_argument("--csv", help="write pnl.csv, cf.csv and bs.csv into this folder")
    ap.add_argument("--tolerance", type=float, help="relative tolerance (default 1e-6)")
    ap.add_argument("--version", action="version", version=version_line())
    a = ap.parse_args(argv)
    tw = build(a.db, by="fy" if a.fy else "period", fy_end_month=a.fy_end_month, tolerance=a.tolerance)
    if a.json:
        Path(a.json).write_text(json.dumps(tw, indent=1, ensure_ascii=False), encoding="utf-8")
    if a.csv:
        d = Path(a.csv)
        d.mkdir(parents=True, exist_ok=True)
        for k in STATEMENTS:
            (d / f"{k}.csv").write_text(csv(tw, k), encoding="utf-8")
    print(text(tw))


def cli() -> None:
    main()


if __name__ == "__main__":
    cli()
