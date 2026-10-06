"""What a DCF value depends on: the graph behind one value cell, classified and bounded so a person can read it.

A value in a workbook (an equity value, an enterprise value) rests on a few kinds of things: the discounting (the
rate, the valuation date, the factors), a terminal value, the cash flows and the rows they are made of, the
calculations and timeline under those, and the typed assumptions at the bottom; plus a bridge (debt, cash, a
distribution payable) that turns a value of the business into a value of the equity. This module draws that
picture from a workbook's model.db. It is the DCF-shaped view of the row dependency graph (edges.py).

How it goes:
  1. dcftrace.trace() follows the cells from the start cell down to the discounting. Each cell on the way is a
     cell node; the cells that sit beside the path (net debt, cash) are the bridge terms.
  2. The discounting's own inputs (cash-flow row, factor row, rate, valuation date, period dates) and the rows the
     cash flow is made of seed the rows; the rows the bridge terms read seed a shallow look behind them.
  3. From the seeds a breadth-first walk goes upstream over the edges the current scenario uses (direct, offset,
     active; modeldash.LIVE). It stops at `depth` rows from a seed and at `max_rows` row nodes, taking the
     shallower rows first, then the discounting's own rows, then the rows with the most inputs. Everything it doesn't show is counted, and collapsed into one group
     node per sheet ("812 more rows on Ops"), so a 400k-formula model stays a page. Check rows (outputs.CHECK) are
     left out and counted; rows a lookup considered but didn't select (inactive) are counted on the row that read them.
  4. Every node gets a class by the rule table below: what the model's own structure says first (the cash-flow
     row of the discounting, its factor row), then what flows from it (what feeds the rate, what only the bridge
     reads), then the labels; a typed input none of those claim is an assumption (one they claim, like a typed
     discount rate, keeps that class and is flagged `input`). Each carries a one-sentence reason naming the
     evidence, so a person can see why and later change it.
  5. Layers: the longest path from the value, so a renderer can put the value at one side and the inputs at the other.

Some rules read more than the labels, each on evidence the model shows:
  - when the discounted row is not a plain same-period sum (`=SUM(a, b:c)*flag`), dcftrace names no parts; the cash-flow
    row's own direct reads (minus the discounting and timeline rows) are then the parts, "read directly".
  - the edges are per row, so a SUMPRODUCT sitting on the cash-flow row itself would give that row the factor, rate and
    date rows although its period columns never read them; those edges are dropped (stats.edges_dropped), unless a
    period column reads what its text can't tell (INDIRECT, OFFSET, a whole column).
  - a pasted input with a total (212 constants, 2 formulas) is an input: constants at least 3x the formulas.
  - a row of constants and formulas that reads only the scenario selector (a short small-integer row that 10+ lookup
    rows, INDEX / CHOOSE / OFFSET / MATCH, read; not a period counter) is a scenario-selected assumption (a structural
    class is kept, flagged input).

An older model.db with no edges table still gets a graph: the trace's cells and the rows the discounting names.

No model (LLM) calls. The result is a dict of plain values that json.dumps takes; modelatlas/depgraph_html.py draws it.

    uv run python -m modelatlas.depgraph out/<dir>/model.db                    # lists the DCF anchors, takes the top one
    uv run python -m modelatlas.depgraph out/<dir>/model.db "Summary!C5"       # from this cell
        --depth N (default 6)  --max-rows N (default 300)  --json path  --html path (default: beside model.db)
"""
import argparse
import json
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from . import __version__, version_line
from . import dcf
from . import dcftrace
from . import outputs
from . import rodb
from . import valuation
from .modeldash import LIVE

CLASSES = [("outcome", "Value outcome"), ("bridge", "Bridge (EV to equity)"), ("discounting", "Discounting"),
           ("terminal", "Terminal value"), ("cashflow", "Cash flow"), ("calculation", "Calculation"),
           ("timeline", "Timeline"), ("assumption", "Assumption")]
BRIDGE_DEPTH = 2  # rows read by a bridge term are looked at this far (the term's inputs and theirs)
SELECTOR_READERS = 10  # lookup rows that must read a row for it to be the scenario selector
SELECTOR_CELLS = 20    # a selector is a switch and the list it picks from, not a row of periods
COUNTER = re.compile(r"\b(?:period|year|quarter|month|counter)s?\b", re.I)  # a period counter's label (not 'flag': a scenario flag is one)
TYPED_RATIO = 3        # a row is typed (an input) when its constants are at least this many times its formulas
BUILDUP_HOPS = 3  # rows this close upstream of the discount rate that read like a rate build-up
MAX_FIELD = 400   # longest pattern / words text kept on a node

# ---- the rule table: label evidence (regex, subclass); first match wins in each table --------------------------
# 'terminal' alone is not enough: a port's "Terminal handling charges" or "Terminal B revenue" is a cash-flow line
TERMINAL = re.compile(r"terminal\s*(?:value|val\b|growth|multiple|cash|year|period|date|\)|$)|\bTV\b|exit multiple|"
                      r"perpetuity", re.I)
TERMINAL_SUB = [(r"growth|perpetuity", "terminal growth"), (r"multiple", "exit multiple"), (r"", "terminal value")]
DISCOUNTING = [  # (label regex, subclass, fewest cells): a row named for the discounting is the discounting, typed or formula
    (r"discount factor|discounting factor", "discount factor", 1),
    (r"discount rate|\bwacc\b", "discount rate", 1),
    (r"mid[- ]?(period|year)|discount(ing)? (convention|timing)|periods? to discount", "timing", 1),
    (r"^valuation date|^date of valuation|^as at date", "valuation date", 1),
    (r"^present value|^pv\b|present value \(|\bpv of\b", "present value", 3),  # a row of per-period values, not one total
    (r"cost of (equity|debt|capital)", "discount rate build-up", 1),
]
BUILDUP = re.compile(r"wacc|cost of (equity|debt|capital)|beta|gearing|premium|tax rate|post-tax|pre-tax|debt margin|"
                     r"risk[- ]free|risk premium|market risk", re.I)
BUILDUP_CORE = re.compile(r"wacc|beta|cost of", re.I)  # a row of the cash flow's lineage is a build-up only if it says so
DEBT_FUNDING = re.compile(r"drawdown|fund(ed|ing)?|facility|repay|refinanc", re.I)  # debt moves, not operating cash
CASHFLOW = re.compile(r"cash ?flow|fcf|ebitda|ebit\b|distribution|dividend|capex|capital expenditure|tax paid|"
                      r"working capital|revenue|opex|operating cost", re.I)
CASHFLOW_ALIAS = {"capex": "capital expenditure", "fcf": "free cash flow", "cashflow": "cash flow", "opex": "operating cost"}
TIMELINE = re.compile(r"\b(?:period|date|year|quarter|month|flag|counter)s?\b|days in", re.I)
ASSUMPTIONS = [  # outputs.ASSUMPTION's vocabulary, split the way a reader asks about it
    (r"wacc|discount|cost of (equity|debt|capital)|\bbeta\b|premium|gearing|risk", "discount rate"),
    (r"valuation date", "valuation date"),
    (r"terminal|perpetuity|exit multiple", "terminal growth"),
    (r"\btax", "tax"),
    (r"\bcpi\b|inflation|indexation|escalat", "inflation / CPI"),
    (r"growth", "growth"),
    (r"volume|traffic|price|toll|tariff|units|demand|quantity|occupancy|throughput", "volume / price"),
    (r"opex|operating|cost|margin|capex|capital expenditure|maintenance|working capital|overhead|salary|wage", "operating"),
    (r"debt|interest|loan|equity|dividend|distribution|drawdown|repayment|facility|financ|cash", "financing"),
]
CALCULATIONS = [  # what a row that is none of the above is, by its label (only to help a reader)
    (r"volume|traffic|price|toll|tariff|units|demand|quantity", "volume / price"),
    (r"index|cpi|inflation|escalat", "indexation"),
    (r"debt|interest|loan|repayment|drawdown|facility", "debt schedule"),
    (r"\btax", "tax"),
    (r"balance|opening|closing", "balance"),
]
BRIDGE_SUB = [(r"debt|borrowing|loan", "debt"), (r"cash", "cash"), (r"deferred|consideration", "deferred consideration"),
              (r"distribution|dividend", "distribution payable"), (r"minorit|non-controlling", "minorities")]


def _typed(m) -> bool:
    """An input row: constants, and at most a total or two among them (at least 3x as many constants as formulas)."""
    return m["n_const"] > 0 and m["n_const"] >= TYPED_RATIO * m["n_formula"]


def _sub(table, text, default):
    for pat, sub in table:
        if re.search(pat, text or "", re.I):
            return sub
    return default


def _rid(sheet: str, row: int) -> str:
    return f"{sheet}!r{row}"


def _clip(s, n=MAX_FIELD):
    s = None if s is None else str(s)
    return s if s is None or len(s) <= n else s[:n - 1] + "…"


def _clean(o):
    """JSON-able: dates as ISO strings, numpy numbers as floats, non-finite numbers as None."""
    if isinstance(o, dict):
        return {str(k): _clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple, set)):
        return [_clean(v) for v in o]
    if isinstance(o, (date, datetime)):
        return o.isoformat()
    if hasattr(o, "item") and not isinstance(o, (str, bytes)):
        o = o.item()
    if isinstance(o, float) and (o != o or o in (float("inf"), float("-inf"))):
        return None
    return o if o is None or isinstance(o, (str, int, float, bool)) else str(o)


def _workbook(db_path: str) -> str:
    """The workbook's name from the model.db folder ('Name__9a325c1f' -> 'Name.xlsx'; the sheets table doesn't keep it)."""
    name = re.sub(r"__[0-9a-f]{8}$", "", Path(db_path).resolve().parent.name)
    return name if re.search(r"\.xls[xm]$", name, re.I) else name + ".xlsx"


# ---- the model, loaded once ---------------------------------------------------------------------------------

class Model:
    """Everything the walk needs from model.db, read once: the live edges both ways (a dict, not a query per row),
    the inactive edges counted, each line item's label and counts, each sheet's timeline row."""

    def __init__(self, db):
        self.db = db
        self.up = defaultdict(list)    # (sheet, row) -> the rows it reads (live edges)
        self.down = defaultdict(list)  # (sheet, row) -> the rows that read it
        self.kind = {}                 # ((src), (dst)) -> edge kind
        self.inactive = Counter()
        cols = [r[1] for r in db.execute("PRAGMA table_info(edges)")]
        self.has_edges = bool(cols)  # an older model.db has no edges table: the graph is then the trace and its seeds
        for ss, sr, ds, dr, k in (db.execute(f"SELECT src_sheet, src_row, dst_sheet, dst_row, "
                                             f"{'kind' if 'kind' in cols else repr('direct')} FROM edges") if cols else ()):
            if k in LIVE:
                self.up[(ss, sr)].append((ds, dr))
                self.down[(ds, dr)].append((ss, sr))
                self.kind[((ss, sr), (ds, dr))] = k
            else:
                self.inactive[(ss, sr)] += 1
        self.meta = {(s, r): {"section": sec or "", "label": (lab or "").strip(), "units": u or "",
                              "n_formula": nf or 0, "n_const": nc or 0}
                     for s, r, sec, lab, u, nf, nc in db.execute(
                         "SELECT sheet, row, section, label, units, n_formula, n_const FROM rows")}
        self.layout = {}
        for s, lay in db.execute("SELECT sheet, layout FROM sheets"):
            try:
                self.layout[s] = json.loads(lay or "{}")
            except ValueError:
                self.layout[s] = {}
        self.names = dcftrace._names(db)
        try:  # a row with no label of its own takes the defined name that points at it ('Scenario' -> Scenario!$F$7)
            named = sorted(db.execute("SELECT name, ref FROM names"))
        except Exception:  # noqa: BLE001 (an older model.db has no names table)
            named = []
        for n, ref in named:
            if n.lower().startswith("_xlnm."):  # print areas and titles aren't names a person gave a row
                continue
            r = dcf._ref((ref or "").lstrip("="), "") if ref and "!" in ref else None
            if r and r[0] and r[1] == r[3] and (r[0], r[1]) in self.meta and not self.meta[(r[0], r[1])]["label"]:
                self.meta[(r[0], r[1])]["label"] = n

    def label(self, key) -> str:
        return self.meta.get(key, {}).get("label", "")

    def is_check(self, key) -> bool:
        return bool(outputs.CHECK.search(self.label(key)))

    def row_of(self, text) -> tuple | None:
        """(sheet, row) of a cell, a range or a named range; also of 'Sheet!r12 label' (a trace's period source)."""
        if not isinstance(text, str) or not text.strip():
            return None
        t = text.strip()
        m = re.match(r"^'?(.+?)'?!r(\d+)(?:\s|$)", t)
        if m:
            return m[1], int(m[2])
        t = self.names.get(t.lower(), t).lstrip("=")
        r = dcf._ref(t, "")
        return (r[0], r[1]) if r and r[0] else None

    def first_formula(self, key):
        r = self.db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND formula IS NOT NULL ORDER BY col LIMIT 1",
                            key).fetchone()
        return r[0] if r else None

    def date_valued(self, key) -> bool:
        lab = self.layout.get(key[0], {}).get("label_col") or 0
        vals = [v for (v,) in self.db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col>? AND value IS NOT NULL "
                                              "ORDER BY col LIMIT 4", (*key, lab))]
        return bool(vals) and all(isinstance(v, str) and dcf._as_date(v) for v in vals)

    def selectors(self) -> set:
        """The scenario selector rows: a short small-integer row (at most SELECTOR_CELLS numbers, 0-99) that
        SELECTOR_READERS or more lookup rows (formulas with INDEX / CHOOSE / OFFSET / MATCH) read, and not a period
        counter (the sheet's timeline row, or a row labelled like a period / year / quarter / counter). Empty when the
        model has none (nothing is then scenario-selected)."""
        if getattr(self, "_selectors", None) is None:
            lookups = set(self.db.execute("SELECT DISTINCT sheet, row FROM cells WHERE formula LIKE '%INDEX(%' OR "
                                          "formula LIKE '%CHOOSE(%' OR formula LIKE '%OFFSET(%' OR formula LIKE '%MATCH(%'"))
            readers = Counter(c for k in lookups for c in self.up.get(k, ()))
            self._selectors = set()
            for k, n in readers.items():
                if (n < SELECTOR_READERS or k not in self.meta or self.layout.get(k[0], {}).get("header_row") == k[1]
                        or COUNTER.search(self.label(k))):
                    continue
                vals = [v for (v,) in self.db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND value IS NOT NULL", k)
                        if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if 2 <= len(vals) <= SELECTOR_CELLS and all(float(v).is_integer() and 0 <= v <= 99 for v in vals):
                    self._selectors.add(k)
        return self._selectors

    def closure(self, seeds):
        """Every row upstream of these (live edges, checks left out) -> (set of rows, check rows met)."""
        seen, checks, stack = set(seeds), set(), list(seeds)
        while stack:
            for c in self.up.get(stack.pop(), ()):
                if c in seen:
                    continue
                if self.is_check(c):
                    checks.add(c)
                    continue
                seen.add(c)
                stack.append(c)
        return seen, checks


# ---- the anchors --------------------------------------------------------------------------------------------

def anchors(db_path: str) -> list[dict]:
    """The DCF value cells of the workbook, as valuation.catalogue orders them (the first usable one is the top).
    The first call on a big workbook can take minutes; the catalogue is saved beside model.db."""
    return [{"cell": b["cell"], "label": b.get("label") or "", "value": _clean(b.get("value")), "ok": bool(b.get("ok")),
             "matches": bool(b.get("matches"))} for b in valuation.catalogue(db_path)]


# ---- the walk -----------------------------------------------------------------------------------------------

def _parts_walk(parts, parent, out, level=1):
    """The core's parts as (parent row, row, sign, level) with the rows named 'Sheet!rN'."""
    for p in parts or []:
        if p.get("external"):
            continue
        m = re.match(r"^(.+)!r(\d+)$", p["row"])
        if not m:
            continue
        key = (m[1], int(m[2]))
        out.append((parent, key, p.get("sign", 1), level))
        _parts_walk(p.get("parts"), key, out, level + 1)


def _tree_cells(tree):
    """The trace tree flattened: [(node, parent cell or None)], and the parent -> child cell pairs."""
    nodes, links, seen = [], [], set()

    def walk(n, parent):
        if parent:
            links.append((parent, n["cell"]))
        if n["cell"] in seen:
            return
        seen.add(n["cell"])
        nodes.append((n, parent))
        for c in n.get("children", []):
            walk(c, n["cell"])
    walk(tree, None)
    return nodes, links


def _scales(db, names, formula: str | None, here: str, cell: str, path: set) -> bool:
    """Whether the formula (on sheet `here`) only multiplies or divides the discounted value by this cell (=EV/Thousand,
    =PV*Share): a units divisor or a share is not a bridge term. False when the cell is a + / - term of its own
    (=EV-Debt, =EV-Debt*Share), sits inside a function, or it can't be told. `path`: the cells on the value's path."""
    targets = [r for r in (dcf._ref(c, "") for c in path) if r]
    me = dcf._ref(cell, "")
    if not formula or not me:
        return False
    body = dcftrace._expand(db, dcftrace._STR.sub('""', formula), names).strip().lstrip("=")

    def at(t, r):
        x = dcf._ref(valuation._strip(t), here)
        return bool(x) and x[0] == r[0] and x[1] <= r[1] <= x[3] and x[2] <= r[2] <= x[4]
    scales = False
    for _, term in valuation._split(body, "+-"):
        factors = [f for _, f in valuation._split(valuation._strip(term), "*/")]
        if not any(at(f, me) for f in factors):
            continue
        if len(factors) == 1 or not any(at(f, r) for f in factors for r in targets):
            return False  # added or taken away (alone or scaled itself): a bridge term
        scales = True
    return scales


_RANGE = re.compile(r"(?:(?:'(?P<q>(?:[^']|'')+)'|(?P<s>[A-Za-z_][\w.]*))!)?(?P<a>\$?[A-Z]{1,3}\$?\d+)(?::(?P<b>\$?[A-Z]{1,3}\$?\d+))?"
                    r"(?![\w(])")
_BLIND = re.compile(r"\b(?:INDIRECT|OFFSET)\(|(?<![\w.$'])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}\b|(?<![\w.$'])\$?\d+:\$?\d+(?![\w.])", re.I)


def _rows_read(db, names, formulas, here: str) -> list | None:
    """The (sheet, first row, last row) of every cell or range in these formulas (named ranges expanded, text left out);
    None when a formula reads what can't be told from its text (INDIRECT, OFFSET, a whole column or row)."""
    out = []
    for f in formulas:
        body = dcftrace._expand(db, dcftrace._STR.sub('""', f), names)
        if _BLIND.search(body):
            return None
        for m in _RANGE.finditer(body):
            if body[:m.start()].endswith(("!", ".", "$")) and not (m["q"] or m["s"]):
                continue
            a = dcf._ref(f"{m['a']}:{m['b']}" if m["b"] else m["a"], "")
            if a:
                sheet = m["q"].replace("''", "'") if m["q"] else (m["s"] or here)
                out.append((sheet, a[1], a[3]))
    return out


class Context:
    """The structural facts the rules read: which rows play which part in the discounting, what is a bridge row,
    which are upstream of the cash flow or the rate."""

    def __init__(self):
        self.roles = {}         # row -> (class, subclass, why): rows the discounting itself names
        self.part_rows = {}     # row -> why: rows the cash flow is made of
        self.part_direct = set()  # of those, the rows read straight off the cash-flow row (dcftrace named no parts)
        self.bridge_only = {}   # row -> the bridge term it feeds
        self.buildup = set()    # rows upstream of the rate that read like a rate build-up
        self.exp_factor = set() # formula rows that raise to / EXP the rate or the factor
        self.cf_lineage = set() # rows upstream of a cash-flow row
        self.terminal_inputs = set()
        self.scenario_selected = {}  # row -> the selector row it reads
        self.timeline_rows = {}  # sheet -> its timeline (period header) row


def _rule_role(ctx, model, key, m, typed):
    return ctx.roles.get(key)


def _rule_bridge(ctx, model, key, m, typed):
    term = ctx.bridge_only.get(key)
    if term:
        return ("bridge", _sub(BRIDGE_SUB, m["label"] + " " + term, "bridge input" if typed else "bridge calculation"),
                f"read only by the bridge term {term}, not by the discounted cash flows")


def _rule_discounting_graph(ctx, model, key, m, typed):
    if key in ctx.buildup:
        return ("discounting", "discount rate build-up", f"within {BUILDUP_HOPS} rows upstream of the discount rate and "
                f"labelled like one of its parts")
    if key in ctx.exp_factor:
        return ("discounting", "discount factor", "raises the discount rate or the factor row to a power (or EXP)")


def _rule_terminal(ctx, model, key, m, typed):
    if TERMINAL.search(m["label"]):
        if re.search(r"\bdate\b", m["label"], re.I) or model.date_valued(key):  # "Terminal Value Date": when, not what
            return ("timeline", "terminal date", "a date labelled for the terminal value")
        return ("terminal", _sub(TERMINAL_SUB, m["label"], "terminal value"), "labelled like the terminal value")
    if key in ctx.terminal_inputs:
        return ("terminal", "terminal input", "a typed input read only by the terminal value rows")


def _rule_discounting_label(ctx, model, key, m, typed):
    for pat, sub, cells in DISCOUNTING:
        if re.search(pat, m["label"], re.I) and m["n_formula"] + m["n_const"] >= cells:
            return ("discounting", sub, f"labelled like the {sub}")


def _rule_part(ctx, model, key, m, typed):
    if key in ctx.part_rows and not typed:
        return ("cashflow", "cash flow part (read directly)" if key in ctx.part_direct else "cash flow part", ctx.part_rows[key])


def _rule_timeline_structure(ctx, model, key, m, typed):
    if ctx.timeline_rows.get(key[0]) == key[1]:
        return ("timeline", "period dates", f"the timeline row of sheet {key[0]}")
    if model.date_valued(key):
        return ("timeline", "dates", "its values are dates")


def _rule_scenario(ctx, model, key, m, typed):
    sel = ctx.scenario_selected.get(key)
    if sel:
        return ("assumption", "scenario-selected", f"a typed value picked by scenario: its only live input is the selector row {_rid(*sel)}")


def _rule_cashflow(ctx, model, key, m, typed):
    hit = CASHFLOW.search(m["label"])
    if hit and not typed and key in ctx.cf_lineage:
        if DEBT_FUNDING.search(m["label"]):  # drawdowns and repayments move debt; they are not operating cash flow
            return ("calculation", "debt schedule", "upstream of the cash flows but labelled like debt funding, not a cash flow line")
        w = hit[0].lower()
        return ("cashflow", CASHFLOW_ALIAS.get(w, w), f"upstream of the discounted cash flows and labelled {w}")


def _rule_timeline_label(ctx, model, key, m, typed):
    hit = TIMELINE.search(m["label"])
    if hit and m["n_formula"] + m["n_const"] >= 3:  # a row of periods, not a single figure that mentions a date
        return ("timeline", hit[0].lower(), f"labelled {hit[0].lower()}")


def _rule_assumption(ctx, model, key, m, typed):
    if typed:
        sub = _sub(ASSUMPTIONS, m["label"], "other")
        return ("assumption", sub, f"a typed input ({m['n_const']} value{'s' if m['n_const'] != 1 else ''}, "
                + (f"{m['n_formula']} formula{'s' if m['n_formula'] != 1 else ''} among them)" if m["n_formula"] else "no formula)"))


def _rule_default(ctx, model, key, m, typed):
    return ("calculation", _sub(CALCULATIONS, m["label"], "calculation"), "a formula row that is none of the above")


# Applied in this order; the first that answers is the class. Add a rule by adding a function and a line here.
RULES = [_rule_role, _rule_bridge, _rule_discounting_graph, _rule_terminal, _rule_discounting_label, _rule_part,
         _rule_timeline_structure, _rule_scenario, _rule_cashflow, _rule_timeline_label, _rule_assumption, _rule_default]


def classify(ctx: Context, model: Model, key: tuple) -> tuple[str, str, str]:
    """(class, subclass, why) for one row."""
    m = model.meta.get(key) or {"label": "", "n_formula": 1, "n_const": 0}
    typed = _typed(m)
    for rule in RULES:
        got = rule(ctx, model, key, m, typed)
        if got:
            sel = ctx.scenario_selected.get(key)
            if sel and "scenario" not in got[2]:  # a structural class stays; the why says what picks its value
                got = (got[0], got[1], f"{got[2]}; picked by scenario (reads only the selector row {_rid(*sel)})")
            return got
    return ("calculation", "calculation", "")


# ---- the layers ---------------------------------------------------------------------------------------------

def _layers(start: str, ids: list[str], edges: list[dict]) -> dict[str, int]:
    """Longest path from the start along src -> dst. Back edges (Excel circularities) are ignored."""
    out = defaultdict(list)
    for e in edges:
        out[e["src"]].append(e["dst"])
    state, order = {}, []
    for root in [start] + ids:
        if root in state:
            continue
        state[root] = 1
        stack = [(root, iter(out[root]))]
        while stack:
            n, it = stack[-1]
            for c in it:
                if c not in state:
                    state[c] = 1
                    stack.append((c, iter(out[c])))
                    break
            else:
                state[n] = 2
                order.append(n)
                stack.pop()
    pos = {n: i for i, n in enumerate(reversed(order))}  # a topological order once back edges are skipped
    layer = {start: 0}
    for n in reversed(order):
        if n not in layer:
            continue
        for c in out[n]:
            if pos[c] > pos[n]:
                layer[c] = max(layer.get(c, 0), layer[n] + 1)
    top = max(layer.values(), default=0)
    return {n: layer.get(n, top + 1) for n in ids}


# ---- build --------------------------------------------------------------------------------------------------

def build(db_path: str, cell: str | None = None, depth: int = 6, max_rows: int = 300, workbook: str | None = None,
          with_anchors: bool = True) -> dict:
    """The dependency graph behind a DCF value cell (the top anchor when cell is None), as a dict of plain values."""
    t0 = time.time()
    anc = []
    if with_anchors or cell is None:
        try:
            anc = anchors(db_path)
        except Exception:  # noqa: BLE001 (an unreadable workbook still gets its graph from the cell)
            anc = []
    if cell is None:
        top = next((a for a in anc if a["ok"]), None)
        if not top:
            raise ValueError("no DCF value cell found in this workbook; give one (Sheet!C5)")
        cell = top["cell"]
    db = rodb.connect(db_path)
    model = Model(db)
    tree = dcftrace.trace(db, cell)
    cores = dcftrace.cores(tree)
    cells, links = _tree_cells(tree)
    cell_row = {n["cell"]: model.row_of(n["cell"]) for n, _ in cells}
    cell_rows = set(cell_row.values())
    start_id = tree["cell"]
    ctx = Context()
    for s, lay in model.layout.items():
        if lay.get("header_row"):
            ctx.timeline_rows[s] = lay["header_row"]

    # -- the discounting's own rows, their roles and the parts of the cash flow
    structural: dict[tuple, int] = {}   # row -> seed depth: the rows the discounting names
    part_edges: dict[tuple, int] = {}   # (src row, dst row) -> sign
    core_links = []                     # (core cell, row, why) edges from the core to the rows it names
    rate_rows, factor_rows, pv_rows, cf_rows = [], [], [], []
    per_core = []                       # (core, cash-flow row, rows that core's discounting names) for the passes below
    for c in cores:
        at = c["cell"]
        m = c.get("method") or {}

        def add(key, role, depth_=1):
            if key and key in model.meta:
                structural.setdefault(key, depth_)
                core_links.append((at, key))
                if role and key not in ctx.roles:
                    ctx.roles[key] = role
            return key
        cf = model.row_of(c.get("cashflow"))
        if cf and cf in model.meta:
            cf_rows.append(cf)
        add(cf, ("cashflow", "discounted cash flow", f"the cash-flow row of the {c['kind']} at {at}"))
        fr = model.row_of(c.get("factor_row")) if c.get("factor_row") and "!" in c["factor_row"] else None
        if fr:
            factor_rows.append(fr)
        add(fr, ("discounting", "discount factor", f"the factor row of the {c['what'].split(' of ')[0]} at {at}"))
        rr = model.row_of(m.get("rate"))
        if rr:
            rate_rows.append(rr)
        add(rr, ("discounting", "discount rate", f"the discount rate cell {m.get('rate')} used by {at}"), 1)
        vr = model.row_of(m.get("valuation_date"))
        add(vr, ("discounting", "valuation date", f"the valuation date cell {m.get('valuation_date')} used by {at}"), 1)
        pr = model.row_of(c.get("pv_row"))
        if pr:
            pv_rows.append(pr)
        add(pr, ("discounting", "present value", f"the present-value row summed at {at}"), 1)
        dr = model.row_of(c.get("dates") or m.get("ends_source"))
        add(dr, ("timeline", "period dates", f"the period dates {at}'s discounting counts from"), 1)
        rows_ = []
        _parts_walk(c.get("parts"), cf, rows_)
        per_core.append((c, cf, fr, [k for k in (fr, rr, vr, dr) if k], not rows_))
        for parent, key, sign, level in rows_:
            if key in model.meta:
                part_edges[(parent, key)] = sign
                structural.setdefault(key, 1 + level)
                ctx.part_rows.setdefault(key, f"a part of the cash-flow row {_rid(*cf)} ({'+' if sign > 0 else '-'})")
        for adj in (c.get("inputs") or {}).get("adjustments", []):
            ar = model.row_of(adj.get("source") or (adj.get("value") if isinstance(adj.get("value"), str) else None))
            if ar and ar in model.meta and ar not in structural:
                structural[ar] = 1
                ctx.bridge_only.setdefault(ar, adj.get("label") or "an adjustment")
                core_links.append((at, ar))

    # -- the rows the factor row reads itself (year fractions, period counts), and the cash flow's direct terms when
    #    dcftrace named no parts: both read off the model's edges, minus rows that already have a part to play
    def is_disc_label(k):
        return any(re.search(pat, model.label(k), re.I) for pat, _, _ in DISCOUNTING)

    def free(k):  # a row not already a discounting / timeline row, not a check, not a date row
        return (k in model.meta and k not in ctx.roles and not model.is_check(k) and ctx.timeline_rows.get(k[0]) != k[1]
                and not model.date_valued(k) and not is_disc_label(k)
                and not TERMINAL.search(model.label(k)))
    for c, cf, fr, _, _ in per_core:
        for k in (model.up.get(fr, ()) if fr else ()):
            if free(k) and k != cf:
                ctx.roles[k] = ("discounting", "discount period", f"read directly by the factor row {_rid(*fr)}: the periods "
                                f"(year fractions from the valuation date) the factors are raised to")
                structural.setdefault(k, 2)
    for c, cf, fr, _, no_parts in per_core:
        for k in (model.up.get(cf, ()) if cf and no_parts else ()):
            if free(k) and k not in ctx.part_rows and k != cf:
                ctx.part_rows[k] = f"read directly by the cash-flow row {_rid(*cf)} (dcftrace names no parts: not a plain sum)"
                ctx.part_direct.add(k)
                structural.setdefault(k, 1)  # the core cell's row reads them: seeds at depth 1, as the cell seeds had them

    # -- the edges the per-row graph invents: the SUMPRODUCT cell sits on the cash-flow row, so the row 'reads' the
    #    factor / rate / dates although its period columns never do
    dropped: set = set()
    for c, cf, fr, named, _ in per_core:
        if cf and model.row_of(c["cell"]) == cf:
            try:
                sheet, row, cols = dcf._row_range(db, c["cashflow"])
                forms = [f for (f,) in db.execute("SELECT DISTINCT formula FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ? "
                                                  "AND formula IS NOT NULL", (sheet, row, cols[0], cols[-1]))]
            except Exception:  # noqa: BLE001 (a range we can't read keeps its edges)
                continue
            read = _rows_read(db, model.names, forms, sheet) if forms else None
            if read is not None:  # None: a period column reads something we can't see (INDIRECT ...): keep the edges
                dropped.update((cf, k) for k in named
                               if k != cf and not any(s == k[0] and lo <= k[1] <= hi for s, lo, hi in read))

    # -- seeds from the cells: what their rows read that the tree doesn't already show
    off_path = {n["cell"] for n, _ in cells if not n.get("on_path") and n["cell"] != start_id}
    # a core's own operands (its cash-flow and factor cells, =SUMPRODUCT(D10:K10,D12:K12)) are the discounting, not bridge terms
    operand_rows = {c["cell"]: {cf, fr} - {None} for c, cf, fr, _, _ in per_core}
    off_path -= {ch for p, ch in links if cell_row.get(ch) in operand_rows.get(p, ())}
    by_cell = {n["cell"]: n for n, _ in cells}
    on_path = {n["cell"] for n, _ in cells if n.get("on_path")}
    for n, parent in cells:  # beside the path but only scaling it (=EV/Thousand): not a bridge term
        if n["cell"] in off_path and parent and _scales(db, model.names, by_cell[parent].get("formula"),
                                                        (cell_row[parent] or ("",))[0], n["cell"], on_path):
            off_path.discard(n["cell"])
    seeds: dict[tuple, dict] = {}   # row -> {"depth", "lim", "bridge"}
    cell_seed_edges = []
    for key, d in structural.items():
        seeds[key] = {"depth": d, "lim": depth, "bridge": False, "struct": True}
    for n, _ in cells:
        ck = cell_row[n["cell"]]
        if not ck or not n.get("formula"):
            continue
        bridge = n["cell"] in off_path
        for dst in model.up.get(ck, ()):
            if dst in cell_rows or model.is_check(dst):
                continue
            cell_seed_edges.append((n["cell"], dst))
            s = seeds.get(dst)
            if s is None:
                seeds[dst] = {"depth": 1, "lim": BRIDGE_DEPTH if bridge else depth, "bridge": bridge, "struct": False,
                              "term": n.get("label") or n["cell"]}
            elif not bridge and s.get("bridge"):
                s.update(bridge=False, lim=depth)
            elif not bridge:
                s["lim"] = depth

    # -- the walk upstream, level by level, under the row budget
    shown: dict[tuple, int] = {}
    expanded: dict[tuple, int] = {}  # row -> the deepest limit it was walked with (a bridge reach first, the main one later)
    pending = defaultdict(dict)   # depth -> row -> lim
    for key, s in seeds.items():
        pending[s["depth"]][key] = s["lim"]
    for d in range(1, depth + 1):
        cand = [k for k, lim in pending.get(d, {}).items() if k not in shown or lim > expanded.get(k, 0)]
        cand.sort(key=lambda k: (not (seeds.get(k) or {}).get("struct"), -len(model.up.get(k, ())), k))
        for k in cand:
            if k not in shown:
                if len(shown) >= max_rows:
                    continue
                shown[k] = d
            lim = pending[d][k]
            expanded[k] = max(lim, expanded.get(k, 0))
            if d < min(lim, depth):
                for c in model.up.get(k, ()):
                    if (c in shown and expanded.get(c, 0) >= lim) or model.is_check(c):
                        continue
                    pending[d + 1][c] = max(lim, pending[d + 1].get(c, 0))
    closure, checks = model.closure(list(seeds))

    # -- structure read off the shown graph: cash-flow lineage, rate build-up, bridge-only rows, factors, terminal inputs
    def reach(starts, hops=None, avoid=()):
        seen = {s: 0 for s in starts if s in shown}
        queue = list(seen)
        while queue:
            k = queue.pop(0)
            if hops is not None and seen[k] >= hops:
                continue
            for c in model.up.get(k, ()):
                if c in shown and c not in seen and c not in avoid:
                    seen[c] = seen[k] + 1
                    queue.append(c)
        return seen
    # a row-level graph joins every cell of a row: a SUMPRODUCT sitting on the cash-flow row gives that row the factor
    # row's edges, so the lineage of the cash flow doesn't go through the discounting's own rows
    disc_rows = {k for k, r in ctx.roles.items() if r[0] in ("discounting", "timeline")}
    ctx.cf_lineage = set(reach(cf_rows, avoid=disc_rows))
    for k in reach(rate_rows, BUILDUP_HOPS):
        if k not in rate_rows and BUILDUP.search(model.label(k)) and (k not in ctx.cf_lineage or BUILDUP_CORE.search(model.label(k))):
            ctx.buildup.add(k)
    main_reach = set(reach([k for k, s in seeds.items() if not s["bridge"]]))
    for k, s in seeds.items():
        if s["bridge"] and k in shown and k not in main_reach:
            ctx.bridge_only.setdefault(k, s["term"])
            for c in model.up.get(k, ()):
                if c in shown and c not in main_reach and c not in structural:
                    ctx.bridge_only.setdefault(c, s["term"])
    for k in shown:
        for key_ in (rate_rows + factor_rows):
            if key_ in model.up.get(k, ()) and k not in ctx.roles:
                f = model.first_formula(k) or ""
                if "^" in f or re.search(r"\bEXP\(", f, re.I):
                    ctx.exp_factor.add(k)
    term_rows = {k for k in shown if TERMINAL.search(model.label(k))}
    for k in shown:
        m = model.meta[k]
        if _typed(m) and k not in term_rows:
            readers = model.down.get(k, ())
            if readers and all(TERMINAL.search(model.label(r)) for r in readers):
                ctx.terminal_inputs.add(k)
    cand = {k: [c for c in model.up.get(k, ()) if not model.is_check(c)] for k in shown
            if model.meta[k]["n_formula"] and model.meta[k]["n_const"] and not _typed(model.meta[k])}
    cand = {k: u[0] for k, u in cand.items() if len(u) == 1}
    if cand and model.has_edges:
        sel = model.selectors()
        ctx.scenario_selected = {k: u for k, u in cand.items() if u in sel}
    for k in list(ctx.bridge_only):  # structure beats the bridge's reach
        if k in ctx.roles:
            del ctx.bridge_only[k]

    # -- the nodes
    nodes, ids = [], set()

    def node(**kw):
        base = {"id": None, "kind": "row", "sheet": None, "row": None, "row_id": None, "cell": None, "label": "", "units": "",
                "section": "", "value": None, "formula": None, "words": None, "pattern": None, "samples": None,
                "n_formula": None, "n_const": None, "input": False, "class": "calculation", "subclass": "", "why": "",
                "layer": 0, "depth": 0, "reads": 0, "read_by": 0, "inactive_hidden": 0, "hidden_upstream": 0, "count": None}
        base.update(kw)
        nodes.append(base)
        ids.add(base["id"])
        return base

    def row_text(key):
        r = db.execute("SELECT patterns, samples FROM rows WHERE sheet=? AND row=?", key).fetchone()
        return (_clip(r[0]), _clip(r[1], 200)) if r else (None, None)

    def row_fields(key):
        m = model.meta.get(key, {"label": "", "units": "", "section": "", "n_formula": 0, "n_const": 0})
        pat, smp = row_text(key)
        return {"sheet": key[0], "row": key[1], "row_id": _rid(*key), "label": model.label(key) or _rid(*key),
                "units": m["units"], "section": m["section"], "pattern": pat, "samples": smp, "n_formula": m["n_formula"],
                "n_const": m["n_const"], "reads": sum(1 for c in model.up.get(key, ()) if not model.is_check(c)),
                "read_by": sum(1 for c in model.down.get(key, ()) if not model.is_check(c)), "inactive_hidden": model.inactive.get(key, 0)}

    core_at = {c["cell"]: c for c in cores}
    cell_class: dict[str, tuple] = {}
    kids = defaultdict(list)
    for p, ch in links:
        kids[p].append(ch)

    def cell_cls(cid):
        """Outcome / core / bridge term / a cell between them / else the class of its row."""
        if cid in cell_class:
            return cell_class[cid]
        n, key = by_cell[cid], cell_row[cid]
        for ch in kids[cid]:
            cell_cls(ch)
        if cid == start_id:
            out = ("outcome", "value outcome", "the cell the graph explains")
        elif cid in core_at:
            c = core_at[cid]
            m = c.get("method") or {}
            out = ("discounting", "present value", f"the {c['what'].split(' of ')[0]} that discounts {c['cashflow']}"
                   + (f" ({m['timing']}-of-period, {m['day_count']})" if m.get("timing") else ""))
        else:
            base = classify(ctx, model, key) if key else ("calculation", "calculation", "")
            if cid in off_path:
                out = base if base[0] in ("discounting", "terminal", "timeline") else (
                    "bridge", _sub(BRIDGE_SUB, n.get("label") or "", "bridge term"),
                    "a term beside the discounted value that turns it into the equity value")
            elif any(cell_class[ch][0] == "bridge" for ch in kids[cid] if ch in cell_class):
                out = ("bridge", "equity bridge", "adds the bridge terms to the discounted value")
            elif base[0] == "calculation" and cid in on_path:
                out = ("discounting", "discounted value", "a cell on the path between the value and the discounting core")
            else:
                out = base
        cell_class[cid] = out
        return out

    for n, _ in cells:
        cid, key = n["cell"], cell_row[n["cell"]]
        cls, sub, why = cell_cls(cid)
        f = row_fields(key) if key else {}
        v = n.get("value")
        node(id=cid, kind="cell", cell=cid, formula=n.get("formula"), value=v if isinstance(v, (str, int, float)) else None,
             words=_clip(n.get("words")), input=not n.get("formula"),
             **{**f, "label": n.get("label") or f.get("label") or cid, "inactive_hidden": 0}, **{"class": cls, "subclass": sub, "why": why})
    for key, d in sorted(shown.items(), key=lambda kv: (kv[1], kv[0])):
        cls, sub, why = classify(ctx, model, key)
        f = row_fields(key)
        formula = model.first_formula(key)
        words = None
        if formula:
            try:
                words = _clip(dcftrace.words(db, dcftrace._expand(db, formula, model.names), key[0]))
            except Exception:  # noqa: BLE001 (words are a convenience; the formula's rows are still linked)
                words = None
        node(id=_rid(*key), kind="row", words=words, input=_typed(f) or key in ctx.scenario_selected,
             **f, **{"class": cls, "subclass": sub, "why": why}, depth=d)

    # -- the edges
    edges: dict[tuple, dict] = {}

    def edge(src, dst, kind, sign=1):
        if src != dst:
            old = edges.get((src, dst))
            if old is None or kind == "part" or (old["kind"] not in ("trace", "part") and kind == "trace"):
                edges[(src, dst)] = {"src": src, "dst": dst, "kind": kind, "sign": sign}
    for p, ch in links:
        edge(p, ch, "trace")
    for at, key in core_links:
        if key in shown:
            edge(at, _rid(*key), model.kind.get((cell_row[at], key), "direct"))
    for at, key in cell_seed_edges:
        if key in shown:
            edge(at, _rid(*key), model.kind.get((cell_row[at], key), "direct"))
    n_dropped = 0
    for k in shown:
        for c in model.up.get(k, ()):
            if c in shown:
                if (k, c) in dropped:
                    n_dropped += 1
                    continue
                edge(_rid(*k), _rid(*c), model.kind[(k, c)])
    for (a, b), sign in part_edges.items():
        if a in shown and b in shown:
            edge(_rid(*a), _rid(*b), "part", sign)

    # -- what isn't shown: counted per node, collapsed per sheet
    hidden_rows = closure - set(shown)
    per_sheet = Counter(s for s, _ in hidden_rows)
    typed_hidden = Counter(s for s, r in hidden_rows if _typed(model.meta.get((s, r), {"n_const": 0, "n_formula": 1})))
    group_ids = {}
    for sheet, n in sorted(per_sheet.items(), key=lambda kv: (-kv[1], kv[0])):
        gid = f"group:{sheet}"
        group_ids[sheet] = gid
        all_typed = typed_hidden[sheet] == n
        node(id=gid, kind="group", sheet=sheet, label=f"{n} more row{'s' if n != 1 else ''} on {sheet}", count=n,
             input=all_typed, **{"class": "assumption" if typed_hidden[sheet] * 2 > n else "calculation",
                                                  "subclass": "collapsed", "why": "rows upstream of the rows shown, left out "
                                                  + ("to keep the page readable (mostly typed inputs)" if typed_hidden[sheet] * 2 > n
                                                     else "to keep the page readable")})
    by_id = {n["id"]: n for n in nodes}
    for k in shown:
        hid = [c for c in model.up.get(k, ()) if c not in shown and not model.is_check(c)]
        by_id[_rid(*k)]["hidden_upstream"] = len(hid)
        for c in hid:
            edge(_rid(*k), group_ids[c[0]], "collapsed")
    for k in sorted(hidden_rows):  # a group is reached through the groups above it too (sorted: the same layers every run)
        for c in model.up.get(k, ()):
            if c in hidden_rows and c[0] != k[0]:
                edge(group_ids[k[0]], group_ids[c[0]], "collapsed")

    # -- the layers and the totals
    el = list(edges.values())
    layer = _layers(start_id, [n["id"] for n in nodes], el)
    for n in nodes:
        n["layer"] = layer[n["id"]]
    count = Counter(n["class"] for n in nodes if n["kind"] != "group")
    sheet_shown = Counter(s for s, _ in shown)
    upstream = Counter(s for s, _ in closure)
    start_node = by_id[start_id]
    out_cores = []
    for c in cores:
        m = c.get("method") or {}
        fr = model.row_of(c.get("factor_row")) if c.get("factor_row") and "!" in c["factor_row"] else None
        out_cores.append({
            "cell": c["cell"], "what": c["what"], "cashflow": c["cashflow"], "cashflow_label": c.get("cashflow_label"),
            "factor_row": _rid(*fr) if fr else None, "kind": c.get("kind"),
            "method": {"rate": m.get("rate"), "rate_value": c.get("rate_value", c.get("rate")),
                       "valuation_date": m.get("valuation_date"), "valuation_date_value": c.get("valuation_date_value"),
                       "timing": m.get("timing"), "day_count": m.get("day_count"), "terminal_date": m.get("terminal_date")},
            "periods": c.get("periods"), "first_period": c.get("first_period"), "last_period": c.get("last_period"),
            "pv": c.get("pv"), "undiscounted": c.get("undiscounted")})
    return _clean({
        "version": __version__,
        "workbook": workbook or _workbook(db_path),
        "start": {"id": start_id, "row_id": start_node["row_id"], "label": start_node["label"], "value": start_node["value"]},
        "anchors": anc,
        "classes": [{"key": k, "name": n} for k, n in CLASSES],
        "nodes": nodes,
        "edges": el,
        "cores": out_cores,
        "stats": {"nodes": len(nodes), "edges": len(el), "rows_upstream_total": len(closure), "rows_shown": len(shown),
                  "cells_shown": len(cells), "inactive_hidden": sum(by_id[_rid(*k)]["inactive_hidden"] for k in shown),
                  "checks_excluded": len(checks), "edges_dropped": n_dropped, "groups": len(group_ids), "max_depth": depth, "max_rows": max_rows,
                  "build_secs": round(time.time() - t0, 2)},
        "by_class": [{"class": k, "n": count[k]} for k, _ in CLASSES if count[k]],
        "by_sheet": sorted(({"sheet": s, "upstream": upstream[s], "shown": sheet_shown[s]}
                            for s in set(upstream) | set(sheet_shown)), key=lambda d: (-d["upstream"], d["sheet"]))})


# ---- command line -------------------------------------------------------------------------------------------

def main(argv=None) -> None:
    ap = argparse.ArgumentParser(description="The dependency graph behind a DCF value cell.")
    ap.add_argument("db", help="out/<dir>/model.db")
    ap.add_argument("cell", nargs="?", help="the value cell, Sheet!C5 (default: the top DCF anchor)")
    ap.add_argument("--depth", type=int, default=6)
    ap.add_argument("--max-rows", type=int, default=300)
    ap.add_argument("--json", help="also write the graph as JSON here")
    ap.add_argument("--html", help="where to write the page (default: depgraph.html beside model.db)")
    ap.add_argument("--version", action="version", version=version_line())
    a = ap.parse_args(argv)
    if a.cell is None:
        found = anchors(a.db)
        for x in found:
            print(f"  {x['cell']:<18} {x['label'][:50]:<50} {x['value']!s:<14} {'ok' if x['ok'] else 'not usable'}"
                  f"{'' if x['matches'] else ' (does not reproduce)'}")
        top = next((x for x in found if x["ok"]), None)
        if not top:
            sys.exit("no DCF anchor found; give a cell")
        a.cell = top["cell"]
        print(f"from the top anchor {a.cell}")
    g = build(a.db, a.cell, a.depth, a.max_rows)
    s = g["stats"]
    print(f"{s['nodes']} nodes ({s['rows_shown']} rows of {s['rows_upstream_total']} upstream, {s['cells_shown']} cells, "
          f"{s['groups']} groups), {s['edges']} edges, {s['checks_excluded']} check rows left out, {s['build_secs']}s")
    print("by class: " + ", ".join(f"{c['class']} {c['n']}" for c in g["by_class"]))
    if a.json:
        Path(a.json).write_text(json.dumps(g, indent=1), encoding="utf-8")
        print("wrote", a.json)
    try:
        from . import depgraph_html
    except ImportError:
        print("(modelatlas/depgraph_html.py isn't there yet, so no page; use --json)")
        return
    path = Path(a.html) if a.html else Path(a.db).parent / "depgraph.html"
    path.write_text(depgraph_html.render(g), encoding="utf-8")
    print("wrote", path)


def cli() -> None:
    main()


if __name__ == "__main__":
    cli()
