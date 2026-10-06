"""Row-level dependency edges, with OFFSET resolved and lookups narrowed to what the current scenario uses.

Every formula's references become edges from its row to the rows it reads. Some functions get more care,
using the values Excel saved in the workbook:
- OFFSET(ref, rows, cols, [height], [width]): when its arguments evaluate (numbers, cells, named values,
  simple arithmetic), the edge goes to the range it actually points at (kind "offset"), not just `ref`.
- SUMIFS / SUMIF / COUNTIFS / AVERAGEIFS over a block of rows, INDEX(block, MATCH(...)) and
  CHOOSE(k, ...): the rows or arguments the current values select are "active"; the other candidates are
  "inactive" (read to decide, but not used in this scenario).
Anything that can't be evaluated keeps plain "direct" edges, so the edges stay a safe superset.

A (src, dst) pair gets the strongest kind any of its references gives it: direct/offset > active > inactive.
    uv run python -m modelatlas.edges out/<dir>/model.db      # rebuild edges in an existing model.db
"""
import re
import sqlite3
import sys
from collections import defaultdict

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token
from openpyxl.utils import column_index_from_string

MAX_RANGE_ROWS = 300
LOOKUPS = {"SUMIFS", "SUMIF", "COUNTIFS", "COUNTIF", "AVERAGEIFS", "AVERAGEIF"}
_REF = re.compile(r"^(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)!)?(?P<a>\$?[A-Z]{0,3}\$?\d*)(?::(?P<b>\$?[A-Z]{0,3}\$?\d*))?$")
_QUOTED = re.compile(r"'(?:[^']|'')*'")  # a quoted sheet name; '' inside is an escaped apostrophe
_CELL = re.compile(r"^\$?([A-Z]{1,3})?\$?(\d+)?$")
_SHAPE = re.compile(r"\$?[A-Z]{1,3}(?=\$?\d)")  # column letters inside refs, for grouping formulas along a row
RANK = {"direct": 3, "offset": 3, "active": 2, "inactive": 1}


def _unspace(text: str) -> str:
    """Drop spaces outside single quotes, so a sheet name like 'Cash flows' stays intact."""
    out, pos = [], 0
    for m in _QUOTED.finditer(text):
        out.append(text[pos:m.start()].replace(" ", ""))
        out.append(m.group())
        pos = m.end()
    out.append(text[pos:].replace(" ", ""))
    return "".join(out)


class Rng:
    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2

    def rows(self):
        return range(self.r1, min(self.r2, self.r1 + MAX_RANGE_ROWS) + 1)


class Node:  # a function call: name + list of arguments (each a list of tokens / Nodes)
    __slots__ = ("name", "args")

    def __init__(self, name):
        self.name, self.args = name, [[]]


def parse(formula: str):
    """Formula -> list of items (Token or Node), with function arguments split."""
    root, stack = [], []
    cur = root
    try:
        items = Tokenizer(formula).items
    except Exception:
        return None
    for t in items:
        if t.type == Token.FUNC and t.subtype == Token.OPEN:
            n = Node(t.value[:-1].upper().removeprefix("_XLFN."))
            cur.append(n)
            stack.append((cur, n))
            cur = n.args[-1]
        elif t.type == Token.FUNC and t.subtype == Token.CLOSE:
            if not stack:
                return None
            parent, _ = stack.pop()
            cur = parent
        elif t.type == Token.SEP and t.subtype == Token.ARG and stack:
            n = stack[-1][1]
            n.args.append([])
            cur = n.args[-1]
        elif t.type != Token.WSPACE:
            cur.append(t)
    return root


class Model:
    """Saved values and names from model.db, for evaluating arguments."""

    def __init__(self, db: sqlite3.Connection):
        self.vals = {(s, r, c): v for s, r, c, v in db.execute("SELECT sheet, row, col, value FROM cells")}
        self.maxcol = dict(db.execute("SELECT sheet, MAX(col) FROM cells GROUP BY sheet"))
        self.names = {n.lower(): ref for n, ref in db.execute("SELECT name, ref FROM names")}
        self.line_items = {(s, r) for s, r in db.execute("SELECT sheet, row FROM rows")}

    def ref(self, text: str, sheet: str, depth: int = 0) -> Rng | None:
        """A1 reference or defined name -> Rng (None for external links / unparseable)."""
        if "[" in text or depth > 3:
            return None
        if text.lower() in self.names and "!" not in text:
            return self.ref(self.names[text.lower()].lstrip("="), sheet, depth + 1)
        m = _REF.match(_unspace(text))
        if not m:
            return None
        sh = m["sheet"] or sheet
        if sh.startswith("'"):
            sh = sh[1:-1].replace("''", "'")
        a, b = _CELL.match(m["a"] or ""), _CELL.match(m["b"] or m["a"] or "")
        if not a or not b:
            return None
        c1, r1 = a.group(1), a.group(2)
        c2, r2 = b.group(1), b.group(2)
        if not r1 and not c1:
            return None
        mc = self.maxcol.get(sh, 16384)
        c1 = column_index_from_string(c1) if c1 else 1
        c2 = column_index_from_string(c2) if c2 else (mc if not b.group(1) else c1)
        r1 = int(r1) if r1 else 1
        r2 = int(r2) if r2 else 1_048_576
        return Rng(sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    def value(self, rng: Rng):
        return self.vals.get((rng.sheet, rng.r1, rng.c1))

    def evaluate(self, items, sheet):
        """Tiny evaluator for argument expressions: numbers, strings, cells, names, & + - * /, and the
        functions INDEX / MATCH. Returns None if it can't."""
        toks = []
        for it in items:
            if isinstance(it, Node):
                v = self.call(it, sheet)
                if v is None:
                    return None
                toks.append(("v", v))
            elif it.type == Token.OPERAND:
                if it.subtype == Token.NUMBER:
                    toks.append(("v", float(it.value)))
                elif it.subtype == Token.TEXT:
                    toks.append(("v", it.value[1:-1].replace('""', '"')))
                elif it.subtype == Token.RANGE:
                    r = self.ref(it.value, sheet)
                    if r is None or (r.r1, r.c1) != (r.r2, r.c2):
                        return None
                    v = self.value(r)
                    toks.append(("v", 0.0 if v is None else v))
                else:
                    return None
            elif it.type in (Token.OP_IN, Token.OP_PRE) and it.value in "+-*/&":
                toks.append(("op", it.value))
            else:
                return None
        if not toks:
            return None
        try:
            if toks[0] == ("op", "-"):
                toks = [("v", -float(toks[1][1]))] + toks[2:]
            elif toks[0] == ("op", "+"):
                toks = toks[1:]
            acc = toks[0][1]
            for i in range(1, len(toks) - 1, 2):
                op, v = toks[i][1], toks[i + 1][1]
                if op == "&":
                    acc = f"{_text(acc)}{_text(v)}"
                elif op == "+":
                    acc = float(acc) + float(v)
                elif op == "-":
                    acc = float(acc) - float(v)
                elif op == "*":
                    acc = float(acc) * float(v)
                elif op == "/":
                    acc = float(acc) / float(v)
            return acc
        except (TypeError, ValueError, ZeroDivisionError, IndexError):
            return None

    def call(self, n: Node, sheet):
        if n.name == "MATCH" and len(n.args) >= 2:
            key = self.evaluate(n.args[0], sheet)
            rng = self.single_range(n.args[1], sheet)
            mt = self.evaluate(n.args[2], sheet) if len(n.args) > 2 and n.args[2] else 1
            if key is None or rng is None or mt != 0:
                return None
            cells = _cells(rng)
            for i, rc in enumerate(cells):
                if _equal(self.vals.get((rng.sheet, *rc)), key):
                    return float(i + 1)
            return None
        if n.name == "INDEX" and len(n.args) >= 2:
            rng = self.single_range(n.args[0], sheet)
            k = self.evaluate(n.args[1], sheet) if n.args[1] else None
            if rng is None or k is None:
                return None
            cells = _cells(rng)
            k = int(k)
            return self.vals.get((rng.sheet, *cells[k - 1])) if 1 <= k <= len(cells) else None
        return None

    def single_range(self, items, sheet) -> Rng | None:
        if len(items) == 1 and not isinstance(items[0], Node) and items[0].type == Token.OPERAND \
                and items[0].subtype == Token.RANGE:
            return self.ref(items[0].value, sheet)
        return None


def _cells(rng: Rng):
    if rng.r1 == rng.r2:
        return [(rng.r1, c) for c in range(rng.c1, rng.c2 + 1)]
    return [(r, rng.c1) for r in range(rng.r1, rng.r2 + 1)]


def _num(v):
    """Excel coerces numeric text ("1") to a number for CHOOSE / INDEX positions."""
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            return None
    return v


def _text(v):
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def _equal(a, b) -> bool:
    if isinstance(a, str) or isinstance(b, str):
        return str(a if a is not None else "").strip().lower() == str(b if b is not None else "").strip().lower()
    try:
        return a is not None and abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return False


def _criterion(value, crit) -> bool:
    """Excel criteria: '>=5', '<>x', 'text' (case-insensitive, * and ? wildcards), or a plain value."""
    if isinstance(crit, str):
        m = re.match(r"^(>=|<=|<>|>|<|=)?(.*)$", crit, re.S)
        op, rhs = m[1] or "=", m[2]
        try:
            num = float(rhs)
        except ValueError:
            num = None
        if num is not None and isinstance(value, (int, float)) and not isinstance(value, bool):
            return {"=": value == num, "<>": value != num, ">": value > num, "<": value < num,
                    ">=": value >= num, "<=": value <= num}[op]
        if op in ("=", "<>"):
            pat = "^" + re.escape(rhs.lower()).replace(r"\*", ".*").replace(r"\?", ".") + "$"
            hit = value is not None and re.match(pat, str(value).lower()) is not None
            return hit if op == "=" else not hit
        return False
    return _equal(value, crit)


def formula_edges(model: Model, sheet: str, row: int, formula: str) -> dict[tuple[str, int], str] | None:
    """(dst_sheet, dst_row) -> kind for one formula cell, or None if it can't be parsed."""
    tree = parse(formula)
    if tree is None:
        return None
    out: dict[tuple[str, int], str] = {}

    def add(rng: Rng | None, kind: str, rows=None):
        if rng is None:
            return
        for r in (rows if rows is not None else rng.rows()):
            key = (rng.sheet, r)
            if RANK[kind] > RANK.get(out.get(key), 0):
                out[key] = kind

    def plain(items):
        for it in items:
            if isinstance(it, Node):
                node(it)
            elif it.type == Token.OPERAND and it.subtype == Token.RANGE:
                add(model.ref(it.value, sheet), "direct")

    def node(n: Node):
        if n.name == "OFFSET" and len(n.args) >= 2:
            base = model.single_range(n.args[0], sheet)
            vals = [model.evaluate(a, sheet) if a else 0.0 for a in n.args[1:3]]
            if base is not None and all(isinstance(v, (int, float)) for v in vals):
                h = model.evaluate(n.args[3], sheet) if len(n.args) > 3 and n.args[3] else base.r2 - base.r1 + 1
                w = model.evaluate(n.args[4], sheet) if len(n.args) > 4 and n.args[4] else base.c2 - base.c1 + 1
                if isinstance(h, (int, float)) and isinstance(w, (int, float)) and h >= 1 and w >= 1:
                    r1, c1 = base.r1 + int(vals[0]), base.c1 + int(vals[1])
                    add(Rng(base.sheet, r1, c1, r1 + int(h) - 1, c1 + int(w) - 1), "offset")
                    for a in n.args[1:]:
                        plain(a)
                    return
            for a in n.args:
                plain(a)
            return
        if n.name in LOOKUPS and _lookup(n):
            return
        if n.name == "INDEX" and _index(n):
            return
        if n.name == "CHOOSE" and _choose(n):
            return
        for a in n.args:
            plain(a)

    def _lookup(n: Node) -> bool:
        a = n.args
        if n.name in ("SUMIF", "COUNTIF", "AVERAGEIF"):
            if len(a) < 2:
                return False
            pairs, target = [(a[0], a[1])], (a[2] if len(a) > 2 else a[0])
        elif n.name == "COUNTIFS":
            pairs, target = list(zip(a[0::2], a[1::2])), None
        else:
            pairs, target = list(zip(a[1::2], a[2::2])), a[0]
        rngs = [model.single_range(p[0], sheet) for p in pairs]
        crits = [model.evaluate(p[1], sheet) for p in pairs]
        tgt = model.single_range(target, sheet) if target is not None else rngs[0]
        if any(r is None for r in rngs) or tgt is None or any(c is None for c in crits):
            return False
        if tgt.r1 == tgt.r2 or any(r.r2 - r.r1 != tgt.r2 - tgt.r1 or r.c1 != r.c2 for r in rngs):
            return False  # only a block of rows picked by a one-column criterion narrows rows
        hits = [k for k in range(tgt.r2 - tgt.r1 + 1)
                if all(_criterion(model.vals.get((r.sheet, r.r1 + k, r.c1)), c) for r, c in zip(rngs, crits))]
        for r in [tgt, *rngs]:
            add(r, "inactive")
            add(r, "active", rows=[r.r1 + k for k in hits])
        for p in pairs:
            plain(p[1])
        return True

    def _index(n: Node) -> bool:
        rng = model.single_range(n.args[0], sheet)
        if rng is None or rng.r1 == rng.r2 or rng.c1 != rng.c2 or len(n.args) < 2 or not n.args[1]:
            return False
        k = _num(model.evaluate(n.args[1], sheet))
        if not isinstance(k, (int, float)) or not 1 <= int(k) <= rng.r2 - rng.r1 + 1:
            return False
        add(rng, "inactive")
        add(rng, "active", rows=[rng.r1 + int(k) - 1])
        for a in n.args[1:]:
            _match_args(a)
        return True

    def _match_args(items):
        """A MATCH inside INDEX: its lookup range is read in full to find the key; the key cell is direct."""
        for it in items:
            if isinstance(it, Node) and it.name == "MATCH" and len(it.args) >= 2:
                plain(it.args[0])
                rng = model.single_range(it.args[1], sheet)
                pos = model.call(it, sheet)
                if rng is not None and pos:
                    add(rng, "inactive")
                    add(rng, "active", rows=[_cells(rng)[int(pos) - 1][0]])
                else:
                    plain(it.args[1])
            elif isinstance(it, Node):
                node(it)
            elif it.type == Token.OPERAND and it.subtype == Token.RANGE:
                add(model.ref(it.value, sheet), "direct")

    def _choose(n: Node) -> bool:
        k = _num(model.evaluate(n.args[0], sheet))
        if not isinstance(k, (int, float)) or not 1 <= int(k) <= len(n.args) - 1:
            return False
        for items in n.args[0:1]:  # the index expression: plain refs direct, lookups narrowed
            _match_args(items) if any(isinstance(i, Node) for i in items) else plain(items)
        for i, items in enumerate(n.args[1:], start=1):
            if i == int(k):
                plain(items)
            else:
                before = dict(out)
                plain(items)
                for key, kind in list(out.items()):  # an unused branch can only add candidates
                    if kind != before.get(key):
                        prev = before.get(key)
                        out[key] = prev if prev and RANK[prev] > RANK["inactive"] else "inactive"
        return True

    plain(tree)
    return out


def build(db: sqlite3.Connection) -> dict:
    """Rebuild the edges table (src_sheet, src_row, dst_sheet, dst_row, kind). Formulas along a row that
    differ only in column letters are analysed once (first and last cell of each such group)."""
    import json
    model = Model(db)
    # Formulas in the label area (the label column and left of it) and the units column build names and
    # numbering, not values; like build_map, leave them out.
    skip = {}
    for sheet, lay in db.execute("SELECT sheet, layout FROM sheets"):
        lay = json.loads(lay or "{}")
        skip[sheet] = (lay.get("label_col") or 0, lay.get("units_col"))
    groups: dict[tuple, list[tuple[int, str]]] = defaultdict(list)
    for sheet, row, col, f, v in db.execute(
            "SELECT sheet, row, col, formula, value FROM cells WHERE formula IS NOT NULL"):
        label_col, units_col = skip.get(sheet, (0, None))
        numeric = isinstance(v, (int, float)) and not isinstance(v, bool)
        if (col < label_col or col == label_col or col == units_col) and not (numeric and col >= label_col):
            continue  # a label / units formula (numbers in the label or units column are data, as in build_map)
        if (sheet, row) in model.line_items:
            groups[(sheet, row, _SHAPE.sub("#", f))].append((col, f))
    edges: dict[tuple, str] = {}
    unparsed = 0
    for (sheet, row, _), cells in groups.items():
        cells.sort()
        for col, f in {cells[0], cells[-1]}:
            got = formula_edges(model, sheet, row, f)
            if got is None:
                unparsed += 1
                continue
            for (ds, dr), kind in got.items():
                if (ds, dr) == (sheet, row) or (ds, dr) not in model.line_items:
                    continue
                key = (sheet, row, ds, dr)
                if RANK[kind] > RANK.get(edges.get(key), 0):
                    edges[key] = kind
    cols = [r[1] for r in db.execute("PRAGMA table_info(edges)")]
    if "kind" not in cols:
        db.execute("DROP TABLE IF EXISTS edges")
        db.execute("CREATE TABLE edges(src_sheet TEXT, src_row INT, dst_sheet TEXT, dst_row INT, kind TEXT)")
    else:
        db.execute("DELETE FROM edges")
    db.executemany("INSERT INTO edges VALUES (?,?,?,?,?)", [(*k, v) for k, v in edges.items()])
    db.execute("CREATE INDEX IF NOT EXISTS ix_e1 ON edges(src_sheet,src_row)")
    db.execute("CREATE INDEX IF NOT EXISTS ix_e2 ON edges(dst_sheet,dst_row)")
    db.commit()
    kinds = defaultdict(int)
    for v in edges.values():
        kinds[v] += 1
    return {"edges": len(edges), "groups": len(groups), "unparsed": unparsed, **kinds}


if __name__ == "__main__":
    import time
    t = time.time()
    with sqlite3.connect(sys.argv[1]) as con:
        print(build(con), f"{time.time() - t:.1f}s")
