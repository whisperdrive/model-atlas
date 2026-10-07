"""Statement detection by identities (modelatlas/statements.py) on the three-way fixture and a minimal workbook.

  base        every block bound (P&L, balance sheet, cash flow, debt, fixed assets, sources and uses, capex, revenue,
              equity, distributions); every identity holds in the saved values; the model's own checks, mined from the
              workbook, say holds and name the same rows; lineage holds
  unbalanced  the balance sheet identity fails and so does the model's own check (a finding, not `unbound`); the
              distributions row that never reaches retained earnings is a finding; every other identity holds
  moved       rows inserted, the P&L sheet renamed, 'Turnover' for revenue, no balance check and no Checks sheet: the
              same blocks are bound by structure and by which identities hold
  stale       the saved results were computed at a tariff the Inputs sheet no longer shows: a `stale` finding
  assets      revenue consolidated from two sheets: the segments are bound to the consolidation, with its parts
  minimal     a hand-built workbook: balance sheet triple found by value, a model check overrules nothing it should not,
              a quarterly sheet against an annual one is `unbound` ('periodicity differs'), and a workbook with no
              balance sheet comes out unbound without a crash
  ontology    docs/model_ontology.md is what modelatlas/ontology.py generates

    uv run python tests/check_statements.py
"""
import json
import shutil
import sqlite3
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

import xlsxwriter  # noqa: E402

from modelatlas import build_map  # noqa: E402
import make_threeway_model  # noqa: E402
from modelatlas import ontology  # noqa: E402
from modelatlas import statements  # noqa: E402

BLOCKS = {"income_statement", "balance_sheet", "cash_flow", "debt", "fixed_assets", "sources_uses", "capex", "revenue",
          "equity", "distributions"}


def _detect(xlsx: Path, tmp: Path, name: str, **kw) -> dict:
    out = build_map.main(str(xlsx), str(tmp / f"{name}__0a1b2c3d"))
    res = statements.detect(out["db"], **kw)
    assert json.loads(json.dumps(res)) == res, "the result is not plain JSON"
    return res


def _ids(res) -> dict:
    return {i["key"]: i for i in res["identities"]}


def _block(res, kind) -> dict:
    return next(b for b in res["blocks"] if b["type"] == kind)


def _row(res, kind, role) -> str:
    r = _block(res, kind)["rows"][role]
    return f"{r['sheet']}!r{r['row']}"


def base(res):
    assert {b["type"] for b in res["blocks"] if b["bound"] and b["complete"]} == BLOCKS, \
        [(b["type"], b["unbound"]) for b in res["blocks"] if not b["complete"]]
    st = {k: i["status"] for k, i in _ids(res).items()}
    assert set(st.values()) == {"holds"}, {k: v for k, v in st.items() if v != "holds"}
    assert len(st) == len(ontology.IDENTITIES)
    assert _row(res, "balance_sheet", "assets") == "BalanceSheet!r10"
    assert _row(res, "balance_sheet", "total_le") == "BalanceSheet!r18"
    assert _row(res, "cash_flow", "cf_closing") == "CashFlow!r24" and _row(res, "cash_flow", "cf_net") == "CashFlow!r23"
    assert _row(res, "income_statement", "ebitda") == "PnL!r9" and _row(res, "income_statement", "npat") == "PnL!r15"
    assert _row(res, "debt", "debt_closing") == "Debt!r22" and _row(res, "fixed_assets", "fa_closing") == "Capex!r15"
    assert _row(res, "sources_uses", "sources") == "Debt!r11" and _row(res, "sources_uses", "uses") == "Debt!r15"
    assert _row(res, "revenue", "rev_total") == "Revenue!r11"
    assert _row(res, "distributions", "dcf_distributions") == "DCF!r7"
    assert _row(res, "cash_flow", "cf_distributions") == "CashFlow!r19"
    assert _row(res, "equity", "re_distributions") == "Equity!r10"
    assert all(i["periods_checked"] >= 1 for i in res["identities"])
    assert _ids(res)["bs_balance"]["periods_checked"] == 15 and _ids(res)["bs_balance"]["failing_periods"] == []
    # the model's own checks, mined: every one says holds, the leaves name the rows the identities were bound to
    ck = {(c["sheet"], c["row"]): c for c in res["checks"]}
    assert {c["verdict"] for c in ck.values()} == {"holds"}, ck
    assert {(s, r) for s, r in ck} >= {("BalanceSheet", 19), ("Checks", 7), ("Checks", 8), ("Checks", 9), ("Checks", 10),
                                       ("Checks", 11), ("Checks", 12), ("Debt", 16)}
    assert ck[("Checks", 11)]["kind"] == "aggregate" and ck[("BalanceSheet", 19)]["kind"] == "leaf"
    assert _ids(res)["bs_total_le"]["given_by"] == "BalanceSheet!r19" and _ids(res)["bs_total_le"]["model_says"] == "holds"
    assert _ids(res)["cf_cash_tie"]["given_by"] == "Checks!r8"
    assert _ids(res)["su_balance"]["given_by"] == "Debt!r16" and _ids(res)["debt_tie"]["given_by"] == "Checks!r10"
    assert {l["status"] for l in res["lineage"]} == {"holds"}, res["lineage"]
    assert [f for f in res["findings"] if f["kind"] != "binding_unbound"] == [], res["findings"]
    assert res["timeline"]["PnL"]["total_column"] == 19 and res["timeline"]["CashFlow"]["total_column"] == 19
    assert res["timeline"]["BalanceSheet"]["total_column"] is None and res["timeline"]["PnL"]["periods"] == 15
    assert {s["label"] for s in _block(res, "revenue")["rows"]["rev_segments*"]} == {"Residential revenue", "Non-residential revenue"}
    assert {c["kind"] for c in res["corkscrews"]} == {"cash", "debt", "fixed_assets", "retained"}
    assert all(c["status"] == "holds" and c["tie"] for c in res["corkscrews"])


def unbalanced(res):
    ids = _ids(res)
    st = {k: i["status"] for k, i in ids.items()}
    assert st["bs_balance"] == "fails" and st["bs_total_le"] == "fails", st
    assert ids["bs_total_le"]["model_says"] == "fails" and ids["bs_total_le"]["given_by"] == "BalanceSheet!r19"
    assert ids["bs_balance"]["max_residual"] > 1 and len(ids["bs_balance"]["failing_periods"]) == 15
    assert not [k for k, v in st.items() if v == "unbound"], st
    assert {k for k, v in st.items() if v != "holds"} == {"bs_balance", "bs_total_le"}, st
    kinds = {f["kind"] for f in res["findings"]}
    assert "model_fails_identity" in kinds and "stale" not in kinds
    assert any(f["kind"] == "model_fails_identity" and "total liabilities and equity" in f["text"] for f in res["findings"])
    assert any(f["kind"] == "pattern_break" and "Equity!r10" in f["text"] for f in res["findings"]), res["findings"]
    ck = {(c["sheet"], c["row"]): c["verdict"] for c in res["checks"]}
    assert ck[("BalanceSheet", 19)] == "fails" and ck[("Checks", 7)] == "fails" and ck[("Checks", 12)] == "fails"
    assert ck.get(("Checks", 11)) == "fails", ck   # a failing 'All checks' (a SUM of checks) is still a check
    assert next(c for c in res["checks"] if (c["sheet"], c["row"]) == ("Checks", 11))["kind"] == "aggregate"
    assert ck[("Checks", 8)] == "holds" and ck[("Debt", 16)] == "holds"
    assert _block(res, "balance_sheet")["bound"] and _row(res, "balance_sheet", "assets") == "BalanceSheet!r10"


def moved(res):
    assert {b["type"] for b in res["blocks"] if b["bound"] and b["complete"]} == BLOCKS
    assert {i["status"] for i in res["identities"]} == {"holds"}, {k: i["status"] for k, i in _ids(res).items()}
    assert _block(res, "income_statement")["sheet"] == "Income statement"
    assert _block(res, "income_statement")["rows"]["revenue"]["label"] == "Turnover"
    assert _block(res, "income_statement")["rows"]["npat"]["label"] == "Net profit after tax"
    assert _block(res, "revenue")["rows"]["rev_total"]["label"] == "Revenue total"
    ids = _ids(res)
    assert ids["bs_balance"]["status"] == "holds" and ids["bs_balance"]["given_by"] is None
    assert "no model check names them" in _block(res, "balance_sheet")["rows"]["assets"]["how"]
    assert not any(c["sheet"] == "Checks" for c in res["checks"])
    assert not any(c["sheet"] == "BalanceSheet" for c in res["checks"])
    assert ids["su_balance"]["given_by"] == "Debt!r16"   # the check that was not removed still arbitrates
    assert {l["status"] for l in res["lineage"]} == {"holds"}


def stale(res):
    assert {i["status"] for i in res["identities"]} == {"holds"}, "a stale cache is consistent with itself"
    st = [f for f in res["findings"] if f["kind"] == "stale"]
    assert st and any("Revenue!D8" in f["text"] and "2.1" in f["text"] and "2.4" in f["text"] for f in st), res["findings"]


def assets(res):
    assert {i["status"] for i in res["identities"]} == {"holds"}, {k: i["status"] for k, i in _ids(res).items()}
    segs = _block(res, "revenue")["rows"]["rev_segments*"]
    assert [s["label"] for s in segs] == ["Residential revenue", "Non-residential revenue"], segs
    for s in segs:
        assert {p["sheet"] for p in s["consolidation_of"]} == {"RevenueNorth", "RevenueSouth"}, s
    assert {c["sheet"] for c in res["consolidations"]} == {"Revenue"} and len(res["consolidations"]) >= 2
    assert _row(res, "revenue", "rev_total") == "Revenue!r10"


# ---- a minimal workbook, built here -----------------------------------------------------------------------------------

def _years(n=6):
    return [date(2026 + i, 6, 30) for i in range(n)]


def minimal(path: Path, check=False, no_bs=False, quarterly_cf=False, flag=False, line_items=False, net_assets=False,
            imbalance=0.0, extra=False, faint_equity=False, plug=False, dates=None):
    """Sheet M: revenue + costs = EBITDA, + tax = NPAT; a cash corkscrew and a retained earnings corkscrew; a balance sheet
    (cash, other assets; debt; share capital, retained earnings) that balances. Labels are plain on purpose.
      flag          the balance check carries a typed enable flag (1) after the last period, as models often do
      line_items    two rows named with check-ish words that are line items: 'Loan balances' (an IF giving numbers) and
                    'Deposit balances basis' (an IF giving text)
      net_assets    a 'Net assets' row = total assets - total liabilities (assets = liabilities + net assets by construction)
      imbalance     other assets overstated by this much: the balance sheet does not balance
      extra         rows for the stale test: 1e9-scale operands whose difference is saved as 0, and -x^2
      faint_equity  total equity is a few cents against assets of millions
      plug          total equity = total assets - total liabilities, so the balance sheet balances whatever the rest says
      dates         the period end dates (default: six financial years ending 30 June); the model has len(dates) periods"""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("M")
    fmt = wb.add_format({"num_format": "dd-mmm-yy"})
    ys = list(dates) if dates else _years(6)
    n = len(ys)
    cols = [chr(ord("D") + i) for i in range(n)]
    ws.write(2, 0, "Period ending")
    for c, d in zip(cols, ys):
        ws.write_datetime(f"{c}3", __import__("datetime").datetime(d.year, d.month, d.day), fmt)
    rev = [100.0 + 10 * i for i in range(n)]
    opx = [-(40.0 + 3 * i) for i in range(n)]
    ebitda = [a + b for a, b in zip(rev, opx)]
    tax = [-0.3 * e for e in ebitda]
    npat = [a + b for a, b in zip(ebitda, tax)]
    def put(row, label, values, formulas=None):
        ws.write(f"A{row}", label)
        ws.write(f"B{row}", "$m")
        for i, c in enumerate(cols):
            f = formulas(i, c) if formulas else None
            if f:
                ws.write_formula(f"{c}{row}", f, None, values[i])
            else:
                ws.write_number(f"{c}{row}", values[i])
    put(5, "Sales", rev)
    put(6, "Costs", opx)
    put(7, "Gross margin", ebitda, lambda i, c: f"={c}5+{c}6")
    put(8, "Income tax", tax, lambda i, c: f"=-0.3*MAX(0,{c}7)")
    put(9, "Profit for the year", npat, lambda i, c: f"={c}7+{c}8")
    close_cash = []
    opening = [100.0]
    for i in range(n):
        close_cash.append(opening[i] + npat[i])
        opening.append(close_cash[i])
    put(12, "Opening cash", opening[:n], lambda i, c: "=100" if i == 0 else f"={cols[i - 1]}14")
    put(13, "Net cash flow", npat, lambda i, c: f"={c}9")
    put(14, "Closing cash", close_cash, lambda i, c: f"={c}12+{c}13")
    re_open = [50.0]
    re_close = []
    for i in range(n):
        re_close.append(re_open[i] + npat[i])
        re_open.append(re_close[i])
    if faint_equity:
        put(17, "Plant", [1e6 + i for i in range(n)])
        put(18, "Receivables", [0.02] * n)
        put(19, "Total assets", [1e6 + i + 0.02 for i in range(n)], lambda i, c: f"=SUM({c}17:{c}18)")
        put(21, "Loans", [1e6 + i for i in range(n)])
        put(22, "Total liabilities", [1e6 + i for i in range(n)], lambda i, c: f"=SUM({c}21:{c}21)")
        put(24, "Share capital", [0.02] * n)
        put(25, "Total equity", [0.02] * n, lambda i, c: f"=SUM({c}24:{c}24)")
        no_bs = True
    if not no_bs:
        put(17, "Cash", close_cash, lambda i, c: f"={c}14")
        put(18, "Other assets", [50.0 + imbalance] * n)
        ta = [c + 50 + imbalance for c in close_cash]
        put(19, "Total assets", ta, lambda i, c: f"=SUM({c}17:{c}18)")
        put(21, "Debt", [80.0] * n)
        put(22, "Total liabilities", [80.0] * n, lambda i, c: f"=SUM({c}21:{c}21)")
        if net_assets:
            put(23, "Net assets", [a - 80.0 for a in ta], lambda i, c: f"={c}19-{c}22")
        put(24, "Share capital", [20.0] * n)
        put(25, "Opening retained earnings", re_open[:n], lambda i, c: "=50" if i == 0 else f"={cols[i - 1]}27")
        put(26, "NPAT", npat, lambda i, c: f"={c}9")
        put(27, "Closing retained earnings", re_close, lambda i, c: f"=SUM({c}25:{c}26)")
        put(28, "Retained earnings", re_close, lambda i, c: f"={c}27")
        if plug:
            put(29, "Total equity", [t - 80.0 for t in ta], lambda i, c: f"={c}19-{c}22")
        else:
            put(29, "Total equity", [20 + r for r in re_close], lambda i, c: f"={c}24+{c}28")
        if check:
            put(31, "Balance check", [0.0] * n, lambda i, c: f"={c}19-({c}22+{c}29)")
            if flag:
                ws.write_number(f"{chr(ord('D') + n)}31", 1)
        if line_items:
            put(33, "Loan balances", [max(c, 0.0) for c in close_cash], lambda i, c: f"=IF({c}14>{c}21,{c}14,0)")
            ws.write("A34", "Deposit balances basis")
            for c in cols:
                ws.write_formula(f"{c}34", f'=IF({c}14>0,"Fixed","Floating")', None, "Fixed")
    if extra:
        big = 1000000000.1 + 0.2
        put(36, "Big A", [1000000000.3] * n)
        put(37, "Big B1", [1000000000.1] * n)
        put(38, "Big B2", [0.2] * n)
        put(39, "Big B", [big] * n, lambda i, c: f"={c}37+{c}38")
        put(40, "Big difference", [0.0] * n, lambda i, c: f"={c}36-{c}39")
        put(41, "Three", [3.0] * n)
        put(42, "Minus three squared", [9.0] * n, lambda i, c: f"=-{c}41^2")
    if quarterly_cf:
        ws2 = wb.add_worksheet("Q")
        qs = [__import__("datetime").datetime(2026 + (2 + 3 * i) // 12, (2 + 3 * i) % 12 + 1, 28) for i in range(8)]
        ws2.write(2, 0, "Period ending")
        for i, d in enumerate(qs):
            ws2.write_datetime(2, 3 + i, d, fmt)
        ws2.write(4, 0, "Net cash flow")
        for i in range(8):
            ws2.write_number(4, 3 + i, 1.0)
    wb.close()


def minimal_checks(tmp):
    p = tmp / "mini.xlsx"
    minimal(p)
    res = _detect(p, tmp, "mini")
    ids = _ids(res)
    st = {k: v["status"] for k, v in ids.items()}
    assert st["bs_balance"] == "holds" and st["bs_total_le"] == "unbound" and st["cf_roll"] == "holds", st
    assert st["pnl_ebitda"] == "holds" and st["pnl_ebit"] == "unbound" and st["re_roll"] == "holds", st
    assert st["cf_cash_tie"] == "holds" and st["re_tie"] == "holds" and st["re_npat"] == "holds", st
    assert _block(res, "balance_sheet")["rows"]["assets"]["how"].startswith("assets = liabilities + equity in all 6 periods")
    assert _row(res, "income_statement", "revenue") == "M!r5"   # 'Sales' by vocabulary, no 'Revenue' label needed
    assert not [f for f in res["findings"] if f["kind"] in ("model_fails_identity", "stale")]
    # a model check naming the pair is the arbiter and the source of the binding
    p = tmp / "mini_check.xlsx"
    minimal(p, check=True)
    res = _detect(p, tmp, "mini_check")
    assert [c["verdict"] for c in res["checks"]] == ["holds"] and res["checks"][0]["kind"] == "leaf"
    # tolerance is configurable: a loose absolute tolerance still holds, a zero one cannot fail a clean model by float noise
    assert _ids(statements.detect(res_db(tmp, "mini_check"), tolerance=(1e-3, 1e-3)))["bs_balance"]["status"] == "holds"
    # no balance sheet: unbound, not a crash, and the P&L still found
    p = tmp / "mini_nobs.xlsx"
    minimal(p, no_bs=True)
    res = _detect(p, tmp, "mini_nobs")
    assert not _block(res, "balance_sheet")["bound"] and _ids(res)["bs_balance"]["status"] == "unbound"
    assert "no sheet has three totals" in _ids(res)["bs_balance"]["why"] and _ids(res)["pnl_ebitda"]["status"] == "holds"
    # a quarterly sheet against an annual one: the cross-sheet identity is unbound with the reason
    p = tmp / "mini_q.xlsx"
    minimal(p, quarterly_cf=True)
    db = build_map.main(str(p), str(tmp / "mini_q__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    r = statements.evaluate(d.b, [(("M", 9), 1, 0), (("Q", 5), -1, 0)], set(), d.tol)
    assert r["status"] == "unbound" and "periodicity differs" in r["why"], r
    # the same sheet in two versions of the model: nothing differs, the result is deterministic
    a, b = statements.detect(db), statements.detect(db)
    a["stats"]["secs"] = b["stats"]["secs"] = 0
    assert a == b


def review_fixes(tmp):
    """Each case once failed: a fix in modelatlas/statements.py, and what it would have reported without it."""
    # a check's verdict is what its formulas compute, not a typed enable flag beside them (was: the model's check fails)
    p = tmp / "mini_flag.xlsx"
    minimal(p, check=True, flag=True)
    res = _detect(p, tmp, "mini_flag")
    assert [c["verdict"] for c in res["checks"]] == ["holds"], res["checks"]
    assert not [f for f in res["findings"] if f["kind"] == "model_fails_identity"], res["findings"]
    # line items named with check-ish words are not checks (was: two spurious 'the model's check ... fails')
    p = tmp / "mini_items.xlsx"
    minimal(p, check=True, line_items=True)
    res = _detect(p, tmp, "mini_items")
    assert [(c["row"], c["verdict"]) for c in res["checks"]] == [(31, "holds")], res["checks"]
    assert not [f for f in res["findings"] if f["kind"] == "model_fails_identity"], res["findings"]
    # net assets = assets - liabilities is no balance sheet evidence: an unbalanced model fails, by label, not 'holds'
    p = tmp / "mini_na.xlsx"
    minimal(p, net_assets=True, imbalance=5.0)
    res = _detect(p, tmp, "mini_na")
    assert _ids(res)["bs_balance"]["status"] == "fails", _ids(res)["bs_balance"]
    assert _row(res, "balance_sheet", "equity") == "M!r29", _block(res, "balance_sheet")["rows"]
    assert "label only" in _block(res, "balance_sheet")["rows"]["equity"]["how"]
    assert abs(_ids(res)["bs_balance"]["max_residual"] - 5.0) < 1e-9
    p = tmp / "mini_na_ok.xlsx"   # the same layout, balanced: found by value, equity is total equity, not net assets
    minimal(p, net_assets=True)
    res = _detect(p, tmp, "mini_na_ok")
    assert _ids(res)["bs_balance"]["status"] == "holds" and _row(res, "balance_sheet", "equity") == "M!r29"
    assert "no model check names them" in _block(res, "balance_sheet")["rows"]["equity"]["how"]
    # a triple whose third row is below the triple's own tolerance is no evidence (was: bound 'by value')
    p = tmp / "mini_faint.xlsx"
    minimal(p, faint_equity=True)
    res = _detect(p, tmp, "mini_faint")
    assert "label only" in _block(res, "balance_sheet")["rows"]["assets"]["how"], _block(res, "balance_sheet")["rows"]
    # stale: float noise of 1e9-scale operands, and Excel's -x^2 = x^2, are not stale (was: two stale findings)
    p = tmp / "mini_extra.xlsx"
    minimal(p, extra=True)
    db = build_map.main(str(p), str(tmp / "mini_extra__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    d.bind = {"x": ("M", 40), "y": ("M", 42)}
    d.stale()
    assert d.stats["stale_cells_checked"] >= 6 and not [f for f in d.findings if f["kind"] == "stale"], d.findings
    # a four-subtotal P&L chain that is not EBITDA/EBIT/PBT/NPAT keeps its labelled roles (was: every role shifted one
    # up, the overheads row bound as depreciation)
    p = tmp / "gm.xlsx"
    gross_margin_pnl(p)
    res = _detect(p, tmp, "gm")
    rows = {r: f"{v['sheet']}!r{v['row']}" for r, v in _block(res, "income_statement")["rows"].items()}
    assert rows.get("ebitda") == "P!r9" and rows.get("ebit") == "P!r11" and rows.get("npat") == "P!r13", rows
    assert rows.get("depreciation") == "P!r10" and "pbt" not in rows, rows
    # the model's own check says holds but the very rows it compares differ by 5: a finding, not 'holds' (was: holds)
    p = tmp / "mini_loose.xlsx"
    minimal(p, check=True, imbalance=5.0)   # the check's saved value is 0: typed over, switched off or too loose
    res = _detect(p, tmp, "mini_loose")
    bs = _ids(res)["bs_balance"]
    assert bs["status"] == "fails" and bs["model_says"] == "holds" and abs(bs["max_residual"] - 5.0) < 1e-9, bs
    assert any(f["kind"] == "model_fails_identity" and "says it holds" in f["text"] for f in res["findings"])
    # a model.db without the edges table: lineage is unbound, not 'fails' (was: every lineage test failed)
    nodb = tmp / "mini_noedges__0a1b2c3d"
    nodb.mkdir()
    shutil.copy(res_db(tmp, "mini"), nodb / "model.db")
    con = sqlite3.connect(nodb / "model.db")
    con.execute("DROP TABLE edges")
    con.commit()
    con.close()
    res = statements.detect(str(nodb / "model.db"))
    assert {l["status"] for l in res["lineage"]} == {"unbound"}, res["lineage"]
    assert _ids(res)["bs_balance"]["status"] == "holds"


def gross_margin_pnl(path: Path):
    """A P&L whose chain is gross margin, EBITDA, EBIT, NPAT (no PBT line): four subtotals, but not the four the ontology
    names, so the labelled ones must not be shifted along to make room for an unlabelled first one."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("P")
    fmt = wb.add_format({"num_format": "dd-mmm-yy"})
    cols = [chr(ord("D") + i) for i in range(6)]
    ws.write(2, 0, "Period ending")
    for c, d in zip(cols, _years(6)):
        ws.write_datetime(f"{c}3", __import__("datetime").datetime(d.year, d.month, d.day), fmt)
    vals = {}

    def put(row, label, fn):
        ws.write(f"A{row}", label)
        for i, c in enumerate(cols):
            v, f = fn(i, c)
            vals[(row, i)] = v
            ws.write_formula(f"{c}{row}", f, None, v) if f else ws.write_number(f"{c}{row}", v)
    put(5, "Sales", lambda i, c: (100.0 + 10 * i, None))
    put(6, "Cost of sales", lambda i, c: (-30.0 - i, None))
    put(7, "Gross margin", lambda i, c: (vals[(5, i)] + vals[(6, i)], f"={c}5+{c}6"))
    put(8, "Overheads", lambda i, c: (-20.0, None))
    put(9, "EBITDA", lambda i, c: (vals[(7, i)] + vals[(8, i)], f"={c}7+{c}8"))
    put(10, "Amortisation", lambda i, c: (-5.0, None))
    put(11, "EBIT", lambda i, c: (vals[(9, i)] + vals[(10, i)], f"={c}9+{c}10"))
    put(12, "Interest and tax", lambda i, c: (-7.0, None))
    put(13, "NPAT", lambda i, c: (vals[(11, i)] + vals[(12, i)], f"={c}11+{c}12"))
    wb.close()


def res_db(tmp, name):
    return str(tmp / f"{name}__0a1b2c3d" / "model.db")


def real_models():
    """Optional, only where the (git-ignored) out/ folders exist: no balance sheet is unbound, not an error."""
    db = ROOT / "out" / "Riverbend_BP25_client_model__814d7fe1" / "model.db"
    if not db.exists():
        return
    res = statements.detect(str(db))
    assert not _block(res, "balance_sheet")["bound"] and _ids(res)["bs_balance"]["status"] == "unbound"
    assert _ids(res)["pnl_ebitda"]["status"] == "holds" and _ids(res)["pnl_ebit"]["status"] == "unbound"
    assert not [f for f in res["findings"] if f["kind"] in ("model_fails_identity", "stale")]
    assert statements.text(res)


def ontology_doc():
    doc = (ROOT / "docs" / "model_ontology.md").read_text(encoding="utf-8")
    assert doc == ontology.markdown(), "docs/model_ontology.md is out of date: uv run python -m modelatlas.ontology --write docs/model_ontology.md"
    for i in ontology.IDENTITIES:
        assert f"`{i.key}`" in doc


# ---- a small builder for the cases below -------------------------------------------------------------------------------

class MW:
    """Annual timeline in D3.., labels in column A; put() writes a row with its cached values and a formula per period."""

    def __init__(self, path, sheets=("M",), n=6):
        import datetime
        self.wb = xlsxwriter.Workbook(str(path))
        self.n = n
        self.cols = [chr(ord("D") + i) for i in range(n)]
        self.ws = {}
        fmt = self.wb.add_format({"num_format": "dd-mmm-yy"})
        for name in sheets:
            ws = self.wb.add_worksheet(name)
            ws.write(2, 0, "Period ending")
            for c, d in zip(self.cols, _years(n)):
                ws.write_datetime(f"{c}3", datetime.datetime(d.year, d.month, d.day), fmt)
            self.ws[name] = ws

    def put(self, sheet, row, label, values, f=None):
        ws = self.ws[sheet]
        ws.write(f"A{row}", label)
        for i, c in enumerate(self.cols):
            fx = f(i, c) if f else None
            if fx:
                ws.write_formula(f"{c}{row}", fx, None, values[i])
            else:
                ws.write_number(f"{c}{row}", values[i])

    def close(self):
        self.wb.close()


def _prev(m, c):
    return m.cols[m.cols.index(c) - 1]


def unlabelled_bs(path, corkscrew=True, check=False):
    """Three balancing totals with no balance sheet words at all ('Total A', 'Total B', 'Total C'). With a cash corkscrew
    whose closing row is a line under Total A they are a balance sheet; without one they merely balance."""
    m = MW(path)
    n = m.n
    net = [10.0 + i for i in range(n)]
    close, opening = [], [100.0]
    for i in range(n):
        close.append(opening[i] + net[i])
        opening.append(close[i])
    if corkscrew:
        m.put("M", 12, "Opening cash", opening[:n], lambda i, c: "=100" if i == 0 else f"={_prev(m, c)}14")
        m.put("M", 13, "Net cash flow", net)
        m.put("M", 14, "Closing cash", close, lambda i, c: f"={c}12+{c}13")
        m.put("M", 17, "Line 1", close, lambda i, c: f"={c}14")
    else:
        m.put("M", 17, "Line 1", close)
    m.put("M", 18, "Line 2", [50.0] * n)
    ta = [c + 50.0 for c in close]
    m.put("M", 19, "Total A", ta, lambda i, c: f"=SUM({c}17:{c}18)")
    m.put("M", 21, "Line 3", [80.0] * n)
    m.put("M", 22, "Total B", [80.0] * n, lambda i, c: f"=SUM({c}21:{c}21)")
    m.put("M", 24, "Line 4", [t - 80.0 for t in ta])
    m.put("M", 25, "Total C", [t - 80.0 for t in ta], lambda i, c: f"=SUM({c}24:{c}24)")
    if check:
        m.put("M", 27, "Balance check", [0.0] * n, lambda i, c: f"={c}19-({c}22+{c}25)")
    m.close()


def enterprise_dcf(path):
    """A cash flow with no distributions row, an equity roll-forward that has one, and an enterprise DCF whose cash-flow
    row ('Free cash flow') equals the operating cash flow row by value."""
    m = MW(path)
    n = m.n
    sales = [100.0 + 10 * i for i in range(n)]
    costs = [-40.0] * n
    ebitda = [a + b for a, b in zip(sales, costs)]
    m.put("M", 5, "Sales", sales)
    m.put("M", 6, "Costs", costs)
    m.put("M", 7, "EBITDA", ebitda, lambda i, c: f"={c}5+{c}6")
    m.put("M", 10, "Cash flow from operations", ebitda, lambda i, c: f"={c}7")
    m.put("M", 11, "Capex", [-10.0] * n)
    m.put("M", 12, "Other investing", [-1.0] * n)
    m.put("M", 13, "Cash flow from investing", [-11.0] * n, lambda i, c: f"=SUM({c}11:{c}12)")
    m.put("M", 14, "Debt drawdown", [5.0] * n)
    m.put("M", 15, "Debt repayment", [-2.0] * n)
    m.put("M", 16, "Cash flow from financing", [3.0] * n, lambda i, c: f"=SUM({c}14:{c}15)")
    net = [e - 11.0 + 3.0 for e in ebitda]
    m.put("M", 18, "Net cash flow", net, lambda i, c: f"={c}10+{c}13+{c}16")
    close, opening = [], [20.0]
    for i in range(n):
        close.append(opening[i] + net[i])
        opening.append(close[i])
    m.put("M", 19, "Opening cash", opening[:n], lambda i, c: "=20" if i == 0 else f"={_prev(m, c)}20")
    m.put("M", 20, "Closing cash", close, lambda i, c: f"={c}19+{c}18")
    df = [0.95 ** (i + 1) for i in range(n)]
    m.put("M", 24, "Discount factor", df)
    m.put("M", 25, "Free cash flow", ebitda, lambda i, c: f"={c}10")
    m.ws["M"].write_formula("D26", f"=SUMPRODUCT(D25:{m.cols[-1]}25,D24:{m.cols[-1]}24)", None,
                            sum(a * b for a, b in zip(ebitda, df)))
    m.ws["M"].write("A26", "Enterprise value")
    re_open, re_close = [0.0], []
    for i in range(n):
        re_close.append(re_open[i] + ebitda[i] - 3.0)
        re_open.append(re_close[i])
    m.put("M", 28, "Opening retained earnings", re_open[:n], lambda i, c: "=0" if i == 0 else f"={_prev(m, c)}31")
    m.put("M", 29, "NPAT", ebitda, lambda i, c: f"={c}7")
    m.put("M", 30, "Distributions", [-3.0] * n)
    m.put("M", 31, "Closing retained earnings", re_close, lambda i, c: f"=SUM({c}28:{c}30)")
    m.close()


def two_balance_sheets(path):
    """A group balance sheet whose cash and senior debt corkscrews tie to its lines, a second debt tranche corkscrew whose
    saved closing balance is off by 1, and a per-asset balance sheet on another sheet."""
    m = MW(path, sheets=("Group", "Asset1"))
    n = m.n
    net = [10.0 + i for i in range(n)]
    cash, co = [], [100.0]
    for i in range(n):
        cash.append(co[i] + net[i])
        co.append(cash[i])
    m.put("Group", 12, "Opening cash", co[:n], lambda i, c: "=100" if i == 0 else f"={_prev(m, c)}14")
    m.put("Group", 13, "Net cash flow", net)
    m.put("Group", 14, "Closing cash", cash, lambda i, c: f"={c}12+{c}13")
    da, dao = [], [100.0]
    for i in range(n):
        da.append(dao[i] - 5.0)
        dao.append(da[i])
    m.put("Group", 16, "Opening debt A", dao[:n], lambda i, c: "=100" if i == 0 else f"={_prev(m, c)}19")
    m.put("Group", 17, "Drawdown A", [0.0] * n)
    m.put("Group", 18, "Repayment A", [-5.0] * n)
    m.put("Group", 19, "Closing debt A", da, lambda i, c: f"=SUM({c}16:{c}18)")
    db, dbo = [], [50.0 + 1.0 - 1.0]
    for i in range(n):
        db.append(dbo[i] - 2.0 + 1.0)      # the saved closing balance is 1 above what its formula gives
        dbo.append(db[i])
    m.put("Group", 21, "Opening debt B", dbo[:n], lambda i, c: "=50" if i == 0 else f"={_prev(m, c)}23")
    m.put("Group", 22, "Repayment B", [-2.0] * n)
    m.put("Group", 23, "Closing debt B", db, lambda i, c: f"=SUM({c}21:{c}22)")
    m.put("Group", 26, "Cash", cash, lambda i, c: f"={c}14")
    m.put("Group", 27, "Other assets", [300.0] * n)
    ta = [x + 300.0 for x in cash]
    m.put("Group", 28, "Total assets", ta, lambda i, c: f"=SUM({c}26:{c}27)")
    m.put("Group", 30, "Debt", da, lambda i, c: f"={c}19")
    m.put("Group", 31, "Total liabilities", da, lambda i, c: f"=SUM({c}30:{c}30)")
    eq = [t - d for t, d in zip(ta, da)]
    m.put("Group", 33, "Share capital", eq)
    m.put("Group", 34, "Total equity", eq, lambda i, c: f"=SUM({c}33:{c}33)")
    m.put("Asset1", 5, "Fixed assets", [200.0 + i for i in range(n)])
    m.put("Asset1", 6, "Total assets", [200.0 + i for i in range(n)], lambda i, c: f"=SUM({c}5:{c}5)")
    m.put("Asset1", 8, "Bank loan", [120.0] * n)
    m.put("Asset1", 9, "Total liabilities", [120.0] * n, lambda i, c: f"=SUM({c}8:{c}8)")
    m.put("Asset1", 11, "Share capital", [80.0 + i for i in range(n)])
    m.put("Asset1", 12, "Total equity", [80.0 + i for i in range(n)], lambda i, c: f"=SUM({c}11:{c}11)")
    m.close()


def circular_and_seeded(path):
    """Stale-test guards and the pattern-break exemptions, on one sheet with nine periods."""
    m = MW(path, n=9)
    n = m.n
    # a circular reference: interest on the average balance, the closing balance adds interest. Saved as Excel's
    # iteration left it: close to, but not exactly, what the saved inputs give
    m.put("M", 5, "Opening balance", [100.0] * n)
    m.put("M", 6, "Interest", [-5.5] * n, lambda i, c: f"=-0.05*({c}5+{c}7)/2")
    m.put("M", 7, "Closing balance", [94.5] * n, lambda i, c: f"={c}5+{c}6")
    # a difference below 1e-3 of the largest operand is not a stale cache
    m.put("M", 10, "Base", [1000.0] * n)
    m.put("M", 11, "Scaled", [1000.4] * n, lambda i, c: f"={c}10*1")
    # seeded rows: typed first period then formulas (no break); the first populated cell may not be the first column
    m.put("M", 14, "Seeded", [50.0] + [51.0] * (n - 1), lambda i, c: None if i == 0 else f"={_prev(m, c)}14+1")
    ws = m.ws["M"]
    ws.write("A15", "Seeded late")
    for i, c in enumerate(m.cols):
        if i == 1:
            ws.write_number(f"{c}15", 7.0)
        elif i > 1:
            ws.write_formula(f"{c}15", f"={_prev(m, c)}15+1", None, 7.0 + i - 1)
    # a typed number in the middle of a row of formulas: a break
    m.put("M", 17, "Hardcoded", [10.0 + i for i in range(n)], lambda i, c: None if i in (0, 5) else f"={_prev(m, c)}17+1")
    m.close()


def single_block(path):
    """Sources and uses in one column (an amount) with a second numeric column beside it (a share of the total)."""
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("SU")
    rows = [("Equity", 300.0, 0.6), ("Debt", 200.0, 0.4)]
    ws.write("A3", "Sources")
    for r, (lab, amt, sh) in enumerate(rows, start=4):
        ws.write(f"A{r}", lab)
        ws.write_number(f"C{r}", amt)
        ws.write_number(f"D{r}", sh)
    ws.write("A6", "Total sources")
    ws.write_formula("C6", "=SUM(C4:C5)", None, 500.0)
    ws.write_formula("D6", "=SUM(D4:D5)", None, 1.0)
    ws.write("A8", "Uses")
    for r, (lab, amt, sh) in enumerate([("Capex", 450.0, 0.85), ("Costs", 50.0, 0.1)], start=9):
        ws.write(f"A{r}", lab)
        ws.write_number(f"C{r}", amt)
        ws.write_number(f"D{r}", sh)
    ws.write("A11", "Total uses")
    ws.write_formula("C11", "=SUM(C9:C10)", None, 500.0)
    ws.write_formula("D11", "=SUM(D9:D10)", None, 0.95)
    wb.close()

def kinds_and_stats(results):
    """Item 1: every identity carries an evidential kind and each role how it was bound; text() prints the two apart."""
    for name, res in results.items():
        assert all(i["kind"] in ("test", "structural") and i["kind_basis"] for i in res["identities"]), name
        n = sum(sum(res["stats"][k].values()) for k in ("tests", "structural"))
        assert n == len(ontology.IDENTITIES), (name, res["stats"])
        t = statements.text(res)
        assert t.index("Tests of the model") < t.index("Structure confirmed"), name
    ids = _ids(results["base"])
    assert ids["bs_balance"]["kind"] == "test" and ids["bs_balance"]["roles"]["assets"]["bound_by"] == "check"
    assert ids["su_balance"]["kind"] == "test" and ids["bs_total_le"]["kind"] == "test"
    # the P&L chain is the formulas themselves; a corkscrew's roll is its closing row's own formula
    assert {ids[k]["kind"] for k in ("pnl_ebitda", "pnl_npat", "cf_roll", "debt_roll", "re_roll")} == {"structural"}
    assert ids["cf_cash_tie"]["kind"] == "structural" and ids["cf_cash_tie"]["roles"]["bs_cash"]["bound_by"] in ("link", "value_search")
    assert results["base"]["stats"]["tests"]["holds"] == 3
    # with no model check the balance sheet is found by value: it holds because it was chosen to, and says so
    mv = _ids(results["moved"])
    assert mv["bs_balance"]["kind"] == "structural" and mv["bs_balance"]["roles"]["assets"]["bound_by"] == "value_search"
    # a failing balance sheet named by a check is a failing TEST
    ub = _ids(results["unbalanced"])
    assert ub["bs_balance"]["status"] == "fails" and ub["bs_balance"]["kind"] == "test"
    assert results["unbalanced"]["stats"]["tests"]["fails"] == 2 and results["base"]["stats"]["structural"]["fails"] == 0
    # every block lists its alternates (none on the clean fixtures)
    assert all(b["alternates"] == [] for b in results["base"]["blocks"])


def bs_anchor(tmp):
    """Item 2: a balance sheet needs corroboration, not words."""
    p = tmp / "unl.xlsx"
    unlabelled_bs(p)
    res = _detect(p, tmp, "unl")
    bs = _block(res, "balance_sheet")
    assert bs["bound"] and _row(res, "balance_sheet", "assets") == "M!r19", bs
    assert _row(res, "balance_sheet", "liabilities") == "M!r22" and _row(res, "balance_sheet", "equity") == "M!r25"
    a = bs["rows"]["assets"]
    assert a["bound_by"] == "value_search" and "corkscrew" in a["how"] and "M!r14" in a["how"] and "M!r17" in a["how"], a
    ids = _ids(res)
    assert ids["bs_balance"]["status"] == "holds" and ids["bs_balance"]["kind"] == "structural"
    # a coincidental triple: balances, but nothing says it is a balance sheet
    p = tmp / "unl_none.xlsx"
    unlabelled_bs(p, corkscrew=False)
    res = _detect(p, tmp, "unl_none")
    assert not _block(res, "balance_sheet")["bound"], _block(res, "balance_sheet")["rows"]
    assert _ids(res)["bs_balance"]["status"] == "unbound"
    # a model check that compares exactly the three rows is corroboration too (and the model's verdict)
    p = tmp / "unl_chk.xlsx"
    unlabelled_bs(p, corkscrew=False, check=True)
    res = _detect(p, tmp, "unl_chk")
    ids = _ids(res)
    assert ids["bs_balance"]["status"] == "holds" and ids["bs_balance"]["kind"] == "test", ids["bs_balance"]
    assert ids["bs_balance"]["given_by"] == "M!r27" and ids["bs_balance"]["roles"]["assets"]["bound_by"] == "check"
    # three closing balances that a 'check summary' adds up are a roll-forward's totals, not a balance sheet
    p = tmp / "unl_closings.xlsx"
    closings_check(p)
    res = _detect(p, tmp, "unl_closings")
    assert not _block(res, "balance_sheet")["bound"], _block(res, "balance_sheet")["rows"]


def closings_check(path):
    m = MW(path)
    n = m.n
    for r0, name, s0, k in ((5, "A", 100.0, -5.0), (10, "B", 60.0, -2.0)):
        o, cl = [s0], []
        for i in range(n):
            cl.append(o[i] + k)
            o.append(cl[i])
        m.put("M", r0, f"Opening {name}", o[:n], lambda i, c, r0=r0, s0=s0: f"={s0:g}" if i == 0 else f"={_prev(m, c)}{r0 + 2}")
        m.put("M", r0 + 1, f"Repayment {name}", [k] * n)
        m.put("M", r0 + 2, f"Closing balance {name}", cl, lambda i, c, r0=r0: f"=SUM({c}{r0}:{c}{r0 + 1})")
    ta = [100.0 - 5.0 * (i + 1) + 60.0 - 2.0 * (i + 1) for i in range(n)]
    m.put("M", 16, "Closing balance", ta, lambda i, c: f"={c}7+{c}12")
    m.put("M", 18, "Tranche reconciliation", [0.0] * n, lambda i, c: f"={c}16-({c}7+{c}12)")
    m.close()


def distributions_tie(tmp):
    """Item 3: an enterprise DCF's cash-flow row equals an operating row; that is not a distributions tie."""
    p = tmp / "ent.xlsx"
    enterprise_dcf(p)
    res = _detect(p, tmp, "ent")
    ids = _ids(res)
    assert ids["dist_cf_eq"]["status"] == "unbound" and ids["dist_cf_eq"]["why"], ids["dist_cf_eq"]
    assert ids["dist_dcf_cf"]["status"] == "unbound" and "M!r25" in ids["dist_dcf_cf"]["why"], ids["dist_dcf_cf"]
    cf = _block(res, "cash_flow")
    assert "cf_distributions" not in cf["rows"] and "cf_distributions" in cf["unbound"], cf
    assert _row(res, "equity", "re_distributions") == "M!r30"   # found, and not compared with an operating row
    assert not [f for f in res["findings"] if f["kind"] == "model_fails_identity"], res["findings"]


def several(tmp):
    """Item 4: the best balance sheet and corkscrew are bound; the others are listed, and a failing roll is a finding."""
    p = tmp / "two.xlsx"
    two_balance_sheets(p)
    res = _detect(p, tmp, "two")
    assert _row(res, "balance_sheet", "assets") == "Group!r28" and _row(res, "debt", "debt_closing") == "Group!r19"
    bs = _block(res, "balance_sheet")
    assert [(a["sheet"], a["status"]) for a in bs["alternates"]] == [("Asset1", "holds")], bs["alternates"]
    assert bs["alternates"][0]["assets"]["row"] == 6
    debt = _block(res, "debt")
    assert [(a["closing"]["row"], a["status"], a["tie"]) for a in debt["alternates"]] == [(23, "fails", None)], debt["alternates"]
    assert {c["closing"]["row"]: c["status"] for c in res["corkscrews"] if c["kind"] == "debt"} == {19: "holds", 23: "fails"}
    cf = [f for f in res["findings"] if f["kind"] == "corkscrew_fails"]
    assert len(cf) == 1 and "Group!r23" in cf[0]["text"] and "largest residual 1" in cf[0]["text"], cf
    assert _ids(res)["debt_roll"]["status"] == "holds"
    t = statements.text(res)
    assert "also found: corkscrew Group!r23" in t and "also found: balance sheet on Asset1" in t


def plug(tmp):
    """Item 5: equity computed as assets minus liabilities cannot test anything."""
    p = tmp / "plug.xlsx"
    minimal(p, plug=True, imbalance=5.0)
    res = _detect(p, tmp, "plug")
    ids = _ids(res)
    reason = "equity is a plug (assets − liabilities), so the identity cannot test the model"
    assert ids["bs_balance"]["status"] == "unbound" and ids["bs_balance"]["why"] == reason, ids["bs_balance"]
    assert ids["bs_total_le"]["status"] == "unbound" and ids["bs_total_le"]["why"] == reason
    pf = [f for f in res["findings"] if f["kind"] == "plugged_balance_sheet"]
    assert len(pf) == 1 and "M!r29" in pf[0]["text"] and "Total equity" in pf[0]["text"], res["findings"]
    assert not [f for f in res["findings"] if f["kind"] == "model_fails_identity"]
    # the balanced model with a memo 'net assets' row beside real equity is not a plug
    p = tmp / "plug_memo.xlsx"
    minimal(p, net_assets=True)
    assert not [f for f in _detect(p, tmp, "plug_memo")["findings"] if f["kind"] == "plugged_balance_sheet"]


def rebind(tmp):
    """Item 6: our binding fails while the model's own check holds."""
    p = tmp / "mini_check.xlsx"
    minimal(p, check=True)
    db = build_map.main(str(p), str(tmp / "rb__0a1b2c3d"))["db"]
    ident = ontology.identity("bs_balance")
    d = statements.Detector(db)
    d.run()
    assert d.bind["equity"] == ("M", 29) and d.results["bs_balance"]["status"] == "holds"
    d.bind["equity"] = ("M", 24)          # our rows differ from the check's: share capital alone is not the equity total
    rec = d.run_identity(ident)
    assert rec["status"] == "holds" and d.bind["equity"] == ("M", 29), rec   # the check's rows hold: rebound to them
    assert d.by["equity"] == "check" and "overruled" in d.how["equity"] and rec["kind"] == "test"
    # the check's rows fail too (its saved value was typed over): not 'holds', not silent
    p = tmp / "mini_loose.xlsx"
    minimal(p, check=True, imbalance=5.0)
    db = build_map.main(str(p), str(tmp / "rb2__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    d.run()
    d.findings.clear()
    d.bind["equity"] = ("M", 24)
    rec = d.run_identity(ident)
    assert rec["status"] == "unbound" and rec["why"].startswith("model's own check holds; binding rejected"), rec
    f = [x for x in d.findings if x["kind"] == "check_and_binding_disagree"]
    assert len(f) == 1 and "M!r19" in f[0]["text"] and "M!r29" in f[0]["text"], d.findings


def stale_guards(tmp):
    """Item 7."""
    p = tmp / "guards.xlsx"
    circular_and_seeded(p)
    db = build_map.main(str(p), str(tmp / "guards__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    d.bind = {"x": ("M", 7), "y": ("M", 11)}
    d.stale()
    assert not [f for f in d.findings if f["kind"] == "stale"], d.findings
    assert d.stats["stale_rows_circular_skipped"] == 1 and d.stats["stale_rows_differing"] == 0, d.stats
    assert d.stats["stale_cells_checked"] >= 18, d.stats
    # the lookback is the whole upstream, and capped (with a note)
    d = statements.Detector(db)
    d.bind = {"x": ("M", 7)}
    d.STALE_CAP = 4
    d.stale()
    assert d.stats["stale_cells_checked"] == 4 and any("stale test stopped after 4 cells" in c for c in d.stats["capped"]), d.stats
    d = statements.Detector(db)
    d.bind = {"x": ("M", 7)}
    d.stale()
    assert d.stats["stale_cells_checked"] >= 10 and not d.stats["capped"]
    # a real difference on a row that is not circular is still reported (three hops upstream of the bound row)
    p = tmp / "mini_extra2.xlsx"
    minimal(p, extra=True)
    db = build_map.main(str(p), str(tmp / "me2__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    d.bind = {"x": ("M", 9)}
    d.stale()
    assert d.stats["stale_cells_checked"] >= 12 and d.stats["stale_rows_circular_skipped"] == 0


def pattern_exemptions(tmp):
    """Item 8: a typed first period is a seed, a typed number in the middle of the row is a break."""
    db = res_db(tmp, "guards")
    d = statements.Detector(db)
    d.bind = {"a": ("M", 14), "b": ("M", 15), "c": ("M", 17)}
    d.pattern_breaks()
    pb = [f["text"] for f in d.findings if f["kind"] == "pattern_break"]
    assert len(pb) == 1 and pb[0].startswith("M!I17 (Hardcoded)"), pb


def minor(tmp):
    """Item 9: _triples reads the caller's absolute tolerance; a single-column block uses its value column."""
    p = tmp / "mini_tol.xlsx"
    minimal(p, imbalance=3.0)
    db = build_map.main(str(p), str(tmp / "mini_tol__0a1b2c3d"))["db"]
    tight = statements.detect(db)
    loose = statements.detect(db, tolerance=(1e-6, 5.0))
    assert "label only" in _block(tight, "balance_sheet")["rows"]["assets"]["how"]
    assert _ids(tight)["bs_balance"]["status"] == "fails"
    assert _block(loose, "balance_sheet")["rows"]["assets"]["how"].startswith("assets = liabilities + equity")
    assert _ids(loose)["bs_balance"]["status"] == "holds"
    p = tmp / "su1.xlsx"
    single_block(p)
    db = build_map.main(str(p), str(tmp / "su1__0a1b2c3d"))["db"]
    d = statements.Detector(db)
    r = statements.evaluate(d.b, [(("SU", 6), 1, 0), (("SU", 11), -1, 0)], set(), d.tol, "single")
    assert r["status"] == "holds" and r["periods_checked"] == 1 and r["scope_used"] == "single", r


def main():
    tmp = Path(tempfile.mkdtemp(prefix="statements_"))
    try:
        paths = make_threeway_model.build(tmp)
        results = {}
        for name in ("base", "unbalanced", "moved", "stale", "assets"):
            results[name] = _detect(Path(paths[name]), tmp, name)
        base(results["base"])
        unbalanced(results["unbalanced"])
        moved(results["moved"])
        stale(results["stale"])
        assets(results["assets"])
        for name in ("base", "unbalanced", "moved", "assets"):
            assert not [f for f in results[name]["findings"] if f["kind"] == "stale"], name
        kinds_and_stats(results)
        print("ok: three-way fixture, 5 variants")
        minimal_checks(tmp)
        review_fixes(tmp)
        print("ok: minimal workbook")
        bs_anchor(tmp)
        distributions_tie(tmp)
        several(tmp)
        plug(tmp)
        rebind(tmp)
        stale_guards(tmp)
        pattern_exemptions(tmp)
        minor(tmp)
        print("ok: kinds, balance sheet anchor, distribution ties, several statements, plug, rebind, stale guards, "
              "seeded rows, tolerances")
        real_models()
        ontology_doc()
        print("ok: ontology doc, real-model smoke")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
