"""The dependency graph behind a DCF value (modelatlas/depgraph.py) on the synthetic sample workbooks.

  Riverbend   from the Summary's equity value: the factor row, the rate and the valuation date are the discounting,
              the terminal value row and its growth rate are terminal, the rows the cash flow adds up are cash flow,
              the typed inputs at the bottom are assumptions, net debt is the bridge; every edge joins two nodes, the
              value is at layer 0, and the result is plain JSON
  bounded     max_rows=5 keeps five rows and collapses the rest into one group per sheet, the counts adding up to
              what was left out
  overlay     XNPV discountings: the rate cells and the date row they count from are discounting
  older db    a model.db with no edges table: the trace and the discounting's rows, no crash
  rules       layers end on circular graphs; 'Terminal handling charges' isn't a terminal value; a units divisor beside the
              value (=EV/Thousand) isn't a bridge term, a cell added or taken away is; a row of constants with a total
              is an input; mid-period / timing rows are discounting; debt funding isn't capex; the build-up words
  shapes      a small synthetic workbook (xlsxwriter): a discounted row that is a sum times a flag (its direct terms
              are cash flow parts), a SUMPRODUCT on the cash-flow row itself (no edge from the row to the factor row,
              counted in stats.edges_dropped), the year-fraction row the factor row reads, a cell on the path between
              the value and the core, a row picked by a scenario selector (only when 10+ lookups read one; a period
              counter read by as many is not one), debt funding upstream of the cash flow, a tax rate read by the rate
              and the cash flow (not a build-up) beside a risk-free rate (a build-up), the SUMPRODUCT's own operand
              cells (cash flow and factors, not bridge terms); what the period columns read is told from their text,
              and nothing is dropped when it can't be (INDIRECT, OFFSET, whole columns)

    uv run python tests/check_depgraph.py
"""
import json
import re
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import xlsxwriter  # noqa: E402
from datetime import date  # noqa: E402

from modelatlas import build_map  # noqa: E402
from modelatlas import depgraph  # noqa: E402
from modelatlas import rodb  # noqa: E402

PACK = ROOT / "tests" / "sample_models"


def _db(name: str) -> str:
    out = Path(tempfile.mkdtemp(prefix="depgraph_"))
    return build_map.main(str(PACK / f"{name}.xlsx"), str(out / f"{name}__0a1b2c3d"))["db"]  # out/<name>__<sha8>/


def riverbend() -> str:
    path = _db("Riverbend_BP25_with_overlay")
    g = depgraph.build(path, "Summary!C5")
    by = {n["id"]: n for n in g["nodes"]}

    def is_(i, cls, sub=None, inp=None):
        n = by[i]
        assert n["class"] == cls and (sub is None or n["subclass"] == sub) and (inp is None or n["input"] is inp), \
            (i, n["class"], n["subclass"], n["input"], n["why"])
    assert g["start"]["id"] == "Summary!C5" and g["start"]["label"] == "Equity value", g["start"]
    assert by["Summary!C5"]["class"] == "outcome" and by["Summary!C5"]["layer"] == 0
    is_("DCF!r9", "discounting", "discount factor")
    is_("Val_Inputs!r5", "discounting", "discount rate", True)
    is_("Val_Inputs!r4", "discounting", "valuation date", True)
    is_("DCF!r6", "terminal")
    is_("Val_Inputs!r6", "terminal", "terminal growth", True)
    for r in (6, 7, 8):
        is_(f"CashFlow!r{r}", "cashflow")
    is_("DCF!D12", "discounting", "present value")
    is_("DCF!D13", "bridge")
    is_("Val_Inputs!r7", "bridge", None, True)
    reached = [n for n in g["nodes"] if n["sheet"] == "Inputs" and n["kind"] == "row"]
    assert reached and all(n["class"] in ("assumption", "timeline") for n in reached), [(n["id"], n["class"]) for n in reached]
    assert any(n["class"] == "assumption" and n["input"] for n in reached)
    assert [c["key"] for c in g["classes"]] == ["outcome", "bridge", "discounting", "terminal", "cashflow", "calculation",
                                                "timeline", "assumption"]
    ids = {n["id"] for n in g["nodes"]}
    assert all(e["src"] in ids and e["dst"] in ids for e in g["edges"]), "an edge to a node that isn't there"
    assert all(n["layer"] >= 0 for n in g["nodes"])
    assert {e["kind"] for e in g["edges"]} >= {"trace", "part"}, {e["kind"] for e in g["edges"]}
    assert g["cores"] and g["cores"][0]["cell"] == "DCF!D12" and g["cores"][0]["factor_row"] == "DCF!r9"
    assert g["cores"][0]["method"]["rate"] == "Val_Inputs!C5", g["cores"][0]["method"]
    assert g["workbook"] == "Riverbend_BP25_with_overlay.xlsx", g["workbook"]
    assert g["anchors"] and any(a["cell"] == "Summary!C5" for a in g["anchors"])
    assert json.loads(json.dumps(g)) == g, "the graph isn't plain JSON"
    assert sum(c["n"] for c in g["by_class"]) == g["stats"]["rows_shown"] + g["stats"]["cells_shown"]
    top = depgraph.build(path, with_anchors=True)  # no cell: the top anchor
    assert top["start"]["id"] == top["anchors"][0]["cell"]
    return path


def bounded(path: str) -> None:
    g = depgraph.build(path, "Summary!C5", max_rows=5)
    rows = [n for n in g["nodes"] if n["kind"] == "row"]
    groups = [n for n in g["nodes"] if n["kind"] == "group"]
    assert len(rows) <= 5 and g["stats"]["rows_shown"] == len(rows), len(rows)
    assert groups and all(n["id"] == f"group:{n['sheet']}" for n in groups)
    assert sum(n["count"] for n in groups) == g["stats"]["rows_upstream_total"] - g["stats"]["rows_shown"], g["stats"]
    ids = {n["id"] for n in g["nodes"]}
    assert all(e["src"] in ids and e["dst"] in ids for e in g["edges"])
    assert any(e["kind"] == "collapsed" for e in g["edges"])
    assert all(n["layer"] >= 0 for n in g["nodes"]) and g["nodes"][0]["layer"] == 0
    shallow = depgraph.build(path, "Summary!C5", depth=2)
    assert max(n["depth"] for n in shallow["nodes"] if n["kind"] == "row") <= 2


def overlay() -> None:
    path = _db("trace_overlay")
    g = depgraph.build(path, "Valuation summary!F16")
    by = {n["id"]: n for n in g["nodes"]}
    assert g["cores"], "no discounting found"
    assert any(c["kind"] == "xnpv" for c in g["cores"]), [c["kind"] for c in g["cores"]]
    for i in ("Inputs!r5", "Inputs!r6", "Cash flows!r11"):
        assert by[i]["class"] == "discounting", (i, by[i]["class"], by[i]["why"])
    assert by["Valuation summary!F15"]["class"] == "bridge"
    assert by["Valuation summary!F10"]["class"] == "discounting"
    assert json.loads(json.dumps(g)) == g


def older(path: str) -> None:
    old = Path(tempfile.mkdtemp(prefix="depgraph_old_")) / "Riverbend_old__0a1b2c3d" / "model.db"
    old.parent.mkdir()
    shutil.copy(path, old)
    with sqlite3.connect(old) as c:
        c.execute("DROP TABLE edges")
    g = depgraph.build(str(old), "Summary!C5", with_anchors=False)
    by = {n["id"]: n for n in g["nodes"]}
    assert by["DCF!r9"]["class"] == "discounting" and by["Val_Inputs!r5"]["class"] == "discounting", g["nodes"]
    ids = set(by)
    assert all(e["src"] in ids and e["dst"] in ids for e in g["edges"]) and g["stats"]["groups"] == 0


def rules() -> None:
    m = lambda f, c: {"label": "", "n_formula": f, "n_const": c}  # noqa: E731
    assert depgraph._typed(m(2, 212)) and depgraph._typed(m(0, 5)) and depgraph._typed(m(1, 3))
    assert not depgraph._typed(m(1, 2)) and not depgraph._typed(m(214, 0)) and not depgraph._typed(m(0, 0))
    for lab, hit in (("Mid-year discounting switch", True), ("Mid-year convention", True), ("Discounting timing", True),
                     ("Periods to discount", True), ("Mid year revenue", True), ("Revenue", False)):
        assert any(re.search(pat, lab, re.I) for pat, sub, _ in depgraph.DISCOUNTING if sub == "timing") is hit, lab
    for lab, hit in (("Corporate tax rate", True), ("Post-tax cost of debt", True), ("Debt margin", True), ("Risk-free rate", True),
                     ("Market risk premium", True), ("Tax paid", False), ("EBITDA margin", False), ("Risk adjustment", False),
                     ("Unlevered beta", True)):
        assert bool(depgraph.BUILDUP.search(lab)) is hit, lab
    for lab, hit in (("Capex drawdown", True), ("Debt funded capex", True), ("Facility repayment", True), ("Refinancing", True),
                     ("Growth capex", False), ("Maintenance capex", False)):
        assert bool(depgraph.DEBT_FUNDING.search(lab)) is hit, lab
    E = lambda *p: [{"src": a, "dst": b} for a, b in p]  # noqa: E731
    lay = depgraph._layers("S", ["S", "A", "B", "C", "D", "X"],
                           E(("S", "A"), ("A", "B"), ("B", "A"), ("B", "C"), ("C", "A"), ("S", "D"), ("D", "D")))
    assert lay["S"] == 0 and all(isinstance(v, int) and v >= 0 for v in lay.values()) and lay["C"] > lay["B"] > lay["A"], lay
    for lab, hit in (("Terminal value", True), ("EV incl. terminal", True), ("Terminal growth rate", True),
                     ("Terminal Val Date", True), ("Include terminal value", True), ("Terminal handling charges", False),
                     ("Terminal B revenue", False)):
        assert bool(depgraph.TERMINAL.search(lab)) is hit, lab
    db = sqlite3.connect(":memory:")
    db.execute("CREATE TABLE names(name TEXT, ref TEXT)")
    db.execute("INSERT INTO names VALUES ('Thousand', 'Units!$C$3')")
    names = {"thousand": "Units!$C$3"}
    assert depgraph._scales(db, names, "=Value!D20/Thousand", "S", "Units!C3", {"Value!D20"})
    assert depgraph._scales(db, names, "=D12*Units!C3", "S", "Units!C3", {"S!D12"})
    assert not depgraph._scales(db, names, "=SUM(F18:F21)", "S", "S!F19", {"S!F18"})
    assert not depgraph._scales(db, names, "=D12-D13", "S", "S!D13", {"S!D12"})
    assert not depgraph._scales(db, names, "=D12-D13*Units!C3", "S", "S!D13", {"S!D12"})
    read = lambda *f: depgraph._rows_read(db, {"df": "'Disc rates'!$D$12:$K$12"}, list(f), "Flows")  # noqa: E731
    assert read("=SUM(D5,D6)*INDEX(DF,1)", "=(D5+D6)*'It''s'!D7") == [("Flows", 5, 5), ("Flows", 6, 6), ("Disc rates", 12, 12),
                                                                     ("Flows", 5, 5), ("Flows", 6, 6), ("It's", 7, 7)]
    assert read('=IF(D3>"10:20",1,0)') == [("Flows", 3, 3)]
    for blind in ('=D5*INDIRECT("R12C"&COLUMN(),FALSE)', "=D5*OFFSET(D11,1,0)", "=SUM(D:D)", "=SUM(12:12)"):
        assert read(blind) is None, blind  # can't tell what it reads: the edges stay


def shapes() -> None:
    out = Path(tempfile.mkdtemp(prefix="depgraph_shapes_"))
    n, vd = 8, date(2025, 6, 30)
    ends = [date(2025, 9, 30), date(2025, 12, 31), date(2026, 3, 31), date(2026, 6, 30), date(2026, 9, 30),
            date(2026, 12, 31), date(2027, 3, 31), date(2027, 6, 30)]
    wb = xlsxwriter.Workbook(out / "shapes.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp, cases, fl, summ = (wb.add_worksheet(s) for s in ("Inputs", "Cases", "Flows", "Summary"))
    for r, lab, v in ((3, "Valuation date", None), (4, "Discount rate", None), (7, "Scenario", 1), (15, "Risk-free rate", 0.04),
                      (16, "Tax rate", 0.25), (17, "Mid-year discounting switch", 0)):
        inp.write(r, 0, lab)
        if v is not None:
            inp.write_number(r, 2, v)
    inp.write_datetime(3, 2, vd, dt)
    inp.write_formula("C5", "=C16+C17/10", None, 0.065)  # the rate reads the risk-free rate and the tax rate
    inp.write_number(7, 3, 2)
    inp.write_number(7, 4, 3)
    for k in range(12):  # twelve lookup rows read the selector (a typed row of values and one INDEX)
        inp.write(20 + k, 0, f"Case driver {k}")
        for j in range(3):
            inp.write_number(20 + k, 4 + j, 0.1 * (j + 1))
        inp.write_formula(20 + k, 3, f"=INDEX(E{21 + k}:G{21 + k},$C$8)", None, 0.1)
    cases.write(13, 0, "Growth pick")  # one constant and formulas that only CHOOSE by the selector
    cases.write_number(13, 2, 0.05)
    for c in "DE":
        cases.write_formula(f"{c}14", "=CHOOSE(Inputs!$C$8,0.1,0.2,0.3)", None, 0.1)
    fl.write(2, 1, "Period ending")
    labels = {4: "Revenue", 5: "Growth capex", 6: "Active flag", 7: "Capex debt drawdown", 9: "Cash flow", 10: "Years",
              11: "Discount factor"}
    for r, lab in labels.items():
        fl.write(r, 1, lab)
    fl.write(3, 1, "Pasted debt funded (paste)")
    for k in range(n):
        c = xlsxwriter.utility.xl_col_to_name(3 + k)
        fl.write_datetime(2, 3 + k, ends[k], dt)
        fl.write_number(3, 3 + k, 0.3)
        fl.write_formula(f"{c}5", f"=100*(1+Cases!$D$14)+{c}4*0", None, 110)
        fl.write_formula(f"{c}6", f"=-{c}5*0.1*(1-Inputs!$C$17)-{c}8", None, -5)
        fl.write_formula(f"{c}7", f"=IF({c}3>Inputs!$C$4,1,0)", None, 1)
        fl.write_formula(f"{c}8", f"={c}5*0.3", None, 33)
        fl.write_formula(f"{c}10", f"=SUM({c}5,{c}6)*{c}7", None, 105)
        yrs = (ends[k] - vd).days / 365  # end of period (the mid-period switch is off), so dcftrace reads the rate back
        fl.write_formula(f"{c}11", f"=({c}3-Inputs!$C$4)/365-0.5*Inputs!$C$18", None, yrs)
        fl.write_formula(f"{c}12", f"=1/(1+Inputs!$C$5)^{c}11", None, 1 / 1.065 ** yrs)
        fl.write_number(13, 3 + k, k + 1)  # a period counter that a dozen lookups read: not a scenario selector
    fl.write(13, 1, "Period number")
    prof = wb.add_worksheet("Profiles")
    for k in range(12):
        prof.write(4 + k, 1, f"Profile {k}")
        prof.write_formula(4 + k, 3, f"=INDEX(Flows!$D$14:$K$14,Flows!D14)", None, 1)
    fl.write_formula("C4", "=SUM(D4:K4)", None, 2.4)  # a pasted row of constants with a total
    fl.write_formula("C10", "=SUMPRODUCT(D10:K10,D12:K12)", None, 780.0)  # the total sits on the cash-flow row
    summ.write(3, 1, "Equity PV")
    summ.write_formula("C4", "=Flows!C10", None, 780.0)
    summ.write(4, 1, "Equity value")
    summ.write_formula("C5", "=C4-C6", None, 700.0)
    summ.write(5, 1, "Debt")
    summ.write_number(5, 2, 80.0)
    wb.close()
    path = build_map.main(str(out / "shapes.xlsx"), str(out / "shapes__0a1b2c3d"))["db"]
    g = depgraph.build(path, "Summary!C5", with_anchors=False)
    by = {x["id"]: x for x in g["nodes"]}

    def is_(i, cls, sub=None, inp=None, why=None):
        x = by[i]
        assert x["class"] == cls and (sub is None or x["subclass"] == sub) and (inp is None or x["input"] is inp) and \
            (why is None or why in x["why"]) and x["why"], (i, x["class"], x["subclass"], x["input"], x["why"])
    assert any(c["kind"] == "sumproduct" and c["cell"] == "Flows!C10" and c["method"]["rate"] for c in g["cores"]), g["cores"]
    # the SUMPRODUCT's own operand cells are the cash flow and the factors, not bridge terms
    is_("Flows!D10", "cashflow")
    is_("Flows!D12", "discounting", "discount factor")
    # 1: the discounted row is a sum times a flag, so its direct reads are the parts
    for i in ("Flows!r5", "Flows!r6", "Flows!r7"):
        is_(i, "cashflow", "cash flow part (read directly)", why="read directly")
    # 2: the SUMPRODUCT sits on the cash-flow row; the row's period columns read none of the discounting rows
    assert ("Flows!r10", "Flows!r12") not in {(e["src"], e["dst"]) for e in g["edges"]}
    assert ("Flows!C10", "Flows!r12") in {(e["src"], e["dst"]) for e in g["edges"]}, "the core still names its factor row"
    assert g["stats"]["edges_dropped"] >= 1, g["stats"]
    # 8: the row the factor row reads (years from the valuation date), 7: the cell between the value and the core
    is_("Flows!r11", "discounting", "discount period", why="factor row")
    is_("Summary!C4", "discounting", "discounted value")
    # 4: a row picked by the selector is an assumption; the selector itself is read by 12 lookup rows (a period counter
    #    read by as many is not a selector)
    is_("Cases!r14", "assumption", "scenario-selected", True, why="selector")
    assert depgraph.Model(rodb.connect(path)).selectors() == {("Inputs", 8)}
    # 6: debt funding upstream of the cash flow is not capex
    is_("Flows!r8", "calculation", "debt schedule", why="debt funding")
    # 3: a pasted row with a total is an input, not a part of the cash flow
    is_("Flows!r4", "assumption", None, True)
    # 9: a tax rate read by the discount rate and by the cash flow is an assumption, not a rate build-up; a risk-free
    #    rate the discount rate reads is
    is_("Inputs!r17", "assumption", "tax")
    is_("Inputs!r16", "discounting", "discount rate build-up")
    # 5: mid-period discounting is a timing row
    is_("Inputs!r18", "discounting", "timing")
    # nothing changes without a selector: the sample workbook has no scenario-selected rows
    assert not any(x["subclass"] == "scenario-selected" for x in depgraph.build(_db("Riverbend_BP25_with_overlay"), "Summary!C5")["nodes"])
    print("shapes: ok")


def cli(path: str) -> None:
    out = Path(tempfile.mkdtemp(prefix="depgraph_cli_"))
    depgraph.main([path, "Summary!C5", "--max-rows", "50", "--json", str(out / "g.json"), "--html", str(out / "g.html")])
    assert json.loads((out / "g.json").read_text())["start"]["id"] == "Summary!C5"


if __name__ == "__main__":
    p = riverbend()
    bounded(p)
    overlay()
    older(p)
    rules()
    shapes()
    cli(p)
    print("ok")
