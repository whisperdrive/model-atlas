"""The diagnostic runner (modelatlas/diagnose.py): reports that carry the shapes of models and nothing that names them.

  scrubbing    one pass over a formula: sheets -> S<k>, defined names -> NAME<k>, strings, big or non-round numbers and
               non-Excel functions replaced, R1C1 references and small structural numbers kept; machine paths removed
  graphs       strongly connected components and the longest path of a small sheet graph
  run          the three-way fixture variants and the Riverbend workbook, built and run: a report per model, the
               summary lists them, the shapes JSON and Markdown hold none of the fixture's sheet names, labels, file
               stems or the company name; identities with kinds and statuses (unbalanced: bs_balance fails); the
               complexity profile and the fixture hints for the failing identity, with scrubbed patterns
  registry     names kept in out/registry.db for a model are part of what the self-scan looks for; a registry that
               cannot be read blocks the report
  messages     an exception message keeps only known safe words: labels, text-cell words, CJK, dates, amounts, paths
               and defined names never pass
  gate         per value, not per joined line ('rev_total', 'revenue' is not "total revenue"); a key outside the schema
               blocks even when it is a lowercase word, and does not exempt itself; non-ASCII blocks; one-word labels,
               text cells and sections are scanned for
  planted      a fixture copy with sheets named like functions, a label word plus a code, defined names that look like
               cells, a CJK label, a section only in rows.section, a title cell, strings, external, structured and
               UDF references in formulas, and a registry target: the report is written and none of it is in it
  unreadable   a model.db that cannot be opened is a recorded failure; the run goes on
  blocked      a stage that emits a label: no shapes report, a local _blocked note naming it, blocked in the summary
  failing      a stage that raises is recorded (scrubbed message and traceback), the other stages and models still run
  full         level full writes real names under diag/full only
  tests        the test-suite runner on one fast script

    uv run python tests/check_diagnose.py
"""
import contextlib
import io
import json
import re
import shutil
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from modelatlas import diagnose  # noqa: E402
import make_threeway_model  # noqa: E402

SHEETS = ["BalanceSheet", "CashFlow", "PnL", "Checks", "Revenue", "Equity", "Debt", "Capex", "Costs", "DCF", "Inputs"]
LABELS = ["total revenue", "closing cash", "senior debt", "harbourline", "total assets", "balance check"]
STEMS = ["threeway_model", "threeway_unbalanced", "threeway_moved", "threeway_stale", "threeway_assets", "riverbend_bp25_client_model"]


def run(argv) -> str:
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        assert diagnose.main(argv) == 0
    return buf.getvalue()


def read(p: Path) -> str:
    return p.read_text(encoding="utf-8")


def pattern_scrub() -> None:
    a = diagnose.Anon(["Revenue", "Cash flows", "O'Brien"], ["Tax_rate", "Debt_margin"])
    got = a.pattern('=SUM(Revenue!R[-2]C:R[-1]C)*Tax_rate+\'Cash flows\'!R5C3+\'O\'\'Brien\'!RC[1]+"Acme Ltd"+0.0847+12+1234567+365+0.5'
                    '+Debt_margin+MYFIRMFN(R1C1)+Other!R2C2+IF(TRUE,#REF!,1000)')
    assert got == ('=SUM(S1!R[-2]C:R[-1]C)*NAME2+S2!R5C3+S3!RC[1]+"…"+#+12+#+365+0.5+NAME1+FN1(R1C1)+X1!R2C2+IF(TRUE,#REF!,1000)'), got
    # a sheet name that is part of a defined name does not leak through the name
    b = diagnose.Anon(["Debt"], ["Debt_margin"])
    assert b.pattern("=Debt_margin*Debt!R1C1") == "=NAME1*S1!R1C1"
    # an unknown identifier is a name, never kept
    assert diagnose.Anon([], []).pattern("=Secret_Thing+1") == "=NAME1+1"
    # the machine
    s = diagnose.scrub_env(f"{ROOT}/modelatlas/edges.py and {Path.home()}/x and {ROOT.as_posix()}")
    assert str(Path.home()) not in s and str(ROOT) not in s and s.startswith("modelatlas/edges.py"), s
    assert diagnose.sign_runs([1.5, 2, 0, None, "x", -3, -1]) == "+x2 0x1 _x1 tx1 -x2"
    assert diagnose.residual_bucket(0, 5) == "0" and diagnose.residual_bucket(3e-7, 1) == "<1e-6"
    assert diagnose.residual_bucket(500, 1000) == ">=1e-1" and diagnose.residual_bucket(1, 0) == "scale_zero"
    assert diagnose.reason_codes("periodicity differs (a, b); x: no capex row; y: is a plug (z)") == ["no_candidate", "periodicity", "plug"] \
        or "periodicity" in diagnose.reason_codes("periodicity differs (a, b)")
    assert diagnose.reason_codes("rev: not a sum of segments") == ["structure_mismatch"]
    # only Excel's own error literals survive a '#'; the rest of a '#WORD' is a name
    assert a.pattern("=#ACME+#N/A+#DIV/0!+R1C1#") == "=#NAME3+#N/A+#DIV/0!+R1C1#", a.pattern("=#ACME+#N/A+#DIV/0!+R1C1#")
    assert diagnose.Anon([], []).func("_xlfn._xlws.SORT") == "SORT" and diagnose.Anon([], []).func("_xludf.Wombat") == "FN1"
    # names that look like cells, dotted names, structured and external references, LAMBDA parameters, a sheet called SUMIF
    c = diagnose.Anon(["SUMIF", "Kestrel Ops"], ["TAX1", "Emu.Rate"])
    got = c.pattern("=TAX1*Emu.Rate+SUMIF!R1C1+SUMIF(R1C1:R2C2,1)+'Kestrel Ops'!RC+Tbl_Bilby[[#This Row],[Numbat]]"
                    "+[1]Koala!R1C1+'C:\\Deals\\[Galah.xlsx]Cockatoo'!R1C1+LAMBDA(Possum,Possum*2)(3)+INDIRECT(\"Platypus!A1\")")
    for w in ("TAX1", "Emu", "Kestrel", "Bilby", "Numbat", "Koala", "Galah", "Cockatoo", "Possum", "Platypus", "Deals"):
        assert w not in got, (w, got)
    assert got.startswith("=NAME2*NAME1+S1!R1C1+SUMIF(") and "S2!RC" in got and "[#This Row]" in got, got
    print("scrubbing: ok")


def messages() -> None:
    """An exception message is projected onto known safe words."""
    a = diagnose.Anon(["Revenue", "Kestrel Ops", "BS"], ["Val_date2", "TAX1", "Emu.Rate"])
    for raw, gone in [("no row for Wombat under Quokka Holdings at 2025-06-30, Revenue XQZ", ("Wombat", "Quokka", "XQZ", "2025", "06-30")),
                      ("could not convert Emu.Rate value 31/12/2027 for BS!C5 Kestrel Ops!D4 in 12.5%", ("Emu", "Rate", "2027", "12/", "Kestrel", "12.5")),
                      ("no such column: Wombat", ("Wombat",)), ("Revenue!R12C4 = 1,234,567 AUD", ("AUD", "234", "Revenue")),
                      ("label \u9577\u6c5f\u96fb\u529b on Val_date2 and TAX1", ("\u9577", "Val_date2", "TAX1", "date2")),
                      (r"[Errno 2] No such file or directory: D:\Deals\Acme\x.db", ("Deals", "Acme")),
                      ("KeyError: ('Kestrel Ops', 12)", ("Kestrel",))]:
        got = a.message(raw)
        for g in gone:
            assert g not in got, (g, got)
        assert not diagnose.non_ascii(got), got
    for keep in ("list index out of range", "database is locked", "no DCF value cell found in this workbook; give one (Sheet!C5)",
                 "division by zero"):
        assert a.message(keep) == keep, (keep, a.message(keep))
    assert a.message("could not convert BS!C5") == "could not convert S3!C5"
    # a frame outside the repo and Python is its file name only
    assert diagnose.code_path("/srv/Deals/Wombat Pty/run.py") == "<other>/run.py"
    assert diagnose.code_path(str(ROOT / "modelatlas" / "statements.py")) == "modelatlas/statements.py"
    print("messages: ok")


def gate_rules() -> None:
    """Per value grams, a fixed schema of keys, one-word terms, non-ASCII."""
    sc = diagnose.Scanner({"wombatco"}, {"total revenue", "revenue xqz"}, terms={"wombat", "xqz", "quokka"})
    ctx = type("C", (), {"scanner": sc, "registry_unreadable": False})()
    rep = {"stages": {"hints": {"result": {"hints": [{"rows": [{"label_vocab": ["rev_total", "revenue"]}]}]}}}}
    md = "- revenue (label) S6 r7: vocab ['rev_segments*', 'rev_total', 'revenue']\n"
    assert diagnose.gate(ctx, rep, md) == [], diagnose.gate(ctx, rep, md)       # not "total revenue"
    assert diagnose.gate(ctx, {"note": "Total revenue"}, "") == ["total revenue"]
    assert "wombat" in diagnose.gate(ctx, {"note": "a Wombat"}, "")                 # a one-word label
    assert "xqz" in diagnose.gate(ctx, {"note": "the XQZ line"}, "")
    assert any("unexpected key" in h for h in diagnose.gate(ctx, {"stages": {"wombat": 1}}, ""))
    assert any("unexpected key" in h for h in diagnose.gate(ctx, {"stages": {"quokka": "quokka"}}, ""))
    assert diagnose.gate(ctx, {"note": "S1 \u9577\u6c5f"}, "") == ["<non-ASCII text>"]
    assert diagnose.gate(ctx, {"note": "\"…\""}, "") == []                          # our own ellipsis is fine
    ctx.registry_unreadable = True
    assert diagnose.gate(ctx, {"stages": {}}, "")
    # traceback lines are code: scanned for names, not for single words
    tb = {"stages": {"census": {"error": {"traceback": ["modelatlas/x.py:1 in f | wombat = 1"], "site": "modelatlas/x.py:1 in f"}}}}
    assert diagnose.gate(type("C", (), {"scanner": sc})(), tb, "") == []
    print("gate: ok")


def graphs() -> None:
    adj = {0: {1}, 1: {2}, 2: {0, 3}, 3: {4}}
    comps = diagnose.scc(5, adj)
    assert sorted(len(c) for c in comps) == [1, 1, 3], comps
    assert diagnose.longest_path(adj, comps) == 2   # {0,1,2} -> 3 -> 4
    assert diagnose.longest_path({}, diagnose.scc(3, {})) == 0
    print("graphs: ok")


def make_workbooks(tmp: Path) -> Path:
    wb = tmp / "wb"
    make_threeway_model.build(wb)
    pack = ROOT / "tests" / "sample_models" / "Riverbend_BP25_client_model.xlsx"
    if not pack.exists():
        try:
            import make_sample_models
            make_sample_models.main()
        except Exception as e:  # noqa: BLE001
            print(f"  (no Riverbend workbook: {type(e).__name__})")
    if pack.exists():
        shutil.copy(pack, wb / pack.name)
    return wb


def shapes_clean(rep_dir: Path, token: str, db: Path) -> None:
    """The written files hold none of the model's names: the scanner's own set, plus the names a reader would look for."""
    text = read(rep_dir / token / "report.json") + "\n" + read(rep_dir / token / "report.md")
    low = text.lower()
    ctx = diagnose.Ctx(db, db.parent.parent, type("A", (), {"depth": 6, "max_rows": 300, "quiet": True})())
    try:
        assert ctx.scanner.hits_many(text.splitlines()) == [], ctx.scanner.hits_many(text.splitlines())
    finally:
        ctx.close()
    for s in SHEETS:
        assert s not in text, f"sheet name {s} in the shapes of {token}"
        assert not re.search(rf"\b{s}!", text, re.I), f"a reference to {s} in {token}"
    for lab in LABELS + STEMS:
        assert lab not in low, f"{lab} in the shapes of {token}"
    assert str(Path.home()) not in text and str(ROOT) not in text and "/Users/" not in text


def planted(out: Path, tmp: Path) -> None:
    """A fixture copy with the client-text cases planted in model.db and the registry: the shapes report is written
    (no false block) and holds none of them."""
    src = next(out.glob("threeway_model__*"))
    pout = tmp / "out_planted"
    dst = pout / "Wombat_Kestrel_deal__aaaabbbb"
    shutil.copytree(src, dst)
    db = sqlite3.connect(dst / "model.db")

    def rename(old, new):
        for t, cols in (("sheets", ["sheet"]), ("cells", ["sheet"]), ("rows", ["sheet"]), ("edges", ["src_sheet", "dst_sheet"])):
            for col in cols:
                db.execute(f"UPDATE {t} SET {col}=? WHERE {col}=?", (new, old))
        db.execute("UPDATE rows SET patterns=replace(patterns, ?, ?)", (old + "!", "'" + new + "'!" if " " in new else new + "!"))
        db.execute("UPDATE cells SET formula=replace(formula, ?, ?) WHERE formula IS NOT NULL", (old + "!", "'" + new + "'!" if " " in new else new + "!"))
    rename("Capex", "SUMIF")
    rename("Costs", "Cons")
    rename("Equity", "Kestrel Ops")
    db.execute("UPDATE rows SET label='Revenue XQZ' WHERE sheet='Revenue' AND row=9")
    db.execute("UPDATE rows SET label='Wombat' WHERE sheet='Revenue' AND row=10")
    db.execute("UPDATE rows SET label=? WHERE sheet='Revenue' AND row=7", ("\u9577\u6c5f\u96fb\u529b",))
    db.execute("UPDATE rows SET section='Quokka Holdings' WHERE sheet='Debt'")
    db.execute("INSERT INTO names(name, ref) VALUES ('Val_date2', 'Inputs!$C$3'), ('TAX1', 'Inputs!$C$4'), ('Emu.Rate', 'Inputs!$C$5')")
    db.execute("UPDATE rows SET patterns=? WHERE sheet='Revenue' AND row=11",
               ('=SUM(R[-2]C:R[-1]C)*Val_date2+TAX1*Emu.Rate+INDIRECT("Platypus!A1")&"Acme Pty 2031"+[1]Koala!R1C1'
                "+Tbl_Bilby[Numbat]+WOMBATFN(R1C1)+_xludf.Echidna(1)+#BANDICOOT+45678+0.0731+IF(#REF!,#DIV/0!,1) x15 (D11..R11)",))
    # header cells that are also words of the report's own vocabulary (error literals, schema) do not block it
    for i, word in enumerate(["Ref", "Div", "Null", "Type", "KeyError"]):
        db.execute("INSERT INTO cells(sheet,row,col,addr,formula,value) VALUES ('Inputs',2,?,?,NULL,?)", (i + 1, f"X{i}", word))
    db.execute("INSERT INTO cells(sheet,row,col,addr,formula,value) VALUES ('Revenue',99,3,'C99','=CASSOWARY(1)',1)")
    db.execute("INSERT INTO cells(sheet,row,col,addr,formula,value) VALUES ('Inputs',1,1,'A1',NULL,'Lyrebird Infrastructure Fund')")
    db.commit()
    db.close()
    reg = sqlite3.connect(pout / "registry.db")
    reg.execute("CREATE TABLE files(id INTEGER PRIMARY KEY, sha256 TEXT, filename TEXT, out_dir TEXT, db_path TEXT, "
                "source_path TEXT, target_name TEXT, project_name TEXT, error TEXT)")
    reg.execute("INSERT INTO files(sha256, filename, out_dir, db_path, source_path, target_name, project_name, error) "
                "VALUES (?,?,?,?,?,?,?,?)",
                ("aaaabbbb" + "0" * 56, "Wombat_Kestrel_deal.xlsx", str(dst), str(dst / "model.db"),
                 "D:\\Deals\\Galah Partners\\Wombat_Kestrel_deal.xlsx", "Tasman Ports Ltd", "Project Wallaby",
                 "Traceback (most recent call last): KeyError: 'x' in openpyxl reader"))
    reg.commit()
    reg.close()
    rep = tmp / "diag_planted"
    stages = diagnose.STAGES

    def raises(ctx):   # an exception's type and site are code: a failed app build in the registry does not block them
        raise KeyError("x")
    diagnose.STAGES = [(n, raises if n == "anchors" else f) for n, f in stages]
    try:
        run([str(pout), "--report", str(rep), "--quiet"])
    finally:
        diagnose.STAGES = stages
    assert not (rep / "_blocked").exists(), read(next((rep / "_blocked").glob("*.txt")))
    assert json.loads(read(rep / "Maaaabbbb" / "report.json"))["stages"]["anchors"]["error"]["type"] == "KeyError"
    text = "\n".join(read(f) for f in rep.rglob("*") if f.is_file())
    text += json.dumps(json.loads(read(rep / "Maaaabbbb" / "report.json")), ensure_ascii=False)
    for w in ["SUMIF!", "Kestrel", "XQZ", "Wombat", "\u9577", "Quokka", "Val_date2", "TAX1", "Emu", "Platypus", "Acme", "2031",
              "Koala", "Bilby", "Numbat", "WOMBATFN", "Echidna", "BANDICOOT", "45678", "0.0731", "CASSOWARY", "Lyrebird",
              "Tasman", "Wallaby", "Galah"]:
        assert w.lower() not in text.lower(), f"planted {w} reached the shapes"
    # the forbidden set holds what the planted db has outside the label column
    ctx = diagnose.Ctx(dst / "model.db", pout, type("A", (), {"depth": 6, "max_rows": 300, "quiet": True})())
    try:
        for probe, hit in [("a Wombat", "wombat"), ("Quokka", "quokka"), ("Lyrebird", "lyrebird"), ("the XQZ row", "xqz"),
                           ("Tasman Ports Ltd", "tasman ports ltd"), ("project wallaby", "project wallaby"),
                           ("galah partners", "galah partners")]:
            assert hit in ctx.scanner.hits(probe), (probe, ctx.scanner.hits(probe))
    finally:
        ctx.close()
    # a registry that is there but cannot be read blocks (its names could not be scanned for)
    (pout / "registry.db").write_bytes(b"not a database at all" * 10)
    rep2 = tmp / "diag_planted2"
    run([str(pout), "--report", str(rep2), "--quiet"])
    assert (rep2 / "_blocked" / "Maaaabbbb.txt").exists() and not (rep2 / "Maaaabbbb" / "report.json").exists()
    # a model.db that cannot be opened is a recorded failure, and the run goes on
    bad = pout / "Broken_model__ccccdddd"
    bad.mkdir()
    (bad / "model.db").write_bytes(b"this is not sqlite" * 100)
    (pout / "registry.db").unlink()
    rep3 = tmp / "diag_planted3"
    run([str(pout), "--report", str(rep3), "--quiet"])
    s3 = json.loads(read(rep3 / "summary.json"))
    by = {m["token"]: m for m in s3["models"]}
    assert by["Mccccdddd"]["stages_failed"] == 1 and by["Maaaabbbb"]["stages_failed"] == 0, s3["models"]
    assert any(e["stage"] == "open" for e in s3["exceptions"]) and (rep3 / "Maaaabbbb" / "report.json").exists()
    assert "Broken" not in read(rep3 / "summary.md") and "Broken" not in read(rep3 / "summary.json")
    print("planted: ok")


def main() -> None:
    pattern_scrub()
    messages()
    gate_rules()
    graphs()
    tmp = Path(tempfile.mkdtemp(prefix="diagnose_"))
    wb = make_workbooks(tmp)
    out, rep = tmp / "out", tmp / "diag"
    log = run([str(out), "--workbooks", str(wb), "--report", str(rep), "--quiet"])
    dbs = sorted(out.glob("*/model.db"))
    assert len(dbs) >= 5, dbs
    tokens = {db.parent.name: diagnose.folder_token(db.parent.name) for db in dbs}
    # nothing the run prints names a folder, file or sheet
    for name in tokens:
        assert name.split("__")[0].lower() not in log.lower(), "the run printed a file name"

    # -- a report per model; the summary lists them
    summary = json.loads(read(rep / "summary.json"))
    md = read(rep / "summary.md")
    assert {m["token"] for m in summary["models"]} == set(tokens.values())
    assert all(not m["blocked"] for m in summary["models"]), summary["models"]
    for folder, tok in tokens.items():
        assert (rep / tok / "report.json").exists() and (rep / tok / "report.md").exists(), tok
        assert tok in md
        assert not (rep / "full").exists()
        shapes_clean(rep, tok, out / folder / "model.db")
    env = json.loads(read(rep / "environment.json"))
    assert env["python"] and "openpyxl" in env["packages"] and "git" in env
    import platform
    assert platform.node() not in read(rep / "environment.json") or len(platform.node()) < 4

    # -- identities, complexity, hints
    by_stem = {folder.split("__")[0]: tok for folder, tok in tokens.items()}
    base = json.loads(read(rep / by_stem["threeway_model"] / "report.json"))
    unb = json.loads(read(rep / by_stem["threeway_unbalanced"] / "report.json"))
    ids = {i["key"]: i for i in base["stages"]["statements"]["result"]["identities"]}
    assert ids["bs_balance"]["status"] == "holds" and ids["bs_balance"]["kind"] == "test"
    assert ids["pnl_ebitda"]["status"] == "holds" and ids["pnl_ebitda"]["kind"] == "structural"
    assert ids["bs_balance"]["residual"] in ("0", "<1e-9"), ids["bs_balance"]
    uid = {i["key"]: i for i in unb["stages"]["statements"]["result"]["identities"]}
    assert uid["bs_balance"]["status"] == "fails" and uid["bs_balance"]["kind"] == "test", uid["bs_balance"]
    assert uid["bs_balance"]["residual"] == ">=1e-1" and uid["bs_balance"]["check_row"] is True
    assert uid["bs_balance"]["failing_periods"] > 0 and uid["bs_balance"]["periods_checked"] == 15
    for st in unb["stages"].values():
        assert st["status"] in ("ok", "skipped"), st
    cx = unb["stages"]["complexity"]["result"]
    assert cx["families_listed"] > 0 and cx["families_per_formula_cell"] > 0 and cx["size_class"] == "small"
    sg = cx["sheet_graph"]
    assert sg["nodes"] == 11 and sg["edges"] > 0 and sg["longest_path"] >= 3 and sg["source_sheets"] >= 1 and sg["sink_sheets"] >= 1
    assert 0 < cx["row_graph"]["cone_share"] <= 1 and cx["row_graph"]["max_depth"] >= 3
    assert cx["dcf"]["anchors"] == 1 and cx["dcf"]["core_kinds"] == ["sumproduct"], cx["dcf"]
    assert any(f["fn"] in ("SUM", "SUMPRODUCT", "IF", "MAX") for f in cx["functions_top"])
    assert cx["timeline"]["periodicities"] == ["annual"] and cx["timeline"]["first_decades"] == ["2020s"]
    dg = unb["stages"]["depgraph"]["result"]
    assert dg["stats"]["nodes"] > 0 and dg["by_class"] and all(re.match(r"S\d+$", b["sheet"]) for b in dg["by_sheet"])
    assert {r["sheet"] for r in unb["sheets"]} == {f"S{i}" for i in range(1, 12)} and all(r["role"] for r in unb["sheets"])
    hint = next(h for h in unb["stages"]["hints"]["result"]["hints"] if h["kind"] == "identity" and h["key"] == "bs_balance")
    assert hint["status"] == "fails" and hint["check_row"] is True and hint["periodicity"] == ["annual"]
    assert len(hint["rows"]) >= 3 and all(re.match(r"S\d+$", r["sheet"]) for r in hint["rows"])
    pats = [p["r1c1"] for r in hint["rows"] for p in r["patterns"]]
    assert pats and all(p.startswith("=") for p in pats) and not any(s in p for p in pats for s in SHEETS), pats
    assert any("S" in p and "!" in p for p in pats) or any("R[" in p for p in pats)
    assert all(set(r["signs"].split()[0][:1]) <= set("+-0t_") for r in hint["rows"] if r["signs"])
    assert hint["rows"][0]["role"] and hint["rows"][0]["label_words"] >= 1 and hint["rows"][0]["row_offset"] is not None
    # a role is a vocabulary name and a word count, never the label
    assert not any("label" in r and isinstance(r.get("label"), str) for r in hint["rows"])
    moved = json.loads(read(rep / by_stem["threeway_moved"] / "report.json"))
    assert moved["stages"]["statements"]["result"]["blocks"], "moved variant has blocks"
    if "Riverbend_BP25_client_model" in by_stem:
        rb = json.loads(read(rep / by_stem["Riverbend_BP25_client_model"] / "report.json"))
        assert rb["stages"]["depgraph"]["status"] == "skipped", rb["stages"]["depgraph"]   # no DCF anchor in this one

    # -- the registry's names are part of the scan
    folder = next(f for f in tokens if f.startswith("threeway_model__"))
    reg = sqlite3.connect(out / "registry.db")
    reg.execute("""CREATE TABLE files(id INTEGER PRIMARY KEY, sha256 TEXT, filename TEXT, out_dir TEXT, db_path TEXT,
                   source_path TEXT, target_name TEXT, project_name TEXT)""")
    reg.execute("INSERT INTO files(sha256, filename, out_dir, db_path, target_name, project_name) VALUES (?,?,?,?,?,?)",
                (folder.split("__")[1] + "0" * 56, "Wombat Pty Ltd financials.xlsx", str(out / folder), str(out / folder / "model.db"),
                 "Wombat Pty Ltd", "Project Quokka"))
    reg.commit()
    reg.close()
    ctx = diagnose.Ctx(out / folder / "model.db", out, type("A", (), {"depth": 6, "max_rows": 300, "quiet": True})())
    try:
        assert ctx.registry and "project quokka" in ctx.scanner.hits("see project quokka")
        assert "wombat pty ltd" in ctx.scanner.hits("the Wombat Pty Ltd file")
    finally:
        ctx.close()
    (out / "registry.db").unlink()
    planted(out, tmp)

    # -- a leak is blocked (a stage emits a label)
    orig = diagnose.stage_complexity

    def leaky(ctx):
        r = orig(ctx)
        r["top_families"][0]["r1c1"] = "Total revenue"
        return r
    base_folder = next(f for f in tokens if f.startswith("threeway_model__"))
    leak_rep = tmp / "diag_leak"
    stages = diagnose.STAGES
    diagnose.STAGES = [(n, leaky if n == "complexity" else f) for n, f in stages]
    try:
        run([str(out), "--report", str(leak_rep), "--only", base_folder, "--quiet"])
    finally:
        diagnose.STAGES = stages
    tok = tokens[base_folder]
    note = leak_rep / "_blocked" / f"{tok}.txt"
    assert note.exists() and "total revenue" in read(note) and "LOCAL ONLY" in read(note)
    assert not (leak_rep / tok / "report.json").exists() and not (leak_rep / tok / "report.md").exists()
    s = json.loads(read(leak_rep / "summary.json"))
    assert s["models"][0]["blocked"] is True and "total revenue" not in read(leak_rep / "summary.md").lower()
    # a leak in a key is caught too, and an earlier run's report does not survive a block
    shutil.copytree(rep / tok, leak_rep / tok)
    diagnose.STAGES = [(n, (lambda c: {"Total revenue": 1}) if n == "census" else f) for n, f in stages]
    try:
        run([str(out), "--report", str(leak_rep), "--only", base_folder, "--quiet"])
    finally:
        diagnose.STAGES = stages
    assert not (leak_rep / tok / "report.json").exists(), "an old report outlived a block"

    # -- a stage that raises is recorded and the run goes on
    def broken(ctx):
        raise ValueError("no row for 'Total revenue' on Revenue at row 12 under " + str(Path.home()))
    fail_rep = tmp / "diag_fail"
    diagnose.STAGES = [(n, broken if n == "statements" else f) for n, f in stages]
    try:
        run([str(out), "--report", str(fail_rep), "--quiet"])
    finally:
        diagnose.STAGES = stages
    fs = json.loads(read(fail_rep / "summary.json"))
    assert len(fs["models"]) == len(dbs) and all(m["stages_failed"] == 1 for m in fs["models"]), fs["models"]
    assert fs["exceptions"] and fs["exceptions"][0]["count"] == len(dbs) and fs["exceptions"][0]["type"] == "ValueError"
    assert fs["exceptions"][0]["site"].startswith("tests/check_diagnose.py:") and "in broken" in fs["exceptions"][0]["site"]
    r = json.loads(read(fail_rep / tokens[base_folder] / "report.json"))
    e = r["stages"]["statements"]["error"]
    assert r["stages"]["statements"]["status"] == "failed" and r["stages"]["complexity"]["status"] == "ok"
    assert e["type"] == "ValueError" and "Revenue" not in e["message"] and "Total" not in e["message"], e["message"]
    assert str(Path.home()) not in json.dumps(e) and any("check_diagnose.py" in f and "/Users" not in f for f in e["traceback"])
    assert any(h["kind"] == "exception" and h["stage"] == "statements" for h in r["stages"]["hints"]["result"]["hints"])
    assert "Exceptions by site" in read(fail_rep / "summary.md") and "tests/check_diagnose.py" in read(fail_rep / "summary.md")
    shapes_clean(fail_rep, tokens[base_folder], out / base_folder / "model.db")

    # -- level full: real names under full/ only
    full_rep = tmp / "diag_full"
    printed = run([str(out), "--report", str(full_rep), "--level", "full", "--only", base_folder, "--quiet"])
    assert "DO NOT SHARE" in printed
    fr = full_rep / "full" / tok / "report.json"
    assert fr.exists() and "BalanceSheet" in read(fr) and "Total revenue" in read(fr) and "DO NOT SHARE" in read(fr)
    shapes_clean(full_rep, tok, out / base_folder / "model.db")
    assert "BalanceSheet" not in read(full_rep / tok / "report.json")

    # -- the test-suite runner on one fast script
    res = diagnose.run_tests(only={"check_edges.py"})
    assert len(res) == 1 and res[0]["status"] == "pass" and res[0]["script"] == "check_edges.py" and len(res[0]["tail"]) <= 15
    assert str(ROOT) not in json.dumps(res) and str(Path.home()) not in json.dumps(res)
    print("ok")


if __name__ == "__main__":
    main()
