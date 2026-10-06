"""Row-level edges (modelatlas/edges.py) into sheets whose names have a space or an apostrophe.

  quoted names   'Disc rates'!D12 and 'It''s'!B3 keep their sheet name (spaces are dropped only outside quotes), so the
                 edge lands on the row it reads; SUMIFS over 'Disc rates'!$5:$5 / an OFFSET based on it / an INDEX-MATCH
                 into it give their kinds too
  no ghosts      no edge goes to a sheet that doesn't exist

    uv run python tests/check_edges.py
"""
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import xlsxwriter  # noqa: E402

from modelatlas import build_map  # noqa: E402
from modelatlas import edges  # noqa: E402


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="edges_"))
    wb = xlsxwriter.Workbook(out / "e.xlsx")
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 1, "Rate")
    inp.write_number(3, 3, 0.08)
    dr = wb.add_worksheet("Disc rates")
    for r in range(3, 14):  # rows 4..14, labels in B, numbers in D:F
        dr.write(r, 1, f"Rate {r + 1}")
        for c in range(3, 6):
            dr.write_number(r, c, 0.01 * (r + 1) + c)
    dr.write(4, 1, "Key row")  # row 5 is the one SUMIFS / OFFSET / INDEX read
    for c, k in zip(range(3, 6), (1, 2, 3)):
        dr.write_number(4, c, k)
    its = wb.add_worksheet("It's")
    for r in range(2, 5):
        its.write(r, 1, f"Item {r + 1}")
        its.write_number(r, 3, 10.0 * r)
    fl = wb.add_worksheet("Flows")
    fl.write(2, 1, "Period")
    for c, k in zip(range(3, 6), (1, 2, 3)):
        fl.write_number(2, c, k)
    rows = [
        ("Direct quoted", "='Disc rates'!D12", 0.01 * 12 + 3),
        ("Escaped quote", "='It''s'!B3", 0),
        ("Spaced ref", "= 'Disc rates'!D12 + Inputs!D4", 0.01 * 12 + 3 + 0.08),
        ("Sumifs", "=SUMIFS('Disc rates'!D12:D14,'Disc rates'!$5:$5,D3)", 0),
        ("Offset", "=OFFSET('Disc rates'!D5,7,0)", 0.01 * 12 + 3),
        ("Index match", "=INDEX('Disc rates'!D12:D14,MATCH(D3,'Disc rates'!D5:F5,0))", 0),
    ]
    for i, (label, f, v) in enumerate(rows):
        fl.write(4 + i, 1, label)
        fl.write_formula(4 + i, 3, f, None, v)
        fl.write_formula(4 + i, 4, f.replace("D3", "E3").replace("D12", "E12").replace("D5", "E5"), None, v)
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "e.xlsx"), str(out / "db"))["db"])
    label_row = {(s, lab): r for s, r, lab in db.execute("SELECT sheet, row, label FROM rows")}
    got = {}
    for ss, sr, ds, dr_, k in db.execute("SELECT * FROM edges"):
        got.setdefault(next(l for (s, l), r in label_row.items() if s == ss and r == sr), []).append((ds, dr_, k))
    sheets = {s for (s,) in db.execute("SELECT sheet FROM sheets")}
    assert {"Disc rates", "It's"} <= sheets, sheets
    ghosts = [e for v in got.values() for e in v if e[0] not in sheets]
    assert not ghosts, ghosts

    def has(label, sheet, row, kind=None):
        return any(d[0] == sheet and d[1] == row and (kind is None or d[2] == kind) for d in got.get(label, []))

    assert has("Direct quoted", "Disc rates", 12, "direct"), got
    assert has("Escaped quote", "It's", 3, "direct"), got
    assert has("Spaced ref", "Disc rates", 12, "direct") and has("Spaced ref", "Inputs", 4, "direct"), got
    assert any(d[0] == "Disc rates" for d in got.get("Sumifs", [])), got
    assert has("Offset", "Disc rates", 12, "offset"), got
    assert has("Index match", "Disc rates", 12) and has("Index match", "Disc rates", 5), got
    # the unit pieces
    m = edges.Model(db)
    r = m.ref("'It''s'!$B$3", "Flows")
    assert (r.sheet, r.r1, r.c1) == ("It's", 3, 2), vars_of(r)
    r = m.ref("'Disc rates'!$5:$5", "Flows")
    assert (r.sheet, r.r1, r.r2) == ("Disc rates", 5, 5)
    assert edges._unspace("'A b'!C3 + 'It''s x'!D4 + E5") == "'A b'!C3+'It''s x'!D4+E5"
    print("edges: ok", {k: len(v) for k, v in got.items()})


def vars_of(r):
    return (r.sheet, r.r1, r.c1, r.r2, r.c2)


if __name__ == "__main__":
    main()
