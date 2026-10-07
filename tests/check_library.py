"""Ingestion (modelatlas/library.py): a dropped workbook is registered, built on the worker thread and removed again.

  add and build   a workbook becomes a queued row, then done, with its model.db, sheets and line items filled
  same bytes      a second add of the same fingerprint is "exists" and leaves no temp file
  bad file        junk bytes saved as .xlsx end in status error with a short message, and the next file still builds
  rejected        a name that is not .xlsx/.xlsm, an empty file, a malformed fingerprint
  crash           a row left 'building' goes back to queued on start() and is built
  remove          deletes the row, the upload and the model folder; refused while building
  names           folders, control characters, Windows device names and over-long names are cut to a file name
  not ours        Remove of a row that never built leaves a folder of the same name it did not adopt
  self-scan       the diagnostics of an uploaded model scan for the upload's own file name (out/atlas.db)

    uv run python tests/check_library.py
"""
import hashlib
import logging
import shutil
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import xlsxwriter  # noqa: E402

from modelatlas import library  # noqa: E402


def workbook(path: Path, n: int) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Calc")
    ws.write(0, 1, "Revenue")
    for c in range(3, 8):
        ws.write_number(0, c, 100 * n + c)
        ws.write_formula(1, c, f"={'ABCDEFGH'[c]}1*2", None, 200 * n + 2 * c)
    ws.write(1, 1, "Doubled")
    wb.close()
    return path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def wait(fid: int, want=("done", "error"), secs: float = 15.0) -> dict:
    t0 = time.time()
    while time.time() - t0 < secs:
        r = library.get(fid)
        if r and r["status"] in want:
            return r
        time.sleep(0.1)
    raise AssertionError(f"row {fid} did not reach {want}: {library.get(fid)}")


def raises(fn, exc, status=None):
    try:
        fn()
    except exc as e:
        if status is not None:
            assert e.status == status, (e.status, status)
        return
    raise AssertionError("did not raise")


def add(tmp: Path, src: Path, name: str | None = None):
    t = tmp / f"t{time.time_ns()}"
    shutil.copy(src, t)
    return library.add_upload(t, name or src.name, sha(src)), t


def main() -> None:
    logging.getLogger("modelatlas.library").setLevel(logging.CRITICAL)   # the bad file's traceback is expected
    tmp = Path(tempfile.mkdtemp(prefix="library_"))
    library.OUT = tmp / "out"
    library.ROOT = tmp
    library.UPLOADS = tmp / "uploads"
    library.OUT.mkdir()
    try:
        w1, w2 = workbook(tmp / "one.xlsx", 1), workbook(tmp / "two.xlsm", 2)

        # add and build
        (status, rec), t1 = add(tmp, w1)
        assert status == "queued" and rec["status"] in ("queued", "building"), rec
        assert not t1.exists(), "the temp file is moved, not copied"
        up = library.UPLOADS / sha(w1)[:12] / "one.xlsx"
        assert up.is_file(), up
        done = wait(rec["id"])
        assert done["status"] == "done", done
        out = Path(done["out_dir"])
        assert out == library.OUT / f"one__{sha(w1)[:8]}", out
        assert (out / "model.db").is_file() and done["sheets"] == 1 and done["line_items"] >= 2, done
        assert done["build_secs"] is not None and done["pct"] == 100 and done["error"] is None, done
        pub = library.public(done)
        assert "source_path" not in pub and pub["out_dir"] == out.name, pub

        # same bytes
        (status, again), t = add(tmp, w1, "renamed.xlsx")
        assert status == "exists" and again["id"] == rec["id"] and not t.exists(), (status, again)
        assert len(library.all_files()) == 1 and library.by_sha(sha(w1))["id"] == rec["id"]

        # a corrupt workbook, then a good one: the worker keeps serving
        junk = tmp / "junk.xlsx"
        junk.write_bytes(b"this is not a zip file at all" * 20)
        (status, bad), _ = add(tmp, junk)
        (_, good), _ = add(tmp, w2)
        b = wait(bad["id"])
        assert b["status"] == "error" and b["error"] and "\n" not in b["error"] and str(tmp) not in b["error"], b
        for e in (FileNotFoundError(2, "No such file", "/Users/me/Deals/Quokka Holdings/model.xlsx"),
                  OSError("cannot open C:\\Deals\\Quokka\\model.xlsx now"), OSError("share \\\\srv\\Quokka\\m.xlsx"),
                  ValueError("bad value at /home/me/Quokka/x.xlsx")):
            short = library._short(e)
            assert "Quokka" not in short and "Deals" not in short and "<path>" in short and "\n" not in short, short
        g = wait(good["id"])
        assert g["status"] == "done" and (Path(g["out_dir"]) / "model.db").is_file(), g

        # rejected
        raises(lambda: add(tmp, w1, "x.xls"), library.UploadRejected, 400)
        raises(lambda: add(tmp, w1, "x.xlsx.exe"), library.UploadRejected, 400)
        empty = tmp / "empty.xlsx"
        empty.write_bytes(b"")
        raises(lambda: library.add_upload(empty, "empty.xlsx", "a" * 64), library.UploadRejected, 400)
        raises(lambda: library.add_upload(shutil.copy(w1, tmp / "t_bad"), "x.xlsx", "XYZ"), library.UploadRejected, 400)
        assert not (tmp / "t_bad").exists(), "a rejected temp file is deleted"
        old = library.max_bytes
        library.max_bytes = lambda: 10
        try:
            raises(lambda: add(tmp, w1, "big.xlsx"), library.UploadRejected, 413)
        finally:
            library.max_bytes = old
        assert not list(library.UPLOADS.glob(".tmp*"))

        # a build left running by a crash is queued again by start() and built
        crashed = tmp / "crash.xlsx"
        workbook(crashed, 3)
        library._write("INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status, step, pct, started_at) "
                       "VALUES (?,?,?,?,?,?,?,?,?)", sha(crashed), "crash.xlsx", crashed.stat().st_size, time.time(),
                       str(crashed), "building", "Reading sheet 1", 40.0, time.time())
        cid = library.by_sha(sha(crashed))["id"]
        library.start()
        c = wait(cid)
        assert c["status"] == "done" and c["sheets"] == 1, c
        # and an unfinished folder from that crash is cleared, not trusted
        half = library.out_folder(c)
        library.remove(cid)
        assert not half.exists()
        half.mkdir()
        (half / "model.db").write_bytes(b"half")
        library._write("INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status) VALUES (?,?,?,?,?,?)",
                       sha(crashed), "crash.xlsx", 1, time.time(), str(crashed), "queued")
        cid = library.by_sha(sha(crashed))["id"]
        library._WAKE.set()
        c = wait(cid)
        assert c["status"] == "done" and c["sheets"] == 1, c

        # rebuild, then remove
        with library._LOCK:   # the worker cannot claim the row until the folder has been looked at
            r = library.rebuild(rec["id"])
            assert r["status"] == "queued" and not out.exists(), r
        again = wait(rec["id"])
        assert again["status"] == "done" and (out / "model.db").is_file(), again
        library._write("UPDATE files SET status='building' WHERE id=?", rec["id"])
        raises(lambda: library.remove(rec["id"]), library.Busy)
        raises(lambda: library.rebuild(rec["id"]), library.Busy)
        library._write("UPDATE files SET status='done' WHERE id=?", rec["id"])
        assert library.remove(rec["id"]) and library.get(rec["id"]) is None
        assert not out.exists() and not up.exists() and not up.parent.exists()
        assert library.remove(rec["id"]) is False
        assert (Path(g["out_dir"]) / "model.db").is_file(), "removing one workbook leaves the others"

        # another process registered the same bytes between the check and the insert: "exists", the moved file removed
        w6 = workbook(tmp / "six.xlsx", 6)
        library._write("INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status) VALUES (?,?,?,?,?,?)",
                       sha(w6), "six.xlsx", 1, time.time(), str(tmp / "elsewhere.xlsx"), "error")
        real, calls = library.by_sha, []
        library.by_sha = lambda h: None if not calls.append(h) and len(calls) == 1 else real(h)
        try:
            (status, raced), t6 = add(tmp, w6, "six.xlsx")
        finally:
            library.by_sha = real
        assert status == "exists" and raced["sha256"] == sha(w6) and not t6.exists(), (status, raced)
        assert not (library.UPLOADS / sha(w6)[:12] / "six.xlsx").exists(), "the losing copy is not left in uploads/"
        library.remove(raced["id"])

        # names
        s = library.safe_name
        assert s("../../x.xlsx") == "x.xlsx" and s("a/b\\c.xlsx") == "c.xlsx" and s("..\\..\\win.xlsx") == "win.xlsx"
        assert s("a\x00b\x1f\x7f.xlsx") == "ab.xlsx" and s('q"u:e?.xlsx') == "que.xlsx", s('q"u:e?.xlsx')
        long = s("n" * 500 + ".xlsx")
        assert len(long) == library.NAME_CAP and long.endswith(".xlsx"), len(long)
        assert s("") == "workbook" and s("..") == "workbook" and s(None) == "workbook"
        for dev in ("CON.xlsx", "nul.xlsm", "COM1.xlsx", "lpt9.xlsx", "AUX .xlsx", "Con.v2.xlsx"):
            assert s(dev).startswith("_"), (dev, s(dev))
        assert s("CONTRACT.xlsx") == "CONTRACT.xlsx" and s("console.xlsx") == "console.xlsx"
        assert s("Prévision 2025 – 東京.xlsx") == "Prévision 2025 – 東京.xlsx", s("Prévision 2025 – 東京.xlsx")
        assert s("evil\u202excod.xlsx") == "evilxcod.xlsx" and s("a\u200bb.xlsx") == "ab.xlsx", s("evil\u202excod.xlsx")
        assert len(library.out_folder({"filename": "x" * 200 + ".xlsx", "sha256": "0" * 64}).name) <= library.FOLDER_STEM_CAP + 10

        # not ours: a queued row has no folder yet, so Remove leaves a same-named folder alone (and its upload goes)
        w4 = workbook(tmp / "four.xlsx", 4)
        theirs = library.OUT / f"four__{sha(w4)[:8]}"
        theirs.mkdir()
        (theirs / "keep.txt").write_text("not the library's")
        library._write("INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status) VALUES (?,?,?,?,?,?)",
                       sha(w4), "four.xlsx", 1, time.time(), str(w4), "error")
        assert library.remove(library.by_sha(sha(w4))["id"]) and (theirs / "keep.txt").is_file(), "a folder it never built is kept"
        shutil.rmtree(theirs)

        # self-scan: the uploaded file's own name is part of what diagnose looks for, read from out/atlas.db; the folder
        # name is cut at FOLDER_STEM_CAP, so a name at the end of a long file name is only in atlas.db
        from modelatlas import diagnose
        client = "Project valuation model final version for the investment committee review Quokka Holdings.xlsx"
        w5 = workbook(tmp / "five.xlsx", 5)
        (status, q), _ = add(tmp, w5, client)
        q = wait(q["id"])
        assert q["status"] == "done", q
        folder = Path(q["out_dir"])
        assert "quokka" not in folder.name.lower(), folder.name
        ctx = diagnose.Ctx(folder / "model.db", library.OUT, None)
        try:
            assert any(r.get("filename") == client for r in ctx.registry), ctx.registry
            assert not ctx.registry_unreadable
            assert "quokka" in ctx.scanner.hits("quokka") and ctx.scanner.hits("Quokka Holdings"), ctx.scanner.hits("Quokka Holdings")
            assert not any("python" in h or "uploads" in h for h in ctx.scanner.hits("python uploads")), "only the file name, not our folders"
        finally:
            ctx.close()
        args = type("A", (), {"report": str(tmp / "diag"), "level": "shapes", "depth": 6, "max_rows": 300, "tests": False, "only": None})()
        row, _, _ = diagnose.process(folder / "model.db", library.OUT, args, None)
        assert not row.get("blocked"), row
        for f in ("report.md", "report.json"):
            text = (tmp / "diag" / row["token"] / f).read_text(encoding="utf-8").lower()
            assert "quokka" not in text and "holdings" not in text, f
        print("check_library: ok")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
