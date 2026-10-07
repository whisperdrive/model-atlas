"""The library: workbooks dropped on the dashboard, and the worker that builds each into a model database.

    <OUT>/atlas.db                          the registry (table files); OUT is ATLAS_OUT or ./out, the folder the dashboard serves
    <ROOT>/uploads/<sha12>/<file name>      the saved upload (git-ignored)
    <OUT>/<stem>__<sha8>/                   its model database, built by build_map.main on one background thread

A workbook is identified by the SHA-256 of its bytes, so dropping the same file twice finds the first. No model calls
are made here: the build is build_map.py, as on the command line. Nothing in an existing database is overwritten:
every upload gets a folder named by its own hash.
"""
import gc
import logging
import os
import re
import shutil
import sqlite3
import threading
import time
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(os.environ.get("ATLAS_OUT", "out")).resolve()
UPLOADS = ROOT / "uploads"
EXTS = (".xlsx", ".xlsm")
SHA_RE = re.compile(r"^[0-9a-f]{64}$")
NAME_CAP = 120
FOLDER_STEM_CAP = 60   # out/<stem>__<sha8>/ stays well inside Windows' 260-character paths
RESERVED = re.compile(r"^(CON|PRN|AUX|NUL|COM[0-9\u00b9\u00b2\u00b3]|LPT[0-9\u00b9\u00b2\u00b3]|CONIN\$|CONOUT\$)$", re.I)
PROGRESS_EVERY = 0.25   # seconds between progress writes

log = logging.getLogger("modelatlas.library")
_LOCK = threading.RLock()   # every write to atlas.db, and the worker's start
_WAKE = threading.Event()
_thread: threading.Thread | None = None
_current: int | None = None   # the row the worker is building now
_ready: set[str] = set()      # registry files whose schema has been created in this process

COLS = ("id", "sha256", "filename", "size", "uploaded_at", "source_path", "status", "step", "pct", "error",
        "started_at", "processed_at", "out_dir", "sheets", "line_items", "build_secs")


class UploadRejected(ValueError):
    """An upload that will not be taken; .status is the HTTP status for it (400 or 413)."""

    def __init__(self, msg: str, status: int = 400):
        super().__init__(msg)
        self.status = status


class Busy(RuntimeError):
    """The workbook is being built right now."""


def max_bytes() -> int:
    try:
        return int(float(os.environ.get("ATLAS_MAX_UPLOAD_MB", "500")) * 1024 * 1024)
    except ValueError:
        return 500 * 1024 * 1024


def too_big() -> str:
    return f"That file is over the {max_bytes() / (1024 * 1024):.4g} MB limit"


# ---- the registry ------------------------------------------------------------------------------------------------

def db_path() -> Path:
    return OUT / "atlas.db"


def _connect() -> sqlite3.Connection:
    path = db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path, timeout=30, check_same_thread=False)
    db.row_factory = sqlite3.Row
    if str(path) not in _ready:
        with _LOCK:
            db.execute("PRAGMA journal_mode=WAL")
            db.execute("""CREATE TABLE IF NOT EXISTS files(
                id INTEGER PRIMARY KEY, sha256 TEXT UNIQUE, filename TEXT, size INT, uploaded_at REAL, source_path TEXT,
                status TEXT, step TEXT, pct REAL, error TEXT, started_at REAL, processed_at REAL, out_dir TEXT,
                sheets INT, line_items INT, build_secs REAL)""")
            db.commit()
            _ready.add(str(path))
    return db


def _one(sql: str, *args) -> dict | None:
    db = _connect()
    try:
        r = db.execute(sql, args).fetchone()
        return dict(r) if r else None
    finally:
        db.close()


def _write(sql: str, *args) -> int:
    with _LOCK:
        db = _connect()
        try:
            cur = db.execute(sql, args)
            db.commit()
            return cur.rowcount
        finally:
            db.close()


def all_files() -> list[dict]:
    """Every row, newest first, with the full paths (see public() for what goes to a browser)."""
    db = _connect()
    try:
        return [dict(r) for r in db.execute("SELECT * FROM files ORDER BY uploaded_at DESC, id DESC")]
    finally:
        db.close()


def get(fid: int) -> dict | None:
    return _one("SELECT * FROM files WHERE id=?", fid)


def by_sha(sha: str) -> dict | None:
    return _one("SELECT * FROM files WHERE sha256=?", sha)


def public(rec: dict | None) -> dict | None:
    """The row as a browser sees it: no path on this machine, the model folder by its name only."""
    if rec is None:
        return None
    out = {k: v for k, v in rec.items() if k not in ("source_path", "out_dir")}
    out["out_dir"] = Path(rec["out_dir"]).name if rec.get("out_dir") else None
    return out


# ---- names -------------------------------------------------------------------------------------------------------

def safe_name(name: str) -> str:
    """A file name that is only a file name: no folders, no control characters, no characters Windows refuses, capped."""
    name = str(name or "").replace("\\", "/").split("/")[-1]
    name = re.sub(r"[\x00-\x1f\x7f<>:\"|?*]", "", name)
    name = "".join(c for c in name if unicodedata.category(c) not in ("Cc", "Cf", "Cs", "Co", "Cn")).strip(" .")
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    ext = ("." + ext) if ext else ""
    if len(stem) + len(ext) > NAME_CAP:
        stem = stem[:max(1, NAME_CAP - len(ext))]
    if not stem.strip():
        return "workbook" + ext
    if RESERVED.match(stem.split(".")[0].rstrip(" ")):   # CON.xlsx, nul.xlsm, COM1.v2.xlsx: Windows devices, not files
        stem = "_" + stem
    return stem + ext


def check_name(name: str) -> str:
    """The safe name, or UploadRejected for anything that is not .xlsx/.xlsm."""
    safe = safe_name(name)
    if not safe.lower().endswith(EXTS):
        raise UploadRejected("Only .xlsx and .xlsm workbooks can be added (save a copy from Excel first)")
    return safe


def out_folder(rec: dict) -> Path:
    stem = re.sub(r"[^\w.\- ]+", "_", Path(rec["filename"]).stem)[:FOLDER_STEM_CAP].strip(" ._") or "workbook"
    return OUT / f"{stem}__{rec['sha256'][:8]}"


# ---- adding, removing --------------------------------------------------------------------------------------------

def add_upload(tmp_path, filename: str, sha256: str) -> tuple[str, dict]:
    """Take a finished temp file: ("exists", row) when its hash is known (the temp file is deleted), else it is moved
    to uploads/<sha12>/<safe name>, a queued row is inserted and the worker is woken: ("queued", row)."""
    tmp = Path(tmp_path)
    try:
        sha = str(sha256).lower()
        if not SHA_RE.match(sha):
            raise UploadRejected("Bad fingerprint")
        safe = check_name(filename)
        size = tmp.stat().st_size
        if size > max_bytes():
            raise UploadRejected(too_big(), 413)
        if size == 0:
            raise UploadRejected("That file is empty")
        with _LOCK:
            known = by_sha(sha)
            if known:
                tmp.unlink(missing_ok=True)
                return "exists", known
            dest = UPLOADS / sha[:12] / safe
            dest.parent.mkdir(parents=True, exist_ok=True)
            os.replace(tmp, dest)
            db = _connect()
            try:
                cur = db.execute("""INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status, step, pct)
                                    VALUES (?,?,?,?,?,?,?,?)""",
                                 (sha, safe, size, time.time(), str(dest), "queued", "Waiting for the worker", 0.0))
                db.commit()
                fid = cur.lastrowid
            except sqlite3.IntegrityError:   # another process registered the same bytes a moment ago
                known = by_sha(sha)
                if not known:
                    raise
                if Path(known.get("source_path") or "") != dest:
                    dest.unlink(missing_ok=True)
                return "exists", known
            finally:
                db.close()
        _ensure()
        _WAKE.set()
        return "queued", get(fid)
    except BaseException:
        tmp.unlink(missing_ok=True)   # a rejected upload leaves nothing behind
        raise


def rebuild(fid: int) -> dict | None:
    """Queue the build again (the model folder is cleared first). Refused while it is being built."""
    with _LOCK:
        rec = get(fid)
        if not rec:
            return None
        if rec["status"] == "building":
            raise Busy("It is being built right now")
        _rmtree(_owned_folder(rec), must=True)
        _write("""UPDATE files SET status='queued', step='Waiting for the worker', pct=0, error=NULL, started_at=NULL,
                  processed_at=NULL, out_dir=NULL, sheets=NULL, line_items=NULL, build_secs=NULL WHERE id=?""", fid)
        rec = get(fid)   # read before the worker is woken: the answer is 'queued', not whatever it has got to
    _ensure()
    _WAKE.set()
    return rec


def remove(fid: int) -> bool:
    """Delete the row, the saved upload and the model folder. Refused while it is being built."""
    with _LOCK:
        rec = get(fid)
        if not rec:
            return False
        if rec["status"] == "building":
            raise Busy("It is being built right now")
        _rmtree(_owned_folder(rec), must=True)
        if rec.get("source_path"):
            _rmtree(Path(rec["source_path"]).parent, under=UPLOADS, must=True)
        _write("DELETE FROM files WHERE id=?", fid)
    return True


def _owned_folder(rec: dict) -> Path | None:
    """The folder this row built into, only if it is one of ours: recorded on the row (a row that never started a build
    owns none, so a folder of the same name it has not adopted is left alone), directly under OUT, named
    <stem>__<its sha8>."""
    if not rec.get("out_dir"):
        return None
    p = Path(rec["out_dir"])
    return p if p.parent == OUT and p.name.endswith("__" + rec["sha256"][:8]) else None


def _rmtree(p: Path | None, under: Path | None = None, must: bool = False) -> None:
    """Delete a folder (only one directly under `under`, when given). must: a folder that is still there afterwards
    (on Windows, a file another program, or a connection not yet collected, holds open) raises Busy."""
    if p is None or not p.exists():
        return
    if under is not None and p.parent != under:
        return
    gc.collect()   # a sqlite connection or openpyxl zip only waiting for the collector still holds its file on Windows
    shutil.rmtree(p, ignore_errors=True)
    if must and p.exists():
        raise Busy("Some of its files are in use (open in another program, or being read); try again in a moment")


# ---- the worker --------------------------------------------------------------------------------------------------

def _ensure() -> None:
    if _thread is None or not _thread.is_alive():
        start()


def start() -> None:
    """Start the one worker thread (once), and put any row left 'building' by a crash back to queued."""
    global _thread
    with _LOCK:
        _write("""UPDATE files SET status='queued', step='Waiting for the worker', pct=0 WHERE status='building' AND id IS NOT ?""",
               _current)
        if _thread is None or not _thread.is_alive():
            _thread = threading.Thread(target=_loop, name="atlas-library-worker", daemon=True)
            _thread.start()
    _WAKE.set()


def _loop() -> None:
    global _current
    while True:
        try:
            rec = _claim()
        except Exception:   # noqa: BLE001 (a locked or unreadable registry: try again, never end the thread)
            log.exception("library: the registry could not be read")
            time.sleep(2)
            continue
        if rec is None:
            _WAKE.wait(2)
            _WAKE.clear()
            continue
        _current = rec["id"]
        try:
            _build(rec)
        except Exception as e:   # noqa: BLE001 (the worker must outlive any one workbook)
            log.exception("library: build of row %s failed", rec["id"])
            try:
                _fail(rec["id"], e)
            except Exception:   # noqa: BLE001 (the registry is locked: the row stays 'building' until the next start)
                log.exception("library: the failure of row %s could not be recorded", rec["id"])
        finally:
            _current = None
            gc.collect()   # the build's openpyxl and sqlite handles on the upload and model.db, released now


def _claim() -> dict | None:
    with _LOCK:
        rec = _one("SELECT * FROM files WHERE status='queued' ORDER BY id LIMIT 1")
        if rec and not _write("UPDATE files SET status='building', step='Starting', pct=0, error=NULL, started_at=? "
                              "WHERE id=? AND status='queued'", time.time(), rec["id"]):
            return None   # another process took it first (one server per out/ is the design; this keeps two honest)
    return rec


def _short(e: BaseException) -> str:
    """One line for the page: what kind of failure, never a traceback or a path."""
    msg = str(e).replace("\n", " ").strip()
    msg = re.sub(r"(['\"])[^'\"]*[\\/][^'\"]*\1", "<path>", msg)   # a quoted path, spaces and all
    msg = re.sub(r"(?:[A-Za-z]:[\\/]|\\\\|(?<![\w.])/)[^\s'\"]+", "<path>", msg)[:200]
    kind = type(e).__name__
    if "zip" in kind.lower() or "BadZipFile" in kind:
        return "Not a readable Excel workbook (the file is not a valid .xlsx/.xlsm)"
    return f"{kind}: {msg}" if msg else kind


def _fail(fid: int, e: BaseException) -> None:
    _write("UPDATE files SET status='error', step=NULL, error=?, processed_at=? WHERE id=?", _short(e), time.time(), fid)


def _counts(model_db: Path) -> tuple[int, int]:
    from . import rodb
    db = rodb.connect(model_db, timeout=10)
    try:
        return (db.execute("SELECT COUNT(*) FROM sheets").fetchone()[0], db.execute("SELECT COUNT(*) FROM rows").fetchone()[0])
    finally:
        db.close()


def _build(rec: dict) -> None:
    from . import build_map
    fid, t0 = rec["id"], time.time()
    folder = out_folder(rec)
    other = [f for f in all_files() if f["id"] != fid and f.get("out_dir") and Path(f["out_dir"]) == folder]
    if other:
        raise RuntimeError("The model folder belongs to another workbook")
    src = Path(rec["source_path"] or "")
    if not src.is_file():
        raise FileNotFoundError("The saved upload is missing")
    complete = (folder / "model.db").is_file() and (folder / "map.txt").is_file()   # map.txt is written last
    _write("UPDATE files SET out_dir=? WHERE id=?", str(folder), fid)
    if folder.exists() and not complete:   # a build that died half way: start clean
        shutil.rmtree(folder)
        if folder.exists():
            raise RuntimeError("An unfinished model folder could not be cleared")
    last = [0.0]

    def progress(frac: float, msg: str) -> None:
        now = time.time()
        if now - last[0] < PROGRESS_EVERY and frac < 1:
            return
        last[0] = now
        _write("UPDATE files SET pct=?, step=? WHERE id=?", round(max(0.0, min(1.0, frac)) * 100, 1), str(msg)[:200], fid)

    if complete:   # the same bytes were built here already (its name carries the hash): take it as it is
        progress(1.0, "Model database already built")
    else:
        build_map.main(str(src), str(folder), progress=progress)
    sheets, items = _counts(folder / "model.db")
    _write("""UPDATE files SET status='done', step=NULL, pct=100, error=NULL, sheets=?, line_items=?, build_secs=?,
              processed_at=? WHERE id=?""", sheets, items, round(time.time() - t0, 1), time.time(), fid)
