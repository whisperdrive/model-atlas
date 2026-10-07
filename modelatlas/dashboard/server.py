"""Model dashboard: a read-only view over out/registry.db and each processed workbook's model.db.
    uv run atlas-dashboard                               then open http://localhost:8001
    uv run uvicorn modelatlas.dashboard.server:app --port 8001      the same, with uvicorn's own options

Reads the out/ folder of the current directory (or the folder named by ATLAS_OUT): every out/<name>/model.db is listed
as a command-line build; out/registry.db is read too if there is one, but nothing here needs it. Every database is
opened with mode=ro, so this never takes a write lock or creates tables.
"""
import json
import os
import re
import sys
import sqlite3
import threading
from collections import OrderedDict
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, Query
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from .. import NOTICE, __version__, depgraph, depgraph_html, diagnose, ontology, rodb, statements, version_line

ROOT = Path(__file__).resolve().parents[2]  # the project root
OUT = Path(os.environ.get("ATLAS_OUT", "out")).resolve()  # ./out under the current directory, or ATLAS_OUT
REGISTRY = OUT / "registry.db"
LIVE = ("direct", "offset", "active")  # edge kinds the current scenario actually uses (see modelatlas/edges.py)

app = FastAPI()


@app.exception_handler(sqlite3.Error)
def _db_error(_request, exc: sqlite3.Error):
    """A model.db that is locked, being rebuilt or not a database: a short answer, never a traceback."""
    return JSONResponse({"detail": f"The model database could not be read ({type(exc).__name__}: {exc})"}, status_code=503)


def _ro(path: Path | str) -> sqlite3.Connection:
    db = rodb.connect(path, timeout=5, check_same_thread=False)   # a quoted URI: a folder name may hold '#' or '%'
    db.row_factory = sqlite3.Row
    return db


def _q(db: sqlite3.Connection, sql: str, *args) -> list[dict]:
    return [dict(r) for r in db.execute(sql, args)]


def _tables(db: sqlite3.Connection) -> set[str]:
    return {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def _json(s):
    try:
        return json.loads(s) if s else None
    except ValueError:
        return None


@app.get("/")
def index():
    return FileResponse(Path(__file__).parent / "index.html")


@app.get("/api/version")
def version():
    return {"name": "Model Atlas", "version": __version__, "notice": NOTICE}


@app.get("/api/portfolio")
def portfolio():
    files, usage = [], None
    if REGISTRY.exists():
        files, usage = _registry_rows()
    return {"files": files + _command_line_dbs(files), "usage": usage}


def _registry_rows():
    with _ro(REGISTRY) as db:
        have = _tables(db)
        files = _q(db, """SELECT id, filename, size, uploaded_at, status, step, error, processed_at, sheets, line_items,
                          build_secs, target_name, project_name, valuation_date, identity_confirmed, previous_id,
                          diff_summary, db_path FROM files ORDER BY uploaded_at DESC""") if "files" in have else []
        usage = None
        if "usage" in have:
            sums = "COUNT(*) AS calls, COALESCE(SUM(input_tokens),0) AS input_tokens, " \
                   "COALESCE(SUM(output_tokens),0) AS output_tokens, COALESCE(SUM(cost_usd),0) AS cost_usd"
            usage = {
                "total": _q(db, f"SELECT {sums}, MIN(ts) AS since, MAX(ts) AS last FROM usage")[0],
                "by_day_purpose": _q(db, f"""SELECT date(ts,'unixepoch','localtime') AS day, purpose, {sums}
                                             FROM usage GROUP BY day, purpose ORDER BY day"""),
                "by_model": _q(db, f"SELECT model, {sums} FROM usage GROUP BY model ORDER BY cost_usd DESC"),
                "by_file": _q(db, f"SELECT file_id, {sums} FROM usage WHERE file_id IS NOT NULL GROUP BY file_id"),
            }
    for f in files:
        f["db_ok"] = bool(f["db_path"]) and Path(f["db_path"]).exists()
        f["source"] = "upload"
        f["_path"] = _resolved(f.pop("db_path"))
    return files, usage


def _resolved(p) -> Path | None:
    try:
        return Path(p).resolve() if p else None
    except OSError:
        return None


def _command_line_dbs(registered: list[dict]) -> list[dict]:
    """out/*/model.db that no registry row points at (built by modelatlas/build_map.py or modelatlas/diagnose.py)."""
    known = {f.pop("_path") for f in registered}
    found = []
    for path in OUT.glob("*/model.db"):
        if path.resolve() in known:
            continue
        try:
            st = path.stat()
            with _ro(path) as db:
                sheets = db.execute("SELECT COUNT(*) FROM sheets").fetchone()[0]
                items = db.execute("SELECT COUNT(*) FROM rows").fetchone()[0]
        except (sqlite3.Error, OSError):
            continue
        name = path.parent.name
        found.append({"id": "dir:" + name, "filename": re.sub(r"__[0-9a-f]{8}$", "", name), "size": st.st_size,
                      "uploaded_at": st.st_mtime, "status": "built", "step": None, "error": None,
                      "processed_at": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                      "sheets": sheets, "line_items": items, "build_secs": None, "target_name": None,
                      "project_name": None, "valuation_date": None, "identity_confirmed": None, "previous_id": None,
                      "diff_summary": None, "db_ok": True, "source": "command line"})
    return sorted(found, key=lambda f: -f["uploaded_at"])


def _model_db(fid: str) -> tuple[dict, Path]:
    """fid is a registry row id ("7") or a database built from the command line ("dir:<folder under out/>")."""
    if fid.startswith("dir:"):
        name = fid[4:]
        if not name or name in (".", "..") or "/" in name or "\\" in name or ".." in name or "\0" in name:
            raise HTTPException(400, "Bad folder name")
        path = OUT / name / "model.db"
        if not path.is_file():
            raise HTTPException(404, "No such workbook")
        return {"id": fid, "filename": re.sub(r"__[0-9a-f]{8}$", "", name), "status": "built", "source": "command line",
                "target_name": None, "project_name": None, "valuation_date": None}, path
    if not fid.isdigit():
        raise HTTPException(400, "Bad workbook id")
    fid = int(fid)
    if not REGISTRY.exists():
        raise HTTPException(404, "No registry yet")
    with _ro(REGISTRY) as db:
        rec = db.execute("SELECT id, filename, status, target_name, project_name, valuation_date, db_path "
                         "FROM files WHERE id=?", (fid,)).fetchone()
    if not rec:
        raise HTTPException(404, "No such workbook")
    rec = dict(rec)
    rec["source"] = "upload"
    path = Path(rec.pop("db_path") or "")
    if not rec["status"] == "done" or not path.is_file():
        raise HTTPException(409, f"{rec['filename']} has no model database yet (status: {rec['status']})")
    return rec, path


@app.get("/api/workbook/{fid}")
def workbook(fid: str):
    rec, path = _model_db(fid)
    with _ro(path) as db:
        have = _tables(db)
        sheets = _q(db, "SELECT rowid AS pos, sheet, state, layout, summary FROM sheets ORDER BY rowid")
        for s in sheets:
            s["layout"] = _json(s["layout"]) or {}
        per_sheet = {r["sheet"]: r for r in _q(db, """SELECT sheet, COUNT(*) AS rows, SUM(n_formula) AS formulas,
                                                       SUM(n_const) AS consts FROM rows GROUP BY sheet""")}
        for s in sheets:
            s.update({k: (per_sheet.get(s["sheet"]) or {}).get(k) or 0 for k in ("rows", "formulas", "consts")})
        live = ",".join(f"'{k}'" for k in LIVE)
        edges = {"kinds": [], "flows": [], "most_read": [], "most_reading": []}
        if "edges" in have:
            edges["kinds"] = _q(db, "SELECT kind, COUNT(*) AS n FROM edges GROUP BY kind ORDER BY n DESC")
            edges["flows"] = _q(db, f"""SELECT src_sheet, dst_sheet, SUM(kind IN ({live})) AS live,
                                        SUM(kind NOT IN ({live})) AS other FROM edges
                                        WHERE src_sheet <> dst_sheet GROUP BY src_sheet, dst_sheet""")
            top = """SELECT e.{a}_sheet AS sheet, e.{a}_row AS row, r.label, r.units, COUNT(*) AS n,
                            COUNT(DISTINCT e.{b}_sheet) AS sheets
                     FROM edges e LEFT JOIN rows r ON r.sheet = e.{a}_sheet AND r.row = e.{a}_row
                     WHERE e.kind IN ({live}) GROUP BY e.{a}_sheet, e.{a}_row ORDER BY n DESC LIMIT 15"""
            edges["most_read"] = _q(db, top.format(a="dst", b="src", live=live))
            edges["most_reading"] = _q(db, top.format(a="src", b="dst", live=live))
        names = _q(db, "SELECT name, ref FROM names ORDER BY name") if "names" in have else []
        sections = _q(db, """SELECT sheet, section, COUNT(*) AS rows FROM rows WHERE section IS NOT NULL
                             AND section <> '' GROUP BY sheet, section ORDER BY sheet, MIN(row)""")
    return {**rec, "sheets": sheets, "edges": edges, "names": names, "sections": sections,
            "totals": {"sheets": len(sheets), "rows": sum(s["rows"] for s in sheets),
                       "formulas": sum(s["formulas"] for s in sheets), "consts": sum(s["consts"] for s in sheets),
                       "edges": sum(k["n"] for k in edges["kinds"]), "names": len(names)}}


@app.get("/api/workbook/{fid}/rows")
def rows(fid: str, sheet: str | None = None, q: str | None = None, limit: int = 200,
         row: int | None = Query(None, ge=1, le=1_048_576)):   # Excel's last row; a bigger int would overflow SQLite
    _, path = _model_db(fid)
    where, args = [], []
    if sheet:
        where.append("r.sheet = ?"); args.append(sheet)
    if row is not None:
        where.append("r.row = ?"); args.append(row)
    if q:
        like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
        where.append("(r.label LIKE ? ESCAPE '\\' OR r.section LIKE ? ESCAPE '\\')"); args += [like] * 2
    live = ",".join(f"'{k}'" for k in LIVE)
    with _ro(path) as db:
        # edges has no index, so pick the page of rows first and count edges only for those.
        out = _q(db, f"""WITH m AS (SELECT * FROM rows r {"WHERE " + " AND ".join(where) if where else ""}
                                    ORDER BY (SELECT rowid FROM sheets s WHERE s.sheet = r.sheet), r.row LIMIT ?)
                         SELECT m.sheet, m.row, m.section, m.label, m.units, m.n_formula, m.n_const, m.samples,
                           (SELECT COUNT(*) FROM edges e WHERE e.src_sheet=m.sheet AND e.src_row=m.row
                              AND e.kind IN ({live})) AS reads,
                           (SELECT COUNT(*) FROM edges e WHERE e.dst_sheet=m.sheet AND e.dst_row=m.row
                              AND e.kind IN ({live})) AS read_by
                         FROM m""", *args, min(max(limit, 1), 1000))
    return {"rows": out}


# ---- statements, profile, ontology: rule-based, no model calls; results kept in memory per (file, mtime) ----------

_CACHE: "OrderedDict[tuple, dict]" = OrderedDict()   # (kind, path, mtime_ns) -> result, least recently used first
_CACHE_MAX = 16          # results kept; a detect on a large model is a few MB, so this stays bounded in a long run
_CACHE_LOCK = threading.Lock()
_KEY_LOCKS: dict[tuple, threading.Lock] = {}  # one per key being computed, so the page's two requests share one detect


def _cached(kind: str, path: Path, make):
    try:
        key = (kind, str(path), path.stat().st_mtime_ns)
    except OSError:
        raise HTTPException(404, "No such workbook")
    with _CACHE_LOCK:
        if key in _CACHE:
            _CACHE.move_to_end(key)
            return _CACHE[key]
        lock = _KEY_LOCKS.setdefault(key, threading.Lock())
    with lock:
        try:
            with _CACHE_LOCK:
                if key in _CACHE:
                    return _CACHE[key]
            val = make()
            with _CACHE_LOCK:
                for old in [k for k in _CACHE if k[:2] == key[:2]]:   # an older build of the same file
                    del _CACHE[old]
                _CACHE[key] = val
                while len(_CACHE) > _CACHE_MAX:
                    _CACHE.popitem(last=False)
            return val
        finally:
            with _CACHE_LOCK:
                _KEY_LOCKS.pop(key, None)


def _detect(path: Path) -> dict:
    return _cached("statements", path, lambda: statements.detect(str(path)))


@app.get("/api/workbook/{fid}/statements")
async def workbook_statements(fid: str):
    _, path = _model_db(fid)
    return await run_in_threadpool(_detect, path)


@app.get("/api/workbook/{fid}/profile")
async def workbook_profile(fid: str):
    """The complexity profile at level full (real sheet names: this dashboard is local)."""
    _, path = _model_db(fid)
    return await run_in_threadpool(_cached, "profile", path, lambda: diagnose.profile_full(str(path), _detect(path)))


@app.get("/api/ontology")
def ontology_json():
    return ontology.as_json()


# ---- value dependency graph (modelatlas/depgraph.py): read-only, rule-based, no model calls -----------------------

@app.get("/api/workbook/{fid}/depgraph/anchors")
async def depgraph_anchors(fid: str):
    _, path = _model_db(fid)
    # valuation.catalogue can take minutes the first time on a big workbook (then it is cached beside model.db),
    # so it runs in the thread pool and the server keeps answering.
    return {"anchors": await run_in_threadpool(depgraph.anchors, str(path))}


def _graph(fid: str, cell: str | None, depth: int, max_rows: int) -> dict:
    rec, path = _model_db(fid)
    cell = (cell or "").strip() or None
    if cell:  # build() draws a graph around any text it is given, so check the cell is in the workbook first
        sheet, _, addr = cell.rpartition("!")
        sheet = sheet[1:-1].replace("''", "'") if sheet.startswith("'") and sheet.endswith("'") else sheet
        with _ro(path) as db:
            found = db.execute("SELECT 1 FROM cells WHERE sheet = ? AND addr = ? COLLATE NOCASE LIMIT 1",
                               (sheet, addr.replace("$", ""))).fetchone()
        if not found:
            raise HTTPException(400, f"No cell {cell} in this workbook; give one as Sheet!C5")
    try:
        return depgraph.build(str(path), cell=cell, depth=depth, max_rows=max_rows,
                              workbook=rec["filename"], with_anchors=False)
    except ValueError as e:
        raise HTTPException(400, str(e))


@app.get("/api/workbook/{fid}/depgraph")
async def depgraph_json(fid: str, cell: str | None = None, depth: int = Query(6, ge=1, le=30),
                        max_rows: int = Query(300, ge=10, le=3000)):
    return await run_in_threadpool(_graph, fid, cell, depth, max_rows)


@app.get("/api/workbook/{fid}/depgraph.html")
async def depgraph_page(fid: str, cell: str | None = None, depth: int = Query(6, ge=1, le=30),
                        max_rows: int = Query(300, ge=10, le=3000)):
    def page():
        g = _graph(fid, cell, depth, max_rows)
        g["generated"] = datetime.now().isoformat(timespec="seconds")
        return depgraph_html.render(g)
    return HTMLResponse(await run_in_threadpool(page))


def main() -> None:
    if "--version" in sys.argv[1:]:
        print(version_line())
        return
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8001)
