"""Render a dependency graph (modelatlas/depgraph.py) as one self-contained HTML report: how a DCF value outcome in an
Excel model is built, as a classified upstream graph plus tables, so a reviewer can see at a glance which assumptions,
cash flows and discounting steps sit behind the number and open any item to see its formula and neighbours.

The page needs no server and no network beyond web fonts: the graph is embedded as JSON, the layout (layers, a
barycentre ordering to cut crossings) and the pan / zoom / highlight run in a few hundred lines of inline script, and the
tables are written as plain HTML so they read without script. No model calls.

    uv run python -m modelatlas.depgraph_html graph.json out.html
"""
import html
import json
import re
import sys

from . import NOTICE, __version__

# Class colours follow the dashboard's categorical order; "calculation" (the bulk of any graph) takes the neutral
# gray so the interesting classes stand out. Classes the graph names beyond these fall back to the spare slot.
_SLOT = {"outcome": "--s1", "bridge": "--s2", "cashflow": "--s3", "timeline": "--s4", "assumption": "--s5",
         "discounting": "--s7", "terminal": "--s8", "calculation": "--gray"}
_SPARE = ["--s6", "--s1", "--s2", "--s3"]
_FILLER = 0.16  # how much of the class colour tints a node fill

esc = lambda s: html.escape("" if s is None else str(s), quote=True)  # noqa: E731


def _num(v) -> str:
    """HTML-safe: a number with separators and up to 4 decimals; anything else escaped (a value can be client text)."""
    if v is None or isinstance(v, bool):
        return ""
    if isinstance(v, (int, float)):
        s = f"{v:,.4f}".rstrip("0").rstrip(".")
        return "0" if s in ("-0", "") else s
    return esc(v)


def _value(n: dict) -> str:
    """Number with separators and up to 4 decimals; a rate-like value inside (-1, 1) reads as a percent."""
    v = n.get("value")
    if isinstance(v, (int, float)) and not isinstance(v, bool) and -1 < v < 1 and "rate" in (n.get("subclass") or "").lower():
        return _num(round(v * 100, 4)) + "%"
    return _num(v)


def _ref(n: dict) -> str:
    if n.get("kind") == "group":
        return ""
    return n.get("cell") or (f"r{n['row']}" if n.get("row") is not None else n["id"])


def _cls_css(classes: list[dict]) -> str:
    out, spare = [], iter(_SPARE)
    for c in classes:
        slot = _SLOT.get(c["key"]) or next(spare, "--s6")
        out.append(f'.k-{re.sub(r"[^A-Za-z0-9_-]", "_", c["key"])} {{ --c: var({slot}); }}')
    return "\n".join(out)


def _header(g: dict, stats: dict) -> str:
    st = g.get("start") or {}
    pills = ""
    anchors = g.get("anchors") or []
    if len(anchors) > 1:
        items = []
        for a in anchors:
            cur = a.get("cell") == st.get("id")
            items.append(f'<span class="pill{" ok" if cur else ""}" title="{"Shown here" if cur else "Rebuild from " + esc(a.get("cell"))}"'
                         f'{" aria-current=\"true\"" if cur else ""}>{esc(a.get("label") or a.get("cell"))} <span class="mono">{esc(a.get("cell"))}</span></span>')
        pills = f'<div class="anchors" aria-label="Other value outcomes in this workbook">{"".join(items)}</div>'
    bits = [f'{_num(stats.get("rows_shown"))} of {_num(stats.get("rows_upstream_total"))} upstream rows shown',
            f'{_num(stats.get("cells_shown"))} cells', f'{_num(stats.get("inactive_hidden"))} inactive hidden',
            f'{_num(stats.get("groups"))} groups']
    if stats.get("build_secs") is not None:
        bits.append(f'built in {_num(stats["build_secs"])} s')
    return f"""<header>
  <div class="hd"><div class="wb mono">{esc(g.get("workbook"))}</div>
  <h1>{esc(st.get("label"))} <span class="big num">{_num(st.get("value"))}</span> <span class="mono muted">{esc(st.get("id"))}</span></h1>
  <p class="muted stats">{" · ".join(bits)}</p>{pills}</div>
  <button class="btn" id="theme" type="button" aria-label="Switch light or dark theme">Theme</button>
</header>"""


def _cores(g: dict) -> str:
    cards = []
    for c in g.get("cores") or []:
        m = c.get("method") or {}
        rv = m.get("rate_value")
        rate = f'{esc(m.get("rate"))} = {_num(round(rv * 100, 4))}%' if isinstance(rv, (int, float)) else esc(m.get("rate"))
        rows = [("Cash flow", f'{esc(c.get("cashflow_label"))} <span class="mono muted">{esc(c.get("cashflow"))}</span>'),
                ("Discount factor row", f'<span class="mono">{esc(c.get("factor_row"))}</span>'),
                ("Rate", f'<span class="mono">{rate}</span>'), ("Valuation date", f'<span class="mono">{esc(m.get("valuation_date"))}</span>'),
                ("Timing", esc(m.get("timing"))), ("Day count", esc(m.get("day_count"))),
                ("Periods", f'{_num(c.get("periods"))}: {esc(c.get("first_period"))} to {esc(c.get("last_period"))}'),
                ("Present value", f'<b class="num">{_num(c.get("pv"))}</b>'), ("Undiscounted", f'<span class="num">{_num(c.get("undiscounted"))}</span>')]
        cards.append(f'<div class="card"><h3>{esc(c.get("what"))}</h3><p class="sub mono">{esc(c.get("cell"))}</p><dl>'
                     + "".join(f"<dt>{k}</dt><dd>{v}</dd>" for k, v in rows if v not in ("", None)) + "</dl></div>")
    if not cards:
        return ""
    return f'<section><h2>How the value is discounted</h2><div class="cores">{"".join(cards)}</div></section>'


def _items(g: dict) -> str:
    nodes = g.get("nodes") or []
    parts = []
    for c in g.get("classes") or []:
        ns = sorted((n for n in nodes if n.get("class") == c["key"]),
                    key=lambda n: (n.get("layer") or 0, n.get("sheet") or "", n.get("row") or 0, n["id"]))
        if not ns:
            continue
        body = []
        for n in ns:
            if n.get("kind") == "group":
                lab, val, why = f'{_num(n.get("count"))} more rows', "", ""
                sub = ""
            else:
                lab, sub, why = n.get("label"), n.get("subclass"), n.get("why")
                val = _value(n) or (f'<span class="mono muted">{esc((n.get("pattern") or "")[:80])}</span>' if n.get("pattern") else "")
            inp = ' <span class="pill" title="Hard-coded input">input</span>' if n.get("input") else ""
            hid = f' <span class="pill" title="Candidates not selected by the active scenario">+{_num(n["inactive_hidden"])} not selected</span>' if n.get("inactive_hidden") else ""
            body.append(f'<tr class="link" tabindex="0" data-id="{esc(n["id"])}"><td>{esc(n.get("sheet"))}</td><td class="mono">{esc(_ref(n))}</td>'
                        f'<td>{esc(lab)}{inp}{hid}</td><td>{esc(sub)}</td><td class="num">{val}</td><td class="num">{_num(n.get("reads"))}</td>'
                        f'<td class="num">{_num(n.get("read_by"))}</td><td class="why">{esc(why)}</td></tr>')
        head = ("<th>Sheet</th><th>Row / cell</th><th>Label</th><th>Subclass</th><th class='num'>Value / pattern</th>"
                "<th class='num'>Reads</th><th class='num'>Read by</th><th>Why</th>")
        op = " open" if len(ns) <= 60 else ""
        parts.append(f'<details class="card k-{re.sub(r"[^A-Za-z0-9_-]", "_", c["key"])}"{op}><summary><i class="sw"></i>{esc(c["name"])} '
                     f'<span class="pill">{len(ns)}</span></summary><div class="scroll"><table><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div></details>')
    return f'<section><h2>Items by class</h2><p class="muted sub">Select a row to find it in the graph.</p><div class="stack">{"".join(parts)}</div></section>'


def _sheets(g: dict) -> str:
    rows = g.get("by_sheet") or []
    if not rows:
        return ""
    top = max((r.get("upstream") or 0) for r in rows) or 1
    body = "".join(f'<tr><td>{esc(r["sheet"])}</td><td class="num">{_num(r.get("upstream"))}</td><td class="num">{_num(r.get("shown"))}</td>'
                   f'<td class="barcell"><div class="bar2" role="img" aria-label="{_num(r.get("upstream"))} upstream, {_num(r.get("shown"))} shown">'
                   f'<i style="width:{100 * (r.get("upstream") or 0) / top:.1f}%"></i><b style="width:{100 * (r.get("shown") or 0) / top:.1f}%"></b></div></td></tr>' for r in rows)
    return (f'<section><h2>Sheets</h2><div class="card scroll"><table><thead><tr><th>Sheet</th><th class="num">Upstream rows</th><th class="num">Shown</th>'
            f'<th>Shown (dark) of upstream (light)</th></tr></thead><tbody>{body}</tbody></table></div></section>')


def render(graph: dict) -> str:
    g = dict(graph)
    known = {c.get("key") for c in g.get("classes") or []}
    extra = sorted({n.get("class") for n in g.get("nodes") or []} - known - {None})
    if extra:  # a class a node uses but the legend doesn't name still gets a colour, a table and a filter
        g["classes"] = list(g.get("classes") or []) + [{"key": k, "name": str(k)} for k in extra]
    stats = dict(g.get("stats") or {})
    stats.setdefault("nodes", len(g.get("nodes") or []))
    stats.setdefault("edges", len(g.get("edges") or []))
    blob = json.dumps(g, ensure_ascii=True, separators=(",", ":")).replace("</", "<\\/").replace("<!--", "<\\u0021--")
    subs = dict((("__TITLE__", esc((g.get("start") or {}).get("label") or "Dependency graph") + " · " + esc(g.get("workbook"))),
                 ("__CLASSCSS__", _cls_css(g.get("classes") or [])), ("__HEADER__", _header(g, stats)),
                 ("__CORES__", _cores(g)), ("__ITEMS__", _items(g)), ("__SHEETS__", _sheets(g)),
                 ("__FOOT__", f"Generated by Model Atlas {esc(g.get('version') or __version__)}"
                              + (f" · {esc(g['generated'])}" if g.get("generated") else "")
                              + f"<br><span class=\"notice\">{esc(NOTICE)}</span>"),
                 ("__GRAPH__", blob)))
    # one pass: a placeholder spelled inside a client label must not be filled in by a later replacement
    return re.sub("|".join(map(re.escape, subs)), lambda m: subs[m[0]], _PAGE)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("usage: depgraph_html.py graph.json out.html")
        return 2
    with open(argv[1], encoding="utf-8") as f:
        out = render(json.load(f))
    with open(argv[2], "w", encoding="utf-8") as f:
        f.write(out)
    print(f"wrote {argv[2]} ({len(out) // 1024} kB)")
    return 0


_PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&display=swap">
<style>
:root {
  --ground: #f4f6f4; --surface: #ffffff; --ink: #1b2320; --muted: #5d6a64; --rule: #d8ded9;
  --accent: #1d6b47; --accent-soft: #e3efe7; --code-bg: #eef1ee;
  --s1: #2a78d6; --s2: #eb6834; --s3: #1baf7a; --s4: #eda100; --s5: #e87ba4; --s6: #008300; --s7: #4a3aa7; --s8: #e34948; --gray: #6b7771;
  --sans: "IBM Plex Sans", system-ui, -apple-system, sans-serif;
  --mono: "IBM Plex Mono", ui-monospace, Menlo, monospace;
  color-scheme: light;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --ground: #111513; --surface: #181e1b; --ink: #e2e9e5; --muted: #93a39b; --rule: #2b3430;
    --accent: #58bd8a; --accent-soft: #1d3328; --code-bg: #1f2622;
    --s1: #3987e5; --s2: #d95926; --s3: #199e70; --s4: #c98500; --s5: #d55181; --s6: #008300; --s7: #9085e9; --s8: #e66767; --gray: #93a39b;
    color-scheme: dark;
  }
}
:root[data-theme="dark"] {
  --ground: #111513; --surface: #181e1b; --ink: #e2e9e5; --muted: #93a39b; --rule: #2b3430;
  --accent: #58bd8a; --accent-soft: #1d3328; --code-bg: #1f2622;
  --s1: #3987e5; --s2: #d95926; --s3: #199e70; --s4: #c98500; --s5: #d55181; --s6: #008300; --s7: #9085e9; --s8: #e66767; --gray: #93a39b;
  color-scheme: dark;
}
body { --c: var(--gray); }
__CLASSCSS__
* { box-sizing: border-box; }
html { background: var(--ground); }
body { margin: 0; background: var(--ground); color: var(--ink); font: 15px/1.55 var(--sans); overflow-x: hidden; }
a { color: var(--accent); }
button, input { font: inherit; color: inherit; }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }
.muted { color: var(--muted); }
.num { font-variant-numeric: tabular-nums; text-align: right; white-space: nowrap; }
.mono { font-family: var(--mono); font-size: 12.5px; }
header { display: flex; align-items: flex-start; gap: 12px; padding: 14px 24px; border-bottom: 1px solid var(--rule); background: var(--surface); }
.hd { flex: 1; min-width: 0; }
.wb { color: var(--muted); }
h1 { font-size: 20px; font-weight: 600; margin: 2px 0; }
h1 .big { font-size: 26px; margin: 0 6px; }
.stats { margin: 2px 0 0; font-size: 13px; }
.anchors { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 8px; }
.btn { background: var(--surface); border: 1px solid var(--rule); border-radius: 6px; padding: 4px 10px; cursor: pointer; font-size: 13px; }
.btn:hover { border-color: var(--accent); }
.btn[aria-pressed="true"] { background: var(--accent-soft); border-color: var(--accent); }
main { max-width: 1240px; margin: 0 auto; padding: 20px 24px 40px; display: grid; gap: 24px; }
section h2 { font-size: 16px; font-weight: 600; margin: 0 0 8px; }
.sub { font-size: 13px; color: var(--muted); margin: 0 0 8px; }
.card { background: var(--surface); border: 1px solid var(--rule); border-radius: 10px; padding: 14px 16px; min-width: 0; }
.card h3 { font-size: 14px; margin: 0; }
.cores { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 380px), 1fr)); gap: 16px; }
dl { display: grid; grid-template-columns: max-content 1fr; gap: 3px 16px; margin: 6px 0 0; font-size: 13.5px; }
dt { color: var(--muted); } dd { margin: 0; min-width: 0; overflow-wrap: anywhere; }
.pill { display: inline-block; font-size: 11.5px; padding: 1px 8px; border-radius: 99px; background: var(--code-bg); color: var(--muted); white-space: nowrap; }
.pill.ok { background: var(--accent-soft); color: var(--accent); }
.scroll { overflow-x: auto; }
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th { text-align: left; font-weight: 500; font-size: 12px; color: var(--muted); padding: 6px 10px; border-bottom: 1px solid var(--rule); white-space: nowrap; }
th.num { text-align: right; }
td { padding: 6px 10px; border-bottom: 1px solid var(--rule); vertical-align: top; }
tr:last-child td { border-bottom: 0; }
tr.link { cursor: pointer; } tr.link:hover td, tr.sel td { background: var(--accent-soft); }
td.why { color: var(--muted); min-width: 200px; }
td.barcell { width: 40%; vertical-align: middle; }
.bar2 { position: relative; height: 10px; } .bar2 i, .bar2 b { position: absolute; left: 0; top: 0; height: 10px; border-radius: 0 4px 4px 0; min-width: 2px; }
.bar2 i { background: color-mix(in srgb, var(--s1) 28%, var(--surface)); } .bar2 b { background: var(--s1); height: 6px; top: 2px; }
.stack { display: grid; gap: 10px; }
details summary { cursor: pointer; font-weight: 600; font-size: 14px; }
details[open] summary { margin-bottom: 8px; }
.sw { display: inline-block; width: 12px; height: 12px; border-radius: 3px; background: var(--c); margin-right: 8px; vertical-align: -1px; }
/* graph */
.tools { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-bottom: 8px; }
.tools input[type=search] { background: var(--surface); border: 1px solid var(--rule); border-radius: 6px; padding: 4px 8px; min-width: 200px; }
.chk { font-size: 13px; display: inline-flex; gap: 6px; align-items: center; }
.legend { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 8px; }
.legend .btn { display: inline-flex; align-items: center; gap: 6px; }
.legend .btn[aria-pressed="false"] { opacity: .55; text-decoration: line-through; }
.gwrap { display: flex; height: 70vh; min-height: 420px; border: 1px solid var(--rule); border-radius: 10px; overflow: hidden; background: var(--surface); }
.gview { position: relative; flex: 1; min-width: 0; }
#svg { width: 100%; height: 100%; display: block; cursor: grab; touch-action: none; }
#svg.drag { cursor: grabbing; }
.axis { position: absolute; left: 12px; top: 8px; font-size: 12px; color: var(--muted); pointer-events: none; background: var(--surface); padding: 0 4px; border-radius: 4px; }
.zoom { position: absolute; right: 10px; bottom: 10px; display: flex; gap: 4px; }
#panel { width: 340px; flex: none; border-left: 1px solid var(--rule); padding: 14px; overflow-y: auto; font-size: 13px; }
#panel h3 { margin: 6px 0 4px; font-size: 15px; }
#panel h4 { margin: 12px 0 4px; font-size: 12px; font-weight: 500; color: var(--muted); }
#panel pre { margin: 0; background: var(--code-bg); border-radius: 6px; padding: 6px 8px; font: 12px/1.45 var(--mono); white-space: pre-wrap; overflow-wrap: anywhere; }
#panel ul { list-style: none; padding: 0; margin: 0; display: grid; gap: 2px; }
#panel li button { all: unset; cursor: pointer; display: block; width: 100%; padding: 2px 4px; border-radius: 4px; color: var(--accent); }
#panel li button:hover, #panel li button:focus-visible { background: var(--accent-soft); outline: 2px solid var(--accent); outline-offset: -1px; }
.cpill { display: inline-block; font-size: 11.5px; padding: 1px 8px; border-radius: 99px; background: color-mix(in srgb, var(--c) 22%, var(--surface)); border: 1px solid var(--c); color: var(--ink); }
.node { cursor: pointer; }
.node rect.box { fill: color-mix(in srgb, var(--c) 16%, var(--surface)); stroke: var(--c); stroke-width: 1.2; }
.node.outcome rect.box { fill: color-mix(in srgb, var(--c) 34%, var(--surface)); stroke-width: 2.5; }
.node.input rect.box { stroke-dasharray: 4 2.5; }
.node rect.back { fill: var(--surface); stroke: var(--c); stroke-width: 1; }
.node text { fill: var(--ink); font-family: var(--sans); pointer-events: none; }
.node text.lab { font-size: 11px; font-weight: 500; } .node text.sh { font-size: 9.5px; fill: var(--muted); }
.node .mark { fill: var(--c); } .node .badge rect { fill: var(--ink); } .node .badge text { fill: var(--surface); font-size: 9px; font-weight: 600; }
.node.faded { opacity: .12; } .node.sel rect.box { stroke: var(--ink); stroke-width: 3; }
.node:focus-visible rect.box { stroke: var(--accent); stroke-width: 3; }
.edge { fill: none; stroke: var(--muted); stroke-opacity: .35; stroke-width: 1; }
.edge.part { stroke-opacity: .6; stroke-width: 1.6; } .edge.collapsed { stroke-dasharray: 2 3; }
.edge.faded { stroke-opacity: .05; } .edge.hot { stroke: var(--ink); stroke-opacity: .75; stroke-width: 1.8; }
footer .notice { display: inline-block; max-width: 760px; margin-top: 6px; }
footer { text-align: center; color: var(--muted); font-size: 12.5px; padding: 0 24px 28px; }
@media (max-width: 900px) { .gwrap { flex-direction: column; height: auto; } .gview { height: 60vh; } #panel { width: auto; border-left: 0; border-top: 1px solid var(--rule); max-height: 50vh; } }
</style>
</head>
<body>
__HEADER__
<main>
__CORES__
<section>
  <h2>The graph</h2>
  <p class="sub">What the value reads, layer by layer: assumptions on the left, flowing right into the outcome. Dashed outline is a hard-coded input; the +N badge counts candidates the active scenario did not select. Select a node to light its upstream chain; wheel zooms, drag pans.</p>
  <div class="tools">
    <input type="search" id="q" placeholder="Search label or cell" aria-label="Search label or cell"><span class="muted" id="qn" aria-live="polite"></span>
    <label class="chk"><input type="checkbox" id="hidecalc"> Hide calculation rows</label>
  </div>
  <div class="legend" id="legend" role="group" aria-label="Filter by class"></div>
  <div class="gwrap">
    <div class="gview"><svg id="svg" role="img" aria-label="Dependency graph"><g id="vp"><g id="ge"></g><g id="gn"></g></g></svg>
      <div class="axis">assumptions &rarr; value</div>
      <div class="zoom"><button class="btn" id="zin" type="button" aria-label="Zoom in">+</button><button class="btn" id="zout" type="button" aria-label="Zoom out">&minus;</button><button class="btn" id="fit" type="button">Fit to view</button></div>
    </div>
    <aside id="panel" aria-live="polite"><p class="muted">Select a node to see what it is, how it is calculated and what it reads.</p></aside>
  </div>
</section>
__ITEMS__
__SHEETS__
</main>
<footer>__FOOT__</footer>
<script type="application/json" id="graph">__GRAPH__</script>
<script>
"use strict";
const $ = s => document.querySelector(s);
const G = JSON.parse($("#graph").textContent);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const cls = k => "k-" + String(k).replace(/[^A-Za-z0-9_-]/g, "_");
const fmt = (v, n) => {
  if (v == null || typeof v === "boolean") return "";
  if (typeof v !== "number") return String(v);
  if (v > -1 && v < 1 && /rate/i.test((n && n.subclass) || "")) return (+(v * 100).toFixed(4)).toLocaleString("en-US", { maximumFractionDigits: 4 }) + "%";
  return v.toLocaleString("en-US", { maximumFractionDigits: 4 });
};
const ref = n => n.kind === "group" ? "" : (n.cell || (n.row != null ? "r" + n.row : n.id));

/* theme */
const root = document.documentElement;
try { const t = localStorage.getItem("depgraph-theme"); if (t) root.dataset.theme = t; } catch (e) {}
$("#theme").addEventListener("click", () => {
  const dark = root.dataset.theme ? root.dataset.theme === "dark" : matchMedia("(prefers-color-scheme: dark)").matches;
  root.dataset.theme = dark ? "light" : "dark";
  try { localStorage.setItem("depgraph-theme", root.dataset.theme); } catch (e) {}
});

/* model */
const nodes = G.nodes, byId = new Map(nodes.map(n => [n.id, n]));
const reads = new Map(nodes.map(n => [n.id, []])), readBy = new Map(nodes.map(n => [n.id, []]));
const edges = G.edges.filter(e => byId.has(e.src) && byId.has(e.dst) && e.src !== e.dst);
edges.forEach(e => { reads.get(e.src).push(e.dst); readBy.get(e.dst).push(e.src); });
const classes = G.classes || [], className = Object.fromEntries(classes.map(c => [c.key, c.name]));
const start = byId.get(G.start && G.start.id) || byId.get(G.start && G.start.row_id) || nodes[0];

/* layout: columns by layer, outcome at the right; barycentre sweeps inside each column */
const W = 164, H = 26, PITCH = 34, COLW = 210, MAXC = 26;
const known = nodes.map(n => n.layer).filter(Number.isFinite), maxL = Math.max(0, ...known.map(l => Math.max(0, Math.round(l))));
const layerOf = n => Number.isFinite(n.layer) ? Math.min(maxL, Math.max(0, Math.round(n.layer))) : (n.kind === "group" ? maxL : 0);
const flip = !!start && layerOf(start) > maxL / 2;
const colOf = n => flip ? layerOf(n) : maxL - layerOf(n);
const cols = Array.from({ length: maxL + 1 }, () => []);
nodes.slice().sort((a, b) => (a.sheet || "").localeCompare(b.sheet || "") || (a.row || 0) - (b.row || 0)).forEach(n => cols[colOf(n)].push(n));
const maxH = Math.max(1, ...cols.map(c => c.length)) * PITCH;
const place = () => cols.forEach((c, ci) => { const off = (maxH - c.length * PITCH) / 2; c.forEach((n, i) => { n.x = ci * COLW; n.y = off + i * PITCH; }); });
place();
for (let s = 0; s < 4; s++) {
  const order = s % 2 ? cols.map((_, i) => cols.length - 1 - i) : cols.map((_, i) => i);
  order.forEach(ci => {
    const sc = new Map();
    cols[ci].forEach((n, i) => {
      const nb = reads.get(n.id).concat(readBy.get(n.id)).filter(m => colOf(byId.get(m)) !== ci);
      sc.set(n.id, nb.length ? nb.reduce((a, m) => a + byId.get(m).y, 0) / nb.length : n.y);
    });
    cols[ci].sort((a, b) => sc.get(a.id) - sc.get(b.id));
    place();
  });
}

/* draw */
const NS = "http://www.w3.org/2000/svg";
const mk = (tag, attrs, parent) => { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; };
const trunc = s => s.length > MAXC ? s.slice(0, MAXC - 1) + "…" : s;
const ge = $("#ge"), gn = $("#gn");
const epath = e => {
  const d = byId.get(e.dst), s = byId.get(e.src);
  const x0 = d.x + W, y0 = d.y + H / 2, x1 = s.x, y1 = s.y + H / 2, k = Math.max(30, Math.abs(x1 - x0) * 0.45);
  return `M${x0},${y0}C${x0 + k},${y0} ${x1 - k},${y1} ${x1},${y1}`;
};
const eEls = edges.map(e => { const p = mk("path", { d: epath(e), class: "edge " + (e.kind || "") }, ge); return p; });
const nEls = new Map();
nodes.forEach(n => {
  const g = mk("g", { class: `node ${cls(n.class)} ${n.kind} ${n.class === "outcome" ? "outcome" : ""} ${n.input ? "input" : ""}`, transform: `translate(${n.x},${n.y})`, tabindex: 0, role: "button", "data-id": n.id }, gn);
  const label = n.kind === "group" ? (n.label || "") : (n.label || n.id);
  mk("title", {}, g).textContent = `${label}\n${n.id}${n.input ? "\nhard-coded input" : ""}${n.inactive_hidden ? `\n+${n.inactive_hidden} not selected` : ""}`;
  if (n.kind === "group") { mk("rect", { class: "back", x: 8, y: -4, width: W - 8, height: H, rx: 5 }, g); mk("rect", { class: "back", x: 4, y: -2, width: W - 4, height: H, rx: 5 }, g); }
  mk("rect", { class: "box", width: W, height: H, rx: 5 }, g);
  const tx = n.input ? 16 : 7;
  if (n.input) mk("path", { class: "mark", d: "M8,8 l4,5 l-4,5 l-4,-5z" }, g);
  mk("text", { class: "lab", x: tx, y: 11 }, g).textContent = trunc(label);
  mk("text", { class: "sh", x: tx, y: 21.5 }, g).textContent = n.kind === "group" ? "grouped" : (n.sheet || "") + (ref(n) ? " " + ref(n) : "");
  if (n.inactive_hidden) {
    const b = mk("g", { class: "badge", transform: `translate(${W - 4},-5)` }, g);
    const t = "+" + n.inactive_hidden, w = 9 + t.length * 5.4;
    mk("rect", { x: -w, y: -4, width: w, height: 11, rx: 5.5 }, b); mk("text", { x: -w / 2, y: 4.5, "text-anchor": "middle" }, b).textContent = t;
  }
  g.addEventListener("click", ev => { ev.stopPropagation(); select(n.id); });
  g.addEventListener("keydown", ev => { if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); select(n.id); } });
  nEls.set(n.id, g);
});

/* state: selection, class filter, search, hide calculation */
let sel = null, anc = null; const off = new Set(); let q = "", hideCalc = false;
const ancestors = id => { const seen = new Set([id]), st = [id]; while (st.length) for (const m of reads.get(st.pop())) if (!seen.has(m)) { seen.add(m); st.push(m); } return seen; };
const matches = n => !q || (n.label || "").toLowerCase().includes(q) || n.id.toLowerCase().includes(q);
function refresh() {
  const vis = new Map();
  nodes.forEach(n => {
    const ok = !off.has(n.class) && !(hideCalc && n.class === "calculation") && matches(n) && (!anc || anc.has(n.id));
    vis.set(n.id, ok);
    const el = nEls.get(n.id); el.classList.toggle("faded", !ok); el.classList.toggle("sel", n.id === sel); el.setAttribute("aria-pressed", n.id === sel);
  });
  edges.forEach((e, i) => {
    const on = vis.get(e.src) && vis.get(e.dst);
    eEls[i].classList.toggle("faded", !on); eEls[i].classList.toggle("hot", !!(anc && on));
  });
  document.querySelectorAll("tr.link").forEach(r => r.classList.toggle("sel", r.dataset.id === sel));
  $("#qn").textContent = q ? nodes.filter(matches).length + " match" + (nodes.filter(matches).length === 1 ? "" : "es") : "";
}
function panel(n) {
  const list = ids => ids.length ? "<ul>" + ids.map(i => { const m = byId.get(i); return `<li><button type="button" data-go="${esc(i)}">${esc(m.label || i)} <span class="muted">${esc(m.sheet)} ${esc(ref(m))}</span></button></li>`; }).join("") + "</ul>" : '<p class="muted">None</p>';
  const row = (h, v, mono) => v === "" || v == null ? "" : `<h4>${h}</h4>${mono ? `<pre>${esc(v)}</pre>` : `<div>${esc(v)}</div>`}`;
  const val = fmt(n.value, n);
  $("#panel").innerHTML = `<span class="cpill ${cls(n.class)}">${esc(className[n.class] || n.class)}${n.subclass ? " · " + esc(n.subclass) : ""}</span>
    <h3>${esc(n.label)}</h3><div class="mono muted">${esc(n.id)}</div>
    ${n.kind === "group" ? `<p>${esc(n.count)} more rows on ${esc(n.sheet)} sit behind this group; raise the row limit to expand it.</p>` : ""}
    ${row("Why it is here", n.why)}${val ? `<h4>Value${n.units ? " (" + esc(n.units) + ")" : ""}</h4><div class="num" style="text-align:left;font-size:16px;font-weight:600">${esc(val)}</div>` : ""}
    ${n.input ? "<h4>Hard-coded input</h4><div>Typed into the workbook, not calculated.</div>" : ""}
    ${row("Formula", n.formula, 1)}${row("In words", n.words, 1)}${row("Pattern", n.pattern, 1)}${row("Samples", n.samples, 1)}
    ${row("Section", n.section)}
    ${n.n_formula != null || n.n_const != null ? `<h4>Cells</h4><div>${esc(n.n_formula ?? 0)} formula, ${esc(n.n_const ?? 0)} constant</div>` : ""}
    ${n.inactive_hidden ? `<h4>Not selected</h4><div>+${esc(n.inactive_hidden)} candidates the active scenario did not select are hidden.</div>` : ""}
    ${n.hidden_upstream ? `<h4>Hidden upstream</h4><div>${esc(n.hidden_upstream)} further rows upstream are not drawn.</div>` : ""}
    <h4>Reads (${reads.get(n.id).length}${n.reads != null && n.reads !== reads.get(n.id).length ? " of " + esc(n.reads) : ""})</h4>${list(reads.get(n.id))}
    <h4>Read by (${readBy.get(n.id).length}${n.read_by != null && n.read_by !== readBy.get(n.id).length ? " of " + esc(n.read_by) : ""})</h4>${list(readBy.get(n.id))}`;
  $("#panel").querySelectorAll("[data-go]").forEach(b => b.addEventListener("click", () => select(b.dataset.go, true)));
}
function select(id, centre) {
  if (id == null || id === sel) { sel = null; anc = null; $("#panel").innerHTML = '<p class="muted">Select a node to see what it is, how it is calculated and what it reads.</p>'; refresh(); return; }
  sel = id; anc = ancestors(id); panel(byId.get(id)); refresh();
  if (centre) { const n = byId.get(id); view.x = $("#svg").clientWidth / 2 - (n.x + W / 2) * view.k; view.y = $("#svg").clientHeight / 2 - (n.y + H / 2) * view.k; apply(); }
}

/* legend */
const counts = {}; nodes.forEach(n => counts[n.class] = (counts[n.class] || 0) + 1);
$("#legend").innerHTML = classes.map(c => `<button class="btn ${cls(c.key)}" type="button" aria-pressed="true" data-k="${esc(c.key)}"><i class="sw" style="margin:0"></i>${esc(c.name)} <span class="muted">${counts[c.key] || 0}</span></button>`).join("");
document.querySelectorAll("#legend button").forEach(b => b.addEventListener("click", () => {
  const k = b.dataset.k; off.has(k) ? off.delete(k) : off.add(k); b.setAttribute("aria-pressed", !off.has(k)); refresh();
}));
$("#q").addEventListener("input", e => { q = e.target.value.trim().toLowerCase(); refresh(); });
$("#hidecalc").addEventListener("change", e => { hideCalc = e.target.checked; refresh(); });

/* pan / zoom */
const svg = $("#svg"), vp = $("#vp"), view = { x: 0, y: 0, k: 1 };
const apply = () => vp.setAttribute("transform", `translate(${view.x},${view.y}) scale(${view.k})`);
const bounds = { w: (cols.length - 1) * COLW + W + 20, h: maxH + 20 };
function fit(floor) {   /* the Fit button passes nothing (it shrinks as far as it must); the first view passes a floor */
  const vw = svg.clientWidth, vh = svg.clientHeight; view.k = Math.max(floor > 0 ? floor : 0.12, Math.min(vw / bounds.w, vh / bounds.h, 1) || 1);  /* a hidden pane has no size */
  view.x = (vw - bounds.w * view.k) / 2 + 10 * view.k; view.y = (vh - bounds.h * view.k) / 2 + 10 * view.k; apply();
}
function home() {   /* readable start: outcome at the right edge, at least 0.6x (0.8x on a small graph, whatever the pane) */
  const vw = svg.clientWidth, vh = svg.clientHeight, floor = nodes.length < 40 ? 0.8 : 0.6;
  const k = Math.max(floor, Math.min(vw / bounds.w, vh / bounds.h, 1));
  if (!start || (k * bounds.w <= vw && k * bounds.h <= vh)) return fit(floor);
  view.k = k; view.x = vw - 24 - (start.x + W) * k; view.y = vh / 2 - (start.y + H / 2) * k; apply();
}
const zoomAt = (f, cx, cy) => {
  const k = Math.min(3, Math.max(0.12, view.k * f)), r = k / view.k;
  view.x = cx - (cx - view.x) * r; view.y = cy - (cy - view.y) * r; view.k = k; apply();
};
svg.addEventListener("wheel", e => { e.preventDefault(); const b = svg.getBoundingClientRect(); zoomAt(Math.exp(-e.deltaY * 0.0015), e.clientX - b.left, e.clientY - b.top); }, { passive: false });
let drag = null;
svg.addEventListener("pointerdown", e => { drag = { x: e.clientX, y: e.clientY, vx: view.x, vy: view.y, moved: false }; });
svg.addEventListener("pointermove", e => {
  if (!drag) return; const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
  if (!drag.moved && Math.hypot(dx, dy) < 4) return;
  if (!drag.moved) { drag.moved = true; svg.classList.add("drag"); svg.setPointerCapture(e.pointerId); }
  view.x = drag.vx + dx; view.y = drag.vy + dy; apply();
});
svg.addEventListener("pointerup", e => { const moved = drag && drag.moved; drag = null; svg.classList.remove("drag"); if (!moved && e.target === svg || (!moved && e.target.closest && !e.target.closest(".node"))) select(null); });
$("#zin").onclick = () => zoomAt(1.4, svg.clientWidth / 2, svg.clientHeight / 2);
$("#zout").onclick = () => zoomAt(1 / 1.4, svg.clientWidth / 2, svg.clientHeight / 2);
$("#fit").onclick = () => fit();
addEventListener("resize", () => { if (!sel) home(); });

/* tables jump to the graph */
document.querySelectorAll("tr.link").forEach(r => {
  const go = () => { select(r.dataset.id, true); $(".gwrap").scrollIntoView({ block: "center" }); };
  r.addEventListener("click", go); r.addEventListener("keydown", e => { if (e.key === "Enter") go(); });
});
home(); refresh();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    sys.exit(main(sys.argv))
