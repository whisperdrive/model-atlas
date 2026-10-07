# Model Atlas

Reads a large Excel financial model into a SQLite map, with no formula engine and no model (LLM) calls. From saved
values and formulas alone it:

- builds a row-level map of the workbook and the edges between line items;
- traces a DCF value to the rows it rests on and puts each one in a class;
- finds the model's statements and schedules by the identities they satisfy, and tests those identities;
- lays the three statements out in one fixed format, from the model's own rows or derived from its schedules;
- draws the dependency graph as a standalone report;
- serves a read-only dashboard over the built databases;
- runs the whole analysis over many models as anonymised diagnostic runs, for a machine that cannot send anything out.

Nothing here recalculates a workbook. Every number is the one Excel saved; the checks are arithmetic over those values.

## Install and run order
```bash
uv sync
uv run atlas-build "path/to/model.xlsx"      # row map + SQLite store into out/<name>/ (run once per workbook)
uv run atlas-graph out/<name>/model.db        # DCF dependency graph, writes out/<name>/depgraph.html
uv run atlas-statements out/<name>/model.db   # statements, identities and findings
uv run atlas-threeway out/<name>/model.db [--fy]   # P&L, cash flow and balance sheet, checks and residuals
uv run atlas-dashboard                        # dashboard on http://localhost:8001: or drop the workbook on it instead of atlas-build
uv run atlas-diagnose out/ --report diag/     # anonymised report per model
```
Check which version produced a file or report with `uv run atlas-build --version` (every command takes `--version`, as
does `uv run python -m modelatlas`, which also lists the commands). The version is shown in the dashboard header, the
graph report footer, the statements text and the JSON of the graph, statements and diagnostic runs. Releases: bump
`__version__` in `modelatlas/__init__.py`; pyproject reads it from there.

Each command has a `python -m` equivalent: `modelatlas.build_map`, `modelatlas.depgraph`, `modelatlas.statements`,
`modelatlas.threeway`, `modelatlas.diagnose`, `modelatlas.census` (per-sheet formula and constant counts, `atlas-census`) and
`modelatlas.edges` (rebuilds the edges of an older model.db in place). Works on `.xlsx` and `.xlsm`; `.xlsb` and `.xls`
have no formulas to read here. Outputs go to `out/<workbook name>/` (`map.txt`, `model.db`, and `depgraph.html` once
the graph has run). `out/` is git-ignored.

## How it adapts to a workbook (modelatlas/layout.py)
Per sheet, from cached values: the **timeline** is the row with the most dates (longest run of date columns,
periodicity from the date gaps); the **label column** is the most text-heavy column left of it; the **units column**
is a column right of that with short, repeated text. Runs of 50+ same-shaped formula-free rows are summarised as one
TABLE line. The headline output is the line item with the biggest upstream tree that nothing depends on, skipping
model-check rows.

## Edges (modelatlas/edges.py)
Edges link each line item to the rows it reads, with a kind: `direct` (a plain reference), `offset` (the range an
OFFSET actually points at, worked out from saved values), `active` (the row a SUMIFS / INDEX-MATCH / CHOOSE selects in
the current scenario) or `inactive` (a candidate it considers but doesn't select). `trace` follows the active path and
counts the inactive candidates. `uv run python -m modelatlas.edges out/<dir>/model.db` upgrades an older model.db in
place.

## The dependency graph (modelatlas/depgraph.py, depgraph_html.py)
`dcf.py` redoes a valuation's discounting in plain Python from the workbook's own cash flows, rate and dates;
`dcftrace.py` finds the discounting in the formulas (a present-value row, XNPV, NPV, SUMPRODUCT with a factor row) and
follows it to its inputs; `valuation.py` finds the DCFs in a workbook by their labels and checks that the redo
reproduces the saved value. The graph starts from a DCF value cell (the top one it found, or one you name) and draws
what the value rests on. Each item is in one class: value outcome, bridge (EV to equity), discounting, terminal value,
cash flow, calculation, timeline or assumption, and each carries a one-sentence reason for its class. An input that is
typed in keeps its class and is flagged as an input.

The walk is bounded by depth (default 6) and a row budget (default 300); everything beyond is counted and collapsed
into one group per sheet, so the picture stays readable on a 400,000-formula model. The report is one self-contained
HTML page that can be saved from the browser.
```bash
uv run atlas-graph out/<dir>/model.db ["Sheet!A1"] [--depth N] [--max-rows N] [--json path] [--html path]
```

## Statements and identities (modelatlas/statements.py, ontology.py)
`statements.py` finds the model's statements and schedules (P&L, balance sheet, cash flow, debt, fixed assets, sources
and uses, revenue, equity, distributions) by the identities that hold in its saved values, reading the model's own
check rows first and using labels last. Each row it binds carries its reason and how it was found. Every identity is
reported as holds, fails or unbound, and split into:

- **tests of the model**: the rows were named by a check or by labels, so the identity can fail;
- **structure confirmed**: it holds by construction (a subtotal that adds its parts), so it says nothing about the
  model's correctness.

Findings list where the model fails an identity, where saved results are stale, where a balance sheet is a plug and
where a typed number breaks a row of formulas. The identities and roles are written out in `docs/model_ontology.md`,
which `modelatlas/ontology.py` generates (`uv run python -m modelatlas.ontology --write docs/model_ontology.md`) and a
test keeps in step.

## Three statements (modelatlas/threeway.py)
`threeway.py` lays the P&L, cash flow and balance sheet out from the blocks `statements.py` bound, in one fixed order of
lines (so two models' statements line up), costs and outflows negative. Where the model has a statement its own rows are
**extracted**; where it does not, the lines are **derived** from the schedules (revenue build, fixed-assets, debt and
equity roll-forwards, capex) and from the other statements, or are `missing` (shown as not found, never invented). A
statement is `extracted`, `derived`, `partial` or `none`; every line says where it came from (source rows, the formula of
a derived line, notes such as "sign flipped" or "opening cash not found; cumulative from the first period").
Identities between the lines are checks (`holds`, `fails`, or `unbound` when a line is missing, or is derived rather
than the model's own row: computed from the other lines it cannot differ, and taken from elsewhere it does not test the
model). A CSV text cell that would start a formula (`=`, `+`, `-`, `@`) is written with a leading apostrophe.

A balance sheet built from schedules does not know its opening balances, so it is compared by how it **moves**: the
change in assets less the change in liabilities and equity, each built from the roll-forward movements. The residual
analysis splits any movement that is not zero into what could explain it (profit against operating cash, depreciation
in the schedule against the P&L, capex against fixed-asset additions, debt flows against the debt balance, distributions
paid against distributions that reach retained earnings), each with a one-sentence reading, and says when a comparison is
no test because one side was taken from the other's own schedule. It runs on an extracted balance sheet that does not
balance too. `by="fy"` (`--fy`) aggregates to financial years: flows sum, stocks take the last period of the year
(opening cash the first), a year with fewer periods than a full year is marked partial, checks and residuals are
recomputed. The dashboard shows it as section 5 (`#/wb/<id>/statements`), with a period / financial-year toggle and a
CSV download per statement (`/api/workbook/<id>/threeway.csv?statement=pnl&by=fy`).
```bash
uv run python -m modelatlas.threeway out/<dir>/model.db [--fy] [--fy-end-month N] [--json path] [--csv dir]
```

## The dashboard (modelatlas/dashboard/)
```bash
uv run atlas-dashboard                                              # http://localhost:8001
uv run uvicorn modelatlas.dashboard.server:app --port 8001          # the same, with uvicorn's options
```
It reads the `out/` folder of the directory it is started in, or the folder named by `ATLAS_OUT`. Drop an `.xlsx` or
`.xlsm` on the portfolio page (or choose files): the browser fingerprints it (SHA-256) and asks whether the server
already has it, otherwise it is uploaded, saved under `uploads/<sha12>/` (git-ignored) and built into
`out/<stem>__<sha8>/` by one background worker, with the row showing queued, building (progress and step) and done
or error. The registry of dropped workbooks is `out/atlas.db`; **Rebuild** and **Remove** on a row clear and rebuild,
or delete, only that workbook's upload and folder, and a build never goes into an existing database. A build the
server died in the middle of is queued again at the next start. The size limit is 500 MB (`ATLAS_MAX_UPLOAD_MB`).
Each workbook page has **Run diagnostics (shapes only)**, which writes the anonymised report to `diag/<token>/` and
links `report.md`; if the self-scan blocks it, the page says so and the reason stays in the local
`diag/_blocked/<token>.txt`, which is never served. Reading is otherwise as below. Every
`out/<name>/model.db` is listed as a command-line build; an `out/registry.db`, if one happens to be there, is read as
well, but nothing needs it. A workbook opens as one page of five sections, each deep-linkable (`#/wb/<id>/holds`):
**What is this model?** (size, sheets and their roles, profile, line-item search); **How is the value built?** (the
dependency graph); **Does it hold together?** (findings, statement blocks and how each row was bound, identities as tests
versus structure confirmed, the model's own checks); **What should I look at?** (pattern breaks, stale cells, scenario
inputs); **The three statements** (the P&L, cash flow and balance sheet in a fixed layout, with checks and residuals). `#/ontology` lists the blocks, roles and identities the detector works from. Every database is opened read-only.

## Diagnostic runs (modelatlas/diagnose.py)
Runs the pipeline (census, formula families, sheet and row graphs, the DCF dependency graph, statement identities) over
many models and writes reports that carry the shapes of the models and nothing that names them. It is for a machine
where real client models are and nothing can be sent out: a person moves the report files by hand, so each one has to
be safe to move.
```bash
uv run atlas-diagnose out/                       # every out/*/model.db
uv run atlas-diagnose --workbooks <dir>          # build each .xlsx/.xlsm there first, then run
uv run atlas-diagnose out/ --report diag/ --level shapes|full --tests --only <folder substring>
```
Level `shapes` (the default) writes tokens (`M<sha8>` for the workbook, `S1..Sn` for sheets), R1C1 patterns with sheet
names, defined names, strings and non-round numbers replaced, sign patterns instead of values and relative residual
buckets, to `diag/<token>/report.md` and `report.json`. Before writing, a self-scan looks for the model's own sheet
names, labels, sections, units, text cells, defined names, file and registry names (and each of their non-generic
words) in that text; one hit, a key outside the fixed schema, any non-ASCII character or an unreadable registry, and the
report is not written, and `diag/_blocked/<token>.txt` (local only) names what leaked. Exception messages keep only
known safe words; everything else in them becomes an ellipsis.

Level `full` also writes real names under `diag/full/`, for use on the same machine only. **Share only
`diag/summary.md` and `diag/<token>/report.*`**, never `diag/full/` or `diag/_blocked/`. Each failing identity or
exception carries a fixture hint (patterns, row offsets, periodicity, whether a check row exists, the exception
site), enough to rebuild the case with a synthetic workbook. `--tests` also runs every `tests/check_*.py`, and
`diag/` is git-ignored.

The self-scan also covers this package itself: run the diagnostics over the repo's own synthetic models to see a clean
report before pointing it at real ones.

## Fixtures and tests
All test workbooks are synthetic and generated into `tests/sample_models/` (git-ignored):
```bash
uv run python tests/make_threeway_model.py      # a three-statement infrastructure model ("Harbourline"), four altered variants, and a schedules-only one
uv run python tests/make_sample_models.py       # a toll road ("Riverbend", code name "Kestrel"): client model, the same with valuation sheets inside, the next year's model
uv run python tests/make_trace_workbook.py      # a quarterly overlay with valuations by XNPV, SUMPRODUCT, a PV row and an NPV
for t in tests/check_*.py; do uv run python "$t"; done
```
`tests/fixtures/depgraph_fixture.json` is a small saved graph for the HTML report's check. Checks that smoke-test a
real model in `out/` skip when it is absent.

## Limits
- Layout detection is heuristic: a category column can be taken for units, and a sheet with several side-by-side
  blocks gets a single label column. One timeline per sheet; a sheet with no date row is a plain list.
- A mostly-typed row (mostly constants, a few formulas) is treated as an assumption, not a calculation.
- Flag rows (0/1 switches) can be counted as parts of a cash flow when they sit in its formula.
- A check row where 1 (or TRUE) means OK, rather than 0, is read as failing: the verdict assumes a difference that
  should be zero.
- A check named only by a weak word (balances, ok, equals) that computes a non-zero number through ROUND is dropped
  rather than reported, so a failing check of that shape goes unseen.
- A second balance sheet on the same sheet as the first is listed as an alternate, not bound.
- A balance sheet whose equity is typed in as a plug (rather than computed as assets less liabilities) is found only
  when the typed numbers make the identity hold; otherwise the identity is unbound.
- Three statements: a line the model's statement lacks is derived only from what the blocks bound (no label search of
  the sheets, except a costs total and the debt schedule's interest row), so an unbound ingredient (a tax row, a capex
  line) stays `missing` and the balance sheet movement then says it is no test. Derived cash is cumulative from the
  first period unless a cash roll-forward gives its opening, so its level differs from the model's by a constant; the
  sources-and-uses block (one column, on no timeline) is not placed in a period. Interest paid is taken as interest
  charged and tax paid as tax charged when the model has no cash rows. One period axis (the periodicity most bound
  rows share); rows on another timeline show only where their dates coincide.
- An identity whose rows are on different periodicities (monthly against annual) is left unbound rather than
  compared.
- The three statements put every line on one timeline, the periodicity most bound rows share; a row on another
  periodicity is sampled at those dates, so a quarterly flow shown against an annual timeline is one quarter, not the
  year. Timelines dated at the start of each period group into financial years correctly only where the start month
  happens to match; the detector does not yet tell start dates from end dates.
- A derived balance sheet assumes share capital does not move, so a model with an equity contribution in its cash
  flow shows a movement-check failure that the note and the unexplained line disclose.
- The report's "Fit to view" button shrinks the graph as far as it must, which on a very large graph is small; zoom in
  from there.
- `outputs.carry` (last year's schedule onto this year's) and the chart payloads of `valuation.validation` /
  `valuation.scenario` belong to the app this was extracted from and are not available here.

## Origin
Extracted from project-bob, whose apps still carry their own copies of these modules for now.

## Licence
MIT, see [LICENSE](LICENSE).

## Disclaimer
Experimental. Model Atlas is a prototype under active development and experimentation. Its maps, traces, classes,
identity checks and reports are produced by automated rules, can be incomplete or wrong, and must be checked by a
qualified person before anyone relies on them. Nothing here is valuation, financial, tax, accounting or legal advice,
or an opinion on any value.

Files stay on the machine they are read on; nothing is sent anywhere. Client files and models remain the property and
confidential information of their owners.

© 2026 Yong Ching Thai. Released under the MIT licence. Third-party names and trademarks belong to their owners and
appear only to identify them; no endorsement is implied.
