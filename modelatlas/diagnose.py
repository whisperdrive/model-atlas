"""Diagnostic runs: the model-analysis pipeline over many models, reported as SHAPES that are safe to share.

The code is developed on synthetic fixtures and run elsewhere on real client models, where nothing may leave the
machine except what a person moves by hand. A report from this module therefore carries what is needed to rebuild a
failure as a synthetic fixture (formula patterns, graph shape, statuses, tracebacks) and nothing that identifies the
client: no file, sheet, label or defined name, no value. No model (LLM) calls.

    uv run python -m modelatlas.diagnose out/                           # every out/*/model.db
    uv run python -m modelatlas.diagnose --workbooks <dir>              # build each .xlsx/.xlsm there first, then run
    uv run python -m modelatlas.diagnose out/ --report diag/ --level shapes|full --tests --only <folder substring>
        --depth 6 --max-rows 300

How a report is made safe (level shapes):
  - it is PROJECTED, not redacted: every field is computed and copied explicitly, so a label cannot ride along in a
    result dict nobody remembered to clean
  - workbook -> M<sha8>; sheets -> S1..Sn in workbook order; defined names -> NAME<k>; functions that are not Excel's
    -> FN<k>; sheets of other files -> X<k>; strings in formulas -> "..."; numbers that are not small and round -> #
  - labels are never written: a role is reported by the vocabulary names that matched and the label's word count
  - values are never written: sign patterns (+ - 0 per period) and relative residual buckets only
  - paths are made relative to the repo; the home directory, user and machine names are removed from every string;
    a traceback frame outside the repo and Python is '<other>/<file>'
  - exception messages are projected too: quoted text, paths, dates, long numbers and non-ASCII replaced, sheet and
    defined names tokenised, and every remaining word that is not a known safe word (the repo's vocabulary, its code's
    identifiers, the words of Python's and sqlite's error messages) replaced by "…"
  - a self-scan then looks for the model's own strings in the final JSON and Markdown text (sheet names, labels,
    sections, units, every text cell, defined names, file and registry names, linked-file names, and each non-generic
    word of these). One hit, a key outside the fixed schema, any non-ASCII character or an unreadable registry, and that
    model's shapes report is NOT written; a local-only diag/_blocked/<token>.txt names the leaked strings instead. Exact
    generic words that the repo's own vocabulary uses (revenue, cash, debt, check ...) and the words of this module's
    fixed text are not scanned for.
  - level full also writes real names and texts, under diag/full/ only. Never share those.
"""
import argparse
import hashlib
import importlib.metadata
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import time
import tokenize
import traceback
import types
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CODE = "modelatlas"  # the package folder under ROOT whose identifiers and source lines count as the repo's own code

from . import NOTICE, __version__, rodb, version_line  # noqa: E402

LIVE = ("direct", "offset", "active")  # edge kinds the current scenario uses (edges.py)
SCHEMA_VERSION = 1
SMALL, MEDIUM = 5_000, 100_000  # formulas: size classes
KEEP_DECIMALS = {0.5, 0.25, 0.1, 0.01, 0.001}
BAD_KEY = re.compile(r"^[a-z][a-z0-9_]{0,40}$")

# Every key a shapes report may have. Keys are schema, never data: a key outside this set blocks the report, and only
# these keys' words are exempt from the self-scan (a key computed from the report itself could exempt a client's word).
SCHEMA_KEYS = frozenset("""
alternates version anchor_sheet anchors blocks bound bound_by build build_secs built by_class by_sheet by_subclass capped
catalogue_cached cells cells_checked cells_shown census check_row check_rows checks_excluded circular_skipped class
closure_rows complete complexity cone_share cone_sheet_tokens cone_sheets consolidations consts context core_kinds cores
corkscrews day_count dcf decade density depgraph edges edges_by_kind edges_dropped error external failing_periods fails
families_left_out families_listed families_note families_per_formula_cell families_scrubbed findings first_decade
first_decades fn formula_cells formula_rows formulas found functions_distinct functions_top groups has_edges_table
has_terminal_date header_row hidden_sheets hints holds how identities identity_kind inactive_hidden key kind kinds
label_column label_vocab label_words lineage linked_files linked_rows longest_families longest_path longest_periods
max_depth max_formula_length max_rows message model_says n names_n names_user nodes non_excel_functions note notes
patterns patterns_more periodicities periodicity periods periods_checked present r1c1 rate_is_name reason reason_codes
reproduce residual result role roles roles_bound roles_unbound row row_graph row_offset rows rows_differing
rows_multi_pattern rows_parsed rows_shown rows_single_cell_break rows_upstream_total runs runs_in_multi_pattern_rows
runs_per_row scc_count scc_largest schema scope secs sheet sheet_graph sheets sheets_involved sheets_n
sheets_with_timeline shown signs sink_sheets site size_class source_sheets stage stages stages_failed stale state
statements stats status structural subclass subtotals tests timeline timing tolerance_relative top_families
total_column traceback type unbound unbound_roles units_column upstream usable verdict verdicts workbook""".split())

# Words of Python's, sqlite's and the OS's own error messages. An exception message keeps a word only if it is one of
# these, a word of the repo's vocabulary or an identifier of modelatlas/ code; every other word becomes "…" (a closed list:
# text from the model, or from a path outside the repo, cannot pass by being missing from a list of the model's words).
ERROR_WORDS = set("""
a an the is are was were be been being has have had not no nor or and but if then else of in on at to into onto from by
with without for as than that this these those it its there here which what when where while who whom only also can
cannot can't could would should must may might will shall do does did done got get gets given give takes take taking
expected expecting unexpected found find missing required optional argument arguments positional keyword keywords
object objects type types instance instances class attribute attributes method methods function functions module
callable iterable iterator subscriptable hashable unhashable mutable immutable frozen sequence mapping dict dictionary
list tuple set frozenset str string bytes bytearray int integer float complex bool boolean none nonetype null
number numbers real value values key keys item items index indices range slice length size empty full too many few
enough more less unpack unpacking assignment support supported unsupported operand operands operation operator
between instance comparison compare division modulo zero by math domain overflow underflow result results large
small recursion depth maximum minimum exceeded limit invalid valid literal base convert converted conversion
decode encode codec codecs byte position ordinal utf ascii charmap continuation start end data unicode character
characters char line lines column columns row rows table tables database databases sqlite sql syntax near
incomplete input output closed close cursor connection readonly read write attempt attempted unable open file files
directory directories path paths permission denied access errno error errors exception exceptions warning
locked busy timeout timed out disk i o io malformed image corrupt not a database no such constraint failed unique
json expecting delimiter property name enclosed double quotes extra memory changed during iteration mutated size furthest
changed generator stopped stop iteration call calling while python object recursion keyerror valueerror typeerror
indexerror attributeerror zerodivisionerror runtimeerror assertionerror oserror operationalerror
""".split())

# Excel's own function names (a name outside this list is not Excel's: it could be a user's, and is replaced by FN<k>)
BUILTIN_FUNCS = set("""ABS ACCRINT ACOS ACOSH ADDRESS AGGREGATE AND ARRAYTOTEXT ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE
AVERAGEA AVERAGEIF AVERAGEIFS BASE BETADIST BIN2DEC BITAND BITOR BITXOR BYCOL BYROW CEILING CEILING.MATH CELL CHAR CHIDIST
CHOOSE CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS COMBIN CONCAT CONCATENATE CONFIDENCE CORREL COS COSH COUNT COUNTA
COUNTBLANK COUNTIF COUNTIFS COUPDAYS COUPNUM COVAR CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS DAYS360 DB DCOUNT
DDB DEC2BIN DEGREES DGET DMAX DMIN DROP DSUM DURATION EDATE EFFECT EOMONTH ERF ERROR.TYPE EVEN EXACT EXP EXPAND EXPON.DIST
FACT FALSE FILTER FIND FINDB FIXED FLOOR FLOOR.MATH FORECAST FORECAST.LINEAR FORMULATEXT FREQUENCY FV FVSCHEDULE GCD GEOMEAN
GETPIVOTDATA GROWTH HLOOKUP HOUR HSTACK HYPERLINK IF IFERROR IFNA IFS INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR
ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT LAMBDA LARGE LCM LEFT LEFTB LEN
LENB LET LINEST LN LOG LOG10 LOGEST LOOKUP LOWER MAKEARRAY MAP MATCH MAX MAXA MAXIFS MDURATION MEDIAN MID MIDB MIN MINA
MINIFS MINUTE MIRR MMULT MOD MODE MONTH MROUND MULTINOMIAL N NA NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV
NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERVALUE OCT2DEC ODD ODDFPRICE OFFSET OR PDURATION PERCENTILE
PERCENTILE.INC PERCENTRANK PERMUT PI PMT POISSON POWER PPMT PRICE PROB PRODUCT PROPER PV QUARTILE QUARTILE.INC QUOTIENT
RADIANS RAND RANDARRAY RANDBETWEEN RANK RANK.EQ RATE RECEIVED REDUCE REPLACE REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN
ROUNDUP ROW ROWS RRI RSQ SCAN SEARCH SEARCHB SEC SECOND SEQUENCE SIGN SIN SINH SKEW SLN SLOPE SMALL SORT SORTBY SQRT
SQRTPI STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP STEYX SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ SUMX2MY2
SWITCH SYD T TAKE TAN TANH TBILLEQ TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TOCOL TODAY TOROW TRANSPOSE
TREND TRIM TRIMMEAN TRUE TRUNC TYPE UNICHAR UNIQUE UPPER VALUE VAR VAR.P VAR.S VARA VARP VDB VLOOKUP VSTACK WEEKDAY WEEKNUM
WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS XIRR XLOOKUP XMATCH XNPV XOR YEAR YEARFRAC YIELD YIELDDISC YIELDMAT ZTEST""".split())

# The fixed text of the reports. It is the same for every model and says nothing about one, so its words are not
# scanned for (a sheet called "Notes" must not block every report that has a "notes" heading).
FAMILIES_NOTE = "a lower bound: build_map keeps 3 patterns per row and counts the rest as more"
ANCHORS_NOTE = "the first call on a large model can take minutes; later runs read the saved catalogue"
MD_PROSE = ("Only the shapes of one model: how many formula families it has, how its sheets feed one another, which "
            "statement identities hold, fail or could not be tested, and where the code stopped. Sheets are S1, S2 in "
            "workbook order; formulas are R1C1 text with sheet names, defined names, strings and non-round numbers "
            "replaced. No file name, sheet name, label, defined name or value is written; values appear only as sign "
            "patterns and relative residual buckets.")
MD_CHECKLIST = ["scan this file once for anything that reads like a business or a person",
                "the file name you share does not carry the workbook's name",
                "this file comes from the shapes level (nothing from the full folder)",
                "if a blocked note exists for a model, nothing was written for it; do not look for it"]
STATIC_TEXT = " ".join([FAMILIES_NOTE, ANCHORS_NOTE, MD_PROSE, NOTICE, "Produced by Model Atlas", *MD_CHECKLIST, """
What this file contains. Review before sharing. Stage status skipped failed ok census complexity depgraph statements hints
anchors fixture exceptions site seconds token size class small medium large formulas families sheets cone share identities
holds fails unbound findings tests structure role src_only mid sink_only isolated check_rows timeline_less unknown
engine periodicity periods decade rows consts yes no none models blocked summary diagnostic run shapes only counts
statuses timings report test suite script exception message type context kind scope residual check row model says reasons
note blocks bound corkscrews subtotals consolidations lineage stale capped classes subclasses cores timing day count
functions longest top edges graph closure depth upstream shown cells header label units column linked files named ranges
hidden built in found usable reproduce catalogue saved read set with number of the and for per from not stage
no usable anchor stage did not complete skipped
annual quarterly monthly daily weekly semiannual direct offset active inactive collapsed scenario selected per_period single cumulative coverage test structural value_search leaf link aggregate
visible hidden veryhidden end start mid other actual ok failed skipped pass fail timeout scale_zero plug rejected_by_check
no_overlap no_numbers derived structure_mismatch no_candidate check_says_holds binding_rejected binding_unbound
check_and_binding_disagree corkscrew_fails model_fails_identity pattern_break plugged_balance_sheet stale withheld matched
self scan repo local python site packages path date furthest
ref div num null spill calc field busy connect getting this row headers totals all"""])
FULL_WARNING = "FULL REPORTS CONTAIN REAL NAMES AND TEXTS FROM THE MODEL. LOCAL USE ONLY. DO NOT SHARE ANYTHING UNDER diag/full/."


class Skip(Exception):
    """A stage that does not apply to this model (not a failure)."""


# ---- words and strings -------------------------------------------------------------------------------------------

def fold(s) -> str:
    """Unicode-folded: NFKC and casefold, so 'ＴＡＸ' and 'tax', 'Straße' and 'STRASSE' compare equal."""
    return unicodedata.normalize("NFKC", str(s)).casefold()


def norm(s: str) -> str:
    """Folded words only, single spaced: 'Total revenue (FY25)' -> 'total revenue fy25'. Any script's letters are
    words (a label in Chinese or Greek is not erased to nothing, which would leave it out of the self-scan)."""
    return re.sub(r"[\W_]+", " ", fold(s)).strip()


def tokens_of(s) -> list[str]:
    """The folded word tokens of a text (letters and digits of any script; '_' and punctuation split)."""
    return re.findall(r"[^\W_]+", fold(s))


def _words_of(text: str) -> set[str]:
    out = set()
    for w in re.findall(r"[^\W\d_]{3,}", fold(text)):
        out.update({w, w + "s", w + "es"})
        if w.endswith("s"):
            out.add(w[:-1])
    return out


def generic_vocabulary() -> tuple[set[str], set[str]]:
    """(words, phrases) the repo's own vocabulary uses: ontology roles, identities and label regexes, depgraph's
    classes and its subclass tables, Excel function names, and this module's fixed text. A model string that is one of
    these words, or exactly one of these phrases, is generic and not scanned for."""
    words, phrases = set(), set()
    try:
        from . import ontology
        for r in ontology.ROLES:
            words |= _words_of(r.name.replace("_", " ") + " " + r.vocab)
            phrases.add(norm(r.name))
        for i in ontology.IDENTITIES:
            words |= _words_of(i.key.replace("_", " "))
            phrases.add(norm(i.key))
        for b in ontology.BLOCKS:
            phrases.add(norm(b))
    except Exception:  # noqa: BLE001 (a missing vocabulary only makes the scan stricter)
        pass
    try:
        from . import depgraph
        for key, name in depgraph.CLASSES:
            phrases.add(norm(name))
            words |= _words_of(key)
            phrases.add(norm(key))
        for table in (depgraph.DISCOUNTING, depgraph.TERMINAL_SUB, depgraph.BRIDGE_SUB, depgraph.ASSUMPTIONS,
                      depgraph.CALCULATIONS):
            for row in table:
                words |= _words_of(row[0])
                phrases.add(norm(str(row[1])) if len(row) > 1 else "")
        for rx in (depgraph.TERMINAL, depgraph.BUILDUP, depgraph.CASHFLOW, depgraph.TIMELINE, depgraph.DEBT_FUNDING):
            words |= _words_of(rx.pattern)
    except Exception:  # noqa: BLE001
        pass
    for f in BUILTIN_FUNCS:
        words.add(f.lower())
        words |= set(tokens_of(f))
    import builtins
    words |= {n.lower() for n in dir(builtins) if n.endswith(("Error", "Exception", "Warning", "Exit", "Interrupt"))}
    words |= {"operationalerror", "databaseerror", "integrityerror", "programmingerror", "interfaceerror", "skip"}
    words |= {w for k in SCHEMA_KEYS for w in [k, *k.split("_")] if len(w) >= 3}
    words |= _words_of(STATIC_TEXT) | _words_of("check checks row rows")
    words |= {w for w in re.findall(r"[a-z0-9_]+", STATIC_TEXT.lower()) if len(w) >= 3}
    words |= {w for p in list(phrases) for w in p.split() if len(w) >= 3}
    phrases.discard("")
    return words, phrases


_GENERIC: tuple | None = None


def generic() -> tuple[set[str], set[str]]:
    global _GENERIC
    if _GENERIC is None:
        _GENERIC = generic_vocabulary()
    return _GENERIC


_CODE_WORDS: set | None = None


def code_words() -> set[str]:
    """The identifiers of the repo's own code (modelatlas/*.py, NAME tokens only: no comments, docstrings or strings),
    whole and split at '_'. Function and module names in a traceback site are these."""
    global _CODE_WORDS
    if _CODE_WORDS is None:
        out = set()
        for f in sorted((ROOT / CODE).glob("*.py")) + sorted((ROOT / "tests").glob("*.py")):
            try:
                with open(f, "rb") as fh:
                    for t in tokenize.tokenize(fh.readline):
                        if t.type == tokenize.NAME:
                            for w in [t.string, *t.string.split("_")]:
                                if len(w) >= 2:
                                    out.add(w.lower())
            except (OSError, SyntaxError, tokenize.TokenError):
                continue
        for f in list((ROOT / CODE).glob("*.py")) + list((ROOT / "tests").glob("*.py")):
            out.update(w.lower() for w in [f.stem, *f.stem.split("_")] if w)
        _CODE_WORDS = out
    return _CODE_WORDS


def safe_words() -> set[str]:
    """Words an exception message may keep: the repo's vocabulary, its code's identifiers and error-message words."""
    return generic()[0] | code_words() | ERROR_WORDS | {"repo", "local", "python", "site", "packages", "other"}


# ---- removing the machine from a string ---------------------------------------------------------------------------

def _machine_words() -> list[str]:
    out = []
    for v in (str(Path.home()), os.environ.get("USERNAME"), os.environ.get("USER"), platform.node(),
              os.environ.get("COMPUTERNAME")):
        if v and len(v) >= 3:
            out.append(v)
            out.append(v.replace("\\", "/"))
            if v == platform.node():
                out.append(v.split(".")[0])
    return sorted(set(out), key=len, reverse=True)


def scrub_env(s: str) -> str:
    """Repo paths made relative; python and site-packages dirs, the home directory, user and machine names removed."""
    if not isinstance(s, str):
        return s
    for root in (str(ROOT), ROOT.as_posix()):
        s = s.replace(root + os.sep, "").replace(root + "/", "").replace(root, "<repo>")
    s = re.sub(r"[^\s\"']*[\\/]site-packages[\\/]", "<site-packages>/", s)
    for base in {sys.prefix, sys.base_prefix, sys.exec_prefix}:
        if base:
            s = s.replace(base, "<python>").replace(base.replace("\\", "/"), "<python>")
    for m in _machine_words():
        s = re.sub(re.escape(m), "<local>", s, flags=re.I)
    return s


def code_path(filename: str) -> str:
    """A traceback frame's file: repo-relative, or under <python>/<site-packages>, or '<other>/<file name>' (a path
    outside these, e.g. a deal folder a script was run from, is never written)."""
    try:
        return Path(filename).resolve().relative_to(ROOT).as_posix()
    except (ValueError, OSError):
        pass
    s = scrub_env(filename).replace("\\", "/")
    if s.startswith(("<python>", "<site-packages>")) or re.fullmatch(r"<[\w .-]+>", s):
        return s
    return "<other>/" + Path(filename.replace("\\", "/")).name


def ascii_only(s: str) -> str:
    """Non-ASCII characters replaced by '?' (code and error text are ASCII; anything else came from data), except the
    '…' this module writes for a replaced string."""
    return "".join(c if ord(c) < 128 or c == "…" else "?" for c in s)


# ---- the scrubber for formulas and messages -------------------------------------------------------------------------

_TOK = re.compile(r"""
   (?P<str>"(?:[^"]|"")*")
 | (?P<qsheet>'(?:[^']|'')+'!)
 | (?P<err>\#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|BUSY!|CONNECT!
            |This\ Row|Headers|Data|Totals|All)(?![\w.]))
 | (?P<sheet>[A-Za-z_\u0080-￿][\w.\u0080-￿]*(?::[A-Za-z_][\w.\u0080-￿]*)?!)
 | (?P<rc>R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?(?![\w(]))
 | (?P<r1>R(?:\[-?\d+\]|\d+)(?![\w(\[!]))
 | (?P<c1>C(?:\[-?\d+\]|\d+)(?![\w(\[!]))
 | (?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+)
 | (?P<id>[A-Za-z_\\\u0080-￿][\w.\u0080-￿]*)
""", re.X)
_SAFE_CHARS = set(" +-*/^&=<>,():;{}[]$%@!#.\\|~")


class Anon:
    """The tokens of one model: sheets S1.., defined names NAME<k>, non-Excel functions FN<k>, other files' sheets X<k>."""

    def __init__(self, sheets: list[str], names: list[str]):
        self.sheets = {n: f"S{i}" for i, n in enumerate(sheets, 1)}
        self._sheets_lower = {n.lower(): t for n, t in self.sheets.items()}
        self.names = {n.lower(): f"NAME{i}" for i, n in enumerate(sorted(set(names), key=str.lower), 1)}
        self.funcs: dict[str, str] = {}
        self.other: dict[str, str] = {}
        self._memo: dict[str, str] = {}

    def sheet_token(self, name: str) -> str:
        t = self.sheets.get(name) or self._sheets_lower.get(name.lower())
        if t:
            return t
        return self.other.setdefault(name.lower(), f"X{len(self.other) + 1}")

    def func(self, name: str) -> str:
        n = re.sub(r"^(?:_xl[a-z]+\.)+", "", name, flags=re.I).upper()
        if n in BUILTIN_FUNCS:
            return n
        return self.funcs.setdefault(n, f"FN{len(self.funcs) + 1}")

    @staticmethod
    def number(text: str) -> str:
        try:
            v = float(text)
        except ValueError:
            return "#"
        if v == int(v) and abs(v) <= 1000:
            return text
        return text if v in KEEP_DECIMALS else "#"

    def pattern(self, text: str) -> str:
        """An R1C1 formula with everything that could be the client's replaced; one pass over the tokens."""
        hit = self._memo.get(text)
        if hit is not None:
            return hit
        out, pos = [], 0
        for m in _TOK.finditer(text):
            for ch in text[pos:m.start()]:
                out.append(ch if ch in _SAFE_CHARS else "?")
            pos, k, t = m.end(), m.lastgroup, m.group()
            if k == "str":
                out.append('"…"')
            elif k == "qsheet":
                out.append(self.sheet_token(t[1:-2].replace("''", "'")) + "!")
            elif k == "sheet":
                out.append(":".join(self.sheet_token(p) for p in t[:-1].split(":")) + "!")
            elif k == "num":
                out.append(self.number(t))
            elif k == "id":
                if text[m.end():m.end() + 1] == "(":
                    out.append(self.func(t))
                elif t.upper() in ("TRUE", "FALSE"):
                    out.append(t.upper())
                else:
                    out.append(self.names.get(t.lower()) or self.names.setdefault(t.lower(), f"NAME{len(self.names) + 1}"))
            else:  # error literal, R1C1 reference
                out.append(t)
        for ch in text[pos:]:
            out.append(ch if ch in _SAFE_CHARS else "?")
        res = "".join(out)
        self._memo[text] = res
        return res

    def message(self, text) -> str:
        """An exception message, projected: quoted text, paths, dates and long numbers replaced, sheet and defined names
        tokenised, non-ASCII dropped, and then every word that is not a known safe word (safe_words) replaced by '…'."""
        return safe_message(str(text), self)


def safe_message(s: str, anon: "Anon | None" = None) -> str:
    s = scrub_env(s)
    s = re.sub(r"'[^'\n]*'|\"[^\"\n]*\"", "'…'", s)
    s = re.sub(r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?", "<date>", s)
    s = re.sub(r"(?:[A-Za-z]:)?(?:[\\/][^\\/\s'\"]+){2,}[\\/]?", "<path>", s)     # any other path
    if anon is not None:
        for n in sorted(anon.sheets, key=len, reverse=True):
            if len(n) >= 2:
                s = re.sub(r"(?<![^\W_])" + re.escape(n) + r"(?![^\W_])", anon.sheets[n], s, flags=re.I)
        for n, t in sorted(anon.names.items(), key=lambda kv: -len(kv[0])):
            s = re.sub(r"(?<![\w.])" + re.escape(n) + r"(?![\w.])", t, s, flags=re.I)
    s = re.sub(r"\d+(?:,\d{3})+(?:\.\d+)?|\d*\.\d+(?:[eE][+-]?\d+)?|\d{4,}", "#", s)
    s = re.sub(r"[^\x00-\x7f…]+", "…", s)
    ok = safe_words()

    def word(m):
        w = m.group()
        if (re.fullmatch(r"(?:S|X|NAME|FN|M)\d+|R(?:\d+|\[-?\d+\])?C(?:\d+|\[-?\d+\])?|\d+", w)
                or (re.fullmatch(r"\$?[A-Z]{1,3}\$?\d{1,7}", w) and s[max(0, m.start() - 1):m.start()] == "!")):
            return w
        parts = [p for p in re.split(r"_+", w.lower()) if p]
        if w.lower() in ok or (parts and all(p in ok or p.isdigit() for p in parts)):
            return w
        return "…"
    s = re.sub(r"\$?[A-Za-z_][A-Za-z0-9_]*", word, s)
    s = re.sub(r"…(?:[ .]?…)+", "…", s)
    return s[:400]


# ---- the self-scan ----------------------------------------------------------------------------------------------

class Scanner:
    """Looks for a model's own strings in text. Three kinds: `subs` (one-word names: sheets, defined names, file name
    parts; matched as substrings of the text once the repo's generic words are blanked out of it, so a sheet called
    "Cons" does not block the word "consolidations"), `grams` (labels and texts of 2+ words or 8+ characters; matched as
    word sequences within one string, generic words and all: "total revenue" is a hit even though both words are
    generic) and `terms` (every word with 3+ letters of any text of the model that is not a generic
    word, e.g. the "Wombat" of a one-word label, or the "xqz" of "Revenue XQZ"; matched as whole words)."""

    def __init__(self, subs: set[str], grams: set[str], extra_words=(), terms: set[str] = frozenset()):
        self.raw = (set(subs), set(grams), set(terms))
        words, phrases = generic()
        self.words = words | set(extra_words)
        self.subs = {s for s in subs if len(s) >= 4 and s not in self.words and s not in phrases}
        self.grams = {g for g in grams if g and not (g in self.words or g in phrases)}
        self.terms = {t for t in terms if t not in self.words}
        self.max_n = min(12, max([len(g.split()) for g in self.grams] + [1]))

    def with_words(self, extra) -> "Scanner":
        """The same strings with more words exempt (the report's schema keys: the same for every model)."""
        return Scanner(self.raw[0], self.raw[1], extra_words=extra, terms=self.raw[2])

    @classmethod
    def union(cls, scanners) -> "Scanner":
        out = cls(set(), set())
        for s in scanners:
            out.subs |= s.subs
            out.grams |= s.grams
            out.terms |= s.terms
            out.words |= s.words
        out.max_n = max([s.max_n for s in scanners] + [1])
        return out

    def hits_many(self, texts, terms=True) -> list[str]:
        """Each string on its own: two neighbouring values must not read as one phrase."""
        found = set()
        for t in texts:
            found.update(self.hits(t, terms=terms))
        return sorted(found)

    def hits(self, text: str, grams=True, terms=True) -> list[str]:
        low = fold(text)
        rest = re.sub(r"[^\W_]+", lambda m: " " if m.group() in self.words else m.group(), low)
        found = {s for s in self.subs if s in rest}
        words = norm(text).split()
        if terms:
            found |= {w for w in words if w in self.terms}
        if grams:
            for size in range(1, self.max_n + 1):
                for i in range(len(words) - size + 1):
                    g = " ".join(words[i:i + size])
                    if g in self.grams:
                        found.add(g)
        return sorted(found)


def _term_words(s) -> set[str]:
    """The words of a text worth scanning for on their own: 3+ letters (a code like 'XQZ' too), not a pure number."""
    return {w for w in norm(str(s or "")).split() if len(re.findall(r"[^\W\d_]", w)) >= 3}


def build_scanner(db, folder: str, registry: list[dict]) -> Scanner:
    """The forbidden strings of one model.db (see the module docstring); the labels and texts are read, never written."""
    subs, grams, terms = set(), set(), set()

    def sub(s, minlen=4):
        s = str(s or "").strip()
        n = norm(s)
        if len(s) >= minlen and len(n) >= 4:
            if " " in n:
                grams.add(n)
                subs.add(n.replace(" ", ""))   # 'cash flows' also as written without the space
            else:
                subs.add(n)
        terms.update(_term_words(s))

    def gram(s):
        if not isinstance(s, str):
            return
        terms.update(_term_words(s))
        for piece in [s, *re.split(r"[,;|()\[\]{}'\"`:/\\]+", s)]:   # also each part of 'Revenue, net (XQZ)'
            n = norm(piece)
            if " " in n or len(n) >= 8:
                grams.add(n)

    have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    for (s, summ) in db.execute("SELECT sheet, summary FROM sheets"):
        sub(s)
        gram(s)
        gram(summ)
    for lab, sec, units in db.execute("SELECT label, section, units FROM rows"):
        for v in (lab, sec, units):
            gram(v)
    for (v,) in db.execute("SELECT DISTINCT value FROM cells WHERE typeof(value) = 'text'"):
        gram(v)       # headers, titles, notes: any text of the model, not only the label column
    for (n, ref) in db.execute("SELECT name, ref FROM names"):
        if n and not str(n).startswith("_xlnm"):
            sub(n, 5)
            terms.update(_term_words(n))
        if isinstance(ref, str):
            for part in re.findall(r"\[([^\]]+)\]", ref):   # a linked file named in a name's reference
                sub(part)
                sub(re.sub(r"\.xls\w*$", "", part, flags=re.I))
    if "extbooks" in have:
        for t, f, sh in db.execute("SELECT target, filename, sheets FROM extbooks"):
            for v in (t, f):
                sub(v)
                sub(re.sub(r"\.xls\w*$", "", str(v or ""), flags=re.I))
                for tok in re.split(r"[_\-\s.\\/]+", str(v or "")):
                    sub(tok)
            try:
                parts = json.loads(sh) if isinstance(sh, str) and sh.lstrip().startswith("[") else re.split(r"[;,|]", str(sh or ""))
            except ValueError:
                parts = re.split(r"[;,|\[\]\"]", str(sh or ""))
            for part in parts:
                sub(part)
    if "extrefs" in have:
        for (s,) in db.execute("SELECT DISTINCT ext_sheet FROM extrefs"):
            sub(s)
    if "extcells" in have:
        for (v,) in db.execute("SELECT DISTINCT value FROM extcells WHERE typeof(value) = 'text'"):
            gram(v)
    stem = re.sub(r"__[0-9a-f]{8}$", "", folder)
    for v in (folder, stem):
        sub(v)
    for tok in re.split(r"[_\-\s.]+", stem):
        sub(tok)
    for rec in registry:
        for key in ("filename", "source_name", "target_name", "project_name"):
            v = rec.get(key)
            if v:
                sub(v)
                gram(v)
                sub(re.sub(r"\.xls\w*$", "", str(v), flags=re.I))
                for tok in re.split(r"[_\-\s.]+", str(v)):
                    sub(tok)
        for part in re.split(r"[\\/]+", str(rec.get("source_path") or "")):   # the deal folders above the file
            sub(part)
        ident = rec.get("identity_json")
        try:
            vals = strings_of(json.loads(ident)) if isinstance(ident, str) and ident.strip() else []
        except ValueError:
            vals = [ident]
        for v in vals + [rec.get("valuation_date")]:   # not `error`: an app traceback's words are code, not the client's
            if isinstance(v, str) and v.strip():
                gram(v)
                sub(v)
    for m in _machine_words():
        sub(m, 3)
    return Scanner(subs, grams, terms=terms)


def strings_of(o, out=None) -> list[str]:
    """Every string VALUE in a JSON structure (keys are checked separately: they must be plain schema words)."""
    out = [] if out is None else out
    if isinstance(o, dict):
        for v in o.values():
            strings_of(v, out)
    elif isinstance(o, (list, tuple)):
        for v in o:
            strings_of(v, out)
    elif isinstance(o, str):
        out.append(o)
    return out


def bad_keys(o, path="") -> list[str]:
    """Keys that are not in the fixed schema (SCHEMA_KEYS). A key is never data."""
    out = []
    if isinstance(o, dict):
        for k, v in o.items():
            if not (BAD_KEY.match(str(k)) and str(k) in SCHEMA_KEYS):
                out.append(f"{path}/{k}")
            out += bad_keys(v, f"{path}/{k}")
    elif isinstance(o, (list, tuple)):
        for v in o:
            out += bad_keys(v, path)
    return out


# ---- a model, opened ---------------------------------------------------------------------------------------------

def folder_token(folder: str) -> str:
    """M<sha8> from the folder's __<sha8> suffix; a folder without one gets the sha1 of its name."""
    m = re.search(r"__([0-9a-f]{8})$", folder)
    return "M" + (m.group(1) if m else hashlib.sha1(folder.encode("utf-8")).hexdigest()[:8])


def registry_rows(out_root: Path, folder: str, sha8: str | None) -> list[dict] | None:
    """The registry rows of this model, from out/registry.db and from out/atlas.db (the workbooks dropped on the
    dashboard, modelatlas/library.py): [] without either; None when one is there but cannot be read."""
    hit = []
    for name in ("registry.db", "atlas.db"):
        reg = out_root / name
        if not reg.exists():
            continue
        try:
            db = rodb.connect(reg)
            try:
                cur = db.execute("SELECT * FROM files")   # every column there is (an older registry lacks some)
                cols = [c[0] for c in cur.description]
                rows = [dict(zip(cols, r)) for r in cur]
            finally:
                db.close()
        except Exception:  # noqa: BLE001 (a registry that is there but cannot be read: its names cannot be scanned for)
            return None
        for r in rows:
            names = {Path(str(r.get("out_dir") or "").replace("\\", "/")).name,
                     Path(str(r.get("db_path") or "").replace("\\", "/")).parent.name}
            if folder in names or (sha8 and str(r.get("sha256") or "").startswith(sha8)):
                if r.get("source_path"):
                    r["source_name"] = Path(str(r["source_path"]).replace("\\", "/")).name
                if name == "atlas.db":
                    # the upload's own name, not the folders above it: those are this repo's uploads/<sha12>/,
                    # not a deal folder, and their words ('python', 'uploads') would block every report
                    r["source_path"] = None
                hit.append(r)
    return hit


class Ctx:
    """What the stages share for one model.db."""

    def __init__(self, db_path: Path, out_root: Path, args):
        self.db_path = Path(db_path)
        self.folder = self.db_path.parent.name
        self.token = folder_token(self.folder)
        m = re.search(r"__([0-9a-f]{8})$", self.folder)
        self.args = args
        self.db = rodb.connect(self.db_path)
        try:
            self._open(out_root, m)
        except BaseException:
            self.close()   # a model.db that cannot be read: the caller records it; the file is not left open
            raise

    def _open(self, out_root: Path, m):
        reg = registry_rows(out_root, self.folder, m.group(1) if m else None)
        self.registry_unreadable = reg is None   # fail closed: the report is blocked (gate)
        self.registry = reg or []
        self.sheet_names = [r[0] for r in self.db.execute("SELECT sheet FROM sheets ORDER BY rowid")]
        names = [r[0] for r in self.db.execute("SELECT name FROM names")]
        self.anon = Anon(self.sheet_names, names)
        self.layouts = {}
        for s, lay in self.db.execute("SELECT sheet, layout FROM sheets"):
            try:
                self.layouts[s] = json.loads(lay) if lay else {}
            except ValueError:
                self.layouts[s] = {}
        self.scanner = build_scanner(self.db, self.folder, self.registry)
        self.have = {r[0] for r in self.db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        self.stash: dict = {}   # raw results for later stages (never written)
        self.full: dict = {}    # real names and texts (level full only)
        self.errors: list[dict] = []

    def tok(self, sheet: str) -> str:
        return self.anon.sheet_token(sheet)

    def close(self):
        try:
            self.db.close()
        except Exception:  # noqa: BLE001
            pass


# ---- errors ------------------------------------------------------------------------------------------------------

def error_record(ctx: "Ctx | None", exc: BaseException, stage: str) -> dict:
    """Type, scrubbed message, scrubbed traceback and the site (the innermost frame in modelatlas/ or tests/)."""
    frames = traceback.extract_tb(exc.__traceback__)
    site = None
    for fr in reversed(frames):
        try:
            rel = Path(fr.filename).resolve().relative_to(ROOT).as_posix()
        except (ValueError, OSError):
            continue
        if rel.startswith(("modelatlas/", "tests/")) and rel != "modelatlas/diagnose.py":
            site = f"{rel}:{fr.lineno} in {fr.name}"
            break
    if site is None and frames:
        fr = frames[-1]
        site = ascii_only(f"{code_path(fr.filename)}:{fr.lineno} in {fr.name}")
    if ctx is not None:
        msg = ctx.anon.message(exc)
        if ctx.scanner.hits(msg):
            msg = "<message withheld: it matched the self-scan>"
    else:
        msg = safe_message(str(exc))

    def in_code(fr):  # the source line of the repo's own code is code; nothing else's is quoted
        try:
            return Path(fr.filename).resolve().relative_to(ROOT / CODE) is not None
        except (ValueError, OSError):
            return False
    tb = [ascii_only(f"{code_path(fr.filename)}:{fr.lineno} in {fr.name}" + (f" | {fr.line}" if fr.line and in_code(fr) else ""))
          for fr in frames]
    return {"stage": stage, "type": type(exc).__name__, "message": msg, "site": site or "unknown", "traceback": tb}


def run_stage(ctx: Ctx, name: str, fn) -> dict:
    t0 = time.time()
    try:
        res = fn(ctx)
        return {"status": "ok", "secs": round(time.time() - t0, 2), "result": res}
    except Skip as e:
        return {"status": "skipped", "secs": round(time.time() - t0, 2), "reason": str(e)}
    except Exception as e:  # noqa: BLE001 (one failure never stops the run)
        rec = error_record(ctx, e, name)
        ctx.errors.append(rec)
        ctx.full.setdefault(name, {})["raw_error"] = {"type": type(e).__name__, "message": str(e),
                                                      "traceback": traceback.format_exc()}
        return {"status": "failed", "secs": round(time.time() - t0, 2), "error": rec}


# ---- helpers on the model ----------------------------------------------------------------------------------------

def q(db, sql, *a):
    return db.execute(sql, a).fetchall()


def decade(iso) -> str | None:
    m = re.match(r"(\d{3})\d", str(iso or ""))
    return m.group(1) + "0s" if m else None


def size_class(formulas: int) -> str:
    return "small" if formulas < SMALL else "medium" if formulas < MEDIUM else "large"


_PATTERN = re.compile(r"(=.*?) x(\d+) \(([A-Z]+)(\d+)(?:\.\.([A-Z]+)(\d+))?\)(?:; |$)")
_MORE = re.compile(r"\+(\d+) more patterns")


def pattern_entries(text: str) -> tuple[list[tuple[str, int]], int]:
    """[(R1C1 pattern, cells)] from a rows.patterns text, and how many more patterns it says it left out."""
    ents = [(m.group(1), int(m.group(2))) for m in _PATTERN.finditer(text or "")]
    more = _MORE.search(text or "")
    return ents, int(more.group(1)) if more else 0


def sheet_graph(db, sheets: list[str]):
    """Data-flow graph between sheets from the live edges: an edge a -> b when b reads a row of a."""
    live = ",".join(repr(k) for k in LIVE)
    flows = q(db, f"SELECT DISTINCT dst_sheet, src_sheet FROM edges WHERE kind IN ({live}) AND src_sheet <> dst_sheet")
    idx = {s: i for i, s in enumerate(sheets)}
    adj = defaultdict(set)
    for a, b in flows:
        if a in idx and b in idx:
            adj[idx[a]].add(idx[b])
    return adj, len(sheets)


def scc(n: int, adj) -> list[list[int]]:
    """Strongly connected components (iterative Tarjan)."""
    index, low, on, stack, comps, counter = {}, {}, set(), [], [], [0]
    for root in range(n):
        if root in index:
            continue
        index[root] = low[root] = counter[0]
        counter[0] += 1
        stack.append(root)
        on.add(root)
        work = [(root, iter(sorted(adj.get(root, ()))))]
        while work:
            v, it = work[-1]
            for w in it:
                if w not in index:
                    index[w] = low[w] = counter[0]
                    counter[0] += 1
                    stack.append(w)
                    on.add(w)
                    work.append((w, iter(sorted(adj.get(w, ())))))
                    break
                if w in on:
                    low[v] = min(low[v], index[w])
            else:
                work.pop()
                if work:
                    p = work[-1][0]
                    low[p] = min(low[p], low[v])
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = stack.pop()
                        on.discard(w)
                        comp.append(w)
                        if w == v:
                            break
                    comps.append(comp)
    return comps


def longest_path(adj, comps) -> int:
    """Edges on the longest path through the condensation of the sheet graph (a cycle counts as one step)."""
    comp_of = {v: i for i, c in enumerate(comps) for v in c}
    cadj = defaultdict(set)
    for v, ws in adj.items():
        for w in ws:
            if comp_of[v] != comp_of[w]:
                cadj[comp_of[v]].add(comp_of[w])
    memo: dict[int, int] = {}
    for start in range(len(comps)):
        stack = [start]
        while stack:
            c = stack[-1]
            if c in memo:
                stack.pop()
                continue
            pend = [d for d in cadj.get(c, ()) if d not in memo]
            if pend:
                stack.extend(pend)
            else:
                memo[c] = 1 + max([memo[d] for d in cadj.get(c, ())], default=-1)
                stack.pop()
    return max(memo.values(), default=0)


def cell_row(cell: str) -> tuple[str, int] | None:
    if not cell or "!" not in cell:
        return None
    sh, addr = cell.rsplit("!", 1)
    m = re.search(r"(\d+)$", addr.replace("$", ""))
    return (sh.strip("'").replace("''", "'"), int(m.group(1))) if m else None


def sign_runs(values) -> str:
    """+ - 0 per period, run-length coded: '+x12 0x3 -x2'. t = text, _ = empty. Magnitudes are never kept."""
    out = []
    for v in values:
        if v is None or v == "":
            s = "_"
        elif isinstance(v, bool) or not isinstance(v, (int, float)):
            s = "t"
        else:
            s = "+" if v > 0 else "-" if v < 0 else "0"
        if out and out[-1][0] == s:
            out[-1][1] += 1
        else:
            out.append([s, 1])
    return " ".join(f"{s}x{n}" for s, n in out[:40])


def row_values(ctx: Ctx, sheet: str, row: int, limit: int = 400) -> list:
    lay = ctx.layouts.get(sheet) or {}
    if lay.get("tl_first") and lay.get("tl_last"):
        rows = q(ctx.db, "SELECT value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ? ORDER BY col",
                 sheet, row, lay["tl_first"], lay["tl_last"])
    else:
        rows = q(ctx.db, "SELECT value FROM cells WHERE sheet=? AND row=? ORDER BY col", sheet, row)
    return [r[0] for r in rows[:limit]]


def residual_bucket(res, scale) -> str | None:
    """The largest residual relative to the largest value of the rows it is made of, as a bucket (never the number)."""
    if res is None:
        return None
    if res == 0:
        return "0"
    if not scale:
        return "scale_zero"
    rel = abs(res) / scale
    for lim, name in ((1e-9, "<1e-9"), (1e-6, "<1e-6"), (1e-3, "<1e-3"), (1e-1, "<1e-1")):
        if rel < lim:
            return name
    return ">=1e-1"


def reason_codes(why: str) -> list[str]:
    """What kind of reason an unbound identity gives, from its text. Only the codes are kept at level shapes."""
    codes = set()
    for part in (why or "").split("; "):
        w = part.lower()
        if "periodicity differs" in w:
            c = "periodicity"
        elif "plug" in w:
            c = "plug"
        elif "binding rejected" in w:
            c = "rejected_by_check"
        elif "capped" in w or "stopped after" in w:
            c = "capped"
        elif "no period with a number" in w:
            c = "no_overlap"
        elif "zero" in w or "no numbers" in w:
            c = "no_numbers"
        elif "derived" in w:
            c = "derived"
        elif "not a sum" in w or "not a link" in w:
            c = "structure_mismatch"
        elif ("not found" in w or "not among" in w or "too few" in w or w.startswith("no ") or ": no " in w):
            c = "no_candidate"
        else:
            c = "other"
        codes.add(c)
    return sorted(codes)


def vocab_names(label) -> list[str]:
    """The ontology roles whose label vocabulary the label matches: the vocabulary's names, never the label."""
    try:
        from . import ontology
        return sorted({r.name for r in ontology.ROLES if ontology.vocab_hit(r.name, label or "")})[:4]
    except Exception:  # noqa: BLE001
        return []


def row_shape(ctx: Ctx, sheet: str, row: int, role=None, bound_by=None, base=None) -> dict:
    """A row, as shape only: patterns (scrubbed), counts, sign pattern, the vocabulary that matched its label."""
    r = q(ctx.db, "SELECT label, n_formula, n_const, patterns FROM rows WHERE sheet=? AND row=?", sheet, row)
    label, nf, nc, pats = r[0] if r else ("", 0, 0, "")
    ents, more = pattern_entries(pats)
    lay = ctx.layouts.get(sheet) or {}
    return {"role": role, "bound_by": bound_by, "sheet": ctx.tok(sheet), "row": row,
            "row_offset": row - base if base is not None else None, "formulas": nf, "consts": nc,
            "label_words": len((label or "").split()), "label_vocab": vocab_names(label),
            "periodicity": lay.get("periodicity"),
            "patterns": [{"r1c1": ctx.anon.pattern(p)[:300], "cells": n} for p, n in ents],
            "patterns_more": more, "signs": sign_runs(row_values(ctx, sheet, row))}


# ---- stage: census -----------------------------------------------------------------------------------------------

def stage_census(ctx: Ctx) -> dict:
    db = ctx.db
    per = {r[0]: r for r in q(db, "SELECT sheet, COUNT(*), SUM(n_formula), SUM(n_const) FROM rows GROUP BY sheet")}
    sheets = []
    for (name, state) in q(db, "SELECT sheet, state FROM sheets ORDER BY rowid"):
        lay = ctx.layouts.get(name) or {}
        r = per.get(name, (name, 0, 0, 0))
        sheets.append({"sheet": ctx.tok(name), "state": state, "rows": r[1], "formulas": r[2] or 0, "consts": r[3] or 0,
                       "periodicity": lay.get("periodicity"), "periods": lay.get("periods") or 0,
                       "header_row": bool(lay.get("header_row")), "label_column": bool(lay.get("label_col")),
                       "units_column": bool(lay.get("units_col")), "first_decade": decade(lay.get("period_start"))})
    names = [r[0] for r in q(db, "SELECT name FROM names")]
    ext = {"present": "extbooks" in ctx.have, "linked_files": 0, "linked_rows": 0}
    if "extbooks" in ctx.have:
        ext["linked_files"] = q(db, "SELECT COUNT(*) FROM extbooks")[0][0]
    if "extrefs" in ctx.have:
        ext["linked_rows"] = q(db, "SELECT COUNT(*) FROM extrefs")[0][0]
    res = {"sheets_n": len(sheets), "hidden_sheets": sum(1 for s in sheets if s["state"] != "visible"),
           "rows": sum(s["rows"] for s in sheets), "formulas": sum(s["formulas"] for s in sheets),
           "consts": sum(s["consts"] for s in sheets), "names_n": len(names),
           "names_user": sum(1 for n in names if not str(n).startswith("_xlnm")), "external": ext,
           "has_edges_table": "edges" in ctx.have, "sheets": sheets}
    ctx.stash["census"] = res
    ctx.full["census"] = {"sheet_names": {ctx.tok(n): n for n in ctx.sheet_names}, "defined_names": names}
    return res


# ---- stage: anchors (the DCF value cells; the first call on a big model takes minutes) -----------------------------

def stage_anchors(ctx: Ctx) -> dict:
    from . import depgraph
    cached = (ctx.db_path.parent / "catalogue.json").exists()
    t0 = time.time()
    anc = depgraph.anchors(str(ctx.db_path))
    ctx.stash["anchors"] = anc
    ctx.stash["top_anchor"] = next((a for a in anc if a["ok"]), None)
    ctx.full["anchors"] = {"cells": [a["cell"] for a in anc], "labels": [a["label"] for a in anc]}
    return {"found": len(anc), "usable": sum(1 for a in anc if a["ok"]), "reproduce": sum(1 for a in anc if a["matches"]),
            "secs": round(time.time() - t0, 2), "catalogue_cached": cached, "note": ANCHORS_NOTE}


# ---- stage: complexity profile -----------------------------------------------------------------------------------------

def stage_complexity(ctx: Ctx) -> dict:
    db = ctx.db
    census = ctx.stash.get("census") or stage_census(ctx)
    fam_cells: Counter = Counter()
    n_rows = hidden_more = multi = runs_multi = breaks = 0
    hist: Counter = Counter()
    for pats in (r[0] for r in db.execute("SELECT patterns FROM rows WHERE n_formula > 0")):
        ents, more = pattern_entries(pats)
        n_rows += 1
        runs = len(ents) + (1 if more else 0)
        hidden_more += more
        hist["4+" if runs >= 4 else str(runs)] += 1
        for p, n in ents:
            fam_cells[p] += n
        if runs > 1:
            multi += 1
            runs_multi += runs
            if any(n == 1 for _, n in ents):
                breaks += 1
    scrubbed: Counter = Counter()
    for p, n in fam_cells.items():
        scrubbed[ctx.anon.pattern(p)] += n
    formulas = census["formulas"]
    top_fam = [{"r1c1": p[:300], "cells": n} for p, n in scrubbed.most_common(12)]
    longest = [{"r1c1": p[:300], "cells": n} for p, n in sorted(scrubbed.items(), key=lambda kv: -len(kv[0]))[:5]]

    fn: Counter = Counter()
    maxlen = 0
    strlit = re.compile(r'"(?:[^"]|"")*"')
    call = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")
    for (f,) in db.execute("SELECT formula FROM cells WHERE formula IS NOT NULL"):
        maxlen = max(maxlen, len(f))
        for m in call.findall(strlit.sub('""', f)):
            fn[ctx.anon.func(m)] += 1

    prof: dict = {"formula_cells": formulas, "formula_rows": n_rows, "size_class": size_class(formulas),
                  "families_listed": len(fam_cells), "families_scrubbed": len(scrubbed), "families_left_out": hidden_more,
                  "families_note": FAMILIES_NOTE,
                  "families_per_formula_cell": round(len(fam_cells) / formulas, 5) if formulas else None,
                  "rows_multi_pattern": multi, "runs_in_multi_pattern_rows": runs_multi,
                  "rows_single_cell_break": breaks, "runs_per_row": [{"runs": k, "rows": v} for k, v in sorted(hist.items())],
                  "top_families": top_fam, "longest_families": longest, "max_formula_length": maxlen,
                  "functions_distinct": len(fn), "functions_top": [{"fn": k, "n": v} for k, v in fn.most_common(30)],
                  "non_excel_functions": len(ctx.anon.funcs)}
    sheets = ctx.sheet_names
    prof["edges_by_kind"], prof["sheet_graph"], prof["row_graph"] = [], None, None
    if "edges" in ctx.have:
        prof["edges_by_kind"] = [{"kind": k, "n": n} for k, n in q(db, "SELECT kind, COUNT(*) FROM edges GROUP BY kind ORDER BY 2 DESC")]
        adj, n = sheet_graph(db, sheets)
        comps = scc(n, adj)
        n_edges = sum(len(v) for v in adj.values())
        indeg, outdeg = Counter(), Counter()
        for a, bs in adj.items():
            for b in bs:
                outdeg[a] += 1
                indeg[b] += 1
        big = max((len(c) for c in comps), default=0)
        prof["sheet_graph"] = {"nodes": n, "edges": n_edges, "density": round(n_edges / (n * (n - 1)), 4) if n > 1 else 0.0,
                               "longest_path": longest_path(adj, comps),
                               "scc_count": sum(1 for c in comps if len(c) > 1), "scc_largest": big if big > 1 else 0,
                               "source_sheets": sum(1 for i in range(n) if indeg[i] == 0),
                               "sink_sheets": sum(1 for i in range(n) if outdeg[i] == 0)}
        roles = {}
        for i in range(n):
            roles[sheets[i]] = ("src_only" if indeg[i] == 0 and outdeg[i] else "sink_only" if outdeg[i] == 0 and indeg[i]
                                else "isolated" if not indeg[i] else "mid")
        ctx.stash["sheet_roles"] = roles
        anchor = ctx.stash.get("top_anchor")
        where = cell_row(anchor["cell"]) if anchor else None
        if where:
            prof["row_graph"] = row_cone(ctx, where, n)
    tl = [s for s in census["sheets"] if s["periods"]]
    prof["timeline"] = {"sheets_with_timeline": len(tl), "periodicities": sorted({s["periodicity"] for s in tl if s["periodicity"]}),
                        "longest_periods": max([s["periods"] for s in tl], default=0),
                        "first_decades": sorted({s["first_decade"] for s in tl if s["first_decade"]})}
    anc = ctx.stash.get("anchors") or []
    dg = ctx.stash.get("depgraph") or {}
    prof["dcf"] = {"anchors": len(anc), "reproduce": sum(1 for a in anc if a["matches"]),
                   "core_kinds": sorted({c["kind"] for c in dg.get("cores", []) if c.get("kind")})}
    return prof


def row_cone(ctx: Ctx, where, n_sheets: int) -> dict:
    """Rows upstream of the top DCF anchor (live edges), how deep, and how many sheets they sit on."""
    adj = defaultdict(list)
    live = ",".join(repr(k) for k in LIVE)
    for ss, sr, ds, dr in ctx.db.execute(f"SELECT src_sheet, src_row, dst_sheet, dst_row FROM edges WHERE kind IN ({live})"):
        adj[(ss, sr)].append((ds, dr))
    seen, frontier, depth = {where}, [where], 0
    while frontier:
        nxt = []
        for node in frontier:
            for d in adj.get(node, ()):
                if d not in seen:
                    seen.add(d)
                    nxt.append(d)
        if nxt:
            depth += 1
        frontier = nxt
    cone = {s for s, _ in seen}
    return {"anchor_sheet": ctx.tok(where[0]), "closure_rows": len(seen) - 1, "max_depth": depth,
            "cone_sheets": len(cone), "cone_share": round(len(cone) / n_sheets, 3) if n_sheets else None,
            "cone_sheet_tokens": sorted(ctx.tok(s) for s in cone)}


# ---- stage: depgraph -------------------------------------------------------------------------------------------------

def stage_depgraph(ctx: Ctx) -> dict:
    from . import depgraph
    if "anchors" not in ctx.stash:
        raise Skip("no usable anchor")
    top = ctx.stash.get("top_anchor")
    if not top:
        raise Skip("no usable anchor")
    g = depgraph.build(str(ctx.db_path), cell=top["cell"], depth=ctx.args.depth, max_rows=ctx.args.max_rows, with_anchors=False)
    ctx.stash["depgraph"] = g
    known = {c[1] for c in depgraph.CLASSES}
    for tab in (depgraph.DISCOUNTING, depgraph.TERMINAL_SUB, depgraph.BRIDGE_SUB, depgraph.ASSUMPTIONS, depgraph.CALCULATIONS):
        known |= {str(t[1]) for t in tab if len(t) > 1}
    sub: Counter = Counter()
    for n in g["nodes"]:
        s = n.get("subclass")
        if n.get("kind") != "group" and s:
            sub[s if s in known or s in ("value outcome", "scenario-selected") else "other"] += 1
    st = g["stats"]
    ctx.full["depgraph"] = {"start_label": g["start"]["label"], "anchor_cell": top["cell"],
                            "sheets": {ctx.tok(b["sheet"]): b["sheet"] for b in g["by_sheet"]}}
    cores = []
    for c in g["cores"]:
        m = c.get("method") or {}
        cores.append({"kind": c.get("kind"), "periods": c.get("periods"),
                      "timing": m.get("timing") if m.get("timing") in ("end", "mid", "start") else ("other" if m.get("timing") else None),
                      "day_count": (m.get("day_count") if m.get("day_count") in ("actual/actual", "actual/365", "30/360")
                                    else ("other" if m.get("day_count") else None)),
                      "rate_is_name": bool(m.get("rate") and "!" not in str(m.get("rate"))),
                      "has_terminal_date": bool(m.get("terminal_date"))})
    return {"stats": {k: st.get(k) for k in ("nodes", "edges", "rows_upstream_total", "rows_shown", "cells_shown",
                                              "inactive_hidden", "checks_excluded", "edges_dropped", "groups", "max_depth",
                                              "max_rows", "build_secs")},
            "by_class": g["by_class"], "by_subclass": [{"subclass": k, "n": v} for k, v in sub.most_common()],
            "by_sheet": [{"sheet": ctx.tok(b["sheet"]), "upstream": b["upstream"], "shown": b["shown"]} for b in g["by_sheet"]],
            "cores": cores, "capped": st.get("rows_shown", 0) >= ctx.args.max_rows and st.get("rows_upstream_total", 0) > st.get("rows_shown", 0)}


# ---- stage: statements -----------------------------------------------------------------------------------------------

def check_blocks(raw: dict) -> set[str]:
    """The ontology blocks that a check row of the model guards (a check mapped to any identity of the block)."""
    from . import ontology
    out = set()
    for c in raw.get("checks", []):
        try:
            if c.get("identity"):
                out.add(ontology.identity(c["identity"]).block)
        except KeyError:
            pass
    return out


def _flat(role_val):
    return role_val if isinstance(role_val, list) else [role_val]


def _role_rows(roles: dict):
    return [(k, x) for k, v in roles.items() for x in _flat(v) if isinstance(x, dict) and "sheet" in x]


def stage_statements(ctx: Ctx) -> dict:
    from . import statements
    r = statements.detect(str(ctx.db_path))
    ctx.stash["statements"] = r
    guarded = check_blocks(r)
    idents = []
    for i in r["identities"]:
        picks = _role_rows(i["roles"])
        scale = 0.0
        for _, x in picks[:30]:
            vals = [abs(v) for v in row_values(ctx, x["sheet"], x["row"]) if isinstance(v, (int, float)) and not isinstance(v, bool)]
            scale = max(scale, max(vals, default=0.0))
        why = i.get("why") or ""
        note = ("check_says_holds" if "although the model's own check" in why else
                "binding_rejected" if "binding rejected" in why else None)
        idents.append({"key": i["key"], "kind": i["kind"], "status": i["status"], "scope": i["scope"],
                       "periods_checked": i["periods_checked"], "failing_periods": len(i.get("failing_periods") or []),
                       "residual": residual_bucket(i.get("max_residual"), scale),
                       "check_row": bool(i.get("given_by")) or i["block"] in guarded, "model_says": i.get("model_says"),
                       "reason_codes": reason_codes(why) if i["status"] == "unbound" else [], "note": note,
                       "bound_by": sorted(set(re.findall(r"check|value_search|label|structure|link", i.get("kind_basis") or ""))),
                       "roles": [{"role": k, "bound_by": x.get("bound_by"), "sheet": ctx.tok(x["sheet"]), "row": x["row"],
                                  "label_words": len((x.get("label") or "").split()), "label_vocab": vocab_names(x.get("label"))}
                                 for k, x in picks[:12]]})
    blocks = []
    for b in r["blocks"]:
        rows = _role_rows(b["rows"])
        blocks.append({"type": b["type"], "bound": b["bound"], "complete": b.get("complete"),
                       "sheet": ctx.tok(b["sheet"]) if b.get("sheet") else None, "roles_bound": len(rows),
                       "roles_unbound": len(b["unbound"]), "alternates": len(b.get("alternates", [])),
                       "bound_by": [{"how": k, "n": v} for k, v in sorted(Counter(x.get("bound_by") or "none" for _, x in rows).items())]})
    cs, chk = r["corkscrews"], r["checks"]
    stats, tl = r["stats"], r.get("timeline") or {}
    res = {"tolerance_relative": (r.get("tolerance") or {}).get("relative"), "identities": idents, "blocks": blocks,
           "corkscrews": {"n": len(cs), "kinds": [{"kind": k, "n": v} for k, v in sorted(Counter(c["kind"] for c in cs).items())],
                          "status": [{"status": k, "n": v} for k, v in sorted(Counter(str(c.get("status")) for c in cs).items())]},
           "subtotals": len(r["subtotals"]), "consolidations": len(r["consolidations"]),
           "check_rows": {"n": len(chk), "verdicts": [{"verdict": k, "n": v} for k, v in sorted(Counter(str(c.get("verdict")) for c in chk).items())],
                          "kinds": [{"kind": k, "n": v} for k, v in sorted(Counter(str(c.get("kind")) for c in chk).items())]},
           "lineage": [{"key": x["key"], "status": x["status"]} for x in r["lineage"]],
           "findings": [{"kind": k, "n": v} for k, v in sorted(Counter(f["kind"] for f in r["findings"]).items())],
           "stale": {"cells_checked": stats.get("stale_cells_checked"), "rows_differing": stats.get("stale_rows_differing"),
                     "circular_skipped": stats.get("stale_rows_circular_skipped")},
           "tests": stats["tests"], "structural": stats["structural"], "rows": stats["rows"], "rows_parsed": stats["rows_parsed"],
           "capped": [m if not ctx.scanner.hits(m) else "<message withheld: it matched the self-scan>"
                      for m in (ctx.anon.message(c) for c in stats.get("capped", []))], "notes": len(stats.get("notes", [])),
           "timeline": [{"sheet": ctx.tok(s), "periodicity": v.get("periodicity"), "periods": v.get("periods"),
                         "decade": decade(v.get("first")), "total_column": bool(v.get("total_column"))} for s, v in tl.items()],
           "secs": stats.get("secs")}
    ctx.full["statements"] = {
        "findings": r["findings"], "reasons": {i["key"]: i["why"] for i in r["identities"] if i["status"] != "holds"},
        "bound_rows": {b["type"]: {k: [{"sheet": x["sheet"], "row": x["row"], "label": x.get("label"), "how": x.get("how")}
                                       for x in _flat(v) if isinstance(x, dict)] for k, v in b["rows"].items()} for b in r["blocks"]},
        "check_rows": [{"sheet": c["sheet"], "row": c["row"], "label": c["label"], "verdict": c.get("verdict")} for c in chk]}
    return res


# ---- stage: fixture hints --------------------------------------------------------------------------------------------

def stage_hints(ctx: Ctx) -> dict:
    """For every identity that fails or is unbound, and every exception: the shape that rebuilds the case."""
    raw = ctx.stash.get("statements")
    census = ctx.stash.get("census") or {}
    hints = []
    if raw:
        blocks = {b["type"]: b for b in raw["blocks"]}
        for i in raw["identities"]:
            if i["status"] not in ("fails", "unbound"):
                continue
            picks = _role_rows(i["roles"])
            if not picks:  # unbound: the rows of its block are what the stage was handling
                b = blocks.get(i["block"])
                picks = _role_rows(b["rows"]) if b else []
            picks = picks[:12]
            base: dict[str, int] = {}
            for _, x in picks:
                base[x["sheet"]] = min(base.get(x["sheet"], x["row"]), x["row"])
            rows = [row_shape(ctx, x["sheet"], x["row"], k, x.get("bound_by"), base[x["sheet"]]) for k, x in picks]
            hints.append({"kind": "identity", "key": i["key"], "status": i["status"], "identity_kind": i["kind"],
                          "scope": i["scope"], "reason_codes": reason_codes(i.get("why")) if i["status"] == "unbound" else [],
                          "check_row": bool(i.get("given_by")) or i["block"] in check_blocks(raw),
                          "model_says": i.get("model_says"), "periodicity": sorted({r["periodicity"] for r in rows if r["periodicity"]}),
                          "sheets_involved": len({r["sheet"] for r in rows}),
                          "unbound_roles": sorted((blocks.get(i["block"]) or {}).get("unbound", {}))[:10], "rows": rows,
                          "site": "modelatlas/statements.py in run_identity"})
    cx = ctx.stash.get("result_complexity") or {}
    for e in ctx.errors:
        hints.append({"kind": "exception", "stage": e["stage"], "type": e["type"], "site": e["site"], "message": e["message"],
                      "context": {"size_class": size_class(census.get("formulas", 0)), "sheets": census.get("sheets_n"),
                                  "periodicities": sorted({s["periodicity"] for s in census.get("sheets", []) if s["periodicity"]}),
                                  "longest_families": (cx.get("longest_families") or [])[:3]}})
    return {"n": len(hints), "hints": hints}


STAGES = [("census", stage_census), ("anchors", stage_anchors), ("depgraph", stage_depgraph),
          ("statements", stage_statements), ("complexity", stage_complexity), ("hints", stage_hints)]
DISPLAY = ["census", "complexity", "depgraph", "statements", "hints", "anchors"]  # the order a reader wants


# ---- report assembly -----------------------------------------------------------------------------------------------

def sheet_roles(ctx: Ctx) -> list[dict]:
    """A structure-only hint per sheet: src_only / mid / sink_only / check_rows / timeline_less, never the name."""
    roles = ctx.stash.get("sheet_roles") or {}
    raw = ctx.stash.get("statements") or {}
    check_rows = Counter(c["sheet"] for c in raw.get("checks", []))
    fcount = dict(q(ctx.db, "SELECT sheet, COUNT(*) FROM rows WHERE n_formula > 0 GROUP BY sheet"))
    out = []
    for s in ctx.sheet_names:
        lay = ctx.layouts.get(s) or {}
        role = roles.get(s, "unknown")
        if fcount.get(s) and check_rows[s] and check_rows[s] * 2 >= fcount[s]:
            role = "check_rows"
        elif not lay.get("periods") and role in ("mid", "sink_only"):
            role = "timeline_less"
        out.append({"sheet": ctx.tok(s), "role": role, "timeline": bool(lay.get("periods"))})
    return out


class _Plain:
    """Stands where Anon does when nothing is to be hidden: the model's own sheet names, patterns and functions."""
    funcs: dict = {}

    @staticmethod
    def sheet_token(name):
        return name

    @staticmethod
    def pattern(text):
        return text

    @staticmethod
    def func(name):
        return re.sub(r"^(?:_xl[a-z]+\.)+", "", name, flags=re.I).upper()


class _LocalCtx(Ctx):
    """The stages' context for a person looking at their own model on their own machine (the dashboard): real names, no
    self-scan, no registry. Never used to write a report."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.folder = self.db_path.parent.name
        self.args = types.SimpleNamespace(depth=6, max_rows=300, quiet=True)
        self.db = rodb.connect(self.db_path, check_same_thread=False)
        try:
            self.sheet_names = [r[0] for r in self.db.execute("SELECT sheet FROM sheets ORDER BY rowid")]
            self.layouts = {}
            for s, lay in self.db.execute("SELECT sheet, layout FROM sheets"):
                try:
                    self.layouts[s] = json.loads(lay) if lay else {}
                except ValueError:
                    self.layouts[s] = {}
            self.have = {r[0] for r in self.db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        except BaseException:
            self.close()   # not a model.db, or locked: the file is not left open
            raise
        self.anon = _Plain()
        self.stash, self.full, self.errors = {}, {}, []

    def tok(self, sheet: str) -> str:
        return sheet


def profile_full(db_path, statements: dict | None = None) -> dict:
    """The complexity profile of one model.db at level full: the stage_complexity result with the model's real sheet
    names, with the census, the DCF anchors and the per-sheet roles beside it. For the local dashboard only; the
    shapes reports never go through here. statements: a statements.detect() result, for the sheets that are check rows."""
    ctx = _LocalCtx(db_path)
    try:
        out: dict = {"census": stage_census(ctx)}
        try:
            out["anchors"] = stage_anchors(ctx)
            stage_depgraph(ctx)
        except Exception as e:  # noqa: BLE001 (no DCF found, or one that cannot be read: the profile still stands)
            out["anchors_error"] = f"{type(e).__name__}: {e}" if not isinstance(e, Skip) else str(e)
        ctx.stash["statements"] = statements or {}
        out["complexity"] = stage_complexity(ctx)
        out["sheets"] = sheet_roles(ctx)
        out["top_anchor"] = ctx.stash.get("top_anchor")
        return out
    finally:
        ctx.close()


def build_report(ctx: Ctx, build_info: dict | None = None) -> dict:
    stages = {}
    for name, fn in STAGES:
        stages[name] = run_stage(ctx, name, fn)
        if stages[name]["status"] == "ok":
            ctx.stash["result_" + name] = stages[name]["result"]
        if not getattr(ctx.args, "quiet", False):
            print(f"  {ctx.token} {name}: {stages[name]['status']} {stages[name]['secs']}s", flush=True)
    return {"schema": SCHEMA_VERSION, "version": __version__, "workbook": ctx.token, "build": build_info,
            "stages": {k: stages[k] for k in DISPLAY if k in stages}, "sheets": sheet_roles(ctx),
            "stages_failed": sum(1 for s in stages.values() if s["status"] == "failed"),
            "secs": round(sum(s["secs"] for s in stages.values()), 2)}


def one(v, default="-"):
    return default if v is None else v


def _table(head, rows):
    return ["| " + " | ".join(head) + " |", "|" + "---|" * len(head)] + ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]


def render_md(rep: dict) -> str:
    st = rep["stages"]
    L = [f"# Shapes of model {rep['workbook']}", "", "## What this file contains", "", f"Produced by Model Atlas {rep.get('version') or __version__}.", "", MD_PROSE, "", "Review before sharing:", ""]
    L += [f"- [ ] {c}" for c in MD_CHECKLIST]
    L += ["", f"Stages failed: {rep['stages_failed']}. Seconds: {rep['secs']}.", "", "## Sheets", ""]
    L += _table(["sheet", "role", "timeline"], [(s["sheet"], s["role"], "yes" if s["timeline"] else "no") for s in rep["sheets"]])
    for name in DISPLAY:
        s = st.get(name)
        if not s:
            continue
        L += ["", f"## Stage {name}: {s['status']} ({s['secs']} s)", ""]
        if s["status"] == "skipped":
            L.append(f"Skipped: {s['reason']}")
        elif s["status"] == "failed":
            e = s["error"]
            L += [f"{e['type']}: {e['message']}", f"site: {e['site']}", "", "```"] + e["traceback"] + ["```"]
        else:
            L += _md_stage(name, s["result"])
    L += ["", "---", "", NOTICE]
    return "\n".join(L) + "\n"


def _md_stage(name, r) -> list[str]:
    L = []
    if name == "census":
        L.append(f"sheets {r['sheets_n']} ({r['hidden_sheets']} hidden), rows {r['rows']}, formulas {r['formulas']}, consts {r['consts']}, "
                 f"named ranges {r['names_user']} (+{r['names_n'] - r['names_user']} built in), linked files {r['external']['linked_files']}")
        L += [""] + _table(["sheet", "rows", "formulas", "consts", "periodicity", "periods", "decade", "header row", "label col", "units col"],
                           [(s["sheet"], s["rows"], s["formulas"], s["consts"], one(s["periodicity"]), s["periods"], one(s["first_decade"]),
                             s["header_row"], s["label_column"], s["units_column"]) for s in r["sheets"]])
    elif name == "anchors":
        L.append(f"found {r['found']}, usable {r['usable']}, reproduce {r['reproduce']}, {r['secs']} s (saved catalogue read: {r['catalogue_cached']})")
        L.append(r["note"])
    elif name == "complexity":
        keys = ("size_class", "formula_cells", "formula_rows", "families_listed", "families_scrubbed", "families_left_out",
                "families_per_formula_cell", "rows_multi_pattern", "runs_in_multi_pattern_rows", "rows_single_cell_break",
                "max_formula_length", "functions_distinct", "non_excel_functions")
        L += [f"- {k}: {r[k]}" for k in keys]
        L.append(f"- {r['families_note']}")
        L.append("- edges by kind: " + ", ".join(f"{e['kind']} {e['n']}" for e in r["edges_by_kind"]))
        for key in ("sheet_graph", "row_graph"):
            if r[key]:
                L.append(f"- {key}: " + ", ".join(f"{k} {v}" for k, v in r[key].items()))
        L.append("- timeline: " + ", ".join(f"{k} {v}" for k, v in r["timeline"].items()))
        L.append("- dcf: " + ", ".join(f"{k} {v}" for k, v in r["dcf"].items()))
        L += ["", "Functions (top 30): " + ", ".join(f"{f['fn']} {f['n']}" for f in r["functions_top"]), "", "Top families:", ""]
        L += _table(["cells", "r1c1"], [(f["cells"], f"`{f['r1c1']}`") for f in r["top_families"]])
        L += ["", "Longest families:", ""] + _table(["cells", "r1c1"], [(f["cells"], f"`{f['r1c1']}`") for f in r["longest_families"]])
    elif name == "depgraph":
        L.append("stats: " + ", ".join(f"{k} {v}" for k, v in r["stats"].items()))
        L.append("classes: " + ", ".join(f"{c['class']} {c['n']}" for c in r["by_class"]))
        L.append("subclasses: " + ", ".join(f"{c['subclass']} {c['n']}" for c in r["by_subclass"]))
        L += [""] + _table(["sheet", "upstream", "shown"], [(b["sheet"], b["upstream"], b["shown"]) for b in r["by_sheet"]])
        L += ["", "cores: " + "; ".join(", ".join(f"{k} {v}" for k, v in c.items()) for c in r["cores"])]
    elif name == "statements":
        L.append(f"tests of the model: {r['tests']}; structure confirmed: {r['structural']}")
        L += [""] + _table(["identity", "kind", "status", "scope", "periods", "residual", "check row", "model says", "reasons", "note"],
                           [(i["key"], i["kind"], i["status"], i["scope"], i["periods_checked"], one(i["residual"]), i["check_row"],
                             one(i["model_says"]), ",".join(i["reason_codes"]) or "-", one(i["note"])) for i in r["identities"]])
        L += ["", "blocks: " + "; ".join(f"{b['type']} bound={b['bound']} roles={b['roles_bound']} unbound={b['roles_unbound']}" for b in r["blocks"])]
        L.append(f"corkscrews {r['corkscrews']['n']} {r['corkscrews']['kinds']}; subtotals {r['subtotals']}; consolidations {r['consolidations']}")
        L.append(f"check rows {r['check_rows']['n']} {r['check_rows']['verdicts']} {r['check_rows']['kinds']}")
        L.append("findings: " + ", ".join(f"{f['kind']} {f['n']}" for f in r["findings"]))
        L.append(f"stale: {r['stale']}; capped: {r['capped']}")
    elif name == "hints":
        L.append(f"{r['n']} hints")
        for h in r["hints"]:
            if h["kind"] == "identity":
                L += ["", f"### identity {h['key']}: {h['status']} ({h['identity_kind']}, scope {h['scope']})",
                      f"reasons {h['reason_codes']}; check row {h['check_row']}; model says {one(h['model_says'])}; periodicity "
                      f"{h['periodicity']}; unbound roles {h['unbound_roles']}", ""]
                for x in h["rows"]:
                    L.append(f"- {x['role']} ({one(x['bound_by'])}) {x['sheet']} r{x['row']} (+{one(x['row_offset'])}): formulas {x['formulas']}, "
                             f"consts {x['consts']}, label words {x['label_words']}, vocab {x['label_vocab']}, signs {x['signs']}")
                    for p in x["patterns"]:
                        L.append(f"  - `{p['r1c1']}` x{p['cells']}")
            else:
                L += ["", f"### exception in {h['stage']} at {h['site']}", f"{h['type']}: {h['message']}", f"context: {h['context']}"]
    return L


# ---- the self-scan gate and the files ------------------------------------------------------------------------------

def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def code_strings(rep) -> set[str]:
    """The strings of a report that are the repo's own code: traceback entries and exception sites (file, line,
    function, and a source line of modelatlas/). Their words are code identifiers, so they are scanned for the model's
    names and labels but not for its single words (a label word "pool" must not block a line `pool = ...`)."""
    out = set()

    def walk(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == "traceback" and isinstance(v, list):
                    out.update(x for x in v if isinstance(x, str))
                elif k == "site" and isinstance(v, str):
                    out.add(v)
                elif k == "type" and isinstance(v, str) and "site" in o:   # an exception's class name
                    out.add(v)
                else:
                    walk(v)
        elif isinstance(o, (list, tuple)):
            for v in o:
                walk(v)
    walk(rep)
    return out


_SCHEMA_WORDS = _words_of(" ".join(k.replace("_", " ") for k in SCHEMA_KEYS)) | {
    w for k in SCHEMA_KEYS for w in [k, *k.split("_")] if len(w) >= 3}


def non_ascii(text: str) -> bool:
    """Anything but ASCII and the '…' this module writes for a replaced string: text from the model or a path."""
    return any(ord(c) > 127 and c != "…" for c in text)


def gate(ctx: Ctx, rep: dict, md: str) -> list[str]:
    """Strings of the model that this report would give away (empty = safe to write). Every string VALUE of the report
    on its own (grams within one value, so 'rev_total', 'revenue' never reads as 'total revenue'); the Markdown, which
    is rendered from those values and fixed text, line by line for substrings; keys against the fixed schema; any
    non-ASCII character anywhere; a registry that could not be read blocks too (its names were not scanned for)."""
    scanner = ctx.scanner.with_words(_SCHEMA_WORDS)
    code = code_strings(rep)
    found = set()
    for v in strings_of(rep):
        found.update(scanner.hits(v, terms=v not in code))
        if non_ascii(v):
            found.add("<non-ASCII text>")
    for line in md.splitlines():
        found.update(scanner.hits(line, grams=False, terms=False))
        if non_ascii(line):
            found.add("<non-ASCII text>")
    found |= {f"<unexpected key {k}>" for k in bad_keys(rep)}
    if getattr(ctx, "registry_unreadable", False):
        found.add("<registry.db or atlas.db could not be read, so the registry's names were not scanned for>")
    return sorted(found)


def summary_row(rep: dict) -> dict:
    st = rep["stages"]
    census = (st.get("census") or {}).get("result") or {}
    cx = (st.get("complexity") or {}).get("result") or {}
    rg = cx.get("row_graph") or {}
    stm = (st.get("statements") or {}).get("result") or {}
    by_kind = {"test": Counter(), "structural": Counter()}
    for i in stm.get("identities", []):
        by_kind[i["kind"]][i["status"]] += 1
    f = lambda c: f"{c['holds']}/{c['fails']}/{c['unbound']}"  # noqa: E731
    return {"token": rep["workbook"], "blocked": False, "size_class": cx.get("size_class") or size_class(census.get("formulas", 0)),
            "formulas": census.get("formulas"), "families": cx.get("families_listed"), "sheets": census.get("sheets_n"),
            "cone_share": rg.get("cone_share"), "anchors": (cx.get("dcf") or {}).get("anchors"),
            "anchors_reproduce": (cx.get("dcf") or {}).get("reproduce"),
            "identities_tests": f(by_kind["test"]) if stm else "-", "identities_structural": f(by_kind["structural"]) if stm else "-",
            "findings": sum(x["n"] for x in stm.get("findings", [])), "stages_failed": rep["stages_failed"], "secs": rep["secs"]}


def process(db_path: Path, out_root: Path, args, build_info=None):
    """One model: run the stages, gate, write. Returns (summary row, exceptions, its scanner)."""
    ctx = Ctx(db_path, out_root, args)
    try:
        rep = build_report(ctx, build_info)
        try:
            md = render_md(rep)
        except Exception as e:  # noqa: BLE001 (a stage that returned an odd shape must not stop the run)
            md = f"# Shapes of model {ctx.token}\n\nThe Markdown could not be rendered ({type(e).__name__}); see report.json.\n"
        leaks = gate(ctx, rep, md)
        report_dir = Path(args.report)
        mine = report_dir / ctx.token
        if leaks:
            for old in ("report.json", "report.md"):  # an earlier run's file must not outlive a block
                (mine / old).unlink(missing_ok=True)
            lines = ["LOCAL ONLY. DO NOT SHARE THIS FILE.", "",
                     f"The shapes report for {ctx.token} was not written: these strings of the model were found in it.", ""]
            lines += [f"  {s}" for s in leaks]
            write_text(report_dir / "_blocked" / f"{ctx.token}.txt", "\n".join(lines) + "\n")
            row = {"token": ctx.token, "blocked": True, "stages_failed": rep["stages_failed"], "secs": rep["secs"]}
            print(f"  {ctx.token}: BLOCKED by the self-scan ({len(leaks)} strings; see _blocked/{ctx.token}.txt)", flush=True)
        else:
            (report_dir / "_blocked" / f"{ctx.token}.txt").unlink(missing_ok=True)
            write_text(mine / "report.json", json.dumps(rep, indent=1, ensure_ascii=True) + "\n")
            write_text(mine / "report.md", md)
            row = summary_row(rep)
        if args.level == "full":
            full = {"warning": FULL_WARNING, "token": ctx.token, "folder": ctx.folder, "registry": ctx.registry,
                    "sheet_names": {ctx.tok(n): n for n in ctx.sheet_names}, "stages": ctx.full, "shapes": rep}
            write_text(report_dir / "full" / ctx.token / "report.json", json.dumps(full, indent=1, ensure_ascii=False, default=str) + "\n")
        excs = [] if leaks else [{"site": e["site"], "type": e["type"], "stage": e["stage"], "token": ctx.token} for e in ctx.errors]
        return row, excs, ctx.scanner
    finally:
        ctx.close()


# ---- the test suite, environment, summary ------------------------------------------------------------------------------

def run_tests(tests_dir: Path = ROOT / "tests", only=None, timeout: int = 600) -> list[dict]:
    """Every tests/check_*.py (or the named scripts) in a fresh interpreter: status, seconds, last 15 lines."""
    out = []
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    for script in sorted(Path(tests_dir).glob("check_*.py")):
        if only and script.name not in only:
            continue
        t0 = time.time()
        try:
            p = subprocess.run([sys.executable, str(script)], cwd=str(ROOT), capture_output=True, text=True, encoding="utf-8",
                               errors="replace", timeout=timeout, env=env)
            status, text = ("pass" if p.returncode == 0 else "fail"), (p.stdout or "") + "\n" + (p.stderr or "")
        except subprocess.TimeoutExpired as e:
            partial = e.stdout.decode("utf-8", "replace") if isinstance(e.stdout, bytes) else (e.stdout or "")
            status, text = "timeout", partial
        tail = [scrub_env(ln) for ln in text.strip().splitlines()[-15:]]
        out.append({"script": script.name, "status": status, "secs": round(time.time() - t0, 1), "tail": tail})
        print(f"  tests: {script.name} {status} {out[-1]['secs']}s", flush=True)
    return out


def environment() -> dict:
    def ver(pkg):
        try:
            return importlib.metadata.version(pkg)
        except importlib.metadata.PackageNotFoundError:
            return None
    git = {"commit": None, "dirty": None}
    exe = shutil.which("git")
    if exe:
        try:
            c = subprocess.run([exe, "rev-parse", "HEAD"], cwd=str(ROOT), capture_output=True, text=True, timeout=10)
            if c.returncode == 0:
                git["commit"] = c.stdout.strip()
                d = subprocess.run([exe, "status", "--porcelain"], cwd=str(ROOT), capture_output=True, text=True, timeout=10)
                git["dirty"] = bool(d.stdout.strip())
        except (OSError, subprocess.TimeoutExpired):
            pass
    return {"python": platform.python_version(), "implementation": platform.python_implementation(),
            "platform": platform.platform(), "machine": platform.machine(),
            "version": __version__, "packages": {p: ver(p) for p in ("openpyxl", "numpy", "fastapi")}, "git": git}


def group_exceptions(excs) -> list[dict]:
    g = defaultdict(list)
    for e in excs:
        g[(e["site"], e["type"])].append(e)
    return [{"site": s, "type": t, "count": len(v), "stage": v[0]["stage"], "tokens": sorted({e["token"] for e in v})}
            for (s, t), v in sorted(g.items(), key=lambda kv: -len(kv[1]))]


def render_summary(rows, excs, tests) -> str:
    L = ["# Diagnostic run summary", "", "Shapes only: tokens, counts, statuses and timings. See each model's report for the shapes.", "",
         f"Models: {len(rows)}; blocked by the self-scan: {sum(1 for r in rows if r['blocked'])}.", ""]
    ok = [r for r in rows if not r["blocked"]]
    L += _table(["token", "size class", "formulas", "families", "sheets", "cone share", "anchors (reproduce)",
                 "tests h/f/u", "structure h/f/u", "findings", "stages failed", "seconds"],
                [(r["token"], r["size_class"], one(r["formulas"]), one(r["families"]), one(r["sheets"]), one(r["cone_share"]),
                  f"{one(r['anchors'])} ({one(r['anchors_reproduce'])})", r["identities_tests"], r["identities_structural"],
                  r["findings"], r["stages_failed"], r["secs"]) for r in ok])
    blocked = [r for r in rows if r["blocked"]]
    if blocked:
        L += ["", "Blocked (no shapes report written):", ""] + _table(["token", "stages failed", "seconds"],
                                                                      [(r["token"], r["stages_failed"], r["secs"]) for r in blocked])
    if tests is not None:
        L += ["", "## Test suite", ""] + _table(["script", "status", "seconds"], [(t["script"], t["status"], t["secs"]) for t in tests])
        for t in tests:
            if t["status"] != "pass":
                L += ["", f"{t['script']}:", "```"] + t["tail"] + ["```"]
    L += ["", "## Exceptions by site", ""]
    grouped = group_exceptions(excs)
    L += [f"- {g['site']}: {g['type']} x{g['count']} (stage {g['stage']}; models {', '.join(g['tokens'])})" for g in grouped] or ["none"]
    return "\n".join(L) + "\n"


# ---- finding and building the models ------------------------------------------------------------------------------------

def find_dbs(out_root: Path, only: str | None) -> list[Path]:
    dbs = sorted(p for p in out_root.glob("*/model.db") if p.is_file())
    return [p for p in dbs if not only or only.lower() in p.parent.name.lower()]


def build_workbooks(wdir: Path, out_root: Path, rebuild=False) -> list[tuple[Path, dict]]:
    """Build each .xlsx/.xlsm in wdir into out/<stem>__<sha8>/ (build_map.main). Returns [(model.db, build info)]."""
    import contextlib
    import io

    from . import build_map
    done = []
    for f in sorted(p for p in Path(wdir).iterdir() if p.suffix.lower() in (".xlsx", ".xlsm") and not p.name.startswith("~$")):
        h = hashlib.sha256()
        with open(f, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        folder = out_root / f"{f.stem}__{h.hexdigest()[:8]}"
        db = folder / "model.db"
        t0 = time.time()
        info = {"built": False, "secs": 0.0, "error": None}
        if db.exists() and not rebuild:
            info["note"] = "model.db was already there"
        else:
            try:
                with contextlib.redirect_stdout(io.StringIO()):  # build_map prints the folder name
                    build_map.main(str(f), str(folder))
                info["built"] = True
            except Exception as e:  # noqa: BLE001
                info["error"] = {"type": type(e).__name__, "site": error_record(None, e, "build")["site"]}
                print(f"  {folder_token(folder.name)}: build failed ({type(e).__name__})", flush=True)
        info["secs"] = round(time.time() - t0, 1)
        if db.exists():
            done.append((db, info))
    return done


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("out", nargs="?", default="out", help="the folder of out/<stem>__<sha8>/model.db")
    ap.add_argument("--workbooks", help="build each .xlsx/.xlsm in this folder into out/ first")
    ap.add_argument("--report", default="diag", help="where the reports go (default diag/)")
    ap.add_argument("--level", choices=("shapes", "full"), default="shapes")
    ap.add_argument("--tests", action="store_true", help="also run every tests/check_*.py")
    ap.add_argument("--only", help="only model folders containing this text")
    ap.add_argument("--depth", type=int, default=6)
    ap.add_argument("--max-rows", type=int, default=300)
    ap.add_argument("--rebuild", action="store_true", help="with --workbooks: rebuild even if model.db exists")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--version", action="version", version=version_line())
    args = ap.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):   # a cp1252 console must not stop a run over a character it cannot show
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    out_root, report = Path(args.out), Path(args.report)
    if args.level == "full":
        print("!" * 78 + f"\n{FULL_WARNING}\n" + "!" * 78, flush=True)
    if args.workbooks:
        jobs = [(db, info) for db, info in build_workbooks(Path(args.workbooks), out_root, args.rebuild)
                if not args.only or args.only.lower() in db.parent.name.lower()]
    else:
        jobs = [(p, None) for p in find_dbs(out_root, args.only)]
    print(f"{len(jobs)} models", flush=True)
    rows, excs, scanners = [], [], []
    for db, info in jobs:
        t0 = time.time()
        try:
            row, ex, sc = process(db, out_root, args, info)
        except Exception as e:  # noqa: BLE001 (a model.db that cannot even be opened)
            tok, rec = folder_token(db.parent.name), error_record(None, e, "open")
            row = {"token": tok, "blocked": False, "size_class": "?", "formulas": None, "families": None, "sheets": None,
                   "cone_share": None, "anchors": None, "anchors_reproduce": None, "identities_tests": "-",
                   "identities_structural": "-", "findings": 0, "stages_failed": 1, "secs": round(time.time() - t0, 1)}
            ex, sc = [{"site": rec["site"], "type": rec["type"], "stage": "open", "token": tok}], None
        print(f"{row['token']}: {'blocked' if row['blocked'] else 'done'} in {row['secs']}s", flush=True)
        rows.append(row)
        excs += ex
        if sc:
            scanners.append(sc)
    tests = None
    if args.tests:
        tests = run_tests()
        if scanners:
            union = Scanner.union(scanners)
            for t in tests:
                t["tail"] = ["<line withheld: it matched the self-scan>" if union.hits(ln) or non_ascii(ln) else ln
                             for ln in t["tail"]]
    if scanners:   # the summary holds tokens, counts and code sites only; checked all the same
        union = Scanner.union(scanners)
        excs = [e if not (union.hits(e["site"], terms=False) or non_ascii(e["site"])) else dict(e, site="<site withheld>")
                for e in excs]
    summary = {"schema": SCHEMA_VERSION, "models": rows, "tests": tests, "exceptions": group_exceptions(excs)}
    write_text(report / "summary.json", json.dumps(summary, indent=1, ensure_ascii=True) + "\n")
    write_text(report / "summary.md", render_summary(rows, excs, tests))
    write_text(report / "environment.json", json.dumps(environment(), indent=1, ensure_ascii=True) + "\n")
    print(f"reports in {report}/ ({sum(1 for r in rows if r['blocked'])} blocked)", flush=True)
    return 0


def cli() -> None:
    sys.exit(main())


if __name__ == "__main__":
    cli()
