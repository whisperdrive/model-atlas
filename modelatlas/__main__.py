"""python -m modelatlas: prints the version (--version) or lists the commands."""
import sys

from . import version_line

USAGE = """Model Atlas: reads an Excel financial model into a map, traces and tests it. No model calls.

commands (each also runs as python -m modelatlas.<module>):
  atlas-build       workbook.xlsx            build the row map and model.db into out/<name>/
  atlas-census      workbook.xlsx            per-sheet formula and constant counts
  atlas-graph       out/<name>/model.db      the DCF dependency graph as an HTML report
  atlas-statements  out/<name>/model.db      statements, identities and findings
  atlas-diagnose    out/ --report diag/      anonymised diagnostic report per model
  atlas-dashboard                            read-only dashboard on http://localhost:8001

every command accepts --version"""


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if argv and argv[0] in ("--version", "-V"):
        print(version_line())
        return 0
    print(version_line() + "\n\n" + USAGE)
    return 0 if not argv or argv[0] in ("-h", "--help") else 2


if __name__ == "__main__":
    sys.exit(main())
