"""Model Atlas: reads a large Excel financial model into a SQLite map without a formula engine or any model calls,
traces a DCF value to the rows it rests on and classifies them, finds the statements and schedules by the identities
their saved values satisfy and tests those, draws the dependency graph as a report, serves a read-only dashboard over
the built databases, and runs the whole analysis over many models as anonymised diagnostic runs."""

__version__ = "0.1.0"  # the one place the version is set; pyproject reads it, every command and report shows it

# One short notice, shown on the dashboard, the dependency graph report, the statements text and the diagnostic reports.
NOTICE = ("Experimental: a prototype whose maps, traces, classes and identity checks come from automated rules, can be "
          "incomplete or wrong, and must be checked by a qualified person. Not valuation, financial, tax, accounting or "
          "legal advice, or an opinion on any value.")


def version_line() -> str:
    return f"model-atlas {__version__}"
