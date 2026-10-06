"""The ontology of a financial model: block types, their roles, and the identities that tie the roles together.

Plain data, no logic about workbooks. modelatlas/statements.py reads it to detect which rows of a model play which role;
docs/model_ontology.md is the same table for a person (python -m modelatlas.ontology --write docs/model_ontology.md).

An identity is `expr = 0` over role names, e.g. `assets - liabilities - equity`. Its scope says how it is tested:
  per_period  the roles' saved values, period by period, aligned by date (a Total column is never a period)
  single      a block laid out in one column (sources and uses): compared column by column, else by their sums
  cumulative  running totals of the roles
  coverage    a statement about the model's structure (every balance has a roll-forward), not about numbers
`flex` names the roles whose sign convention is tested both ways (costs stored as negatives or as positives).
A role ending in `*` stands for a list of rows (the movements of a corkscrew). `vocab` is a label regex: a tie-breaker
and a reason, and (see modelatlas/statements.py) the last resort that names a row when no identity can.

`kind` says what evidence the identity gives when it holds. It is a default: modelatlas/statements.py decides it again for
every result from how the roles were bound (`bound_by`: check | value_search | label | structure | link).
  test        the identity can genuinely fail against the saved values: its rows were named by a model check or by
              labels, so nothing about how they were chosen made the identity true
  structural  the identity holds by construction: one row's own formula is the identity (the P&L subtotal chain, a
              corkscrew's closing row, a total that adds its parts), or a row was chosen BECAUSE its values equal its
              counterpart's (a balance sheet line found by value, a link)
A structural identity still fails when the saved values are stale or a formula was typed over; it just says nothing
about whether the model is right.
"""
import argparse
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class Role:
    name: str
    block: str
    what: str
    vocab: str


@dataclass(frozen=True)
class Identity:
    key: str
    block: str
    title: str
    expr: str
    scope: str = "per_period"
    flex: tuple = ()
    note: str = ""
    kind: str = "test"


BLOCKS = {
    "income_statement": "Income statement (P&L)",
    "balance_sheet": "Balance sheet",
    "cash_flow": "Cash flow statement",
    "debt": "Debt schedule",
    "fixed_assets": "Fixed assets schedule",
    "sources_uses": "Sources and uses",
    "capex": "Capital expenditure",
    "revenue": "Revenue build",
    "equity": "Equity (retained earnings) roll-forward",
    "distributions": "Distributions",
}

_R = Role
ROLES = [
    # income statement: costs are usually negative, so `revenue + opex = ebitda`; the other sign is tested too
    _R("revenue", "income_statement", "revenue line of the P&L", r"\b(revenues?|turnover|sales)\b|total income"),
    _R("opex", "income_statement", "operating costs", r"operating (costs?|expenses?)|opex|\bcosts?\b|expenses?"),
    _R("ebitda", "income_statement", "EBITDA", r"\bebitda\b"),
    _R("depreciation", "income_statement", "depreciation and amortisation", r"deprec|amorti[sz]"),
    _R("ebit", "income_statement", "EBIT", r"\bebit\b|operating (profit|income|result)"),
    _R("interest", "income_statement", "interest / finance cost", r"interest|finance (costs?|charges?)"),
    _R("pbt", "income_statement", "profit before tax", r"\bpbt\b|before tax"),
    _R("tax", "income_statement", "tax expense", r"\btax(es|ation)?\b"),
    _R("npat", "income_statement", "net profit after tax", r"\bnpat\b|net (profit|income|earnings)|profit after tax|"
                                                           r"profit for the (year|period)"),
    # balance sheet
    _R("assets", "balance_sheet", "total assets", r"^(total )?assets\b(?!.*liabilit)"),
    _R("liabilities", "balance_sheet", "total liabilities", r"^(total )?liabilities\b(?!.*equity)"),
    _R("equity", "balance_sheet", "total equity", r"^total equity\b(?!.*liabilit)|^(shareholders'? )?equity$|"
                                                  r"shareholders'? (equity|funds)|net assets"),
    _R("total_le", "balance_sheet", "total liabilities and equity",
       r"liabilities (and|&) (shareholders'? )?equity|equity (and|&) liabilities"),
    _R("bs_cash", "balance_sheet", "cash line", r"^cash( and cash equivalents| at bank)?$"),
    _R("bs_fixed_assets", "balance_sheet", "fixed assets line", r"fixed assets|\bppe\b|property.*(plant|equipment)|"
                                                                r"non-?current assets|infrastructure|intangible"),
    _R("bs_debt", "balance_sheet", "debt line", r"debt|borrowing|loans?\b|facilit|\bnotes\b|bonds"),
    _R("bs_retained", "balance_sheet", "retained earnings line", r"retained|accumulated (profit|earnings)|reserves"),
    # cash flow
    _R("cf_opening", "cash_flow", "opening cash", r"opening|beginning|start of"),
    _R("cfo", "cash_flow", "cash flow from operations", r"operating activities|operations|\bcfo\b|operating cash"),
    _R("cfi", "cash_flow", "cash flow from investing", r"investing|\bcfi\b"),
    _R("cff", "cash_flow", "cash flow from financing", r"financing|\bcff\b"),
    _R("cf_net", "cash_flow", "net cash flow", r"net (cash|increase|change|movement)|net cash ?flow"),
    _R("cf_closing", "cash_flow", "closing cash", r"closing cash|cash at (the )?end|ending cash"),
    _R("cf_capex", "cash_flow", "capital expenditure in investing", r"capex|capital expenditure|purchase of"),
    _R("cf_distributions", "cash_flow", "distributions paid in financing", r"distribution|dividend"),
    # debt
    _R("debt_opening", "debt", "opening debt", r"opening|beginning"),
    _R("debt_movements*", "debt", "drawdowns and repayments", r"draw|repay|advance|issue|redemption|principal|amorti"),
    _R("debt_closing", "debt", "closing debt", r"debt|loan|borrowing|facilit|closing balance|principal|\bnotes\b"),
    # fixed assets
    _R("fa_opening", "fixed_assets", "opening fixed assets", r"opening|beginning"),
    _R("fa_movements*", "fixed_assets", "additions and depreciation", r"addition|capex|purchase|deprec|amorti|dispos"),
    _R("fa_closing", "fixed_assets", "closing fixed assets", r"closing fixed|fixed assets|\bppe\b"),
    _R("fa_additions", "fixed_assets", "additions", r"addition|capex|purchase"),
    _R("fa_depreciation", "fixed_assets", "depreciation charge", r"deprec|amorti"),
    # sources and uses
    _R("sources", "sources_uses", "total sources", r"sources"),
    _R("uses", "sources_uses", "total uses", r"\buses\b|application of funds"),
    # capex
    _R("capex_total", "capex", "total capex", r"total capex|capex total|total capital expenditure|^capex$|"
                                                r"capital expenditure total"),
    # revenue
    _R("rev_segments*", "revenue", "revenue segments", r"revenue|sales"),
    _R("rev_total", "revenue", "total revenue", r"total revenue|revenue total|total sales|^revenue$|turnover"),
    # equity
    _R("re_opening", "equity", "opening retained earnings", r"opening|beginning"),
    _R("re_movements*", "equity", "profit and distributions", r"npat|profit|income|distribution|dividend|transfer"),
    _R("re_closing", "equity", "closing retained earnings", r"closing|ending|retained"),
    _R("re_npat", "equity", "profit entering retained earnings", r"npat|profit|income|earnings"),
    _R("re_distributions", "equity", "distributions out of retained earnings", r"distribution|dividend"),
    # distributions
    _R("dcf_distributions", "distributions", "the cash-flow row the DCF discounts",
       r"distribution|dividend|cash ?flow|fcfe"),
]

_I = Identity
IDENTITIES = [
    _I("pnl_ebitda", "income_statement", "Revenue plus operating costs is EBITDA", "revenue + opex - ebitda",
       flex=("opex",), kind="structural"),
    _I("pnl_ebit", "income_statement", "EBITDA plus depreciation is EBIT", "ebitda + depreciation - ebit",
       flex=("depreciation",), kind="structural"),
    _I("pnl_pbt", "income_statement", "EBIT plus interest is profit before tax", "ebit + interest - pbt",
       flex=("interest",), kind="structural"),
    _I("pnl_npat", "income_statement", "Profit before tax plus tax is NPAT", "pbt + tax - npat", flex=("tax",), kind="structural"),
    _I("bs_balance", "balance_sheet", "Assets equal liabilities plus equity, every period", "assets - liabilities - equity",
       flex=("liabilities", "equity")),
    _I("bs_total_le", "balance_sheet", "Assets equal total liabilities and equity", "assets - total_le"),
    _I("bs_le_sum", "balance_sheet", "Liabilities plus equity is total liabilities and equity",
       "liabilities + equity - total_le", flex=("liabilities", "equity")),
    _I("bs_rollforward", "balance_sheet", "Every balance sheet line has a roll-forward (or never moves)", "coverage",
       scope="coverage", kind="structural"),
    _I("cf_net", "cash_flow", "Operating, investing and financing flows sum to net cash flow",
       "cfo + cfi + cff - cf_net", kind="structural"),
    _I("cf_roll", "cash_flow", "Opening cash plus net cash flow is closing cash", "cf_opening + cf_net - cf_closing", kind="structural"),
    _I("cf_cash_tie", "cash_flow", "Closing cash on the cash flow statement is cash on the balance sheet",
       "cf_closing - bs_cash", kind="structural"),
    _I("debt_roll", "debt", "Opening debt plus drawdowns and repayments is closing debt",
       "debt_opening + debt_movements* - debt_closing", kind="structural"),
    _I("debt_tie", "debt", "Closing debt is the balance sheet debt line", "debt_closing - bs_debt", flex=("bs_debt",), kind="structural"),
    _I("fa_roll", "fixed_assets", "Opening fixed assets plus additions and depreciation is closing",
       "fa_opening + fa_movements* - fa_closing", kind="structural"),
    _I("fa_tie", "fixed_assets", "Closing fixed assets is the balance sheet fixed assets line",
       "fa_closing - bs_fixed_assets", kind="structural"),
    _I("fa_dep", "fixed_assets", "Depreciation in the schedule is depreciation in the P&L (up to sign)",
       "fa_depreciation - depreciation", flex=("depreciation",)),
    _I("su_balance", "sources_uses", "Total sources equal total uses", "sources - uses", scope="single"),
    _I("capex_cf", "capex", "Total capex is the capital expenditure in cash flow from investing (up to sign)",
       "capex_total - cf_capex", flex=("cf_capex",), kind="structural"),
    _I("fa_additions", "capex", "Additions in the fixed assets schedule are minus total capex",
       "fa_additions + capex_total", flex=("capex_total",), kind="structural"),
    _I("rev_sum", "revenue", "The revenue segments sum to total revenue", "rev_segments* - rev_total", kind="structural"),
    _I("rev_pnl", "revenue", "Total revenue is the P&L revenue line", "rev_total - revenue", kind="structural"),
    _I("re_roll", "equity", "Opening retained earnings plus profit and distributions is closing",
       "re_opening + re_movements* - re_closing", kind="structural"),
    _I("re_npat", "equity", "Profit entering retained earnings is the P&L NPAT", "re_npat - npat", kind="structural"),
    _I("re_tie", "equity", "Closing retained earnings is the balance sheet retained earnings line",
       "re_closing - bs_retained", kind="structural"),
    _I("dist_dcf_cf", "distributions", "The DCF's cash-flow row is the distributions paid in the cash flow (up to sign)",
       "dcf_distributions - cf_distributions", flex=("cf_distributions",), kind="structural"),
    _I("dist_cf_eq", "distributions", "Distributions paid in the cash flow are the equity roll-forward's (up to sign)",
       "cf_distributions - re_distributions", flex=("re_distributions",)),
]

_BY_KEY = {i.key: i for i in IDENTITIES}
_ROLE = {r.name: r for r in ROLES}


def identity(key: str) -> Identity:
    return _BY_KEY[key]


def role(name: str) -> Role:
    return _ROLE[name if name in _ROLE else name + "*"]


def coefficients(ident: Identity) -> dict[str, int]:
    """`a + b - c` -> {'a': 1, 'b': 1, 'c': -1}; empty for a coverage identity."""
    if ident.scope == "coverage":
        return {}
    out, sign = {}, 1
    for tok in re.findall(r"[+-]|[A-Za-z_*]+", ident.expr):
        if tok == "+":
            sign = 1
        elif tok == "-":
            sign = -1
        else:
            out[tok] = sign
            sign = 1
    return out


def vocab(name: str):
    return re.compile(role(name).vocab, re.I)


def vocab_hit(name: str, label: str) -> bool:
    return bool(label) and bool(vocab(name).search(label.strip()))


def markdown() -> str:
    lines = ["# Model ontology", "",
             "Generated from `modelatlas/ontology.py` (`uv run python -m modelatlas.ontology --write docs/model_ontology.md`). "
             "An identity is `expression = 0` over roles; `statements.py` finds the rows that play the roles by which "
             "identities hold in the saved values, and uses labels to break ties and, as a last resort, to name a row.", "",
             "Each identity is a *test* of the model when it can genuinely fail given how its rows were chosen (they were "
             "named by a model check or by labels), and *structural* when it holds by construction (one row's own "
             "formula is the identity, or a row was chosen because its values equal its counterpart's). The column "
             "below is the default; `statements.py` decides it again for every result from how each role was bound "
             "(`check`, `value_search`, `label`, `structure` or `link`).", ""]
    for block, title in BLOCKS.items():
        lines += [f"## {title}", "", "| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |",
                  "|---|---|---|---|---|---|"]
        for i in IDENTITIES:
            if i.block == block:
                lines.append(f"| `{i.key}` | {i.title} | `{i.expr}` | {i.scope} | {', '.join(i.flex) or '-'} | {i.kind} |")
        lines += ["", "| Role | Is | Label vocabulary (tie-breaker, last-resort name) |", "|---|---|---|"]
        for r in ROLES:
            if r.block == block:
                lines.append(f"| `{r.name}` | {r.what} | `{r.vocab}` |")
        lines.append("")
    return "\n".join(lines)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--write", help="write the markdown table here")
    a = ap.parse_args()
    if a.write:
        with open(a.write, "w", encoding="utf-8") as f:
            f.write(markdown())
    else:
        print(markdown())
