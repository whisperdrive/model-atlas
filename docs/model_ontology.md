# Model ontology

Generated from `modelatlas/ontology.py` (`uv run python -m modelatlas.ontology --write docs/model_ontology.md`). An identity is `expression = 0` over roles; `statements.py` finds the rows that play the roles by which identities hold in the saved values, and uses labels to break ties and, as a last resort, to name a row.

Each identity is a *test* of the model when it can genuinely fail given how its rows were chosen (they were named by a model check or by labels), and *structural* when it holds by construction (one row's own formula is the identity, or a row was chosen because its values equal its counterpart's). The column below is the default; `statements.py` decides it again for every result from how each role was bound (`check`, `value_search`, `label`, `structure` or `link`).

## Income statement (P&L)

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `pnl_ebitda` | Revenue plus operating costs is EBITDA | `revenue + opex - ebitda` | per_period | opex | structural |
| `pnl_ebit` | EBITDA plus depreciation is EBIT | `ebitda + depreciation - ebit` | per_period | depreciation | structural |
| `pnl_pbt` | EBIT plus interest is profit before tax | `ebit + interest - pbt` | per_period | interest | structural |
| `pnl_npat` | Profit before tax plus tax is NPAT | `pbt + tax - npat` | per_period | tax | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `revenue` | revenue line of the P&L | `\b(revenues?|turnover|sales)\b|total income` |
| `opex` | operating costs | `operating (costs?|expenses?)|opex|\bcosts?\b|expenses?` |
| `ebitda` | EBITDA | `\bebitda\b` |
| `depreciation` | depreciation and amortisation | `deprec|amorti[sz]` |
| `ebit` | EBIT | `\bebit\b|operating (profit|income|result)` |
| `interest` | interest / finance cost | `interest|finance (costs?|charges?)` |
| `pbt` | profit before tax | `\bpbt\b|before tax` |
| `tax` | tax expense | `\btax(es|ation)?\b` |
| `npat` | net profit after tax | `\bnpat\b|net (profit|income|earnings)|profit after tax|profit for the (year|period)` |

## Balance sheet

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `bs_balance` | Assets equal liabilities plus equity, every period | `assets - liabilities - equity` | per_period | liabilities, equity | test |
| `bs_total_le` | Assets equal total liabilities and equity | `assets - total_le` | per_period | - | test |
| `bs_le_sum` | Liabilities plus equity is total liabilities and equity | `liabilities + equity - total_le` | per_period | liabilities, equity | test |
| `bs_rollforward` | Every balance sheet line has a roll-forward (or never moves) | `coverage` | coverage | - | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `assets` | total assets | `^(total )?assets\b(?!.*liabilit)` |
| `liabilities` | total liabilities | `^(total )?liabilities\b(?!.*equity)` |
| `equity` | total equity | `^total equity\b(?!.*liabilit)|^(shareholders'? )?equity$|shareholders'? (equity|funds)|net assets` |
| `total_le` | total liabilities and equity | `liabilities (and|&) (shareholders'? )?equity|equity (and|&) liabilities` |
| `bs_cash` | cash line | `^cash( and cash equivalents| at bank)?$` |
| `bs_fixed_assets` | fixed assets line | `fixed assets|\bppe\b|property.*(plant|equipment)|non-?current assets|infrastructure|intangible` |
| `bs_debt` | debt line | `debt|borrowing|loans?\b|facilit|\bnotes\b|bonds` |
| `bs_retained` | retained earnings line | `retained|accumulated (profit|earnings)|reserves` |

## Cash flow statement

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `cf_net` | Operating, investing and financing flows sum to net cash flow | `cfo + cfi + cff - cf_net` | per_period | - | structural |
| `cf_roll` | Opening cash plus net cash flow is closing cash | `cf_opening + cf_net - cf_closing` | per_period | - | structural |
| `cf_cash_tie` | Closing cash on the cash flow statement is cash on the balance sheet | `cf_closing - bs_cash` | per_period | - | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `cf_opening` | opening cash | `opening|beginning|start of` |
| `cfo` | cash flow from operations | `operating activities|operations|\bcfo\b|operating cash` |
| `cfi` | cash flow from investing | `investing|\bcfi\b` |
| `cff` | cash flow from financing | `financing|\bcff\b` |
| `cf_net` | net cash flow | `net (cash|increase|change|movement)|net cash ?flow` |
| `cf_closing` | closing cash | `closing cash|cash at (the )?end|ending cash` |
| `cf_capex` | capital expenditure in investing | `capex|capital expenditure|purchase of` |
| `cf_distributions` | distributions paid in financing | `distribution|dividend` |

## Debt schedule

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `debt_roll` | Opening debt plus drawdowns and repayments is closing debt | `debt_opening + debt_movements* - debt_closing` | per_period | - | structural |
| `debt_tie` | Closing debt is the balance sheet debt line | `debt_closing - bs_debt` | per_period | bs_debt | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `debt_opening` | opening debt | `opening|beginning` |
| `debt_movements*` | drawdowns and repayments | `draw|repay|advance|issue|redemption|principal|amorti` |
| `debt_closing` | closing debt | `debt|loan|borrowing|facilit|closing balance|principal|\bnotes\b` |

## Fixed assets schedule

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `fa_roll` | Opening fixed assets plus additions and depreciation is closing | `fa_opening + fa_movements* - fa_closing` | per_period | - | structural |
| `fa_tie` | Closing fixed assets is the balance sheet fixed assets line | `fa_closing - bs_fixed_assets` | per_period | - | structural |
| `fa_dep` | Depreciation in the schedule is depreciation in the P&L (up to sign) | `fa_depreciation - depreciation` | per_period | depreciation | test |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `fa_opening` | opening fixed assets | `opening|beginning` |
| `fa_movements*` | additions and depreciation | `addition|capex|purchase|deprec|amorti|dispos` |
| `fa_closing` | closing fixed assets | `closing fixed|fixed assets|\bppe\b` |
| `fa_additions` | additions | `addition|capex|purchase` |
| `fa_depreciation` | depreciation charge | `deprec|amorti` |

## Sources and uses

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `su_balance` | Total sources equal total uses | `sources - uses` | single | - | test |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `sources` | total sources | `sources` |
| `uses` | total uses | `\buses\b|application of funds` |

## Capital expenditure

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `capex_cf` | Total capex is the capital expenditure in cash flow from investing (up to sign) | `capex_total - cf_capex` | per_period | cf_capex | structural |
| `fa_additions` | Additions in the fixed assets schedule are minus total capex | `fa_additions + capex_total` | per_period | capex_total | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `capex_total` | total capex | `total capex|capex total|total capital expenditure|^capex$|capital expenditure total` |

## Revenue build

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `rev_sum` | The revenue segments sum to total revenue | `rev_segments* - rev_total` | per_period | - | structural |
| `rev_pnl` | Total revenue is the P&L revenue line | `rev_total - revenue` | per_period | - | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `rev_segments*` | revenue segments | `revenue|sales` |
| `rev_total` | total revenue | `total revenue|revenue total|total sales|^revenue$|turnover` |

## Equity (retained earnings) roll-forward

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `re_roll` | Opening retained earnings plus profit and distributions is closing | `re_opening + re_movements* - re_closing` | per_period | - | structural |
| `re_npat` | Profit entering retained earnings is the P&L NPAT | `re_npat - npat` | per_period | - | structural |
| `re_tie` | Closing retained earnings is the balance sheet retained earnings line | `re_closing - bs_retained` | per_period | - | structural |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `re_opening` | opening retained earnings | `opening|beginning` |
| `re_movements*` | profit and distributions | `npat|profit|income|distribution|dividend|transfer` |
| `re_closing` | closing retained earnings | `closing|ending|retained` |
| `re_npat` | profit entering retained earnings | `npat|profit|income|earnings` |
| `re_distributions` | distributions out of retained earnings | `distribution|dividend` |

## Distributions

| Identity | Says | Expression = 0 | Scope | Sign tested both ways | Kind |
|---|---|---|---|---|---|
| `dist_dcf_cf` | The DCF's cash-flow row is the distributions paid in the cash flow (up to sign) | `dcf_distributions - cf_distributions` | per_period | cf_distributions | structural |
| `dist_cf_eq` | Distributions paid in the cash flow are the equity roll-forward's (up to sign) | `cf_distributions - re_distributions` | per_period | re_distributions | test |

| Role | Is | Label vocabulary (tie-breaker, last-resort name) |
|---|---|---|
| `dcf_distributions` | the cash-flow row the DCF discounts | `distribution|dividend|cash ?flow|fcfe` |
