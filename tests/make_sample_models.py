"""Write the synthetic Riverbend workbooks the checks use to tests/sample_models/ (git-ignored; re-run to regenerate).

A fictional toll road ("Riverbend Toll Road", code name Project Kestrel) with two years' client models and the
valuation sheets a valuer adds to one:

  Riverbend_BP25_client_model.xlsx    Inputs / Operations / CashFlow
  Riverbend_BP25_with_overlay.xlsx    the same, with Val_Inputs / DCF / Summary sheets in the one file
  Riverbend_BP26_client_model.xlsx    the next year's model, with an extra insurance line

Every formula's value is saved as Excel would, so the workbooks can be read without a formula engine.
    uv run python tests/make_sample_models.py
"""
import sys
from datetime import date
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as COL

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from modelatlas import dcf  # noqa: E402

OUT = ROOT / "tests" / "sample_models"
YEARS = 20
FIRST_COL = 3  # column D


# ---- the client model ------------------------------------------------------------------------------------

def client_numbers(fy0: int, traffic0: float, growth: float, toll0: float, cpi: float, opex_pct: float,
                   capex: float, major: float, insurance: float | None) -> dict:
    ends = [date(fy0 + i, 6, 30) for i in range(YEARS)]
    traffic = [traffic0 * (1 + growth) ** i for i in range(YEARS)]
    toll = [toll0 * (1 + cpi) ** i for i in range(YEARS)]
    revenue = [t * p for t, p in zip(traffic, toll)]
    opex = [-r * opex_pct for r in revenue]
    ins = [-(insurance or 0) * (1 + cpi) ** i for i in range(YEARS)]
    ebitda = [r + o + s for r, o, s in zip(revenue, opex, ins)]
    cap = [-(capex * (1 + cpi) ** i + (major if (fy0 + i) % 5 == 0 else 0)) for i in range(YEARS)]
    tax = [-max(0.0, (e + c) * 0.30) for e, c in zip(ebitda, cap)]
    fcf = [e + c + t for e, c, t in zip(ebitda, cap, tax)]
    return dict(ends=ends, traffic=traffic, toll=toll, revenue=revenue, opex=opex, insurance=ins, ebitda=ebitda,
                capex=cap, tax=tax, fcf=fcf)


def write_client(wb, n: dict, inputs: dict, insurance: bool) -> dict:
    """Inputs / Operations / CashFlow sheets. Returns {line item: (sheet, row)} (1-based rows)."""
    b = wb.add_format({"bold": True})
    pct, num, dt = wb.add_format({"num_format": "0.00%"}), wb.add_format({"num_format": "#,##0.0"}), \
        wb.add_format({"num_format": "dd-mmm-yy"})
    i = wb.add_worksheet("Inputs")
    i.write(0, 0, "Riverbend Toll Road - business plan inputs", b)
    at = {}
    for r, (label, value, unit) in enumerate(inputs["rows"], start=2):
        i.write(r, 0, label)
        if isinstance(value, date):
            i.write_datetime(r, 1, value, dt)
        else:
            i.write(r, 1, value, pct if unit == "%" else num)
        i.write(r, 2, unit)
        at[label] = f"Inputs!$B${r + 1}"

    def timeline(ws, title):
        ws.write(0, 0, title, b)
        ws.write(2, 0, "Period ending", b)
        for k, d in enumerate(n["ends"]):
            ws.write_datetime(2, FIRST_COL + k, d, dt)
        ws.write(3, 0, "Financial year")
        for k, d in enumerate(n["ends"]):
            ws.write(3, FIRST_COL + k, f"FY{d.year % 100:02d}")

    rows = {}
    o = wb.add_worksheet("Operations")
    timeline(o, "Operations")
    op_rows = [("Traffic", "m trips", "traffic"), ("Toll (nominal)", "A$", "toll"), ("Toll revenue", "A$m", "revenue"),
               ("Operating costs", "A$m", "opex")]
    if insurance:
        op_rows.append(("Insurance", "A$m", "insurance"))
    op_rows.append(("EBITDA", "A$m", "ebitda"))
    for k, (label, unit, key) in enumerate(op_rows):
        rows[key] = ("Operations", 6 + k)
    for key, (sheet, r) in rows.items():
        label, unit = next((lab, u) for lab, u, kk in op_rows if kk == key)
        o.write(r - 1, 0, label)
        o.write(r - 1, 1, unit)
    for k in range(YEARS):
        c, p = COL(FIRST_COL + k), COL(FIRST_COL + k - 1)
        rr = {key: r for key, (_, r) in rows.items()}
        o.write_formula(f"{c}{rr['traffic']}", f"={at['Opening traffic']}" if k == 0 else
                        f"={p}{rr['traffic']}*(1+{at['Traffic growth']})", num, n["traffic"][k])
        o.write_formula(f"{c}{rr['toll']}", f"={at['Toll at start of forecast']}" if k == 0 else
                        f"={p}{rr['toll']}*(1+{at['CPI']})", num, n["toll"][k])
        o.write_formula(f"{c}{rr['revenue']}", f"={c}{rr['traffic']}*{c}{rr['toll']}", num, n["revenue"][k])
        o.write_formula(f"{c}{rr['opex']}", f"=-{c}{rr['revenue']}*{at['Operating costs (% of revenue)']}", num,
                        n["opex"][k])
        parts = f"{c}{rr['revenue']}+{c}{rr['opex']}"
        if insurance:
            o.write_formula(f"{c}{rr['insurance']}", f"=-{at['Insurance premium']}*(1+{at['CPI']})^{k}", num,
                            n["insurance"][k])
            parts += f"+{c}{rr['insurance']}"
        o.write_formula(f"{c}{rr['ebitda']}", f"={parts}", num, n["ebitda"][k])

    cf = wb.add_worksheet("CashFlow")
    timeline(cf, "Cash flow")
    cf_rows = [("EBITDA", "ebitda"), ("Capital expenditure", "capex"), ("Tax paid", "tax"),
               ("Unlevered free cash flow", "fcf")]
    for k, (label, key) in enumerate(cf_rows):
        cf.write(5 + k, 0, label)
        cf.write(5 + k, 1, "A$m")
        rows[f"cf_{key}"] = ("CashFlow", 6 + k)
    e_row = rows["ebitda"][1]
    for k in range(YEARS):
        c = COL(FIRST_COL + k)
        fy = n["ends"][k].year
        cf.write_formula(f"{c}6", f"=Operations!{c}{e_row}", num, n["ebitda"][k])
        major = f"+IF(MOD({fy},5)=0,{at['Major maintenance (every 5 years)']},0)"
        cf.write_formula(f"{c}7", f"=-({at['Maintenance capex']}*(1+{at['CPI']})^{k}{major})", num, n["capex"][k])
        cf.write_formula(f"{c}8", f"=-MAX(0,({c}6+{c}7)*{at['Tax rate']})", num, n["tax"][k])
        cf.write_formula(f"{c}9", f"={c}6+{c}7+{c}8", num, n["fcf"][k])
    return rows


# ---- the overlay (valuation) -----------------------------------------------------------------------------

def valuation(n: dict, vd: date, rate: float, g: float, net_debt: float) -> dict:
    ends = {k: e for k, e in enumerate(n["ends"])}
    df = dcf.factors(ends, vd, rate, "end", "actual/actual")
    tv = n["fcf"][-1] * (1 + g) / (rate - g)
    flows = [f + (tv if k == YEARS - 1 else 0) for k, f in enumerate(n["fcf"])]
    ev = sum(flows[k] * df[k] for k in range(YEARS))
    return dict(df=[df[k] for k in range(YEARS)], tv=tv, flows=flows, ev=ev, equity=ev - net_debt)


def write_overlay(wb, n: dict, cf_ref, vd: date, rate: float, g: float, net_debt: float) -> dict:
    """Val_Inputs / DCF / Summary. cf_ref(col_letter) -> formula text reading the client's free cash flow."""
    b = wb.add_format({"bold": True})
    pct, num, dt = wb.add_format({"num_format": "0.00%"}), wb.add_format({"num_format": "#,##0.0"}), \
        wb.add_format({"num_format": "dd-mmm-yy"})
    v = valuation(n, vd, rate, g, net_debt)
    vi = wb.add_worksheet("Val_Inputs")
    vi.write(0, 0, "Valuation assumptions", b)
    vi.write(3, 0, "Valuation date"); vi.write_datetime(3, 2, vd, dt)
    vi.write(4, 0, "Discount rate (post-tax nominal WACC)"); vi.write(4, 2, rate, pct)
    vi.write(5, 0, "Terminal growth rate"); vi.write(5, 2, g, pct)
    vi.write(6, 0, "Net debt at valuation date"); vi.write(6, 1, "A$m"); vi.write(6, 2, net_debt, num)

    d = wb.add_worksheet("DCF")
    d.write(0, 0, "Discounted cash flow", b)
    d.write(2, 0, "Period ending", b)
    for k, e in enumerate(n["ends"]):
        d.write_datetime(2, FIRST_COL + k, e, dt)
    labels = [(5, "Unlevered free cash flow (client model)"), (6, "Terminal value"), (7, "Valuation cash flow"),
              (9, "Discount factor")]
    for r, lab in labels:
        d.write(r - 1, 0, lab)
        d.write(r - 1, 1, "" if r == 9 else "A$m")
    last = COL(FIRST_COL + YEARS - 1)
    for k in range(YEARS):
        c = COL(FIRST_COL + k)
        d.write_formula(f"{c}5", cf_ref(c), num, n["fcf"][k])
        if k == YEARS - 1:
            d.write_formula(f"{c}6", f"={c}5*(1+Val_Inputs!$C$6)/(Val_Inputs!$C$5-Val_Inputs!$C$6)", num, v["tv"])
        else:
            d.write(f"{c}6", 0, num)
        d.write_formula(f"{c}7", f"={c}5+{c}6", num, v["flows"][k])
        d.write_formula(f"{c}9", f"=1/(1+Val_Inputs!$C$5)^YEARFRAC(Val_Inputs!$C$4,{c}$3,1)",
                        wb.add_format({"num_format": "0.0000"}), v["df"][k])
    d.write(11, 0, "Enterprise value"); d.write(11, 1, "A$m")
    d.write_formula("D12", f"=SUMPRODUCT(D7:{last}7,D9:{last}9)", num, v["ev"])
    d.write(12, 0, "Less: net debt"); d.write(12, 1, "A$m")
    d.write_formula("D13", "=-Val_Inputs!$C$7", num, -net_debt)
    d.write(13, 0, "Equity value"); d.write(13, 1, "A$m")
    d.write_formula("D14", "=D12+D13", num, v["equity"])

    lo, hi = valuation(n, vd, rate + 0.0025, g, net_debt), valuation(n, vd, rate - 0.0025, g, net_debt)
    s = wb.add_worksheet("Summary")
    s.write(0, 0, "Valuation summary (A$m)", b)
    for c, h in enumerate(["", "Low", "Preferred", "High"]):
        s.write(2, c, h, b)
    s.write(3, 0, "Enterprise value"); s.write(3, 1, round(lo["ev"], 1), num)
    s.write_formula("C4", "=DCF!D12", num, v["ev"]); s.write(3, 3, round(hi["ev"], 1), num)
    s.write(4, 0, "Equity value"); s.write(4, 1, round(lo["equity"], 1), num)
    s.write_formula("C5", "=DCF!D14", num, v["equity"]); s.write(4, 3, round(hi["equity"], 1), num)
    s.write(5, 0, "Low / high: discount rate +/- 0.25% (pasted from the sensitivity run)")
    return {**v, "low": lo, "high": hi}



def inputs_rows(start: date, traffic0, growth, toll0, cpi, opex_pct, capex, major, insurance=None):
    rows = [("Forecast start", start, "date"), ("Opening traffic", traffic0, "m trips"), ("Traffic growth", growth, "%"),
            ("Toll at start of forecast", toll0, "A$"), ("CPI", cpi, "%"),
            ("Operating costs (% of revenue)", opex_pct, "%"), ("Maintenance capex", capex, "A$m"),
            ("Major maintenance (every 5 years)", major, "A$m"), ("Tax rate", 0.30, "%")]
    if insurance is not None:
        rows.append(("Insurance premium", insurance, "A$m"))
    return {"rows": rows}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    vd, rate, g, net_debt = date(2025, 6, 30), 0.0725, 0.025, 850.0
    prior_in = dict(traffic0=42.0, growth=0.02, toll0=6.50, cpi=0.025, opex_pct=0.18, capex=35.0, major=120.0)
    cur_in = dict(traffic0=43.1, growth=0.021, toll0=6.70, cpi=0.03, opex_pct=0.18, capex=38.0, major=120.0)
    prior = client_numbers(2026, **prior_in, insurance=None)
    current = client_numbers(2027, **cur_in, insurance=4.0)

    files = []
    p = OUT / "Riverbend_BP25_client_model.xlsx"
    wb = xlsxwriter.Workbook(p)
    rows = write_client(wb, prior, inputs_rows(date(2025, 7, 1), **prior_in), insurance=False)
    wb.close()
    files.append(p)

    p = OUT / "Riverbend_BP26_client_model.xlsx"
    wb = xlsxwriter.Workbook(p)
    write_client(wb, current, inputs_rows(date(2026, 7, 1), **cur_in, insurance=4.0), insurance=True)
    wb.close()
    files.append(p)

    fcf_row = rows["cf_fcf"][1]
    p = OUT / "Riverbend_BP25_with_overlay.xlsx"
    wb = xlsxwriter.Workbook(p)
    write_client(wb, prior, inputs_rows(date(2025, 7, 1), **prior_in), insurance=False)
    v = write_overlay(wb, prior, lambda c: f"=CashFlow!{c}{fcf_row}", vd, rate, g, net_debt)
    wb.close()
    files.append(p)

    for f in files:
        print(f.relative_to(ROOT))
    print(f"\nprior: EV {v['ev']:,.0f}, equity {v['equity']:,.0f}, rate {rate:.2%}, TGR {g:.2%}")


if __name__ == "__main__":
    main()
