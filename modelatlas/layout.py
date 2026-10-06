"""Detect each sheet's layout from cached values, so nothing is hard-coded per workbook.

Finds:
  header_row / tl_first..tl_last - a timeline: the row with the most dates, and the span of date columns
  label_col - the text-heavy column (left of the timeline) holding line-item names
  units_col - a column right of label_col with short, repeated text (e.g. "$'000", "%", "1 / 0")
"""
from collections import Counter
from datetime import date, datetime
from statistics import median

from openpyxl.utils import get_column_letter

MIN_TIMELINE = 4   # dates needed in a row to call it a timeline header
SCAN_COLS = 8      # without a timeline, look for labels in the first N used columns


def _is_text(v) -> bool:
    return isinstance(v, str) and v.strip() != ""


def _periodicity(dates: list[date]) -> str:
    if len(dates) < 2:
        return "?"
    gap = median((b - a).days for a, b in zip(dates, dates[1:]))
    for name, days in (("daily", 1), ("weekly", 7), ("monthly", 30), ("quarterly", 91),
                       ("half-yearly", 182), ("annual", 365)):
        if abs(gap - days) <= max(1, days * 0.1):
            return name
    return f"~{gap:g}-day"


def detect_layout(values: list[list]) -> dict:
    """values: calamine to_python(skip_empty_area=False) output (row 1 / col 1 at index 0)."""
    # 1. timeline: row with most date cells
    header_row, date_cols = None, []
    for i, row in enumerate(values):
        cols = [j for j, v in enumerate(row) if isinstance(v, (date, datetime))]
        if len(cols) >= MIN_TIMELINE and len(cols) > len(date_cols):
            header_row, date_cols = i + 1, cols
    layout: dict = {"header_row": header_row, "tl_first": None, "tl_last": None,
                    "periods": 0, "period_start": None, "period_end": None, "periodicity": None}
    if date_cols:
        # keep the longest run of consecutive date columns
        runs, cur = [], [date_cols[0]]
        for c in date_cols[1:]:
            if c == cur[-1] + 1:
                cur.append(c)
            else:
                runs.append(cur)
                cur = [c]
        runs.append(cur)
        run = max(runs, key=len)
        if len(run) >= MIN_TIMELINE:
            dates = [values[header_row - 1][c] for c in run]
            dates = [d.date() if isinstance(d, datetime) else d for d in dates]
            layout.update(tl_first=run[0] + 1, tl_last=run[-1] + 1, periods=len(run),
                          period_start=str(min(dates)), period_end=str(max(dates)),
                          periodicity=_periodicity(dates))
        else:
            layout["header_row"] = None

    # 2. label column: most (and longest) text, left of the timeline
    width = max((len(r) for r in values), default=0)
    first_used = next((j for j in range(width) if any(j < len(r) and r[j] not in ("", None) for r in values)), 0)
    limit = (layout["tl_first"] - 1) if layout["tl_first"] else min(width, first_used + SCAN_COLS)
    stats = {}
    for j in range(limit):
        texts = [r[j].strip() for r in values if j < len(r) and _is_text(r[j])]
        if texts:
            stats[j] = texts
    label_col = units_col = None
    if stats:
        label_col = max(stats, key=lambda j: len(stats[j]) * min(sum(map(len, stats[j])) / len(stats[j]), 20))
        # 3. units column: right of label, short and repetitive text
        best = 0.0
        for j, texts in stats.items():
            if j <= label_col or len(texts) < 0.2 * len(stats[label_col]):
                continue
            avg_len = sum(map(len, texts)) / len(texts)
            distinct = len(Counter(texts)) / len(texts)
            if avg_len <= 12 and distinct <= 0.6 and len(texts) > best:
                units_col, best = j, len(texts)
    layout["label_col"] = label_col + 1 if label_col is not None else None
    layout["units_col"] = units_col + 1 if units_col is not None else None
    return layout


def describe(layout: dict) -> str:
    """One-line human/LLM readable layout summary."""
    L = lambda c: get_column_letter(c) if c else "-"
    parts = [f"labels {L(layout['label_col'])}"]
    if layout["units_col"]:
        parts.append(f"units {L(layout['units_col'])}")
    if layout["tl_first"]:
        parts.append(f"timeline {L(layout['tl_first'])}:{L(layout['tl_last'])} = {layout['periods']} "
                     f"{layout['periodicity']} periods {layout['period_start']}..{layout['period_end']} "
                     f"(dates in row {layout['header_row']})")
    else:
        parts.append("no timeline")
    return ", ".join(parts)
