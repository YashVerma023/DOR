"""DOR.html builder.

Produces the shareable Daily Operations Report: a single self-contained HTML
file (inline CSS, no external assets) holding ONLY the summary data —
KPIs, the per-algo trade value outlier summary, the outlier mix bars, and the
Int / Pos+Int segregation pivot. The full row-level data stays in the Excel
workbook; this file is the client-facing view.
"""

import html
import math
import statistics
from datetime import datetime


def _money(v):
    return f"{float(v):,.0f}"


def _num2(v):
    return f"{float(v):,.2f}"


def _esc(v):
    return html.escape(str(v))


def _kpi_tiles(pairs):
    tiles = "".join(
        f'<div class="kpi"><div class="kpi-label">{_esc(label)}</div>'
        f'<div class="kpi-value">{_esc(value)}</div></div>'
        for label, value in pairs
    )
    return f'<div class="kpi-grid">{tiles}</div>'


def _tv_summary_table(tv_summary):
    body = []
    for s in tv_summary:
        band = f"{_num2(s['band_low'])} – {_num2(s['band_high'])}"
        body.append(
            "<tr>"
            f'<td class="txt">{_esc(s["algo"])}</td>'
            f"<td>{s['rows']}</td>"
            f"<td>{_num2(s['avg_lot_pct'])}</td>"
            f"<td>{_num2(s['std_dev'])}</td>"
            f"<td>{_esc(band)}</td>"
            f'<td class="below">{s["below"]}</td>'
            f"<td>{s['in_range']}</td>"
            f'<td class="above">{s["above"]}</td>'
            "</tr>"
        )
    return f"""
    <table class="tbl">
      <thead><tr>
        <th class="txt">Algo</th><th>Rows</th><th>Avg Lot Pct</th><th>Std Dev</th>
        <th>Band</th><th>Below</th><th>In</th><th>Above</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


def _nice_ticks(vmax, target=5):
    """Round y-axis ticks from 0 up to (at least) vmax."""
    if vmax <= 0:
        vmax = 1.0
    magnitude = 10 ** math.floor(math.log10(vmax / target))
    step = magnitude
    for mult in (1, 2, 2.5, 5, 10):
        step = mult * magnitude
        if vmax / step <= target + 1:
            break
    ticks, t = [], 0.0
    while t < vmax + step * 0.999:
        ticks.append(round(t, 10))
        t += step
    return ticks


def _boxplot_svg(tv_rows):
    """Inline-SVG box plot of lot pct per algo — self-contained (no JS), same
    design as the dashboard chart: gray box/whiskers for the distribution,
    flagged outliers overlaid as blue (below) / orange (above) points with
    native <title> hover tooltips."""
    by_algo = {}
    for r in tv_rows:
        if r.get("lot_pct") is None or not r["algo"]:
            continue
        by_algo.setdefault(str(r["algo"]), []).append(r)
    if not by_algo:
        return ""
    algos = sorted(by_algo, key=lambda a: (0, int(a)) if a.isdigit() else (1, a))
    vmax = max(float(r["lot_pct"]) for rows in by_algo.values() for r in rows)
    ticks = _nice_ticks(vmax * 1.05)
    top = ticks[-1]

    W, H = 1040, 420
    ml, mr, mt, mb = 54, 16, 14, 44
    plot_w, plot_h = W - ml - mr, H - mt - mb

    def ypix(v):
        return mt + plot_h - (v / top) * plot_h

    slot = plot_w / len(algos)
    box_w = min(64.0, slot * 0.5)

    parts = []
    for t in ticks:
        y = ypix(t)
        parts.append(f'<line x1="{ml}" y1="{y:.1f}" x2="{ml + plot_w}" y2="{y:.1f}" '
                     f'stroke="#E5E7EB" stroke-width="1"/>')
        parts.append(f'<text x="{ml - 8}" y="{y:.1f}" text-anchor="end" '
                     f'dominant-baseline="middle" class="bx-tick">{t:g}</text>')

    for i, algo in enumerate(algos):
        rows = by_algo[algo]
        cx = ml + slot * (i + 0.5)
        vals = sorted(float(r["lot_pct"]) for r in rows)
        if len(vals) == 1:
            q1 = med = q3 = vals[0]
        else:
            q1, med, q3 = statistics.quantiles(vals, n=4, method="inclusive")
        iqr = q3 - q1
        lo = min((v for v in vals if v >= q1 - 1.5 * iqr), default=vals[0])
        hi = max((v for v in vals if v <= q3 + 1.5 * iqr), default=vals[-1])

        cap_w = box_w * 0.5
        parts.append(f'<line x1="{cx:.1f}" y1="{ypix(lo):.1f}" x2="{cx:.1f}" y2="{ypix(hi):.1f}" '
                     f'stroke="#6B7684" stroke-width="1.4"/>')
        for v in (lo, hi):
            parts.append(f'<line x1="{cx - cap_w / 2:.1f}" y1="{ypix(v):.1f}" '
                         f'x2="{cx + cap_w / 2:.1f}" y2="{ypix(v):.1f}" '
                         f'stroke="#6B7684" stroke-width="1.4"/>')
        box_h = max(ypix(q1) - ypix(q3), 1.5)
        parts.append(f'<rect x="{cx - box_w / 2:.1f}" y="{ypix(q3):.1f}" width="{box_w:.1f}" '
                     f'height="{box_h:.1f}" fill="#9AA4B1" fill-opacity="0.45" '
                     f'stroke="#6B7684" stroke-width="1.2" rx="2"/>')
        parts.append(f'<line x1="{cx - box_w / 2:.1f}" y1="{ypix(med):.1f}" '
                     f'x2="{cx + box_w / 2:.1f}" y2="{ypix(med):.1f}" '
                     f'stroke="#40474F" stroke-width="2"/>')
        parts.append(f'<text x="{cx:.1f}" y="{H - mb + 20}" text-anchor="middle" '
                     f'class="bx-lab">{_esc(algo)}</text>')

        flagged = [r for r in rows if r["outlier"] in ("Below average range", "Above average range")]
        for j, r in enumerate(flagged):
            color = "#0072B2" if r["outlier"] == "Below average range" else "#E69F00"
            dx = ((j % 7) - 3) * 2.6  # deterministic jitter so stacked points stay visible
            tip = f"{r['user_id']} · {r['server']} · lot pct {float(r['lot_pct']):.2f}"
            parts.append(f'<circle cx="{cx + dx:.1f}" cy="{ypix(float(r["lot_pct"])):.1f}" r="3.5" '
                         f'fill="{color}" fill-opacity="0.85"><title>{_esc(tip)}</title></circle>')

    parts.append(f'<text x="{ml + plot_w / 2:.1f}" y="{H - 6}" text-anchor="middle" class="bx-title">Algo</text>')
    parts.append(f'<text x="14" y="{mt + plot_h / 2:.1f}" text-anchor="middle" class="bx-title" '
                 f'transform="rotate(-90 14 {mt + plot_h / 2:.1f})">Lot Pct</text>')

    legend = """
    <div class="legend">
      <span><i class="dot dot-below"></i>Below range (outlier)</span>
      <span><i class="dot dot-above"></i>Above range (outlier)</span>
      <span class="legend-note">box = middle 50% of users (line = median) · whiskers = 1.5&times;IQR · hover a point for the user</span>
    </div>"""
    svg = (f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="Box plot of lot pct per algo" '
           f'style="width:100%;height:auto;display:block">{"".join(parts)}</svg>')
    return legend + svg


def _metric_cells(r):
    return (
        f"<td>{r['Users']}</td>"
        f"<td>{r['SLHit']}</td>"
        f"<td>{_money(r['MaxLoss'])}</td>"
        f"<td>{_money(r['Allocation'])}</td>"
        f"<td>{_money(r['Realized'])}</td>"
        f"<td>{_money(r['Unrealized'])}</td>"
        f"<td>{_money(r['MTM'])}</td>"
        f"<td>{_num2(r['Return'])}</td>"
        f"<td>{_num2(r['P95'])}</td>"
        f"<td>{_num2(r['P5'])}</td>"
    )


def _pivot_table(pivot_rows):
    """Drill-down pivot: one 'Grand Total' row per algo (collapsed default);
    clicking it reveals the Int / Pos+Int sub-total rows, and clicking those
    reveals the per-server detail. Plain inline JS — no libraries."""
    # Rebuild the hierarchy from the flat rows.
    buckets, order, grand = {}, [], None
    for r in pivot_rows:
        if r["kind"] == "grandtotal":
            grand = r
            continue
        algo = str(r.get("algo_val", ""))
        if algo not in buckets:
            buckets[algo] = {"total": None, "sections": {}, "sec_order": []}
            order.append(algo)
        bucket = buckets[algo]
        if r["kind"] == "algototal":
            bucket["total"] = r
            continue
        section = r["Section"]
        if section not in bucket["sections"]:
            bucket["sections"][section] = {"rows": [], "subtotal": None}
            bucket["sec_order"].append(section)
        if r["kind"] == "subtotal":
            bucket["sections"][section]["subtotal"] = r
        else:
            bucket["sections"][section]["rows"].append(r)

    body = []
    for i, algo in enumerate(order):
        bucket = buckets[algo]
        total = bucket["total"]
        if total is None:
            continue
        akey = f"a{i}"
        body.append(
            f'<tr class="algototal lvl1" data-key="{akey}">'
            f'<td class="txt"><span class="caret">&#9656;</span>Grand Total</td>'
            f'<td class="txt">{_esc(algo)}</td>'
            f'<td class="txt">{_esc(total["SERVER"])}</td>'
            f"{_metric_cells(total)}</tr>"
        )
        for j, section in enumerate(bucket["sec_order"]):
            block = bucket["sections"][section]
            subtotal = block["subtotal"]
            if subtotal is None:
                continue
            skey = f"{akey}s{j}"
            body.append(
                f'<tr class="subtotal lvl2 hidden" data-parent="{akey}" data-key="{skey}">'
                f'<td class="txt"><span class="caret">&#9656;</span>{_esc(section)}</td>'
                f'<td class="txt">{_esc(algo)}</td>'
                f'<td class="txt">{_esc(subtotal["SERVER"])}</td>'
                f"{_metric_cells(subtotal)}</tr>"
            )
            for r in block["rows"]:
                body.append(
                    f'<tr class="lvl3 hidden" data-parent="{skey}">'
                    f'<td class="txt indent">{_esc(section)}</td>'
                    f'<td class="txt">{_esc(algo)}</td>'
                    f'<td class="txt">{_esc(r["SERVER"])}</td>'
                    f"{_metric_cells(r)}</tr>"
                )
    if grand is not None:
        body.append(
            '<tr class="grand">'
            '<td class="txt">Total</td>'
            '<td class="txt">Grand Total</td>'
            f'<td class="txt">{_esc(grand["SERVER"])}</td>'
            f"{_metric_cells(grand)}</tr>"
        )

    return f"""
    <div class="pivot-tools">
      <button type="button" id="pivot-expand">Expand all</button>
      <button type="button" id="pivot-collapse">Collapse all</button>
      <span class="hint">click a row to drill down: Algo &rarr; Int / Pos+Int &rarr; servers</span>
    </div>
    <table class="tbl" id="pivot-table">
      <thead><tr>
        <th class="txt">Type</th><th class="txt">Algo</th><th class="txt">Server</th>
        <th>Users</th><th>SL Hit</th><th>Max Loss</th><th>Allocation</th>
        <th>Realized P&amp;L</th><th>Unrealized P&amp;L</th><th>MTM</th>
        <th>Return %</th><th>95%</th><th>5%</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


_PIVOT_SCRIPT = """
<script>
(function () {
  function childrenOf(key) {
    return document.querySelectorAll('tr[data-parent="' + key + '"]');
  }
  function collapse(tr) {
    tr.classList.remove('open');
    childrenOf(tr.dataset.key).forEach(function (child) {
      child.classList.add('hidden');
      if (child.dataset.key) collapse(child);
    });
  }
  document.querySelectorAll('#pivot-table tr[data-key]').forEach(function (tr) {
    tr.addEventListener('click', function () {
      if (tr.classList.contains('open')) { collapse(tr); return; }
      tr.classList.add('open');
      childrenOf(tr.dataset.key).forEach(function (child) { child.classList.remove('hidden'); });
    });
  });
  var expandAll = document.getElementById('pivot-expand');
  var collapseAll = document.getElementById('pivot-collapse');
  if (expandAll) expandAll.addEventListener('click', function () {
    document.querySelectorAll('#pivot-table tr[data-key]').forEach(function (tr) { tr.classList.add('open'); });
    document.querySelectorAll('#pivot-table tr[data-parent]').forEach(function (tr) { tr.classList.remove('hidden'); });
  });
  if (collapseAll) collapseAll.addEventListener('click', function () {
    document.querySelectorAll('#pivot-table tr[data-key]').forEach(collapse);
  });
})();
</script>
"""


def build_dor_html(report_date, deviation, tv_totals, tv_summary, pivot_stats, pivot_rows,
                   tv_rows=None):
    """Assemble the complete DOR.html document and return it as a string.
    `tv_rows` (the per-user report rows) feeds the box plot; when omitted the
    chart section is skipped."""
    generated = datetime.now().strftime("%d-%b-%Y %H:%M")
    grand = next((r for r in pivot_rows if r["kind"] == "grandtotal"), None)

    tv_kpis = _kpi_tiles([
        ("Users", f"{tv_totals['users']:,}"),
        ("Orders", f"{tv_totals['orders']:,}"),
        ("Total Lots", _num2(tv_totals["lots"])),
        ("Trade Value", _money(tv_totals["trade_value"])),
    ])
    seg_pairs = [
        ("Accounts", f"{pivot_stats['accounts']:,}"),
        ("Positional", f"{pivot_stats['positional']:,}"),
        ("Intraday", f"{pivot_stats['intraday']:,}"),
        ("Noren", f"{pivot_stats['noren']:,}"),
    ]
    if grand:
        seg_pairs += [
            ("Realized P&L", _money(grand["Realized"])),
            ("Unrealized P&L", _money(grand["Unrealized"])),
            ("MTM", _money(grand["MTM"])),
            ("Return %", _num2(grand["Return"])),
        ]
    seg_kpis = _kpi_tiles(seg_pairs)

    boxplot = _boxplot_svg(tv_rows) if tv_rows else ""
    boxplot_section = f"""
<section>
  <h2>Lot Pct by Algo (Box Plot)</h2>
  <div class="section-note">Distribution of each algo&rsquo;s lot pct across users;
  the &plusmn;{deviation:g}&sigma; flagged outliers are shown as points.</div>
  <div class="card">
    {boxplot}
  </div>
</section>""" if boxplot else ""

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Daily Operations Report — {_esc(report_date)}</title>
<style>
  :root {{
    --ink: #1F2937; --muted: #6B7280; --line: #E5E7EB;
    --surface: #FFFFFF; --bg: #F3F4F6;
    --header: #1E293B; --header-2: #64748B;
    --below: #0072B2; --above: #E69F00; --in: #CBD5E1;
    --subtotal: #EDF0F5; --algototal: #DCE3EC; --grand: #D1C9E1;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--ink); font-size: 14px; line-height: 1.45;
  }}
  .masthead {{
    background: var(--header); color: #fff; padding: 26px 20px;
  }}
  .masthead .inner {{ max-width: 1100px; margin: 0 auto; }}
  .masthead h1 {{ font-size: 22px; font-weight: 700; letter-spacing: .2px; }}
  .masthead .sub {{ color: #CBD5E1; margin-top: 4px; font-size: 13px; }}
  .masthead .date {{ float: right; text-align: right; font-size: 13px; color: #E2E8F0; }}
  .masthead .date b {{ display: block; font-size: 18px; color: #fff; }}
  main {{ max-width: 1100px; margin: 22px auto 40px; padding: 0 20px; }}
  section {{ margin-bottom: 28px; }}
  h2 {{
    font-size: 16px; font-weight: 700; margin-bottom: 4px; color: var(--header);
  }}
  .section-note {{ color: var(--muted); font-size: 12.5px; margin-bottom: 12px; }}
  .card {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 10px;
    padding: 16px 18px; box-shadow: 0 1px 2px rgba(15, 23, 42, .05);
  }}
  .kpi-grid {{
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 10px; margin-bottom: 14px;
  }}
  @media (max-width: 640px) {{
    .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
  }}
  .kpi {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 10px;
    padding: 12px 14px;
  }}
  .kpi-label {{ color: var(--muted); font-size: 12px; }}
  .kpi-value {{ font-size: 20px; font-weight: 700; margin-top: 2px; font-variant-numeric: tabular-nums; }}
  .tbl {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  .tbl th {{
    background: var(--header); color: #fff; padding: 8px 10px; text-align: right;
    font-weight: 600; white-space: nowrap; position: sticky; top: 0;
  }}
  .tbl th.txt {{ text-align: left; }}
  .tbl td {{
    padding: 6px 10px; border-bottom: 1px solid var(--line); text-align: right;
    font-variant-numeric: tabular-nums; white-space: nowrap;
  }}
  .tbl td.txt {{ text-align: left; }}
  .tbl tbody tr:hover td {{ background: #F8FAFC; }}
  .tbl tr.subtotal td {{ background: var(--subtotal); font-weight: 600; }}
  .tbl tr.algototal td {{ background: var(--algototal); font-weight: 700; }}
  .tbl tr.grand td {{ background: var(--grand); font-weight: 700; }}
  .tbl td.below {{ color: var(--below); font-weight: 600; }}
  .tbl td.above {{ color: #B45309; font-weight: 600; }}
  .scroll {{ overflow-x: auto; }}
  .legend {{ display: flex; gap: 18px; align-items: center; margin-bottom: 10px;
             color: var(--muted); font-size: 12.5px; flex-wrap: wrap; }}
  .legend .legend-note {{ margin-left: auto; }}
  .dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 3px; margin-right: 6px; }}
  .dot-below {{ background: var(--below); }}
  .dot-in {{ background: var(--in); }}
  .dot-above {{ background: var(--above); }}
  .bx-tick {{ font-size: 11px; fill: var(--muted); }}
  .bx-lab {{ font-size: 12.5px; font-weight: 600; fill: var(--ink); }}
  .bx-title {{ font-size: 12px; fill: var(--muted); }}
  .hidden {{ display: none; }}
  #pivot-table tr[data-key] td {{ cursor: pointer; user-select: none; }}
  .caret {{ display: inline-block; width: 15px; color: var(--muted);
            transition: transform .15s ease; }}
  tr.open .caret {{ transform: rotate(90deg); }}
  #pivot-table td.indent {{ padding-left: 30px; }}
  .pivot-tools {{ display: flex; gap: 8px; align-items: center; margin-bottom: 10px; }}
  .pivot-tools button {{
    border: 1px solid var(--line); background: var(--surface); border-radius: 6px;
    padding: 4px 12px; font-size: 12px; cursor: pointer; color: var(--ink);
  }}
  .pivot-tools button:hover {{ background: #F8FAFC; }}
  .pivot-tools .hint {{ color: var(--muted); font-size: 12px; margin-left: auto; }}
  .footnote {{ color: var(--muted); font-size: 12px; margin-top: 10px; }}
  footer {{
    max-width: 1100px; margin: 0 auto 30px; padding: 0 20px;
    color: var(--muted); font-size: 12px; border-top: 1px solid var(--line); padding-top: 12px;
  }}
  @media print {{
    body {{ background: #fff; }}
    .card, .kpi {{ box-shadow: none; }}
    .pivot-tools {{ display: none; }}
    .masthead {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .tbl th, .tbl tr.subtotal td, .tbl tr.algototal td, .tbl tr.grand td,
    .dot {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  }}
</style>
</head>
<body>
<div class="masthead">
  <div class="inner">
    <div class="date">Report date<b>{_esc(report_date)}</b></div>
    <h1>Daily Operations Report</h1>
    <div class="sub">Trade Value &amp; Intraday / Positional Segregation — summary view</div>
  </div>
</div>
<main>

<section>
  <h2>Intraday / Positional Segregation</h2>
  <div class="section-note">
    Accounts present in the Combined Max Loss file(s) are Positional; their Realized P&amp;L includes the
    (Realized PNL + Net Settlement Value) addon &mdash; Noren accounts take the 1DTE addon only.
    Within each algo, <b>Int</b> lists the intraday accounts per server and <b>Pos+Int</b> the
    positional accounts per server &mdash; the same categorisation as the Excel Segregation sheet.
  </div>
  {seg_kpis}
  <div class="card scroll">
    {_pivot_table(pivot_rows)}
  </div>
  <div class="footnote">
    *P5/P95 are the 5th and 95th percentile of per-user returns, used to represent min/max while
    controlling for data aberrations (intraday pauses/stops, under/over-funded accounts,
    suboptimal allocations etc.). Sub-Total server counts are per section; algo totals count each
    server once.
  </div>
</section>

<section>
  <h2>Trade Value</h2>
  <div class="section-note">
    Completed NIFTY / BANKNIFTY / SENSEX orders only (F&amp;O exchanges NFO / BFO), duplicates
    removed. Outliers are users whose
    lot pct (lots / allocation &times; 100) is more than {deviation:g} standard deviation(s) from
    their algo&rsquo;s average.
  </div>
  {tv_kpis}
  <div class="card scroll">
    {_tv_summary_table(tv_summary)}
  </div>
</section>

{boxplot_section}

</main>
<footer>
  Generated {_esc(generated)} &middot; Daily Operations Report &middot; Full row-level data is
  available in the accompanying Excel workbook.
</footer>
{_PIVOT_SCRIPT}
</body>
</html>"""
