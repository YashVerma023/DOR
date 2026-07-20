"""DOR.html builder.

Produces the shareable Daily Operations Report: a single self-contained HTML
file (inline CSS, no external assets) holding ONLY the summary data —
KPIs, the per-algo trade value outlier summary, the outlier mix bars, and the
Int / Pos+Int segregation pivot. The full row-level data stays in the Excel
workbook; this file is the client-facing view.
"""

import base64
import html
import json
import math
import statistics
from datetime import datetime

from tradevalue import strike_chain


def _money(v):
    return f"{float(v):,.0f}"


def _num2(v):
    return f"{float(v):,.2f}"


def _esc(v):
    return html.escape(str(v))


def _kpi_tiles(pairs):
    """Each pair is (label, value) or (label, value, css_class) — the class
    colours the value (e.g. "pos"/"neg" for profit/loss)."""
    tiles = []
    for pair in pairs:
        label, value = pair[0], pair[1]
        cls = f" {pair[2]}" if len(pair) > 2 and pair[2] else ""
        tiles.append(
            f'<div class="kpi"><div class="kpi-label">{_esc(label)}</div>'
            f'<div class="kpi-value{cls}">{_esc(value)}</div></div>'
        )
    return f'<div class="kpi-grid">{"".join(tiles)}</div>'


def _sign_cls(value):
    """CSS class for a P&L number: profit green, loss red, zero default."""
    return "pos" if float(value) > 0 else ("neg" if float(value) < 0 else "")


def _tv_summary_table(tv_summary):
    body = []
    for s in tv_summary:
        has_stats = s["avg_lot_pct"] is not None
        band = f"{_num2(s['band_low'])} – {_num2(s['band_high'])}" if has_stats else "&mdash;"
        body.append(
            "<tr>"
            f'<td class="txt">{_esc(s["algo"])}</td>'
            f"<td>{s['users']}</td>"
            f"<td>{_num2(s['avg_lot_pct']) if has_stats else '&mdash;'}</td>"
            f"<td>{_num2(s['std_dev']) if has_stats else '&mdash;'}</td>"
            f"<td>{band}</td>"
            f'<td class="below">{s["below"]}</td>'
            f"<td>{s['in_range']}</td>"
            f'<td class="above">{s["above"]}</td>'
            "</tr>"
        )
    return f"""
    <table class="tbl">
      <thead><tr>
        <th class="txt">Algo</th><th>Total Users</th><th>Avg Lot Pct</th><th>Std Dev</th>
        <th>Band</th><th>Below</th><th>In</th><th>Above</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


def _pct_cell(value):
    return _num2(value) if value is not None else "&mdash;"


def _slippage_tables(slippage):
    """Avg slippage per algo (accounts / slippage accounts / avg Realized
    ML % of the slippage accounts, with an overall row) and the major
    slippages — accounts whose Realized ML % is above their algo's average."""
    sum_body = []
    for s in slippage["summary"] + [slippage["overall"]]:
        cls = ' class="grand"' if s["ALGO"] == "Overall" else ""
        sum_body.append(
            f"<tr{cls}>"
            f"<td>{_esc(s['ALGO'])}</td>"
            f"<td>{s['accounts']}</td>"
            f"<td>{s['slipped']}</td>"
            f"<td>{_pct_cell(s['avg_slippage'])}</td>"
            "</tr>"
        )
    summary_tbl = f"""
        <h3>Avg Slippage per Algo</h3>
        <div class="tbl-scroll">
        <table class="tbl">
          <thead><tr>
            <th>Algo</th><th>Accounts</th><th>Slippage Accounts</th><th>Avg Slippage %</th>
          </tr></thead>
          <tbody>{''.join(sum_body)}</tbody>
        </table>
        </div>"""

    if slippage["majors"]:
        maj_body = []
        for r in slippage["majors"]:
            maj_body.append(
                "<tr>"
                f"<td>{_esc(r['ALGO'])}</td>"
                f"<td>{_esc(r['SERVER'])}</td>"
                f"<td>{_esc(r['UserID'])}</td>"
                f"<td>{_money(r['Allocation'])}</td>"
                f"<td>{_money(r['MaxLoss'])}</td>"
                f"{_pnl_td(r['Realized'])}"
                f"<td>{_num2(r['MLPct'])}</td>"
                f'<td class="above">{_num2(r["RealizedMLPct"])}</td>'
                f"<td>{_num2(r['DiffPct'])}</td>"
                f"<td>{_num2(r['AlgoAvgSlippage'])}</td>"
                "</tr>"
            )
        majors_tbl = f"""
        <h3>Major Slippages</h3>
        <div class="tbl-scroll">
        <table class="tbl">
          <thead><tr>
            <th>Algo</th><th>Server</th><th>User ID</th><th>Allocation</th><th>Max Loss</th>
            <th>Realized P&amp;L</th><th>ML %</th><th>Realized ML %</th><th>Diff %</th>
            <th>Algo Avg %</th>
          </tr></thead>
          <tbody>{''.join(maj_body)}</tbody>
        </table>
        </div>"""
    else:
        majors_tbl = ('<h3>Major Slippages</h3>'
                      '<div class="empty-note">No slippage account is above its '
                      'algo&rsquo;s average.</div>')

    return f"""
    <div class="strike-grid">
      <div class="card strike-card">{summary_tbl}</div>
      <div class="card strike-card">{majors_tbl}</div>
    </div>"""


def _outlier_clients_table(outlier_rows):
    """The stricter client-outlier view: band ("Average Lots") and difference
    are whole lots, converted per user from the lot-pct band."""
    if not outlier_rows:
        return '<div class="empty-note">No outlier clients at this deviation.</div>'
    body = []
    for r in outlier_rows:
        band = (f"{int(round(r['band_low_lots'])):,} &ndash; "
                f"{int(round(r['band_high_lots'])):,}")
        cls = "below" if r["outlier"].startswith("Below") else "above"
        body.append(
            "<tr>"
            f"<td>{_esc(r['algo'])}</td>"
            f"<td>{_esc(r['server'])}</td>"
            f"<td>{_esc(r['user_id'])}</td>"
            f"<td>{band}</td>"
            f'<td class="{cls}">{_esc(r["outlier"])}</td>'
            f"<td>{int(round(r['lots'])):,}</td>"
            f"<td>{int(round(r['diff_lots'])):,}</td>"
            "</tr>"
        )
    return f"""
    <table class="tbl">
      <thead><tr>
        <th>Algo</th><th>Server</th><th>User ID</th>
        <th>Average Lots</th><th>Outlier</th><th>Lots Fired</th><th>Diff of Lots</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


def _expiry_label(expiry):
    return expiry.strftime("%d%b%y").upper()


def _strikes_tables(strikes):
    """The two strike summaries side by side, each behind its own expiry +
    index dropdown pair: distinct strikes per (algo, server) — with an "All"
    option that reproduces the unfiltered counts — and the lots-per-strike
    option chain (CE | Strike | PE). Lots are whole numbers."""
    # every (expiry|All, index|All) combination is precomputed, so the table
    # swap is a pure lookup in the page's JS
    seg_set, exp_dates = set(), set()
    for r in strikes["by_algo_server"]:
        for contract in r["contracts"]:
            seg_set.add(contract[0])
            exp_dates.add(contract[1])
    expiry_labels = [_expiry_label(e) for e in sorted(exp_dates)]
    segments = sorted(seg_set)

    def counts_for(exp_label, seg):
        rows, distinct = [], set()
        for r in strikes["by_algo_server"]:
            sel = [c for c in r["contracts"]
                   if (exp_label == "All" or _expiry_label(c[1]) == exp_label)
                   and (seg == "All" or c[0] == seg)]
            if sel:
                rows.append([r["algo"] or "—", r["server"], len(sel)])
                distinct.update(sel)
        return {"rows": rows, "total": len(distinct)}

    as_counts = {e: {s: counts_for(e, s) for s in ["All"] + segments}
                 for e in ["All"] + expiry_labels}

    # server-side render of the default (All / All) view — shown before the
    # script runs and in print
    initial = as_counts["All"]["All"]
    algo_body = [
        "<tr>"
        f'<td class="txt">{_esc(algo)}</td>'
        f'<td class="txt">{_esc(server)}</td>'
        f"<td>{count}</td>"
        "</tr>"
        for algo, server, count in initial["rows"]
    ]
    algo_body.append(
        '<tr class="grand"><td class="txt">Total (distinct)</td><td class="txt"></td>'
        f"<td>{initial['total']}</td></tr>"
    )

    # {expiry label: {index: [[ce, strike, pe], ...]}} — expiry labels in date order
    chains, chain_expiries = {}, []
    for (segment, expiry), rows in strike_chain(strikes["per_strike"]).items():
        label = _expiry_label(expiry)
        if label not in chains:
            chains[label] = {}
            chain_expiries.append((expiry, label))
        chains[label][segment] = [[int(round(ce)), strike, int(round(pe))]
                                  for ce, strike, pe in rows]
    chain_expiries = [label for _, label in sorted(chain_expiries)]

    return f"""
    <div class="strike-grid">
      <div class="card strike-card">
        <h3>Strikes per Algo / Server</h3>
        <div class="chain-controls">
          <label>Expiry <select id="as-expiry"></select></label>
          <label>Index <select id="as-index"></select></label>
        </div>
        <div class="tbl-scroll">
        <table class="tbl">
          <thead><tr><th class="txt">Algo</th><th class="txt">Server</th><th>Strikes Traded</th></tr></thead>
          <tbody id="as-body">{''.join(algo_body)}</tbody>
        </table>
        </div>
      </div>
      <div class="card strike-card">
        <h3>Lots per Strike &mdash; Option Chain</h3>
        <div class="chain-controls">
          <label>Expiry <select id="chain-expiry"></select></label>
          <label>Index <select id="chain-index"></select></label>
        </div>
        <div class="tbl-scroll">
        <table class="tbl">
          <thead><tr><th>CE</th><th>Strike</th><th>PE</th></tr></thead>
          <tbody id="chain-body"></tbody>
        </table>
        </div>
      </div>
    </div>
    <script>
    var CHAINS = {json.dumps(chains)};
    var CHAIN_EXPIRIES = {json.dumps(chain_expiries)};
    var AS_COUNTS = {json.dumps(as_counts)};
    </script>"""


_CHAIN_SCRIPT = """
<script>
(function () {
  var expirySel = document.getElementById('chain-expiry');
  var indexSel = document.getElementById('chain-index');
  var body = document.getElementById('chain-body');
  if (!expirySel || !indexSel || !body) return;
  function fmt(n) { return n.toLocaleString('en-US'); }
  function fillOptions(sel, values) {
    sel.innerHTML = '';
    values.forEach(function (v) {
      var opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      sel.appendChild(opt);
    });
  }
  function render() {
    var rows = (CHAINS[expirySel.value] || {})[indexSel.value] || [];
    var ceTotal = 0, peTotal = 0, out = '';
    rows.forEach(function (r) {
      ceTotal += r[0]; peTotal += r[2];
      out += '<tr><td>' + fmt(r[0]) + '</td><td>' + r[1] + '</td><td>' + fmt(r[2]) + '</td></tr>';
    });
    out += '<tr class="grand"><td>' + fmt(ceTotal) + '</td><td>Total</td><td>' + fmt(peTotal) + '</td></tr>';
    body.innerHTML = out;
  }
  function onExpiry() {
    fillOptions(indexSel, Object.keys(CHAINS[expirySel.value] || {}).sort());
    render();
  }
  fillOptions(expirySel, CHAIN_EXPIRIES);
  expirySel.addEventListener('change', onExpiry);
  indexSel.addEventListener('change', render);
  onExpiry();
})();
(function () {
  var expirySel = document.getElementById('as-expiry');
  var indexSel = document.getElementById('as-index');
  var body = document.getElementById('as-body');
  if (!expirySel || !indexSel || !body) return;
  function esc(v) {
    var d = document.createElement('div');
    d.textContent = String(v);
    return d.innerHTML;
  }
  function fillOptions(sel, values) {
    sel.innerHTML = '';
    values.forEach(function (v) {
      var opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      sel.appendChild(opt);
    });
  }
  function render() {
    var data = (AS_COUNTS[expirySel.value] || {})[indexSel.value] || {rows: [], total: 0};
    var out = '';
    data.rows.forEach(function (r) {
      out += '<tr><td class="txt">' + esc(r[0]) + '</td><td class="txt">' + esc(r[1]) +
             '</td><td>' + r[2] + '</td></tr>';
    });
    out += '<tr class="grand"><td class="txt">Total (distinct)</td><td class="txt"></td><td>' +
           data.total + '</td></tr>';
    body.innerHTML = out;
  }
  fillOptions(expirySel, Object.keys(AS_COUNTS));
  fillOptions(indexSel, Object.keys(AS_COUNTS['All']));
  expirySel.addEventListener('change', render);
  indexSel.addEventListener('change', render);
})();
</script>
"""


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


def _boxplot_svg(boxplot_rows):
    """Inline-SVG box plot of per-USER lot pct per algo (one point per user)
    — self-contained (no JS), same design as the dashboard chart: gray
    box/whiskers for the distribution, flagged outliers overlaid as blue
    (below) / orange (above) points with native <title> hover tooltips."""
    by_algo = {}
    for r in boxplot_rows:
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


def _pnl_td(value):
    """Table cell for a P&L amount — profit green, loss red."""
    cls = _sign_cls(value)
    return f'<td class="{cls}">{_money(value)}</td>' if cls else f"<td>{_money(value)}</td>"


def _metric_cells(r):
    return (
        f"<td>{r['Users']}</td>"
        f"<td>{r['SLHit']}</td>"
        f"<td>{_money(r['MaxLoss'])}</td>"
        f"<td>{_money(r['Allocation'])}</td>"
        f"{_pnl_td(r['Realized'])}"
        f"{_pnl_td(r['Unrealized'])}"
        f"{_pnl_td(r['MTM'])}"
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
                   boxplot_rows=None, strikes=None, outliers=None, outlier_deviation=None,
                   slippage=None, excel_bytes=None, excel_filename=None):
    """Assemble the complete DOR.html document and return it as a string.
    `boxplot_rows` (per-user observations from tradevalue.
    user_lot_observations) feeds the box plot; `strikes` (from
    tradevalue.strike_report) feeds the strikes section; `outliers` (from
    tradevalue.outlier_clients, flagged at `outlier_deviation` std devs)
    feeds the outlier clients section; `slippage` ({"summary", "overall",
    "majors"} from segregate_int_pos_mtm2) feeds the slippage section;
    `excel_bytes`/`excel_filename` embed the full workbook as a download
    button in the masthead. Each section is skipped when its data is
    omitted."""
    generated = datetime.now().strftime("%d-%b-%Y %H:%M")
    grand = next((r for r in pivot_rows if r["kind"] == "grandtotal"), None)

    tv_kpis = _kpi_tiles([
        ("Users", f"{tv_totals['users']:,}"),
        ("Orders", f"{tv_totals['orders']:,}"),
        ("Total Lots", _money(tv_totals["lots"])),
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
            ("Realized P&L", _money(grand["Realized"]), _sign_cls(grand["Realized"])),
            ("Unrealized P&L", _money(grand["Unrealized"]), _sign_cls(grand["Unrealized"])),
            ("MTM", _money(grand["MTM"]), _sign_cls(grand["MTM"])),
            ("Return %", _num2(grand["Return"])),
        ]
    seg_kpis = _kpi_tiles(seg_pairs)

    download_btn = ""
    if excel_bytes:
        b64 = base64.b64encode(excel_bytes).decode("ascii")
        fname = excel_filename or "DOR.xlsx"
        download_btn = (
            f'<a class="dl-btn" download="{_esc(fname)}" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;'
            f'base64,{b64}">&#11015; Download Excel &mdash; full data</a>'
        )

    boxplot = _boxplot_svg(boxplot_rows) if boxplot_rows else ""
    boxplot_section = f"""
<section>
  <h2>Lot Pct by Algo (Box Plot)</h2>
  <div class="section-note">Distribution of each algo&rsquo;s per-user lot pct (one point
  per user); the &plusmn;{deviation:g}&sigma; flagged outliers are shown as points.</div>
  <div class="card">
    {boxplot}
  </div>
</section>""" if boxplot else ""

    if slippage is None:
        slippage_section = ""
    else:
        slippage_note = """
  <div class="section-note">ML % = MAX LOSS / ALLOCATION; Realized ML % =
  |Realized P&amp;L| / ALLOCATION &mdash; plain ratios of allocation, same convention as
  Return %. An account has slippage only when Realized ML % exceeds ML % by at least 0.1
  (1.00 &rarr; 1.09 is not slippage; 1.10 is). Avg Slippage = average Realized ML % of the
  algo&rsquo;s slippage accounts; Major = slippage accounts above their algo&rsquo;s
  average.</div>"""
        slippage_body = (_slippage_tables(slippage) if slippage["overall"]["slipped"]
                         else '<div class="card"><div class="empty-note">'
                              '-- No slippage today --</div></div>')
        slippage_section = f"""
<section>
  <h2>Slippage</h2>
  {slippage_note}
  {slippage_body}
</section>"""

    k2 = outlier_deviation if outlier_deviation is not None else deviation
    outliers_section = f"""
<section>
  <h2>Outlier Clients (&plusmn;{k2:g}&sigma;)</h2>
  <div class="section-note">One row per user: those whose combined lot pct sits more than
  {k2:g} standard deviation(s) from their algo&rsquo;s average. <b>Average Lots</b> is that
  band converted into lots for the user&rsquo;s own allocation; <b>Diff of Lots</b> is how
  many lots beyond the nearest band edge were actually fired.</div>
  <div class="card scroll">
    {_outlier_clients_table(outliers)}
  </div>
</section>""" if outliers is not None else ""

    strikes_section = f"""
<section>
  <h2>Strikes Traded</h2>
  <div class="section-note">A strike is one contract (index + expiry + strike + CE/PE);
  the two exchange symbol formats for the same contract are counted once. Pick an expiry
  and index to see its option chain &mdash; CE / PE lots per strike (whole lots, same
  date-based lot sizes as the trade value rows). Users with no MTM entry inherit their
  server&rsquo;s algo; an algo shows blank only when the whole server has no MTM
  accounts.</div>
  {_strikes_tables(strikes)}
</section>
{_CHAIN_SCRIPT}""" if strikes and strikes.get("per_strike") else ""

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
  /* every table cell — data AND header — is centered (client-facing report) */
  .tbl th {{
    background: var(--header); color: #fff; padding: 8px 10px; text-align: center;
    font-weight: 600; white-space: nowrap; position: sticky; top: 0;
  }}
  .tbl td {{
    padding: 6px 10px; border-bottom: 1px solid var(--line); text-align: center;
    font-variant-numeric: tabular-nums; white-space: nowrap;
  }}
  .tbl tbody tr:hover td {{ background: #F8FAFC; }}
  .tbl tr.subtotal td {{ background: var(--subtotal); font-weight: 600; }}
  .tbl tr.algototal td {{ background: var(--algototal); font-weight: 700; }}
  .tbl tr.grand td {{ background: var(--grand); font-weight: 700; }}
  .tbl td.below {{ color: var(--below); font-weight: 600; }}
  .tbl td.above {{ color: #B45309; font-weight: 600; }}
  .tbl td.pos {{ color: #15803D; font-weight: 600; }}
  .tbl td.neg {{ color: #DC2626; font-weight: 600; }}
  .kpi-value.pos {{ color: #15803D; }}
  .kpi-value.neg {{ color: #DC2626; }}
  .masthead .dl-btn {{
    display: inline-block; margin-top: 12px; padding: 7px 16px; border-radius: 8px;
    background: #334155; border: 1px solid #475569; color: #fff; text-decoration: none;
    font-size: 12.5px; font-weight: 600;
  }}
  .masthead .dl-btn:hover {{ background: #475569; }}
  .scroll {{ overflow-x: auto; }}
  .strike-grid {{ display: grid; grid-template-columns: 1fr 1.4fr; gap: 14px;
                  align-items: start; }}
  @media (max-width: 820px) {{
    .strike-grid {{ grid-template-columns: 1fr; }}
  }}
  /* title + dropdowns stay fixed; only the table body scrolls, and the sticky
     header sits flush at the scrollport top so no row peeks above it */
  .strike-card .tbl-scroll {{ max-height: 420px; overflow: auto; }}
  .strike-card h3 {{ font-size: 13.5px; font-weight: 700; color: var(--header);
                     margin-bottom: 8px; }}
  .chain-controls {{ display: flex; gap: 16px; margin-bottom: 10px; flex-wrap: wrap;
                     color: var(--muted); font-size: 12.5px; }}
  .chain-controls select {{
    margin-left: 6px; padding: 3px 8px; border: 1px solid var(--line); border-radius: 6px;
    background: var(--surface); color: var(--ink); font-size: 12.5px;
  }}
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
  .pivot-tools {{ display: flex; gap: 8px; align-items: center; margin-bottom: 10px; }}
  .pivot-tools button {{
    border: 1px solid var(--line); background: var(--surface); border-radius: 6px;
    padding: 4px 12px; font-size: 12px; cursor: pointer; color: var(--ink);
  }}
  .pivot-tools button:hover {{ background: #F8FAFC; }}
  .pivot-tools .hint {{ color: var(--muted); font-size: 12px; margin-left: auto; }}
  .footnote {{ color: var(--muted); font-size: 12px; margin-top: 10px; }}
  .empty-note {{ color: var(--muted); font-style: italic; text-align: center; padding: 14px; }}
  footer {{
    max-width: 1100px; margin: 0 auto 30px; padding: 0 20px;
    color: var(--muted); font-size: 12px; border-top: 1px solid var(--line); padding-top: 12px;
  }}
  @media print {{
    body {{ background: #fff; }}
    .card, .kpi {{ box-shadow: none; }}
    .pivot-tools, .masthead .dl-btn {{ display: none; }}
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
    {download_btn}
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

{slippage_section}

<section>
  <h2>Trade Value</h2>
  <div class="section-note">
    Completed NIFTY / BANKNIFTY / SENSEX orders only (F&amp;O exchanges NFO / BFO), duplicates
    removed. Outliers are users whose
    lot pct (combined lots across their segments &divide; allocation &times; 100) is more than
    {deviation:g} standard deviation(s) from their algo&rsquo;s average. All columns count
    users, so Below + In + Above = Total Users (users with no usable allocation carry
    no flag).
  </div>
  {tv_kpis}
  <div class="card scroll">
    {_tv_summary_table(tv_summary)}
  </div>
</section>

{boxplot_section}

{outliers_section}

{strikes_section}

</main>
<footer>
  Generated {_esc(generated)} &middot; Daily Operations Report &middot; Full row-level data is
  available in the accompanying Excel workbook.
</footer>
{_PIVOT_SCRIPT}
</body>
</html>"""
