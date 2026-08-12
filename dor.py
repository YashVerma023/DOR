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
from datetime import datetime

from portfolio import DEFAULT_PORTFOLIO_PATTERN, portfolio_json, portfolio_report
from tradevalue import format_crore, format_indian, order_summary_totals


def _money(v):
    """A magnitude in a table: always a WHOLE number, Indian-grouped. Every
    count, quantity, lot, allocation and value uses this — decimals in a
    table are noise. The two exceptions are format_crore() (the compact
    "5.2 Cr" / "52.36 L" form, whose decimals are the precision) and _num2()
    below."""
    return format_indian(v)


def _num2(v):
    """A PERCENTAGE — the only plain number that keeps 2 decimals, because
    the value is a small ratio whose decimals carry the meaning (and the
    slippage rule is defined at 0.1 granularity: 1.09 is not slippage,
    1.10 is). Never use this for a magnitude; use _money()."""
    return format_indian(v, places=2)


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


def _tv_summary_table(tv_summary, user_obs=None, table_id="tv-summary"):
    """The Algo Summary as a drill-down pivot — one row per (algo, Int /
    Pos+Int) group; each row expands into its OUTLIER users — user id,
    server, lots per Cr, lots, and whether they sit below or above the band
    (below users first, worst first). `table_id` keeps ids unique when the
    report carries one table per User MTM file."""
    outliers_by_group = {}
    for o in (user_obs or []):
        if o["outlier"] in ("Below average range", "Above average range"):
            outliers_by_group.setdefault(
                (o["trade_date"], o["algo"], o["user_type"]), []).append(o)

    body = []
    for i, s in enumerate(tv_summary):
        has_stats = s["median_lots_per_cr"] is not None
        # below users first (farthest below leading), then above (farthest
        # above leading) — worst first within each side
        flagged = sorted(
            outliers_by_group.get((s["trade_date"], s["algo"], s["user_type"]), []),
            key=lambda o: ((0, o["lots_per_cr"])
                           if o["outlier"].startswith("Below")
                           else (1, -o["lots_per_cr"])))
        # every row is drillable, even with no outliers — the drill then
        # shows an explicit "none" row, so the interaction stays uniform
        key = f"{table_id}-{i}"
        body.append(
            f'<tr class="algototal lvl1" data-key="{key}">'
            f'<td class="txt"><span class="caret">&#9656;</span>{_esc(s["algo"])}</td>'
            f"<td>{_esc(s['user_type']) or '&mdash;'}</td>"
            f"<td>{s['users']}</td>"
            f"<td>{_money(s['median_lots_per_cr']) if has_stats else '&mdash;'}</td>"
            f'<td class="below">{s["below"]}</td>'
            f"<td>{s['in_range']}</td>"
            f'<td class="above">{s["above"]}</td>'
            "</tr>"
        )
        for o in flagged:
            cls = "below" if o["outlier"].startswith("Below") else "above"
            # a user trading several indexes is judged here on this index's
            # lots alone over their WHOLE allocation — flag that partial
            # exposure so a "Below" verdict is not read as under-trading
            others = o.get("also_trades") or []
            note = (f' <span class="note">&#9888; also trades '
                    f'{_esc(", ".join(others))}</span>') if others else ""
            body.append(
                f'<tr class="lvl3 hidden" data-parent="{key}">'
                f'<td class="txt"></td><td></td>'
                f"<td>{_esc(o['user_id'])}{note}</td>"
                f"<td>{_esc(o['server'])}</td>"
                f"<td>{_money(o['lots'])} lots</td>"
                f'<td colspan="2" class="{cls}">{_esc(o["outlier"])}</td>'
                "</tr>"
            )
        if not flagged:
            body.append(
                f'<tr class="lvl3 hidden" data-parent="{key}">'
                f'<td class="txt"></td>'
                f'<td colspan="6" class="note">No outlier users in this group.</td>'
                "</tr>"
            )
    return f"""
    <table class="tbl drill-tbl" id="{table_id}">
      <thead><tr>
        <th class="txt">Algo</th><th>Type</th><th>Total Users</th><th>Median Lots per Cr</th>
        <th>Below</th><th>In</th><th>Above</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


def _orders_table(order_rows, table_id="orders-summary"):
    """The Orders Summary as a drill-down pivot — one row per (algo, Int /
    Pos+Int) group; each row expands into its SERVERS, same column shape one
    level down (users / total / executed / failed / hedge / VAR)."""
    body = []
    for i, r in enumerate(order_rows):
        key = f"{table_id}-{i}"
        body.append(
            f'<tr class="algototal lvl1" data-key="{key}">'
            f'<td class="txt"><span class="caret">&#9656;</span>{_esc(r["algo"]) or "&mdash;"}</td>'
            f"<td>{_esc(r['user_type']) or '&mdash;'}</td>"
            f"<td>{_money(r['users'])}</td>"
            f"<td>{_money(r['orders'])}</td>"
            f"<td>{_money(r['executed'])}</td>"
            f'<td class="above">{_money(r["failed"])}</td>'
            f"<td>{_money(r.get('pending', 0))}</td>"
            f"<td>{_money(r['hedge'])}</td>"
            f"<td>{_money(r['var'])}</td>"
            "</tr>"
        )
        for s in r["servers"]:
            body.append(
                f'<tr class="lvl3 hidden" data-parent="{key}">'
                f'<td class="txt"></td>'
                f'<td class="txt">{_esc(s["server"]) or "&mdash;"}</td>'
                f"<td>{_money(s['users'])}</td>"
                f"<td>{_money(s['orders'])}</td>"
                f"<td>{_money(s['executed'])}</td>"
                f'<td class="above">{_money(s["failed"])}</td>'
                f"<td>{_money(s.get('pending', 0))}</td>"
                f"<td>{_money(s['hedge'])}</td>"
                f"<td>{_money(s['var'])}</td>"
                "</tr>"
            )
    totals = order_summary_totals(order_rows)
    body.append(
        '<tr class="grand"><td class="txt">Total</td><td class="txt"></td>'
        f"<td>{_money(totals['users'])}</td>"
        f"<td>{_money(totals['orders'])}</td>"
        f"<td>{_money(totals['executed'])}</td>"
        f"<td>{_money(totals['failed'])}</td>"
        f"<td>{_money(totals.get('pending', 0))}</td>"
        f"<td>{_money(totals['hedge'])}</td>"
        f"<td>{_money(totals['var'])}</td></tr>"
    )
    return f"""
    <table class="tbl drill-tbl" id="{table_id}">
      <thead><tr>
        <th class="txt">Algo</th><th>Type</th><th>Total Users</th><th>Total Orders</th>
        <th>Executed</th><th>Failed/Cancelled/Rejected</th><th>Pending</th>
        <th>Hedge</th><th>VAR</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


def _pct_cell(value):
    return _num2(value) if value is not None else "&mdash;"


def _slippage_tables(slippage):
    """Avg slippage per algo (accounts / slippage accounts / avg Realized
    ML % of the slippage accounts, with an overall row) and the major
    slippages — accounts whose Realized ML % is above their algo's average.
    The algo selection is made in the dashboard inputs; the report shows the
    selected algos only."""
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
                f"<td>{_esc(format_crore(r['MaxLoss']))}</td>"
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


def _portfolio_table(report, table_id):
    """One portfolio analysis as a three-level drill-down table:
    algo rows (server COUNT in the Server column) expand into server rows,
    which expand into per-user rows; a Total row closes the table. Same
    markup pattern as the segregation pivot (data-key / data-parent)."""
    body = []
    for i, algo_row in enumerate(report["algos"]):
        akey = f"{table_id}a{i}"
        body.append(
            f'<tr class="algototal lvl1" data-key="{akey}">'
            f'<td class="txt"><span class="caret">&#9656;</span>{_esc(algo_row["algo"])}</td>'
            f"<td>{algo_row['n_servers']}</td>"
            f"<td>{algo_row['users']}</td>"
            f"<td>{algo_row['portfolios']}</td>"
            f"<td>{_money(algo_row['orders'])}</td>"
            f"{_pnl_td(algo_row['pnl'])}</tr>"
        )
        for j, server_row in enumerate(algo_row["server_rows"]):
            skey = f"{akey}s{j}"
            body.append(
                f'<tr class="subtotal lvl2 hidden" data-parent="{akey}" data-key="{skey}">'
                f'<td class="txt"></td>'
                f'<td><span class="caret">&#9656;</span>{_esc(server_row["server"])}</td>'
                f"<td>{server_row['users']}</td>"
                f"<td>{server_row['portfolios']}</td>"
                f"<td>{_money(server_row['orders'])}</td>"
                f"{_pnl_td(server_row['pnl'])}</tr>"
            )
            for u in server_row["user_rows"]:
                body.append(
                    f'<tr class="lvl3 hidden" data-parent="{skey}">'
                    f'<td class="txt"></td><td></td>'
                    f'<td>{_esc(u["user_id"])}</td>'
                    f"<td>{u['portfolios']}</td>"
                    f"<td>{_money(u['orders'])}</td>"
                    f"{_pnl_td(u['pnl'])}</tr>"
                )
    total = report["total"]
    body.append(
        '<tr class="grand"><td class="txt">Total</td><td></td>'
        f"<td>{total['users']}</td><td>{total['portfolios']}</td>"
        f"<td>{_money(total['orders'])}</td>{_pnl_td(total['pnl'])}</tr>"
    )
    return f"""
    <table class="tbl drill-tbl" id="{table_id}">
      <thead><tr>
        <th class="txt">Algo</th><th>Server</th>
        <th>Portfolio Executed Users</th><th>No. of Portfolio</th>
        <th>Total Orders</th><th>PnL</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


# click-to-expand for every .drill-tbl table (Algo Summary, strikes,
# portfolio — prerendered or built later by JS); collapsing a row also
# collapses every open level underneath it. Always emitted.
_DRILL_SCRIPT = """
<script>
(function () {
  function collapseTree(table, key) {
    table.querySelectorAll('tr[data-parent="' + key + '"]').forEach(function (child) {
      child.classList.add('hidden');
      if (child.dataset.key) {
        child.classList.remove('open');
        collapseTree(table, child.dataset.key);
      }
    });
  }
  document.addEventListener('click', function (e) {
    var tr = e.target.closest('.drill-tbl tr[data-key]');
    if (!tr || !tr.dataset.key) return;
    var table = tr.closest('table');
    if (tr.classList.contains('open')) {
      tr.classList.remove('open');
      collapseTree(table, tr.dataset.key);
    } else {
      tr.classList.add('open');
      table.querySelectorAll('tr[data-parent="' + tr.dataset.key + '"]')
        .forEach(function (child) { child.classList.remove('hidden'); });
    }
  });
})();
</script>
"""

_PORTFOLIO_SCRIPT = """
<script>
(function () {
  function esc(v) {
    var d = document.createElement('div');
    d.textContent = String(v);
    return d.innerHTML;
  }
  function fmt(n) { return Math.round(n).toLocaleString('en-IN'); }
  function pnlTd(v) {
    var cls = v > 0 ? ' class="pos"' : (v < 0 ? ' class="neg"' : '');
    return '<td' + cls + '>' + fmt(v) + '</td>';
  }
  var input = document.getElementById('pf-pattern');
  var btn = document.getElementById('pf-analyze');
  var out = document.getElementById('pf-result');
  if (!input || !btn || !out || typeof PF_DATA === 'undefined') return;

  // combobox: every portfolio name goes into the datalist — the browser
  // narrows the dropdown to the names containing whatever is typed
  var datalist = document.getElementById('pf-names');
  if (datalist) {
    PF_DATA.portfolios.slice().sort().forEach(function (name) {
      var opt = document.createElement('option');
      opt.value = name;
      datalist.appendChild(opt);
    });
  }

  function analyse() {
    var pat = input.value.trim().toUpperCase();
    if (!pat) { out.innerHTML = ''; return; }
    var matching = {};
    PF_DATA.portfolios.forEach(function (name, i) {
      if (name.toUpperCase().indexOf(pat) !== -1) matching[i] = true;
    });
    // three levels: algo -> server -> user (same shape as the QS table)
    var algos = {}, algoOrder = [];
    var totUsers = {}, totPf = {}, totOrders = 0, totPnl = 0;
    PF_DATA.rows.forEach(function (r) {
      if (!matching[r[3]]) return;
      var algo = PF_DATA.algos[r[0]], server = PF_DATA.servers[r[1]];
      var user = PF_DATA.users[r[2]];
      if (!algos[algo]) { algos[algo] = {servers: {}, users: {}, pf: {}, orders: 0, pnl: 0}; algoOrder.push(algo); }
      var a = algos[algo];
      if (!a.servers[server]) a.servers[server] = {users: {}, pf: {}, orders: 0, pnl: 0};
      var s = a.servers[server];
      if (!s.users[user]) s.users[user] = {pf: {}, orders: 0, pnl: 0};
      var u = s.users[user];
      u.pf[r[3]] = true; u.orders += r[4]; u.pnl += r[5];
      s.pf[r[3]] = true; s.orders += r[4]; s.pnl += r[5];
      a.users[user.toUpperCase()] = true; a.pf[r[3]] = true; a.orders += r[4]; a.pnl += r[5];
      totUsers[user.toUpperCase()] = true; totPf[r[3]] = true; totOrders += r[4]; totPnl += r[5];
    });
    if (!algoOrder.length) {
      out.innerHTML = '<div class="empty-note">No portfolio name contains &ldquo;' +
        esc(input.value.trim()) + '&rdquo;.</div>';
      return;
    }
    algoOrder.sort(function (x, y) {
      var xn = parseInt(x, 10), yn = parseInt(y, 10);
      if (!isNaN(xn) && !isNaN(yn) && xn !== yn) return xn - yn;
      return x < y ? -1 : (x > y ? 1 : 0);
    });
    var html = '<table class="tbl drill-tbl" id="pf-custom"><thead><tr>' +
      '<th class="txt">Algo</th><th>Server</th>' +
      '<th>Portfolio Executed Users</th><th>No. of Portfolio</th>' +
      '<th>Total Orders</th><th>PnL</th></tr></thead><tbody>';
    algoOrder.forEach(function (algo, i) {
      var a = algos[algo];
      var servers = Object.keys(a.servers).sort();
      var akey = 'pfc' + i;
      html += '<tr class="algototal lvl1" data-key="' + akey + '">' +
        '<td class="txt"><span class="caret">&#9656;</span>' + esc(algo) + '</td>' +
        '<td>' + servers.length + '</td>' +
        '<td>' + Object.keys(a.users).length + '</td><td>' + Object.keys(a.pf).length + '</td>' +
        '<td>' + fmt(a.orders) + '</td>' + pnlTd(a.pnl) + '</tr>';
      servers.forEach(function (server, j) {
        var s = a.servers[server];
        var skey = akey + 's' + j;
        var userNames = Object.keys(s.users);
        html += '<tr class="subtotal lvl2 hidden" data-parent="' + akey + '" data-key="' + skey + '">' +
          '<td class="txt"></td>' +
          '<td><span class="caret">&#9656;</span>' + esc(server) + '</td>' +
          '<td>' + userNames.length + '</td><td>' + Object.keys(s.pf).length + '</td>' +
          '<td>' + fmt(s.orders) + '</td>' + pnlTd(s.pnl) + '</tr>';
        userNames.sort(function (x, y) { return s.users[x].pnl - s.users[y].pnl; });
        userNames.forEach(function (u) {
          var d = s.users[u];
          html += '<tr class="lvl3 hidden" data-parent="' + skey + '">' +
            '<td class="txt"></td><td></td><td>' + esc(u) + '</td>' +
            '<td>' + Object.keys(d.pf).length + '</td><td>' + fmt(d.orders) + '</td>' +
            pnlTd(d.pnl) + '</tr>';
        });
      });
    });
    html += '<tr class="grand"><td class="txt">Total</td><td></td>' +
      '<td>' + Object.keys(totUsers).length + '</td><td>' + Object.keys(totPf).length + '</td>' +
      '<td>' + fmt(totOrders) + '</td>' + pnlTd(totPnl) + '</tr></tbody></table>';
    out.innerHTML = '<h3 class="pf-title">' + esc(input.value.trim()) +
      ' Portfolio Analysis</h3><div class="card scroll">' + html + '</div>';
  }
  btn.addEventListener('click', analyse);
  input.addEventListener('keydown', function (e) { if (e.key === 'Enter') analyse(); });
  input.addEventListener('change', analyse);  // fires when a dropdown name is picked
})();
</script>
"""


def _strikes_tables(strikes, mids=None, expiry_map=None):
    """The two strike summaries side by side, each behind its own expiry +
    index dropdown pair: distinct strikes per (algo, server) — with an "All"
    option that reproduces the unfiltered counts — and the lots-per-strike
    option chain (CE | Strike | PE). Lots are whole numbers.

    A contract is (index, strike, CE/PE) — the expiry is not parsed from the
    symbols. `expiry_map` ({index: expiry text entered in the dashboard})
    labels each index's chain; an index with no entry shows as "NA". `mids`
    maps index -> day-mid price ((day open + day close) / 2, entered in the
    dashboard). When the selected index has a mid, the chain centres on the
    ATM (the traded strike nearest the mid, highlighted) and shows only N
    strikes above and below it — N is the "No. of strikes" input in the
    page, default 10."""
    # every (expiry|All, index|All) combination is precomputed, so the table
    # swap is a pure lookup in the page's JS
    seg_set = {r["segment"] for r in strikes["per_strike"]}
    for r in strikes["by_algo_server"]:
        for contract in r["contracts"]:
            seg_set.add(contract[0])
    label_of = {s: (expiry_map or {}).get(s) or "NA" for s in seg_set}
    expiry_labels = sorted(set(label_of.values()))
    segments = sorted(seg_set)

    def counts_for(exp_label, seg):
        # two levels: per algo (distinct across its servers) with the server
        # rows underneath — [[algo, nServers, algoDistinct, [[server, n]...]]]
        per_algo, distinct = {}, set()
        for r in strikes["by_algo_server"]:
            sel = [c for c in r["contracts"]
                   if (exp_label == "All" or label_of[c[0]] == exp_label)
                   and (seg == "All" or c[0] == seg)]
            if sel:
                bucket = per_algo.setdefault(r["algo"] or "—", {"servers": [], "contracts": set()})
                bucket["servers"].append([r["server"], len(sel)])
                bucket["contracts"].update(sel)
                distinct.update(sel)
        algos = [[algo, len(b["servers"]), len(b["contracts"]), sorted(b["servers"])]
                 for algo, b in per_algo.items()]
        return {"algos": algos, "total": len(distinct)}

    as_counts = {e: {s: counts_for(e, s) for s in ["All"] + segments}
                 for e in ["All"] + expiry_labels}

    # server-side render of the default (All / All) view — shown before the
    # script runs and in print (algo rows expand into their servers)
    initial = as_counts["All"]["All"]
    algo_body = []
    for i, (algo, n_servers, algo_distinct, servers) in enumerate(initial["algos"]):
        key = f"as-a{i}"
        algo_body.append(
            f'<tr class="algototal lvl1" data-key="{key}">'
            f'<td class="txt"><span class="caret">&#9656;</span>{_esc(algo)}</td>'
            f"<td>{n_servers}</td>"
            f"<td>{algo_distinct}</td>"
            "</tr>"
        )
        for server, count in servers:
            algo_body.append(
                f'<tr class="lvl3 hidden" data-parent="{key}">'
                f'<td class="txt"></td>'
                f'<td>{_esc(server)}</td>'
                f"<td>{count}</td>"
                "</tr>"
            )
    algo_body.append(
        '<tr class="grand"><td class="txt">Total (distinct)</td><td class="txt"></td>'
        f"<td>{initial['total']}</td></tr>"
    )

    # {expiry label: {index: {strike: [[algo, catIdx, ce, pe], ...]}}} — the
    # chain's raw material, split by algo and trade kind so the hedge / VAR /
    # sq-off toggles and the algo dropdown can slice it client-side
    # (catIdx: 0=normal, 1=hedge, 2=var, 3=sqoff)
    cat_index = {"normal": 0, "hedge": 1, "var": 2, "sqoff": 3}
    chains, chain_algos = {}, set()
    for r in strikes["per_strike"]:
        label = label_of[r["segment"]]
        cell = (chains.setdefault(label, {}).setdefault(r["segment"], {})
                .setdefault(str(r["strike"]), {}))
        for (algo, cat), lots in r.get("breakdown", {}).items():
            algo_label = str(algo) if str(algo) else "—"
            chain_algos.add(algo_label)
            entry = cell.setdefault(f"{algo_label}|{cat_index[cat]}", [0.0, 0.0])
            entry[0 if r["opt_type"] == "CE" else 1] += round(float(lots), 2)
    chain_data = {
        exp: {seg: {strike: [[k.split("|")[0], int(k.split("|")[1]), v[0], v[1]]
                             for k, v in cell.items()]
                    for strike, cell in seg_d.items()}
              for seg, seg_d in exp_d.items()}
        for exp, exp_d in chains.items()
    }
    chain_expiries = sorted(chains)
    chain_algos = sorted(chain_algos, key=lambda a: (0, int(a)) if a.isdigit() else (1, a))

    return f"""
    <div class="strike-grid">
      <div class="card strike-card">
        <h3>Strikes per Algo / Server</h3>
        <div class="chain-controls">
          <label>Expiry <select id="as-expiry"></select></label>
          <label>Index <select id="as-index"></select></label>
        </div>
        <div class="tbl-scroll">
        <table class="tbl drill-tbl">
          <thead><tr><th class="txt">Algo</th><th>Server</th><th>Strikes Traded</th></tr></thead>
          <tbody id="as-body">{''.join(algo_body)}</tbody>
        </table>
        </div>
      </div>
      <div class="card strike-card">
        <h3>Lots per Strike &mdash; Option Chain</h3>
        <div class="chain-controls">
          <label>Expiry <select id="chain-expiry"></select></label>
          <label>Index <select id="chain-index"></select></label>
          <label>Algo <select id="chain-algo"></select></label>
          <label>Strikes <input id="chain-n" type="number" value="10" min="1"></label>
          <span class="chip-group">
            <label class="chip"><input type="checkbox" id="chain-hedge" checked>Hedge</label>
            <label class="chip"><input type="checkbox" id="chain-var" checked>VAR</label>
            <label class="chip"><input type="checkbox" id="chain-sqoff" checked>Sq-off</label>
          </span>
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
    var CHAIN_D = {json.dumps(chain_data)};
    var CHAIN_EXPIRIES = {json.dumps(chain_expiries)};
    var CHAIN_ALGOS = {json.dumps(chain_algos)};
    var CHAIN_MIDS = {json.dumps(mids or {})};
    var AS_COUNTS = {json.dumps(as_counts)};
    var AS_EXPIRIES = {json.dumps(["All"] + expiry_labels)};
    var AS_INDEXES = {json.dumps(["All"] + segments)};
    </script>"""


_CHAIN_SCRIPT = """
<script>
(function () {
  var expirySel = document.getElementById('chain-expiry');
  var indexSel = document.getElementById('chain-index');
  var body = document.getElementById('chain-body');
  if (!expirySel || !indexSel || !body) return;
  function fmt(n) { return n.toLocaleString('en-IN'); }
  function fillOptions(sel, values) {
    sel.innerHTML = '';
    values.forEach(function (v) {
      var opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      sel.appendChild(opt);
    });
  }
  var nInput = document.getElementById('chain-n');
  var algoSel = document.getElementById('chain-algo');
  var cbHedge = document.getElementById('chain-hedge');
  var cbVar = document.getElementById('chain-var');
  var cbSqoff = document.getElementById('chain-sqoff');
  function render() {
    // aggregate the (algo, trade kind) breakdown under the active filters —
    // normal trades always count; hedge / VAR / sq-off follow their toggles
    var segData = (CHAIN_D[expirySel.value] || {})[indexSel.value] || {};
    var catOn = [true, cbHedge.checked, cbVar.checked, cbSqoff.checked];
    var algoPick = algoSel.value;
    var rows = [];
    Object.keys(segData).forEach(function (strike) {
      var ce = 0, pe = 0;
      segData[strike].forEach(function (e) {
        if (algoPick !== 'All' && e[0] !== algoPick) return;
        if (!catOn[e[1]]) return;
        ce += e[2]; pe += e[3];
      });
      if (ce || pe) rows.push([Math.round(ce), parseInt(strike, 10), Math.round(pe)]);
    });
    rows.sort(function (a, b) { return a[1] - b[1]; });
    // centre on the ATM (traded strike nearest the day-mid) when the index
    // has a mid entered — show N strikes above and below it
    var mid = CHAIN_MIDS[indexSel.value];
    var atmIdx = -1;
    if (mid && rows.length) {
      var n = nInput ? parseInt(nInput.value, 10) : 10;
      if (!n || n < 1) n = 10;
      atmIdx = 0;
      rows.forEach(function (r, i) {
        if (Math.abs(r[1] - mid) < Math.abs(rows[atmIdx][1] - mid)) atmIdx = i;
      });
      var lo = Math.max(0, atmIdx - n);
      var atmStrike = rows[atmIdx][1];
      rows = rows.slice(lo, atmIdx + n + 1);
      atmIdx = rows.findIndex(function (r) { return r[1] === atmStrike; });
    }
    var ceTotal = 0, peTotal = 0, out = '';
    rows.forEach(function (r, i) {
      ceTotal += r[0]; peTotal += r[2];
      out += '<tr' + (i === atmIdx ? ' class="atm"' : '') + '><td>' + fmt(r[0]) +
             '</td><td>' + r[1] + '</td><td>' + fmt(r[2]) + '</td></tr>';
    });
    out += '<tr class="grand"><td>' + fmt(ceTotal) + '</td><td>Total</td><td>' + fmt(peTotal) + '</td></tr>';
    body.innerHTML = out;
  }
  function onExpiry() {
    fillOptions(indexSel, Object.keys(CHAIN_D[expirySel.value] || {}).sort());
    render();
  }
  fillOptions(expirySel, CHAIN_EXPIRIES);
  fillOptions(algoSel, ['All'].concat(CHAIN_ALGOS));
  expirySel.addEventListener('change', onExpiry);
  indexSel.addEventListener('change', render);
  algoSel.addEventListener('change', render);
  [cbHedge, cbVar, cbSqoff].forEach(function (cb) { cb.addEventListener('change', render); });
  if (nInput) nInput.addEventListener('input', render);
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
    var data = (AS_COUNTS[expirySel.value] || {})[indexSel.value] || {algos: [], total: 0};
    var out = '';
    data.algos.forEach(function (a, i) {
      var key = 'as-a' + i;
      out += '<tr class="algototal lvl1" data-key="' + key + '">' +
             '<td class="txt"><span class="caret">&#9656;</span>' + esc(a[0]) + '</td>' +
             '<td>' + a[1] + '</td><td>' + a[2] + '</td></tr>';
      a[3].forEach(function (s) {
        out += '<tr class="lvl3 hidden" data-parent="' + key + '">' +
               '<td class="txt"></td><td>' + esc(s[0]) + '</td><td>' + s[1] + '</td></tr>';
      });
    });
    out += '<tr class="grand"><td class="txt">Total (distinct)</td><td class="txt"></td><td>' +
           data.total + '</td></tr>';
    body.innerHTML = out;
  }
  // explicit arrays, not Object.keys(): an all-digit expiry label would be
  // an integer-like key and hoist itself ahead of "All"
  fillOptions(expirySel, AS_EXPIRIES);
  fillOptions(indexSel, AS_INDEXES);
  expirySel.addEventListener('change', render);
  indexSel.addEventListener('change', render);
  render();
})();
</script>
"""


def _charts_section(charts):
    """One intraday chart PER INDEX that traded — normally NIFTY and SENSEX,
    three when BANKNIFTY is in the book.

    Each chart is a self-contained card carrying its own payload in a
    `data-chart` attribute, so the single script below initialises any number
    of them without ids having to be unique per section."""
    blocks = [_chart_section(c) for c in (charts or []) if c and c.get("series")]
    if not blocks:
        return ""
    return f"""
<section>
  <h2>Intraday</h2>
  <div class="section-note">One chart per index traded. Lines are plotted on the
  <b>close</b> of each bucket. The index is the cash index, which has no continuous ticks
  during the closing auction, so its line ends at 15:29 while order-driven series run to
  15:40 &mdash; the shaded band marks that auction / extended window rather than hiding the
  gap.
  <br>The lower panel is <b>lots fired</b> per bucket.
  <span class="dot-key" style="background:#15803D"></span><b>Completed</b> is every executed
  lot and equals the Trade Value lots for that index;
  <span class="dot-key" style="background:#0891B2"></span>Stoxxo,
  <span class="dot-key" style="background:#EA580C"></span>Hedge and
  <span class="dot-key" style="background:#9333EA"></span>VAR are its three parts, so
  <b>Stoxxo + Hedge + VAR = Completed</b> &mdash; do not add the legend up.
  <span class="dot-key" style="background:#DC2626"></span><b>Failed / cancelled</b> counts
  every order that did not execute, whatever its tag. Orders still live at the close belong
  to none of the five, and square-off orders are excluded &mdash; they close a position
  rather than place one.</div>
  {''.join(blocks)}
</section>
{_CHART_SCRIPT}"""


def _chart_section(chart):
    """One index's card: algo + timeframe selectors, the price panel and the
    lots panel. The payload rides on the element, not in a global, so several
    charts coexist."""
    return f"""
  <h3 class="pf-title">{_esc(chart['index'])}</h3>
  <div class="card chart-card">
    <div class="chain-controls">
      <label>Algo <select class="c-algo"><option value="">All algos</option></select></label>
      <label>Timeframe <select class="c-tf">
        <option value="1">1 MIN</option>
        <option value="5" selected>5 MIN</option>
        <option value="15">15 MIN</option>
        <option value="30">30 MIN</option>
        <option value="60">1 Hour</option>
      </select></label>
      <label class="chip"><input type="checkbox" class="c-log" checked>Log lots</label>
      <span class="c-legend chip-group"></span>
    </div>
    <div class="chart-wrap">
      <svg class="c-svg" viewBox="0 0 1080 560" preserveAspectRatio="xMidYMid meet"></svg>
      <div class="chart-tip"></div>
    </div>
    <script type="application/json" class="c-data">{json.dumps(chart)}</script>
  </div>"""


_CHART_SCRIPT = """
<script>
(function () {
  var NS = 'http://www.w3.org/2000/svg';
  var LINE = ['#1E40AF', '#B45309', '#0F766E', '#9333EA'];
  // `complete` is the TOTAL executed; stoxxo/hedge/var are its three parts, so
  // the legend must not be added up. `failed` is every non-executed lot.
  var CAT_COLOR = {complete: '#15803D', stoxxo: '#0891B2',
                   hedge: '#EA580C', var: '#9333EA', failed: '#DC2626'};
  var CAT_LABEL = {complete: 'Completed (total)', stoxxo: 'Stoxxo',
                   hedge: 'Hedge', var: 'VAR', failed: 'Failed / cancelled'};
  var CAT_ORDER = ['complete', 'stoxxo', 'hedge', 'var', 'failed'];
  // two stacked panels sharing one time axis: prices above, lots below
  var W = 1080, H = 560, L = 76, R = 70, T = 20, B = 44;
  var LOTS_H = 150, GAP = 34;
  var P1B = H - B - LOTS_H - GAP;   // price panel bottom
  var P2T = P1B + GAP;              // lots panel top

  function mins(s) { var p = s.split(':'); return +p[0] * 60 + +p[1]; }
  function label(m) {
    var h = Math.floor(m / 60), n = m % 60;
    return (h < 10 ? '0' : '') + h + ':' + (n < 10 ? '0' : '') + n;
  }
  function el(tag, attrs, text) {
    var n = document.createElementNS(NS, tag);
    for (var k in attrs) n.setAttribute(k, attrs[k]);
    if (text !== undefined) n.textContent = text;
    return n;
  }
  function fmt(v) {
    if (Math.abs(v) >= 1000) return v.toLocaleString('en-IN', {maximumFractionDigits: 0});
    return (Math.round(v * 100) / 100).toString();
  }
  function lotFmt(v) {
    return Math.round(v).toLocaleString('en-IN');
  }
  // bucket 1-minute points to `step`, keeping the LAST value = the close
  function resample(points, step) {
    var out = [], last = null, bucket = null;
    points.forEach(function (p) {
      var b = Math.floor(mins(p[0]) / step) * step;
      if (bucket === null) bucket = b;
      if (b !== bucket) { out.push([bucket, last]); bucket = b; }
      last = p[1];
    });
    if (bucket !== null && last !== null) out.push([bucket, last]);
    return out;
  }

  function initChart(card) {
    var holder = card.querySelector('.c-data');
    if (!holder) return;
    var C = JSON.parse(holder.textContent);
    var svg = card.querySelector('.c-svg');
    var tip = card.querySelector('.chart-tip');
    var wrap = card.querySelector('.chart-wrap');
    var tfSel = card.querySelector('.c-tf');
    var algoSel = card.querySelector('.c-algo');
    var logBox = card.querySelector('.c-log');
    var legend = card.querySelector('.c-legend');
    var LOTS = C.lots || null;

    if (LOTS && LOTS.algos) {
      LOTS.algos.forEach(function (a) {
        var o = document.createElement('option');
        o.value = a; o.textContent = 'Algo ' + a;
        algoSel.appendChild(o);
      });
    } else {
      algoSel.disabled = true;
      if (logBox) logBox.disabled = true;
    }

    function lotsByCat(step, algo) {
      if (!LOTS) return {};
      var acc = {};
      LOTS.rows.forEach(function (r) {
        if (algo && LOTS.algos[r[1]] !== algo) return;
        var b = Math.floor(mins(r[0]) / step) * step;
        var cat = LOTS.cats[r[2]];
        (acc[cat] = acc[cat] || {});
        acc[cat][b] = (acc[cat][b] || 0) + r[3];
      });
      var out = {};
      Object.keys(acc).forEach(function (cat) {
        out[cat] = Object.keys(acc[cat]).map(function (b) {
          return [+b, acc[cat][b]];
        }).sort(function (a, b) { return a[0] - b[0]; });
      });
      return out;
    }

    var state = [], lots = {}, X, Y, LY, x0, x1, lotMax, useLog;

    function draw() {
      var step = +tfSel.value;
      svg.innerHTML = '';
      state = C.series.map(function (s, i) {
        return {name: s.name, axis: s.axis || 'left', color: LINE[i % LINE.length],
                data: resample(s.points, step)};
      });
      var all = [].concat.apply([], state.map(function (s) { return s.data; }));
      if (!all.length) return;
      var xs = all.map(function (d) { return d[0]; });
      x0 = Math.min.apply(null, xs); x1 = Math.max.apply(null, xs);
      if (C.shade) {
        x0 = Math.min(x0, mins(C.shade[0])); x1 = Math.max(x1, mins(C.shade[1]));
      }
      if (x1 === x0) x1 = x0 + 1;

      // ticks snap to a round step (index 50, premium 1) and the range spans
      // exactly four of them, so gridlines land ON the numbers
      var STEP = C.axis_step || {left: 50, right: 1};
      var range = {};
      ['left', 'right'].forEach(function (ax) {
        var vs = state.filter(function (s) { return s.axis === ax; })
          .reduce(function (a, s) {
            return a.concat(s.data.map(function (d) { return d[1]; }));
          }, []);
        if (!vs.length) return;
        var lo = Math.min.apply(null, vs), hi = Math.max.apply(null, vs);
        var st = STEP[ax] || 1;
        var pad = (hi - lo) * 0.08 || st;
        lo -= pad; hi += pad;
        var base = Math.floor(lo / st) * st, top = Math.ceil(hi / st) * st;
        if (top <= base) top = base + st;
        var ts = Math.ceil((top - base) / 4 / st) * st;
        range[ax] = [base, base + ts * 4];
      });

      X = function (m) { return L + (m - x0) / (x1 - x0) * (W - L - R); };
      Y = function (v, ax) {
        var r = range[ax]; if (!r) return P1B;
        return T + (1 - (v - r[0]) / (r[1] - r[0])) * (P1B - T);
      };

      lots = lotsByCat(step, algoSel.value);
      useLog = logBox && logBox.checked;
      lotMax = 0;
      Object.keys(lots).forEach(function (c) {
        lots[c].forEach(function (p) { if (p[1] > lotMax) lotMax = p[1]; });
      });
      LY = function (v) {
        if (lotMax <= 0) return H - B;
        var t = useLog
          ? Math.log10(Math.max(v, 1)) / Math.log10(Math.max(lotMax, 10))
          : v / lotMax;
        return (H - B) - t * (H - B - P2T);
      };

      // ---- panel backgrounds: one card, two clearly bounded regions ----
      svg.appendChild(el('rect', {x: L, y: T, width: W - L - R, height: P1B - T,
                                  fill: '#FFFFFF'}));
      svg.appendChild(el('rect', {x: L, y: P2T, width: W - L - R, height: (H - B) - P2T,
                                  fill: '#F8FAFC', rx: 3}));

      if (C.shade) {
        var sa = X(mins(C.shade[0])), sb = X(mins(C.shade[1]));
        [[T, P1B], [P2T, H - B]].forEach(function (band) {
          svg.appendChild(el('rect', {x: sa, y: band[0], width: Math.max(sb - sa, 1),
                                      height: band[1] - band[0],
                                      fill: '#94A3B8', 'fill-opacity': '.13'}));
        });
        svg.appendChild(el('text', {x: (sa + sb) / 2, y: T + 12, 'text-anchor': 'middle',
                                    'font-size': '10', fill: '#475569'},
                           C.shade_label || ''));
      }

      // ---- price panel gridlines + both axes ----
      for (var g = 0; g <= 4; g++) {
        var y = T + g / 4 * (P1B - T);
        svg.appendChild(el('line', {x1: L, y1: y, x2: W - R, y2: y,
                                    stroke: '#E5E7EB'}));
        ['left', 'right'].forEach(function (ax) {
          var r = range[ax]; if (!r) return;
          var v = r[1] - g / 4 * (r[1] - r[0]);
          svg.appendChild(el('text', {x: ax === 'left' ? L - 8 : W - R + 8, y: y + 4,
                                      'text-anchor': ax === 'left' ? 'end' : 'start',
                                      'font-size': '11', fill: '#6B7280'}, fmt(v)));
        });
      }

      // ---- lots panel: gridlines, ticks, caption ----
      svg.appendChild(el('text', {x: L, y: P2T - 9, 'font-size': '11',
                                  fill: '#475569', 'font-weight': '600'},
                         'Lots fired' + (useLog ? '  (log scale)' : '')));
      if (lotMax > 0) {
        var ticks = useLog ? [lotMax, lotMax / 10, lotMax / 100]
                           : [lotMax, lotMax / 2];
        ticks.forEach(function (v) {
          if (v < 1) return;
          var yy = LY(v);
          svg.appendChild(el('line', {x1: L, y1: yy, x2: W - R, y2: yy,
                                      stroke: '#E2E8F0', 'stroke-dasharray': '2 4'}));
          svg.appendChild(el('text', {x: L - 8, y: yy + 4, 'text-anchor': 'end',
                                      'font-size': '10', fill: '#94A3B8'}, lotFmt(v)));
        });
      }

      // x ticks every 30 minutes, drawn through BOTH panels so they read as one
      for (var m = Math.ceil(x0 / 30) * 30; m <= x1; m += 30) {
        svg.appendChild(el('line', {x1: X(m), y1: T, x2: X(m), y2: P1B,
                                    stroke: '#F1F5F9'}));
        svg.appendChild(el('line', {x1: X(m), y1: P2T, x2: X(m), y2: H - B,
                                    stroke: '#F1F5F9'}));
        svg.appendChild(el('line', {x1: X(m), y1: H - B, x2: X(m), y2: H - B + 4,
                                    stroke: '#9CA3AF'}));
        svg.appendChild(el('text', {x: X(m), y: H - B + 17, 'text-anchor': 'middle',
                                    'font-size': '11', fill: '#6B7280'}, label(m)));
      }
      svg.appendChild(el('line', {x1: L, y1: P1B, x2: W - R, y2: P1B, stroke: '#CBD5E1'}));
      svg.appendChild(el('line', {x1: L, y1: H - B, x2: W - R, y2: H - B, stroke: '#9CA3AF'}));

      // ---- the lines ----
      state.forEach(function (s) {
        if (!s.data.length) return;
        var d = s.data.map(function (p, i) {
          return (i ? 'L' : 'M') + X(p[0]).toFixed(1) + ' ' + Y(p[1], s.axis).toFixed(1);
        }).join(' ');
        svg.appendChild(el('path', {d: d, fill: 'none', stroke: s.color,
                                    'stroke-width': 1.8, 'stroke-linejoin': 'round'}));
      });

      // ---- the dots, on stems so they read as activity off the baseline ----
      var cats = CAT_ORDER.filter(function (c) { return lots[c] && lots[c].length; });
      var spread = Math.min(2.6, (W - L - R) / Math.max(x1 - x0, 1) * 0.5);
      cats.forEach(function (cat, ci) {
        var dx = (ci - (cats.length - 1) / 2) * spread;   // nudge so dots at the
        var col = CAT_COLOR[cat] || '#64748B';            // same minute stay visible
        lots[cat].forEach(function (p) {
          if (!p[1]) return;
          var cx = X(p[0]) + dx, cy = LY(p[1]);
          svg.appendChild(el('line', {x1: cx, y1: H - B, x2: cx, y2: cy,
                                      stroke: col, 'stroke-width': 1,
                                      'stroke-opacity': '.32'}));
          svg.appendChild(el('circle', {cx: cx, cy: cy, r: 3,
                                        fill: col, 'fill-opacity': '.9'}));
        });
      });

      legend.innerHTML = state.map(function (s) {
        return '<label class="chip"><span class="sw2" style="background:' + s.color +
               '"></span>' + s.name + ' <span class="note">(' +
               (s.axis === 'left' ? 'left' : 'right') + ')</span></label>';
      }).concat(cats.map(function (cat) {
        var tot = lots[cat].reduce(function (a, p) { return a + p[1]; }, 0);
        return '<label class="chip"><span class="sw2 rnd" style="background:' +
               (CAT_COLOR[cat] || '#64748B') + '"></span>' +
               (CAT_LABEL[cat] || cat) + ' <span class="note">' + lotFmt(tot) +
               ' lots</span></label>';
      })).join('');

      var hover = el('g', {});
      svg.appendChild(hover);

      svg.onmousemove = function (e) {
        var box = svg.getBoundingClientRect();
        var px = (e.clientX - box.left) / box.width * W;
        var m = x0 + (px - L) / (W - L - R) * (x1 - x0);
        var ref = state[0].data, snap = null, bd = 1e9;
        ref.forEach(function (p) {
          var dd = Math.abs(p[0] - m);
          if (dd < bd) { bd = dd; snap = p[0]; }
        });
        if (snap === null) return;
        hover.innerHTML = '';
        hover.appendChild(el('line', {x1: X(snap), y1: T, x2: X(snap), y2: H - B,
                                      stroke: '#64748B', 'stroke-width': 1,
                                      'stroke-dasharray': '3 3'}));
        var rows = ['<div class="r"><span class="k">Time</span><span class="v">' +
                    label(snap) + '</span></div>'];
        state.forEach(function (s) {
          var hit = null;
          s.data.forEach(function (p) { if (p[0] === snap) hit = p; });
          if (!hit) return;
          hover.appendChild(el('circle', {cx: X(snap), cy: Y(hit[1], s.axis), r: 3.5,
                                          fill: s.color, stroke: '#fff', 'stroke-width': 1.5}));
          rows.push('<div class="r"><span class="k"><span class="sw" style="background:' +
                    s.color + '"></span>' + s.name + '</span><span class="v">' +
                    fmt(hit[1]) + '</span></div>');
        });
        cats.forEach(function (cat) {
          var hit = null;
          lots[cat].forEach(function (p) { if (p[0] === snap) hit = p; });
          if (!hit || !hit[1]) return;
          hover.appendChild(el('circle', {cx: X(snap), cy: LY(hit[1]), r: 5.5,
                                          fill: 'none', stroke: CAT_COLOR[cat],
                                          'stroke-width': 1.6}));
          rows.push('<div class="r"><span class="k"><span class="sw rnd" style="background:' +
                    (CAT_COLOR[cat] || '#64748B') + '"></span>' + (CAT_LABEL[cat] || cat) +
                    '</span><span class="v">' + lotFmt(hit[1]) + ' lots</span></div>');
        });
        tip.innerHTML = rows.join('');
        tip.style.opacity = 1;
        var wpx = wrap.clientWidth, xpx = X(snap) / W * wpx, tw = tip.offsetWidth;
        tip.style.left = (xpx + tw + 24 > wpx ? xpx - tw - 14 : xpx + 14) + 'px';
        tip.style.top = '14px';
      };
      svg.onmouseleave = function () {
        tip.style.opacity = 0;
        hover.innerHTML = '';
      };
    }

    tfSel.addEventListener('change', draw);
    algoSel.addEventListener('change', draw);
    if (logBox) logBox.addEventListener('change', draw);
    draw();
  }

  document.querySelectorAll('.chart-card').forEach(initChart);
})();
</script>
"""


def _pnl_td(value, crore=False):
    """Table cell for a P&L amount — profit green, loss red."""
    text = format_crore(value) if crore else _money(value)
    cls = _sign_cls(value)
    return f'<td class="{cls}">{_esc(text)}</td>' if cls else f"<td>{_esc(text)}</td>"


def _metric_cells(r):
    # MTM is the only P&L figure in the report — its Realized and Unrealized
    # parts live in the Excel workbook's Summary / MTM Data sheets.
    # Allocation is displayed x100 (the stored value is in hundreds);
    # MTM % = MTM / Allocation, on the stored basis.
    return (
        f"<td>{r['Users']}</td>"
        f"<td>{_esc(format_crore(r['MaxLoss']))}</td>"
        f"<td>{_esc(format_crore(r['Allocation'] * 100))}</td>"
        f"{_pnl_td(r['MTM'], crore=True)}"
        f"<td>{_num2(r['Return'])}</td>"
        f"<td>{r['SLHit']}</td>"
    )


def _pivot_table(pivot_rows, table_id="pivot-table"):
    """Drill-down pivot: one 'Grand Total' row per algo (collapsed default);
    clicking it reveals the Int / Pos+Int sub-total rows, and clicking those
    reveals the per-server detail. Plain inline JS — no libraries.
    `table_id` prefixes every drill key so two pivots (one per User MTM
    file) never cross-expand each other."""
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
        akey = f"{table_id}-a{i}"
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
            for k, r in enumerate(block["rows"]):
                dkey = f"{skey}d{k}"
                body.append(
                    f'<tr class="lvl3 hidden" data-parent="{skey}" data-key="{dkey}">'
                    f'<td class="txt indent">{_esc(section)}</td>'
                    f'<td class="txt">{_esc(algo)}</td>'
                    f'<td class="txt"><span class="caret">&#9656;</span>{_esc(r["SERVER"])}</td>'
                    f"{_metric_cells(r)}</tr>"
                )
                # deepest level: the server's accounts, worst realized first
                for u in r.get("user_rows", []):
                    body.append(
                        f'<tr class="lvl4 hidden" data-parent="{dkey}">'
                        f'<td class="txt"></td><td class="txt"></td>'
                        f'<td class="txt">{_esc(u["UserID"])}</td>'
                        f'<td class="note">{_esc(u["Alias"])}</td>'
                        f"<td>{_esc(format_crore(u['MaxLoss']))}</td>"
                        f"<td>{_esc(format_crore(u['Allocation'] * 100))}</td>"
                        f"{_pnl_td(u['MTM'], crore=True)}"
                        f"<td>{_num2(u['Return'])}</td>"
                        f"<td>{u['SLHit']}</td>"
                        "</tr>"
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
    <div class="pivot-tools" data-target="{table_id}">
      <button type="button">Expand all</button>
      <button type="button">Collapse all</button>
      <span class="hint">click a row to drill down: Algo &rarr; Int / Pos+Int &rarr; servers
      &rarr; users</span>
    </div>
    <table class="tbl pivot-tbl" id="{table_id}">
      <thead><tr>
        <th class="txt">Type</th><th class="txt">Algo</th><th class="txt">Server</th>
        <th>Users</th><th>Max Loss</th><th>Allocation</th>
        <th>MTM</th><th>MTM %</th><th>SL Hit</th>
      </tr></thead>
      <tbody>{''.join(body)}</tbody>
    </table>"""


_PIVOT_SCRIPT = """
<script>
(function () {
  document.querySelectorAll('.pivot-tools').forEach(function (tools) {
    var table = document.getElementById(tools.dataset.target);
    if (!table) return;
    function childrenOf(key) {
      return table.querySelectorAll('tr[data-parent="' + key + '"]');
    }
    function collapse(tr) {
      tr.classList.remove('open');
      childrenOf(tr.dataset.key).forEach(function (child) {
        child.classList.add('hidden');
        if (child.dataset.key) collapse(child);
      });
    }
    table.querySelectorAll('tr[data-key]').forEach(function (tr) {
      tr.addEventListener('click', function () {
        if (tr.classList.contains('open')) { collapse(tr); return; }
        tr.classList.add('open');
        childrenOf(tr.dataset.key).forEach(function (child) { child.classList.remove('hidden'); });
      });
    });
    var btns = tools.querySelectorAll('button');
    btns[0].addEventListener('click', function () {
      table.querySelectorAll('tr[data-key]').forEach(function (tr) { tr.classList.add('open'); });
      table.querySelectorAll('tr[data-parent]').forEach(function (tr) { tr.classList.remove('hidden'); });
    });
    btns[1].addEventListener('click', function () {
      table.querySelectorAll('tr[data-key]').forEach(collapse);
    });
  });
})();
</script>
"""


def build_dor_html(report_date, deviation, tv_totals, tv_summary=None, pivot_stats=None,
                   pivot_rows=None, user_obs=None, strikes=None, slippage=None,
                   portfolio=None, mids=None, excel_bytes=None, excel_filename=None,
                   seg_sections=None, tv_sections=None, expiry_map=None,
                   order_rows=None, order_sections=None, dte="0DTE", charts=None):
    """Assemble the complete DOR.html document and return it as a string.
    `user_obs` (per-user observations from tradevalue.user_lot_observations)
    feeds the Algo Summary drill-down (outlier user ids); `strikes` (from
    tradevalue.strike_report) feeds the strikes section; `slippage`
    ({"summary", "overall", "majors"} from summary, all
    algos) feeds the slippage section with its in-page algo filter;
    `portfolio` (groups from portfolio.portfolio_groups) feeds the portfolio
    analysis — the QS table prerendered plus any-pattern analysis in the
    browser; `mids` ({index: day-mid price}) centres the option chain on the
    ATM strike; `expiry_map` ({index: expiry text entered in the dashboard})
    labels the strike chains; `excel_bytes`/`excel_filename` embed the full
    workbook as a download button in the masthead; `dte` ("0DTE" / "1DTE" /
    "4DTE") is the account scope the run covered and is shown in the masthead
    and the segregation note. Each section is skipped when its data is
    omitted.

    When a secondary User MTM is uploaded, pass `seg_sections`
    ([{"label", "stats", "rows"}, ...]) and `tv_sections` ([{"label",
    "summary", "obs"}, ...]) — one entry per MTM file — instead of
    `pivot_stats`/`pivot_rows` and `tv_summary`/`user_obs`; each segregation
    section and Algo Summary table then renders once per file."""
    generated = datetime.now().strftime("%d-%b-%Y %H:%M")
    if seg_sections is None:
        seg_sections = [{"label": None, "stats": pivot_stats, "rows": pivot_rows}]
    if tv_sections is None:
        tv_sections = [{"label": None, "summary": tv_summary, "obs": user_obs}]

    tv_kpis = _kpi_tiles([
        ("Users", _money(tv_totals["users"])),
        ("Orders", _money(tv_totals["orders"])),
        ("Total Lots", _money(tv_totals["lots"])),
        ("Trade Value", format_crore(tv_totals["trade_value"])),
    ])

    seg_blocks = []
    for i, sec in enumerate(seg_sections):
        stats, rows = sec["stats"], sec["rows"]
        grand = next((r for r in rows if r["kind"] == "grandtotal"), None)
        seg_pairs = [("Accounts", _money(stats["accounts"]))]
        # nothing matched the All User sheet -> no Positional / Intraday counts
        if stats.get("positional") is not None:
            seg_pairs += [
                ("Positional", _money(stats["positional"])),
                ("Intraday", _money(stats["intraday"])),
            ]
        if stats.get("unclassified"):
            seg_pairs += [("Unclassified", _money(stats["unclassified"]))]
        if grand:
            seg_pairs += [
                ("Allocation", format_crore(grand["Allocation"] * 100)),
                ("MTM", format_crore(grand["MTM"]), _sign_cls(grand["MTM"])),
                ("MTM %", _num2(grand["Return"])),
            ]
        heading = f" &mdash; {_esc(sec['label'])}" if sec.get("label") else ""
        unc_note = ""
        if stats.get("unclassified"):
            unc_note = (
                f' {stats["unclassified"]} account(s) fall outside the '
                f"{_esc(dte)} scope &mdash; absent from the All User sheet, dropped as "
                "<code>DLR ACC</code> / <code>NOT RUNNING</code>, or running on other "
                "days. They are grouped under <b>Unclassified</b> at the bottom of the "
                "pivot; algo totals cover classified accounts only, the Grand Total "
                "covers everything.")
        seg_blocks.append(f"""
<section>
  <h2>Intraday / Positional Segregation{heading}</h2>
  <div class="section-note">Accounts are typed from the All User sheet&rsquo;s
  <b>Running Type</b> (INT &rarr; Int, POS &rarr; Pos+Int), narrowed to the accounts the
  <b>{_esc(dte)}</b> scope covers. <b>MTM = Realized P&amp;L + Unrealized P&amp;L</b>, computed
  per account from the User MTM sheet; <b>MTM %</b> = MTM / Allocation. The Realized and
  Unrealized parts are in the Excel workbook&rsquo;s <b>Summary</b> and <b>MTM Data</b>
  sheets.</div>
  {_kpi_tiles(seg_pairs)}
  <div class="card scroll">
    {_pivot_table(rows, table_id=f"pivot-table-{i}")}
  </div>
  <div class="footnote">
    *Sub-Total server counts are per section; algo totals count each server once. The
    per-account base table every figure here aggregates (<b>MTM Data</b>) is in the Excel
    workbook.{unc_note}
  </div>
</section>""")
    seg_sections_html = "".join(seg_blocks)

    tv_blocks = []
    multi_note = ""
    if len(tv_sections) > 1 and any(o.get("also_trades")
                                    for sec in tv_sections for o in (sec.get("obs") or [])):
        multi_note = (
            '<div class="footnote">&#9888; Each index is judged '
            "<b>independently</b>. A user trading more than one index appears in "
            "several tables, and in each their Lots per Cr is that index&rsquo;s lots "
            "over their <b>whole</b> allocation &mdash; a partial-exposure figure. Such "
            "a user reads lower than a single-index peer deploying the same capital and "
            "can show as <i>Below average range</i> without under-trading; they are "
            "marked <b>&#9888; also trades &hellip;</b> in the expanded outlier "
            "rows.</div>")
    for i, sec in enumerate(tv_sections):
        heading = (f'<h3 class="pf-title">{_esc(sec["label"])}</h3>'
                   if sec.get("label") else "")
        tv_blocks.append(f"""
  {heading}
  <div class="card scroll">
    {_tv_summary_table(sec["summary"], sec["obs"], table_id=f"tv-summary-{i}")}
  </div>""")
    tv_tables_html = "".join(tv_blocks)

    if order_sections is None and order_rows:
        order_sections = [{"label": None, "rows": order_rows}]
    order_blocks = []
    for i, sec in enumerate(order_sections or []):
        if not sec["rows"]:
            continue
        heading = (f'<h3 class="pf-title">{_esc(sec["label"])}</h3>'
                   if sec.get("label") else "")
        unattributed = next((r for r in sec["rows"] if not r["algo"]), None)
        note = ""
        if unattributed:
            note = (f'<div class="footnote"><b>&mdash;</b> = '
                    f'{unattributed["users"]} user(s) that fired orders but appear in no '
                    "User MTM, so they have no algo and no Int / Pos+Int type (the Trade "
                    "Value Algo Summary leaves them out entirely): "
                    + _esc(", ".join(unattributed.get("user_ids", []))) + "</div>")
        order_blocks.append(f"""
  {heading}
  <div class="card scroll">
    {_orders_table(sec["rows"], table_id=f"orders-summary-{i}")}
  </div>
  {note}""")
    orders_section = f"""
<section>
  <h2>Orders Summary</h2>
  <div class="section-note">Every <b>option</b> order fired, whatever its outcome, per
  algo and Int / Pos+Int type &mdash; click a row to drill into its <b>servers</b>.
  Futures rows are excluded (same scope as the strikes section): a rejected FUT order is
  not index-options activity and would invent an algo row in an index that algo never
  traded. <b>Executed</b> counts the COMPLETE orders;
  <b>Failed/Cancelled/Rejected</b> counts the cancelled and rejected
  ones; <b>Pending</b> counts orders still live at end of day (OPEN / OPEN_PENDING) &mdash;
  live is not failed, so they are kept apart.
  <b>Hedge</b> and <b>VAR</b> count the <b>executed</b> orders tagged
  <code>h_&hellip;</code> and <code>v_&hellip;</code> &mdash; both are slices of
  <b>Executed</b>, never of Total Orders, so a cancelled hedge is counted in Failed with
  every other cancellation. Algo and type
  are resolved exactly as in the Trade Value rows and the tables are split per index the
  same way, so the user counts reconcile &mdash; the one legitimate difference is a user
  whose <i>every</i> order in an index failed: they appear here (orders were fired) but
  not in that index&rsquo;s Trade Value table (nothing executed). Users with no MTM entry
  have neither an algo nor a type and collect in the <b>&mdash;</b> row.</div>
  {''.join(order_blocks)}
</section>""" if order_blocks else ""

    download_btn = ""
    if excel_bytes:
        b64 = base64.b64encode(excel_bytes).decode("ascii")
        fname = excel_filename or "DOR.xlsx"
        download_btn = (
            f'<a class="dl-btn" download="{_esc(fname)}" '
            f'href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;'
            f'base64,{b64}">&#11015; Download Excel &mdash; full data</a>'
        )

    if slippage is None:
        slippage_section = ""
    else:
        slippage_note = """
  <div class="section-note">ML % = MAX LOSS / ALLOCATION; Realized ML % =
  |Realized P&amp;L| / ALLOCATION &mdash; measured on the <b>realized</b> loss, not MTM,
  because a max-loss stop is about what was actually booked. Plain ratios of allocation,
  same convention as MTM %. An account has slippage only when Realized ML % exceeds ML % by at least 0.1
  (1.00 &rarr; 1.09 is not slippage; 1.10 is). Avg Slippage = average Realized ML % of the
  algo&rsquo;s slippage accounts; <b>Major</b> = slippage accounts above their algo&rsquo;s
  average, plus the lone slippage account of an algo &mdash; with one account the average
  <i>is</i> that account, so it is that algo&rsquo;s worst by definition.</div>"""
        no_sl = slippage.get("no_sl") or []
        if no_sl:
            slippage_note += f"""
  <div class="footnote">&#8505; {len(no_sl)} account(s) carry an allocation but
  <b>MAX LOSS = 0</b>, so no stop-loss is configured &mdash; they are <b>excluded</b> from
  this analysis and listed in the Excel <b>no_sl_Acc</b> sheet:
  {_esc(", ".join(n["UserID"] for n in no_sl[:10]))}{" &hellip;" if len(no_sl) > 10 else ""}.</div>"""
        empty_note = slippage.get("empty_note", "-- No slippage today --")
        slippage_body = (_slippage_tables(slippage) if slippage["overall"]["slipped"]
                         else '<div class="card"><div class="empty-note">'
                              f"{_esc(empty_note)}</div></div>")
        slippage_section = f"""
<section>
  <h2>Max SL Slippage Analysis</h2>
  {slippage_note}
  {slippage_body}
</section>"""

    if portfolio:
        pf_report = portfolio_report(portfolio, DEFAULT_PORTFOLIO_PATTERN)
        pf_data = portfolio_json(portfolio)
        portfolio_section = f"""
<section>
  <h2>Portfolio Analysis</h2>
  <div class="section-note">From the Multileg Orders (MLOB), COMPLETED Orders only.
  PnL = &Sigma; sell (Avg Price &times; Filled Qty) &minus; &Sigma; buy (Avg Price &times;
  Filled Qty); Total Orders = the COMPLETE entries of the portfolios. The algo comes from the
  User MTM matching (users with no MTM entry inherit their server&rsquo;s algo). Click a row
  to see its per-user detail. The default view covers every portfolio whose name contains
  &ldquo;{_esc(DEFAULT_PORTFOLIO_PATTERN)}&rdquo;; type any other name fragment below to
  analyse it instantly.</div>
  <h3 class="pf-title">{_esc(DEFAULT_PORTFOLIO_PATTERN)} Portfolio Analysis</h3>
  <div class="card scroll">
    {_portfolio_table(pf_report, "pf-qs")}
  </div>
  <div class="chain-controls pf-controls">
    <label>Analyze another portfolio &mdash; pick a name or type a fragment
      <input id="pf-pattern" list="pf-names" placeholder="e.g. WT3% or PPP"></label>
    <datalist id="pf-names"></datalist>
    <button type="button" id="pf-analyze">Analyze</button>
  </div>
  <div id="pf-result"></div>
</section>
<script>var PF_DATA = {json.dumps(pf_data)};</script>
{_PORTFOLIO_SCRIPT}"""
    else:
        portfolio_section = ""

    charts_section = _charts_section(charts)

    strikes_section = f"""
<section>
  <h2>Strikes Traded</h2>
  <div class="section-note">A strike is one contract (index + strike + CE/PE);
  every exchange symbol format for the same contract is counted once, and each
  index&rsquo;s expiry is the one entered in the dashboard. The algo view is
  the default &mdash; click an algo to see its per-server counts (an algo&rsquo;s count is
  distinct across its servers, not the column sum). Pick an expiry and index to filter, or
  to see the option chain &mdash; CE / PE lots per strike (whole lots, same date-based lot
  sizes as the trade value rows). When the index has a day-mid (entered in the dashboard:
  (day open + day close) / 2), the chain centres on the <b>ATM</b> strike (highlighted) and
  shows only <b>No. of strikes</b> above and below it &mdash; change the number to widen or
  narrow the view. Users with no MTM entry inherit their server&rsquo;s algo.</div>
  {_strikes_tables(strikes, mids, expiry_map)}
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
    display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 10px; margin-bottom: 14px;
  }}
  @media (max-width: 640px) {{
    .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
  }}
  .kpi {{
    background: var(--surface); border: 1px solid var(--line); border-radius: 10px;
    padding: 12px 14px; text-align: center;
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
  .chain-controls {{ align-items: center; }}
  .chain-controls select {{
    margin-left: 6px; padding: 4px 8px; border: 1px solid var(--line); border-radius: 6px;
    background: var(--surface); color: var(--ink); font-size: 12.5px;
  }}
  .chain-controls input:not([type="checkbox"]) {{
    margin-left: 6px; padding: 4px 10px; border: 1px solid var(--line); border-radius: 6px;
    background: var(--surface); color: var(--ink); font-size: 12.5px; width: 220px;
  }}
  .chain-controls #chain-n {{ width: 58px; }}
  .chip-group {{ display: inline-flex; gap: 8px; flex-wrap: wrap; }}
  /* intraday charts — one card per index, svg scales with the card */
  .chart-wrap {{ width: 100%; position: relative; }}
  .c-svg {{ width: 100%; height: auto; display: block; }}
  .chart-card + .pf-title {{ margin-top: 18px; }}
  .dot-key {{ display: inline-block; width: 9px; height: 9px; border-radius: 50%;
              margin: 0 4px 0 6px; vertical-align: middle; }}
  .sw2 {{ display: inline-block; width: 10px; height: 10px; border-radius: 2px;
          margin-right: 5px; vertical-align: middle; }}
  .sw2.rnd {{ border-radius: 50%; width: 9px; height: 9px; }}
  .chart-tip {{
    position: absolute; pointer-events: none; opacity: 0; transition: opacity .08s;
    background: rgba(30,41,59,.96); color: #fff; border-radius: 6px;
    padding: 8px 10px; font-size: 12px; line-height: 1.55; white-space: nowrap;
    box-shadow: 0 4px 14px rgba(0,0,0,.28); z-index: 5;
    font-variant-numeric: tabular-nums;
  }}
  .chart-tip .r {{ display: flex; justify-content: space-between; gap: 18px; }}
  .chart-tip .k {{ color: #CBD5E1; }}
  .chart-tip .v {{ font-weight: 600; }}
  .chart-tip .sw {{ display: inline-block; width: 8px; height: 8px; border-radius: 2px;
                    margin-right: 6px; }}
  .chart-tip .sw.rnd {{ border-radius: 50%; }}
  @media print {{ .chart-tip {{ display: none; }} }}
  .chain-controls label.chip {{
    display: inline-flex; align-items: center; gap: 6px; padding: 4px 12px;
    border: 1px solid var(--line); border-radius: 999px; cursor: pointer;
    background: var(--surface); color: var(--ink); user-select: none;
    transition: background .12s ease, border-color .12s ease;
  }}
  .chain-controls label.chip:hover {{ background: #F8FAFC; }}
  .chain-controls label.chip:has(input:checked) {{
    background: #EDF2FB; border-color: #94A3B8; font-weight: 600;
  }}
  .chain-controls input[type="checkbox"] {{
    accent-color: var(--header); width: 14px; height: 14px; margin: 0;
  }}
  .chain-controls button {{
    border: 1px solid var(--line); background: var(--surface); border-radius: 6px;
    padding: 4px 14px; font-size: 12.5px; cursor: pointer; color: var(--ink);
  }}
  .chain-controls button:hover {{ background: #F8FAFC; }}
  .pf-controls {{ margin-top: 16px; }}
  .tbl tr.atm td {{ background: #FEF3C7; font-weight: 600; }}
  /* option-chain side tints: CE faint blue, PE faint amber (ATM/total exempt) */
  #chain-body tr:not(.atm):not(.grand) td:first-child {{ background: rgba(0, 114, 178, .05); }}
  #chain-body tr:not(.atm):not(.grand) td:last-child {{ background: rgba(230, 159, 0, .06); }}
  .tbl td.note {{ color: var(--muted); font-style: italic; }}
  .pf-title {{ font-size: 13.5px; font-weight: 700; color: var(--header);
               margin: 10px 0 8px; }}
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
  .pivot-tbl tr[data-key] td, .drill-tbl tr[data-key] td {{
    cursor: pointer; user-select: none; }}
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
    .pivot-tools, .masthead .dl-btn, .pf-controls {{ display: none; }}
    .masthead {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    .tbl th, .tbl tr.subtotal td, .tbl tr.algototal td, .tbl tr.grand td,
    .dot {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  }}
</style>
</head>
<body>
<div class="masthead">
  <div class="inner">
    <div class="date">Report date<b>{_esc(report_date)}</b>{_esc(dte)}</div>
    <h1>Daily Operations Report</h1>
    <div class="sub">Trade Value &amp; Intraday / Positional Segregation — summary view</div>
    {download_btn}
  </div>
</div>
<main>

{seg_sections_html}

{slippage_section}

<section>
  <h2>Trade Value</h2>
  <div class="section-note">
    Completed NIFTY / BANKNIFTY / SENSEX orders only (F&amp;O exchanges NFO / BFO), duplicates
    removed. Outliers are users whose
    <b>Lots per Cr</b> (combined lots &divide; normalise, where normalise = allocation /
    1,00,000 &mdash; sub-1-Cr allocations normalise fractionally, e.g. 80,000 &rarr; 0.8)
    is more than {deviation:g} robust deviation(s) (MAD) away from their algo&rsquo;s median. All
    columns count users, so Below + In + Above = Total Users (only users with no usable
    allocation carry no flag). Click an algo row to see its outlier user ids.
  </div>
  {tv_kpis}
  {multi_note}
  {tv_tables_html}
  <div class="footnote">
    *Every account is normalised by allocation / 1,00,000 (e.g. 15,60,000 &rarr; 15.6;
    80,000 &rarr; 0.8), so Lots per Cr is comparable across allocation sizes. Expanded rows
    show each outlier user&rsquo;s id &middot; server &middot; actual lots fired &middot; flag.
  </div>
</section>

{charts_section}

{orders_section}

{portfolio_section}

{strikes_section}

</main>
<footer>
  Generated {_esc(generated)} &middot; Daily Operations Report &middot; Full row-level data is
  available in the accompanying Excel workbook.
</footer>
{_DRILL_SCRIPT}
{_PIVOT_SCRIPT}
</body>
</html>"""
