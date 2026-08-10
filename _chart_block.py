# NOTE: this scratch file has been merged into dor.py and is no longer used.
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
  <br>The lower panel is <b>lots fired</b> per bucket:
  <span class="dot-key" style="background:#15803D"></span>completed,
  <span class="dot-key" style="background:#DC2626"></span>failed / cancelled / rejected,
  <span class="dot-key" style="background:#EA580C"></span>hedge,
  <span class="dot-key" style="background:#9333EA"></span>VAR.
  An order carries a status and a tag independently, so the <b>tag wins</b>: a completed
  hedge counts as hedge, never as completed, which keeps the four groups disjoint and the
  totals reconciled. Square-off orders are excluded &mdash; they close a position rather
  than place one.</div>
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
  var CAT_COLOR = {complete: '#15803D', failed: '#DC2626',
                   hedge: '#EA580C', var: '#9333EA'};
  var CAT_LABEL = {complete: 'Completed', failed: 'Failed',
                   hedge: 'Hedge', var: 'VAR'};
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
      var cats = Object.keys(lots);
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
