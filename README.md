# Daily Operations Report (DOR)

A Streamlit app (plus standalone CLIs) that turns four raw inputs into the daily ops report:

1. **Trade Value analysis** (§1) — per-user lots / trade value from the orderbook, with
   statistical outlier flagging on per-user **Lots per Cr** (split Int / Pos+Int), the
   per-(algo, type) summary (a drill-down that reveals the outlier user ids), and the
   strikes-traded views (algo → server counts + a CE/Strike/PE option chain with ATM
   centring and hedge/VAR/sq-off/algo filters).
2. **Intraday / Positional segregation** (§2) — a pivot of Realized/Unrealized P&L per algo →
   Int / Pos+Int → server, with max-loss addons applied to positional accounts, the worst-10%ile
   exception report, and the **slippage** view (accounts whose realized loss / allocation
   overshoots their max-loss / allocation by ≥ 0.1).

3. **Portfolio analysis** (§4, optional — needs the Multileg Orders / MLOB input) — per
   (algo, server) portfolio summary with per-user drill-down for the "QS" portfolios by
   default, plus instant any-portfolio analysis inside the HTML report itself.

Outputs: one Excel workbook `DOR_<date>.xlsx` (sheets: `tradevalue`, `summary`, `Strikes`,
`Portfolio QS` when the MLOB is given, `Segregation`, `Raw_Data_Per_User`, `Worst 10%ile`,
`Slippage`) and one self-contained, client-shareable `DOR_<date>.html` (all tables
center-aligned, data and headers; numbers in Indian digit grouping, e.g. `16,44,536`).

## Files

| File | Role |
|---|---|
| `app.py` | Streamlit UI — upload the inputs, set the outlier deviation, one **Process** click computes everything. Run with `streamlit run app.py`. |
| `tradevalue.py` | Trade value engine: report rows, algo summary, strikes/option chain (also a CLI: `python tradevalue.py orderbook.csv -s user_mtm.xlsx -d 1`). |
| `segregate_int_pos_mtm2.py` | Segregation engine: pivot, raw per-user sheet, worst 10%ile, slippage (also an interactive CLI). |
| `portfolio.py` | Portfolio analysis engine over the Multileg Orders (MLOB): per-portfolio / per-user PnL, QS summary, any-pattern reports. |
| `dor.py` | Renders the DOR.html summary (inline CSS/SVG/JS, no external assets — dropdown filters, the drill-downs and the in-page portfolio analysis are plain inline JS). |

## Inputs

| Input | Used for |
|---|---|
| Compiled Orderbook (CSV/Excel) | Trade value rows |
| Compiled User MTM (CSV/Excel) | Allocations + algo for trade value; the account universe for segregation |
| Combined Max Loss — 1DTE (required) | Positional classification + realized-P&L addon |
| Combined Max Loss — 4DTE (optional) | Extra addon for non-Noren positional accounts |
| Multileg Orders — MLOB (optional, Excel/CSV) | The Portfolio Analysis (§4) — omitting it just hides that section |
| Outlier deviation `k` (default 1.0) | Width of the Lots-per-Cr range (median ± k × MAD, per algo + type group) — drives the flags and the Algo Summary Range |
| NIFTY / SENSEX day open + day close | Day-mid = (open + close) / 2 — the option chain's ATM strike (§1.9); 0 = not set |
| Segregation algos (multiselect, empty = all) | Which algos the Int / Pos+Int pivot and its KPI summary cover (§2.3) — dashboard and DOR.html; populated after the first Process |
| Slippage algos (multiselect, empty = all) | Which algos the slippage analysis covers (§2.7) — applies to the dashboard and DOR.html; populated after the first Process |

Column matching is tolerant: header names are normalised (lower-cased, non-alphanumerics stripped),
so `user_id`, `User ID` and `UserID` all match.

---

# Part 1 — Trade Value (`tradevalue.py`)

## 1.1 Row filters on the orderbook

A raw order row survives only if **all** of these hold:

- **Exchange ∈ {BFO, NFO}** — F&O only; NSE/BSE/MCX cash rows (e.g. `NIFTYBEES-EQ`) are dropped
  because their quantities aren't lot multiples.
- **Status == `COMPLETE`** — never-traded orders are dropped.
- **Symbol classifies into a segment** (substring match on the upper-cased symbol, checked in this
  order so BANKNIFTY isn't misread as NIFTY):
  `BANKNIFTY`/`BANK NIFTY` → **BANKNIFTY**, else `SENSEX` → **SENSEX**, else `NIFTY` → **NIFTY**,
  else the row is ignored.

## 1.2 Canonical keys

- **User key** — trimmed, upper-cased; **all-digit ids have leading zeros stripped**
  (`04101961` ≡ `4101961`), because Excel/CSV drops them. The report displays the longest
  (zero-padded) form it saw.
- **Server key** — trimmed, upper-cased; `NAN`/`NONE`/`NA` read as blank.

## 1.3 Deduplication

Duplicate orders are dropped on the key **(user key, trade date, order_id, exchg_order_id)**,
keeping the occurrence with the **lowest row id / SNO**. Rows whose order id is missing **or was
mangled into scientific notation by Excel** (e.g. `2.60623E+13` — which would collapse many
distinct orders onto one key) cannot be keyed and are kept as-is.

## 1.4 Lot-size history

`lots` uses a **date-dependent lot size** per segment (the divisor applied to traded quantity):

| Segment | Period | Lot size |
|---|---|---|
| NIFTY | ≤ 26-Dec-2024 | 25 |
| | 27-Dec-2024 → 30-Dec-2025 | 75 (exception: 29/30-Jan-2025 → 25) |
| | ≥ 31-Dec-2025 (and unknown date) | 65 |
| SENSEX | ≤ 03-Jan-2025 (also 30/31-Jan-2025) | 10 |
| | otherwise (and unknown date) | 20 |
| BANKNIFTY | ≤ 26-Feb-2025 | 15 |
| | 27-Feb-2025 → 26-Jun-2025 | 30 |
| | 27-Jun-2025 → 30-Dec-2025 | 35 |
| | ≥ 31-Dec-2025 (and unknown date) | 30 |

## 1.5 Aggregation — one report row per (date, server, user, segment)

All math is done in `Decimal` (no float drift):

```
Order Count = number of orders in the group
Quantity    = Σ qty                      (signed — buys and sells offset)
Lots        = Σ |qty| / lot_size(date)   (absolute — both sides add)
Trade Value = Σ avg_price × qty          (signed)
```

## 1.6 Allocation matching

Each group's user is looked up in the Compiled User MTM (keyed on the canonical user id).
The same user id may legitimately exist on several servers as **distinct accounts** — then the
entry whose **server matches the orderbook row's server** is used; if no server agrees, the row
stays **unmatched (blank Allocation)** rather than guessing the wrong account. Per (user, server)
only the first MTM row is kept.

## 1.7 Lots per Cr and outlier flagging

Every account normalises by `allocation / 1,00,000` (stored value — 1,00,000 stored = 1 Cr in
the ×100 display convention): 15,60,000 → 15.6, and the sub-1-Cr variations 80,000 / 60,000 /
40,000 / 20,000 → 0.8 / 0.6 / 0.4 / 0.2. Only accounts with no usable allocation carry no
metric and a blank flag. The report row's `Normalise` and `Lots per Cr` columns document the
row:

```
Normalise         = Allocation / 1,00,000
Lots per Cr (row) = row Lots / Normalise         (blank when the user has no allocation)
```

**The outlier unit is the USER, not the row.** A report row is per (user, segment), so a user
trading NIFTY and SENSEX would otherwise be judged twice on partial exposures. Per
(date, algo, user), the user's lots are summed across all their segments (and servers) and
normalised by their **total** allocation (one account per server):

```
lots_per_cr (user) = Σ user's lots / ( Σ user's allocations / 1,00,000 )
```

Each (user, server) account is also classified **Int / Pos+Int** using the segregation
classification (§2.1) — the tradevalue rows carry a `Type` column — and outliers are judged
**per (trade date, algo, type) group over these per-user values**, using **robust statistics**
(median / MAD) so that a blowing account cannot widen the very range that is meant to catch it
(with mean ± std-dev, one extreme user inflates the std dev and shields itself):

```
median = the middle per-user lots_per_cr          (n = users with a usable allocation)
MAD    = 1.4826 × median( |lots_per_cr − median| )   (std-dev equivalent, outlier-immune)
range  = [ median − k·MAD , median + k·MAD ]         (k = the chosen deviation, default 1)
```

- `lots_per_cr < mean − k·std` → **"Below average range"**
- `lots_per_cr > mean + k·std` → **"Above average range"**
- otherwise → **"In range"**
- users with no allocation or no algo can't be judged → blank flag

The user's flag is stamped on **each of their report rows** (so a row's own `Lots per Cr` may
sit inside the range while the row is flagged — the judgement is on the user's combined
exposure). The report note under the Algo Summary explains the normalisation.

## 1.8 Algo summary & totals

One summary row per (date, algo, **Int / Pos+Int**) group, **every column counting users**:
`Total Users` (distinct users of the group), `Median Lots per Cr` / `Range` (the
per-user statistics of §1.7, computed within the group), and how many **users** fall Below /
In / Above the range — so `Below + In + Above = Total Users` per row (only users with no
usable allocation carry no flag and fall outside the three columns).

The table is a **drill-down pivot**: clicking an (algo, type) row (DOR.html) or picking one
(dashboard) reveals the **outlier user ids** — user id, server, lots per Cr, lots, and whether
they sit below or above the range (below users first, worst first).

The sum of `Total Users` across algos can differ from the `Users` KPI: unmatched users (no MTM
entry) have no algo to sit under.

Report KPIs: `Users` = distinct user ids, `Orders` = Σ order counts, `Total Lots` = Σ lots,
`Trade Value` = Σ trade values.

## 1.9 Strikes traded (`Strikes` sheet + dashboard + DOR.html)

A **strike** is one option contract: **(index, expiry date, strike price, CE/PE)**. The same
contract circulates under two exchange symbol formats, so both are normalised to that one key
(otherwise a contract traded through different brokers/servers would count twice):

| Format | Example | Reading |
|---|---|---|
| `DDMMMYY` | `NIFTY21JUL2624100PE` | day 21, month JUL, year 26 → 21-Jul-2026, strike 24100, PE |
| `YYMDD` | `NIFTY2672124100PE` | year 26, month 7, day 21 → 21-Jul-2026 (Oct/Nov/Dec are the single letter `O`/`N`/`D`), strike 24100, PE |

Symbols that don't parse as a NIFTY / BANKNIFTY / SENSEX option (futures, equity) are skipped.
Both summaries are computed over the **deduped** orders:

- **Strikes per Algo / Server** — a **two-level drill-down**: the default view is per algo
  (algo | no. of servers | distinct strikes), and clicking an algo (or picking one in the
  dashboard) reveals its per-server counts. The algo comes from the same MTM allocation matching
  as the trade value rows (§1.6); an unmatched user (no MTM entry) inherits their **server's**
  algo — so a server doesn't split into a duplicate blank-algo row; the algo is blank only when
  the whole server has no MTM accounts. Counts are **distinct** contracts at every level — an
  algo's figure is distinct across its servers, and the "Total (distinct)" row across everything
  — never the column sum (two servers can trade the same strike). In the dashboard and DOR.html
  the table has its own expiry + index dropdown pair (default **All / All**).
- **Lots per Strike — option chain** — per contract, `lots = Σ |qty| / lot_size(trade date)` (the
  same lot math as §1.5), presented as an option chain: one row per strike price with
  **CE | Strike | PE** lots side by side. The dashboard and DOR.html put the chain behind an
  expiry + index dropdown pair; the Excel sheet stacks one chain block per (index, expiry).
  When the index's **day open / day close** are entered in the dashboard, the chain centres on
  the **ATM** — the traded strike nearest the day-mid `(open + close) / 2`, highlighted — and
  shows only **No. of strikes** above and below it (an input in the chain card, default 10;
  both the report and the dashboard have it). Without a day-mid the full chain shows, and its
  total ties back to the `Total Lots` KPI exactly. Lots are always shown as **whole numbers**
  (no decimals) — everywhere, including the trade value rows and KPIs.

  The chain also carries **trade-kind and algo filters** (default: everything included). The
  order's `Tag` marks the kind — `h_…` = hedge, `v_…` = VAR, `s_…` = square-off, anything else
  is a normal trade. Three toggles (Hedge / VAR / Sq-off) include or exclude those kinds —
  normal trades always count — and an Algo dropdown restricts the chain to one algo. Both the
  report and the dashboard have the same controls; the per-(algo, kind) split is embedded in
  the page so filtering never recomputes.

# Part 2 — Int / Pos Segregation (`segregate_int_pos_mtm2.py`)

## 2.1 Classification

Every account (row) of the Compiled User MTM is either:

- **Positional** — its User ID appears in **either** Combined Max Loss file (1DTE or 4DTE);
- **Intraday** — otherwise.

A user id on several servers is several **distinct accounts**: max-loss entries are assigned to
exactly **one** compiled row each — matched on server when the id is ambiguous, falling back to
the user's first row when the entry's server is blank/unmatched — so an addon is **never
double-counted**.

## 2.2 Addons and adjusted Realized P&L

From each max-loss file, per (User ID, Server) (duplicates: last row wins):

```
addon = Realized PNL + Net Settlement Value
```

Then per account:

```
Addon Applied = Addon 1DTE               if User Type == "Noren"
              = Addon 1DTE + Addon 4DTE  otherwise (non-Noren positional)
              = 0                        for intraday accounts

Realized P&L (Final) = compiled Realized P&L + Addon Applied     ("AdjRealized")
MTM                  = Realized P&L (Final) + Unrealized P&L
UserReturn           = Realized P&L (Final) / ALLOCATION         (0 when allocation is 0)
```

## 2.3 Pivot aggregation (per algo → Int / Pos+Int → server)

Within each algo the accounts are split into an **Int** block (intraday) and a **Pos+Int** block
(positional), each grouped by server. Per group:

```
No. of Users        = row count
No. of SL Hit Users = count of rows with SL HIT/NOT == 1
MAX LOSS            = Σ MAX LOSS
ALLOCATION          = Σ ALLOCATION
Realized P&L        = Σ Realized P&L (Final)
Unrealized P&L      = Σ Unrealized P&L
MTM                 = Realized + Unrealized
P&L %               = Realized / Allocation        (0 when allocation is 0)
95% (P95) / 5% (P5) = 95th / 5th percentile of UserReturn within the group,
                      computed only over accounts with ALLOCATION > 0
```

P5/P95 stand in for min/max **deliberately** — they control for data aberrations (intraday
pauses/stops, under/over-funded accounts, suboptimal allocations) that would make the true
min/max misleading.

**Server counts:** the Sub-Total row's server count is per section; the algo total counts each
**distinct** server once (a server running both Int and Pos appears in both blocks but is counted
once for the algo); the Grand Total sums the per-algo counts.

**KPI tiles:** `Accounts` = all rows, `Positional` = classified positional, `Intraday` =
accounts − positional, plus `Allocation` (× 100, Cr), `Realized P&L` (Cr) and `P&L %` from the
grand total. Tiles are centered and flow in a single row.

**Algo selection:** the dashboard inputs carry a "Segregation — algos to include" multiselect
(empty = all). The pivot and its KPI tiles — in the dashboard **and** in DOR.html — cover only
the selected algos; the Excel `Segregation` / `Raw_Data_Per_User` / `Worst 10%ile` sheets always
keep every algo (the workbook is the full audit trail).

**Displayed columns:** the dashboard and DOR.html pivot show a trimmed set —
`Algo · Server · Users · SL Hit · Max Loss · Allocation · Realized P&L · P&L %` (plus the
Type/Section row label; the ratio field is titled **P&L %** everywhere, including the Excel
header). `Unrealized P&L`, `MTM` and `P95/P5` stay in the Excel Segregation sheet, where the
full column set and live formulas remain.

**Compact money display:** `Allocation` (× 100 — the stored value is in hundreds) and
`Realized P&L` render in compact Indian units (`403.8 Cr`, `53.5 L`) in the pivot and the KPI
cards; the Trade Value KPI uses the same format. `P&L %` and every calculation keep the stored
basis; the Excel sheet keeps full stored values.

## 2.4 Excel formulas (Segregation sheet)

Aggregate cells are written as live formulas so the sheet stays auditable in Excel:
Sub-Totals use `=SUM(...)` over their data block, Algo/Grand Totals add the sub-total rows,
and every `P&L %` cell is `=IF(G{row}=0,0,H{row}/G{row})` (Realized ÷ Allocation).
P5/P95 are written as computed values. The Realized / Unrealized / MTM columns carry
conditional formatting — profit green, loss red — that keeps recolouring as the formulas
recalculate.

## 2.5 Raw_Data_Per_User sheet

One row per account with the full audit trail: Type, User Type, algo, server, id, alias, SL Hit
(0/1), MAX LOSS, ALLOCATION, Compiled Realized P&L, Addon 4DTE, Addon 1DTE, Addon Applied,
Realized P&L (Final), Unrealized P&L, and `MTM = Final Realized + Unrealized`.

## 2.6 Worst 10%ile sheet

Per **Algo × Int/Pos** group:

- Rank accounts by `UserReturn`; keep those **at or below the group's 10th percentile**.
  Only accounts with `ALLOCATION > 0` are ranked (zero-allocation returns are undefined).
- `Algo Avg Return %` shown next to each user = the **group-level** return
  `Σ Realized (Final) / Σ Allocation` — the same value as that section's sub-total P&L %.
- `Reason` is auto-inferred (first match wins, editable afterwards in Excel):
  1. SL Hit == 1 → **"SL Hit"**
  2. allocation ≤ 0 → **"Zero/low allocation"**
  3. return < 0 → **"Negative return"**
  4. otherwise → **"Low relative return"**

## 2.7 Slippage (`Slippage` sheet)

Evaluated over **all** accounts with `ALLOCATION > 0`, both limits expressed as plain **ratios**
of the account's allocation (positive numbers = loss; the same unit convention as the pivot's
`P&L %` — **not** multiplied by 100):

```
ML %          = MAX LOSS / ALLOCATION
Realized ML % = |compiled Realized P&L| / ALLOCATION
```

An account **has slippage** only when its realized loss overshoots the configured max-loss by
**at least 0.1**:

```
Realized ML % − ML % ≥ 0.1
```

so ML% 1.00 → Realized 1.09 is *not* slippage, 1.10 is; a profit or a loss inside the limit is
never slippage. The **compiled** Realized P&L is used (no positional addons). Per-order slippage
(trigger vs fill) is not derivable from the orderbook: it carries no stop-loss orders.

- **Avg slippage** — per algo + overall: total accounts, slippage accounts, and
  `Avg Slippage % = mean Realized ML %` over the algo's **slippage accounts** ("—" when none).
- **Major slippages** — the slippage accounts whose `Realized ML %` is **greater than their
  algo's Avg Slippage %**, sorted worst-first. The Excel sheet lists **all** slippage accounts
  with the major rows highlighted; the dashboard and DOR.html show the major table alongside
  the per-algo summary.
- **Algo filter** — the algos to analyse are picked in the **dashboard inputs** (multiselect,
  empty = all); both the dashboard section and the DOR.html report show the selection, with the
  Overall row computed over it.
- A day with no slippage accounts shows **"-- No slippage today --"** in all three outputs.

## 2.8 Report date

Inferred from the first non-empty `Date` value in the compiled MTM sheet, formatted `dd-mm-YYYY`.

---

# Part 3 — DOR.html (`dor.py`)

Pure presentation — no new math. It embeds the already-computed numbers into a single
self-contained HTML file: KPI tiles, the per-algo trade-value outlier summary as a drill-down
(§1.8 — algo rows expand into their outlier user ids), the two strikes tables (§1.9, side by
side, algo → server drill; the chain has an ATM window with its No.-of-strikes input), the
slippage summary and major slippages for the algos selected at input time (§2.7), the
portfolio analysis (§4), and the segregation pivot as a click-to-drill-down table
(Algo → Int / Pos+Int → servers → users; the user level shows id, alias, SL hit, max loss,
allocation, realized P&L and P&L %, worst realized first).

Presentation rules: every table is center-aligned — data and headers — and every number uses
**Indian digit grouping** (16,44,536) for the client-facing look. Realized P&L / Unrealized
P&L / MTM values are coloured by sign (profit green, loss red) in the pivot and the KPI tiles.
In the side-by-side cards (strikes, slippage) the card title and the expiry/index dropdowns stay
fixed and only the table body scrolls, with the sticky header flush at the scroll top. The
strikes/chain dropdown filters are precomputed lookups embedded in the page — no recalculation
in the browser. The masthead carries a **Download Excel** button with the full workbook embedded
in the page (base64 data URI — the HTML stays a single self-contained file); row-level data
otherwise stays out of the HTML.

---

# Part 4 — Portfolio Analysis (`portfolio.py`)

Driven by the optional **Multileg Orders (MLOB)** input; shown just above the Strikes section.

## 4.1 Algo column

The MLOB has no algo, so one is derived first: each (User ID, Server) is matched against the
Compiled User MTM exactly like the trade value rows (§1.6); users with no MTM entry inherit
their **server's** algo (same fallback as the strikes table, §1.9).

## 4.2 PnL

Only `Status == COMPLETE` rows count. Per portfolio (and identically per user):

```
sell value = Σ Avg Price × Filled Quantity     over SELL rows
buy value  = Σ Avg Price × Filled Quantity     over BUY rows
PnL        = sell value − buy value
```

## 4.3 Portfolio Summary table

The default **"QS" Portfolio Analysis** covers every portfolio whose `Portfolio Name` contains
`QS` (e.g. `MTWTF_SN_ODTE_1100_WT3%_QS`, `MTWTF_SN_ODTE_QS0_1100_WT3%`). It is a
**three-level drill-down** — the compact algo view is the default, and each level expands
(click, like the segregation pivot) into the next:

```
1. Algo view    algo | no. of servers | users | portfolios | orders | PnL
2. Server view  the algo's servers, same metrics per server
3. User view    the server's users — user id | portfolios | orders | PnL (worst PnL first)
```

with per level:

```
Portfolio Executed Users = distinct users that executed a matching portfolio
                           (at algo level a user on two servers counts once)
No. of Portfolio         = distinct matching portfolio names
Total Orders             = COMPLETE entries (rows) of those portfolios
PnL                      = Σ sell value − Σ buy value
```

The Total row counts distinct users / portfolios over the whole selection. PnL cells are
coloured by sign. The dashboard shows the same three views through algo / server pickers.

## 4.4 Analyse any other portfolio — in the HTML itself

Below the QS table, DOR.html has a **combobox**: a dropdown of every portfolio name that
narrows as you type (`y` → only names containing `y`, `ya` → only `ya`, …). Pick a name to
analyse that portfolio, or type any fragment (e.g. `WT3%`) to aggregate every matching
portfolio — either way the same report renders instantly (the per-(algo, server, user,
portfolio) groups are embedded in the page, so no reprocessing is involved). The dashboard has
the same pair: a type-to-filter name picker plus a fragment box. Outputs: the QS analysis is also written to the workbook as the
**`Portfolio QS`** sheet (summary rows with the user detail underneath, Total at the end).

---

## Notes

- `app.py` caches processing on the file contents + deviation `k` + report schema, and warns when
  the uploaded inputs changed since the last **Process** click. The drill-down picks, the strikes
  filters, the slippage algo selection and the portfolio patterns are all applied at render time
  from the cached results — none of them requires a reprocess.
- Trade value math uses `Decimal` end to end; segregation uses pandas floats with money rounded
  to whole rupees in the outputs.
- Lots are always displayed as whole numbers (no decimals) across all outputs; `Lots per Cr`,
  `P&L %` and `Slippage %` keep 2 decimals. Displayed numbers use Indian digit grouping
  (`16,44,536`) everywhere — the dashboard and DOR.html format directly, and the Excel money
  cells carry an Indian-grouping number format (the cell values stay plain numbers).
