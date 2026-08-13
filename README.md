# Daily Operations Report (DOR)

A Streamlit app (plus standalone CLIs) that turns four raw inputs into the daily ops report:

1. **Trade Value analysis** (§1) — per-user lots / trade value from the orderbook, with
   statistical outlier flagging on per-user **Lots per Cr** (split Int / Pos+Int), the
   per-(algo, type) summary (a drill-down that reveals the outlier user ids), and the
   strikes-traded views (algo → server counts + a CE/Strike/PE option chain with ATM
   centring and hedge/VAR/sq-off/algo filters).
2. **Intraday / Positional segregation** (§2) — a pivot of MTM per algo →
   Int / Pos+Int → server, classified from the **All User** sheet's `Running Type` and
   narrowed by the **DTE** selected in the sidebar, plus the **slippage** view (accounts
   whose realized loss / allocation overshoots their max-loss / allocation by ≥ 0.1).
   `MTM = Realized P&L + Unrealized P&L`, always computed.

3. **Portfolio analysis** (§4, optional — needs the Multileg Orders / MLOB input) — per
   (algo, server) portfolio summary with per-user drill-down for the "QS" portfolios by
   default, plus instant any-portfolio analysis inside the HTML report itself.

Outputs: one Excel workbook `DOR_<DTE>_<date>.xlsx` (sheets: `unclassified`, `tradevalue`,
`summary`, `Strikes`, `Orders`, `Portfolio QS` when the MLOB is given, `Summary`,
`MTM Data`, `Slippage`, `no_sl_Acc`) and one self-contained, client-shareable `DOR_<date>.html` (all tables
center-aligned, data and headers; numbers in Indian digit grouping, e.g. `16,44,536`).
When a **secondary index User MTM** is uploaded (§below) the workbook carries a second set —
`summary 2`, `Summary 2`, `MTM Data 2`, `Slippage 2` — and the
dashboard + DOR.html show the segregation cards/pivot and the Algo Summary once **per MTM file**.

## Files

| File | Role |
|---|---|
| `app.py` | Streamlit UI — upload the inputs, set the outlier deviation, one **Process** click computes everything. Run with `streamlit run app.py`. |
| `tradevalue.py` | Trade value engine: report rows, algo summary, strikes/option chain (also a CLI: `python tradevalue.py orderbook.csv -s user_mtm.xlsx -d 1`). |
| `summary.py` | Summary engine: All User classification, DTE scope, MTM, pivot, `MTM Data`, slippage, `no_sl_Acc` (also an interactive CLI). |
| `marketdata.py` | Index day High/Low + intraday series (yfinance), premium-file parsing. |
| `user_aliases.json` | All User id ↔ MTM id, when the two files name an account differently. |
| `account_aliases.json` | Orderbook base id → MTM account id, per server. |
| `CALCULATION.md` | Every formula and constant, with the reasoning. |
| `portfolio.py` | Portfolio analysis engine over the Multileg Orders (MLOB): per-portfolio / per-user PnL, QS summary, any-pattern reports. |
| `dor.py` | Renders the DOR.html summary (inline CSS/SVG/JS, no external assets — dropdown filters, the drill-downs and the in-page portfolio analysis are plain inline JS). |

## Inputs

| Input | Used for |
|---|---|
| Compiled Orderbook (CSV/Excel) | Trade value rows |
| Compiled User MTM (CSV/Excel) | Allocations + algo for trade value; the account universe for segregation |
| Secondary index User MTM (optional, CSV/Excel) | When two indexes ran on **different servers** (e.g. 8 NIFTY servers + 2 BANKNIFTY servers) the orderbook is combined but the User MTM comes as two files — the second file's accounts join the allocation matching, and its servers get their **own** segregation cards/pivot, Algo Summary and outlier (median ± MAD) bands, so one index's Lots-per-Cr scale never corrupts the other's |
| All User Details (Excel, tab `Main`) — **required** | Int / Pos+Int classification via `Running Type`, DTE scope via `Running Days` |
| ATM premium file per index (optional) | The premium line on the intraday chart |
| Multileg Orders — MLOB (optional, Excel/CSV) | The Portfolio Analysis (§4) — omitting it just hides that section |
| Outlier deviation `k` (default 1.0) | Width of the Lots-per-Cr range (median ± k × MAD, per algo + type group) — drives the flags and the Algo Summary Range |
| NIFTY / SENSEX / BANKNIFTY day open + day close | Day-mid = (open + close) / 2 — the option chain's ATM strike (§1.9); 0 = not set |
| NIFTY / SENSEX / BANKNIFTY expiry (text, e.g. `28JUL26`) | The expiry label of each index's strikes (§1.9). One orderbook holds a **single expiry per index**, and the symbol formats are too ambiguous to parse a date from — so the expiry is entered here, and only the strike is read from the symbol. Left empty → the chain shows `NA` |
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

> ### 🔑 Orderbook dedup key
> ```
> User ID + Order ID + Order Time + Exchg Order ID + Exchange Time + Tag
> ```
> The MLOB uses a **different** key — see §4.0.

Duplicate orders are dropped on the key
**User ID + Order ID + Order Time + Exchg Order ID + Exchange Time + Tag**, keeping the
occurrence with the **lowest row id / SNO**.

Six components rather than four because a compiled orderbook that has been through Excel
cannot be keyed on Order ID alone — in the 11-08-2026 file **21% of Order IDs came back in
scientific notation** (`2.60811E+13`), which collapses thousands of distinct orders onto one
value. The exchange fields survive Excel intact, so the key stays discriminating and **every
row can be keyed** (no bypass path). A corrupted Order ID is logged as a warning; the source
file should be exported with that column as text.

Measured on 11-08: the key collapses 2,28,704 rows, of which 2,28,719 are byte-identical
duplicates — it removes duplicates and essentially nothing else.

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

### What the 1.4826 is

**2.54 converts inches to centimetres. 1.4826 converts MAD to standard deviation.** It is a
unit conversion between two rulers for the same thing, nothing more.

MAD is literally *"the middle distance from the middle"* — half the users are closer than
it, half further — so **MAD covers 50% of the data by definition**, while one standard
deviation covers **68%**. A MAD reading therefore has to be scaled up before it can be read
as a std-dev. For bell-shaped data that factor is always ~1.4826, whatever the units
(σ=20 → ratio 1.4915; σ=3 → 1.4812).

Formally it is `1 / Φ⁻¹(0.75)` = 1/0.6745, and it is **derived in code, not typed**:
`Decimal(str(1 / NormalDist().inv_cdf(0.75)))`. It is the standard convention — R's `mad()`
default and scipy's `median_abs_deviation(scale="normal")` use the same number.

**Without it the band would be a third too narrow.** On 05-08 at k=1: **175 users flagged
without, 118 with** — 57 ordinary users wrongly called outliers, and `k=1` would secretly
mean 0.67.

**Caveat:** the 0.6745 relationship holds for a *normal* distribution only. On flat data the
ratio is ~1.05, not 1.48. Where an algo's users are near-clones the band is legitimately
tiny (algo 8: ±1.9 on a median of 250) and is better read as plain *distance from the
median* than as a true sigma.

- `lots_per_cr < median − k·MAD` → **"Below average range"**
- `lots_per_cr > median + k·MAD` → **"Above average range"**
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

A **strike** is one option contract: **(index, strike price, CE/PE)**. The same contract
circulates under several exchange symbol formats, so all of them must normalise to that one key
(otherwise a contract traded through different brokers/servers would count twice):

| Format | Example | Strike read |
|---|---|---|
| compact `DDMMMYY` | `NIFTY21JUL2624100PE` | 24100 PE |
| compact `YYMDD` | `NIFTY2672124100PE` | 24100 PE (Oct/Nov/Dec are the single letter `O`/`N`/`D`) |
| compact monthly (no day) | `NIFTY26JUL24100PE` | 24100 PE |
| spaced, 4-digit year | `NIFTY 21JUL2026 PE 24100` | 24100 PE |
| spaced, 2-digit year | `NIFTY 21JUL26 24100 PE` | 24100 PE |

The **expiry is deliberately not parsed** — the compact forms are ambiguous
(`NIFTY26JUL22600CE` reads as 26-Jul-2022 strike 600 just as well as Jul-2026 strike 22600),
and one orderbook holds a **single expiry per index** anyway. So each index's expiry is
**entered in the dashboard** and applied as a display label, and only the strike is read from
the symbol — by **value**: index strikes sit in a known band (NIFTY 8k–60k, BANKNIFTY
20k–100k, SENSEX 40k–200k — set far wider than today's levels so drift never breaks parsing),
so the tail digits that read as a plausible level are the strike (tried 5-digit first, then
6 and 4). Symbols that don't parse as a NIFTY / BANKNIFTY / SENSEX option (futures, equity)
are skipped. Both summaries are computed over the **deduped** orders:

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
  expiry + index dropdown pair (the expiry labels are the ones entered in the inputs; an index
  with no entry shows `NA`); the Excel sheet stacks one chain block per index, captioned with
  its entered expiry.
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

# Part 2 — Int / Pos Segregation (`summary.py`)

## 2.1 Classification — from the All User sheet

Classification comes from the **All User Details** workbook, tab `Main`. Six columns are
read: `userId`, `server`, `algo`, `max_loss`, `Running Type`, `Running Days`.

**Step 1 — drop on upload.** A row is dropped when **any** of `server`, `Running Type` or
`Running Days` reads `DLR ACC` or `NOT RUNNING`. Dealer and stopped accounts belong to no
report, whatever the DTE.

**Step 2 — DTE scope.** `Running Days` states the days an account *runs on*, so the scopes
are **cumulative**:

| DTE | Running Days included |
|---|---|
| 0DTE | all — `0DTE` + `1DTE/0DTE` + `DAILY` |
| 1DTE | `1DTE/0DTE` **+ `DAILY`** |
| 4DTE | `DAILY` |

A `DAILY` account trades every day, including 1DTE days.

**Step 3 — type.** `INT` → **Int**, `POS` → **Pos+Int**. Matching is on the canonical user
id alone (`userId` is unique in the sheet). An account the scope does not cover is
**Unclassified** and reported in its own section at the bottom of the pivot, so each Algo
Total covers only that algo's classified accounts and
`Grand Total = Σ Algo Totals + Unclassified`.

> The Combined Max Loss files, the realized-P&L addon, the Noren rule and the hardcoded
> positional algos (19/27) were **removed** — classification is now entirely from
> `Running Type`.

## 2.2 MTM

```
MTM   = Realized P&L + Unrealized P&L      per account, ALWAYS computed
MTM % = MTM / ALLOCATION
```

The compiled sheet ships its own `MTM` column and it is **ignored** — it is only populated
where Unrealized is non-zero (577 of 633 rows blank on 30-07-2026, understating the book by
₹4.17 Cr). A dashboard note reports what was measured: rows differing, blank vs wrong value,
and the rupee gap.

MAX LOSS comes from the compiled MTM's own column; the All User `max_loss` is carried as a
reference column only.

## 2.3 Pivot aggregation (per algo → Int / Pos+Int → server)

Per block: `Users`, `SL Hit`, `MAX LOSS`, `ALLOCATION`, `Realized P&L`, `Unrealized P&L`,
`MTM`, `MTM %`. Allocation is displayed **×100** (stored in hundreds); `MTM %` uses the
stored basis. In Excel `MTM %` is the live formula `=IF(G=0,0,J/G)` so it survives editing.

The P95/P5 percentile columns and the **Worst 10%ile** sheet were removed — neither appeared
anywhere in the HTML report.

## 2.4 Slippage (`Slippage` sheet)

Plain ratios of allocation, not multiplied by 100:

```
ML %          = MAX LOSS / ALLOCATION
Realized ML % = -Realized P&L / ALLOCATION       (positive = loss)
slippage when   Realized ML % - ML % >= 0.1      (1e-9 float guard)
```

Measured on **realized** loss, not MTM, because a max-loss stop is about what was booked.

**Eligibility:** allocation > 0 **and** MAX LOSS > 0. Accounts with an allocation but no
configured stop would have ML % = 0, so any loss past the threshold would read as slippage
against a limit that never existed — they are listed in the **`no_sl_Acc`** sheet and
excluded from the analysis and from the account counts.

**Major slippage** = Realized ML % above that algo's average, **or** the algo has exactly
one slippage account (with one account it *is* the average, so a strict `>` could never
fire).

## 2.5 `MTM Data` sheet

The per-account base table every summary figure aggregates, replacing the old
`Raw_Data_Per_User`:

`UserID · Alias · ALLOCATION · MAX LOSS · Total Orders · Total Lots · SERVER · ALGO ·
Running Type · Running Days · OPERATOR · EXPIRY · Date · Month · Day · INDEX · MTM ·
Realized P&L · Unrealized P&L · SL HIT/NOT`

## 2.6 `unclassified` sheet

First sheet in the workbook: `UserID · Alias · ALGO · SERVER · Running Type · Running Days`
for every account the DTE scope did not cover. Running Type / Running Days read
**`Not Found`** when the row was dropped on upload or is absent from the sheet — a dropped
row is gone and is not resurrected.

## 2.7 Report date

Read from the compiled MTM's `Date` column and pre-filled into the **Market date** input,
which stays editable. It drives the index High/Low fetch, the chain's ATM anchor and the
chart.


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

## 4.0 Where the MLOB is used — and where it is not

**The Multileg Orders sheet feeds the Portfolio Analysis only.** No other figure in the
report touches it:

| Calculation | Source |
|---|---|
| Portfolio summary / `Portfolio QS` sheet / in-HTML any-pattern analysis | **MLOB** |
| Trade Value, lots, Lots per Cr, outliers | orderbook |
| Strikes / option chain | orderbook |
| Orders Summary | orderbook |
| Intraday chart (index, premium, lot dots) | orderbook + market data |
| Segregation pivot, MTM, MTM % | User MTM + All User |
| Slippage, `no_sl_Acc` | User MTM |

Omit the MLOB and the Portfolio section simply disappears; nothing else changes.

**Columns used** (7): `User ID`, `Portfolio Name`, `Transaction`, `Avg Price`,
`Filled Quantity`, `Status`, `Server`. **Read only for the dedup key** (3): `Date`,
`Order ID`, `Symbol`. Everything else (`Leg ID`, `Quantity`, `Tag`, `Remarks`, …) is ignored.

> ### 🔑 MLOB dedup key — deliberately NOT the orderbook key
> ```
> User ID + Date + Order ID + Symbol + Portfolio Name
> ```
> A multileg Order ID is **reused across legs**: on 05-08, 4,504 order ids spanned more than
> one portfolio and 2,213 carried a BUY leg *and* a SELL leg. The orderbook key would merge
> the two sides and corrupt PnL, which is `sell − buy`.

`Filled Quantity` is used rather than `Quantity`, so PnL reflects what actually traded.
Lots are **not** derived from the MLOB — those come from the orderbook.


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
