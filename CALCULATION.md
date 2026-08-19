# DOR — Calculation Reference

Every number the Daily Operations Report produces, the formula behind it, and the
reason for every fixed constant.

Written to be auditable: if a figure in the report looks wrong, this document should
tell you which formula produced it and what it assumes.

---

## Contents

1. [Inputs](#1-inputs)
2. [Identity matching — how a user is recognised across files](#2-identity-matching)
3. [Deduplication](#3-deduplication)
4. [Lot sizes](#4-lot-sizes)
5. [Trade Value](#5-trade-value)
6. [Outlier detection (Lots per Cr)](#6-outlier-detection)
7. [Strike parsing and the option chain](#7-strike-parsing-and-the-option-chain)
8. [Account classification — DTE scope and Running Type](#8-account-classification)
9. [MTM](#9-mtm)
10. [Slippage](#10-slippage)
11. [Orders Summary](#11-orders-summary)
12. [Intraday chart](#12-intraday-chart)
13. [Portfolio analysis](#13-portfolio-analysis)
14. [Index market data](#14-index-market-data)
15. [Constants — the complete list](#15-constants--the-complete-list)
16. [Reconciliation identities](#16-reconciliation-identities)

---

## 1. Inputs

| File | Required | Used for |
|---|---|---|
| Compiled Orderbook (CSV/Excel) | yes | Trade Value, lots, strikes, Orders Summary, chart dots |
| Compiled User MTM (Excel `Sheet1` / CSV) | yes | allocation, algo, MAX LOSS, Realized & Unrealized P&L, report date |
| All User Details (Excel, tab `Main`) | yes | Int / Pos+Int classification |
| Secondary User MTM | optional | a second index set running on its own servers |
| Multileg Orders (MLOB) | optional | Portfolio analysis |
| ATM premium, one per index | optional | the premium line on the chart |
| `aliases.json` | optional | both id maps: All User id ↔ MTM id (string value), and orderbook base id → MTM account id per server (object value) |

**Orderbook row filters**, applied in this order:

1. `Exchange ∈ {BFO, NFO}` — drops cash rows (NSE/BSE) whose quantities are not lot multiples
2. Symbol classifies to NIFTY / BANKNIFTY / SENSEX — anything else is ignored
3. `Status == COMPLETE` — **only** for Trade Value and strikes. The Orders Summary and the
   chart read every status.

---

## 2. Identity matching

The same account is written differently in each file, so every join runs on a canonical key.

### 2.1 `_user_key` — the canonical user id

```
key = str(value).strip().upper()
if key looks like "<digits>.<zeros>":  key = digits        # Excel float artifact
if key.isdigit():                      key = key.lstrip("0") or "0"
```

Two corrections, both caused by Excel:

| Problem | Example | Why |
|---|---|---|
| Leading zeros dropped | `04101961` → `4101961` | Excel types an all-digit id as a number |
| Rendered as a float | `6954037.0` → `6954037` | a numeric column saved and reopened |

Without these, the same account keys differently in two files and silently loses its algo.

### 2.2 `_server_key`

Upper-cased and trimmed; `NAN` / `NONE` / `NA` all read as blank. A server label is only
ever compared to another server label, never parsed.

### 2.3 Base id vs account id — `aliases.json` (object entries)

The orderbook records `TB2433`; the MTM records `TB2433A41`. The same base id is a
**different account on each server**, so the map is keyed on both:

```json
"TB2433": { "VS8": "TB2433A41", "VS29": "TB2433A42" }
```

Applied **before** inference. `add_user_aliases()` can also infer the link by prefix when
exactly one MTM account on that server extends the orderbook id — but inference is a guess
that stops working the day two accounts share a prefix, and then the orders fall into the
unattributed `—` row. Anything listed in the file is never guessed.

Stripping a trailing `A<digits>` would be unsafe: `R7RA1315` would reduce to `R7R` and
collide with real ids. That is why the match is by prefix within a server, or explicit.

### 2.4 Different id entirely — `aliases.json` (string entries)

Some accounts share no prefix at all: the All User sheet says `CC04`, the MTM says
`XLDH142`. Three of the seven known cases carry no trace of the other id anywhere in the
row (`A19_CC03`, `CC_A8_CC10_4C`, `A27_MTPRO`), so the alias column cannot be parsed and
the mapping has to be stated:

```json
{ "CC04": "XLDH142" }        // All User id : MTM id
```

Applied before the row's own id, because the MTM id can itself exist in the All User sheet
as a dropped `DLR ACC` row.

### 2.5 Allocation lookup

`_pick_allocation(allocations, user_key, server)`

- one MTM row for the user → use it
- several rows (the same id on several servers = distinct accounts) → match on the
  orderbook row's server
- no server matches → **unmatched**, blank allocation

Never guesses between accounts.

---

## 3. Deduplication

> ### 🔑 The two dedup keys — at a glance
>
> **Orderbook**
> ```
> User ID + Order ID + Order Time + Exchg Order ID + Exchange Time + Tag
> ```
>
> **Multileg orders (MLOB)** — deliberately different
> ```
> User ID + Date + Order ID + Symbol + Portfolio Name
> ```
>
> They differ because a multileg Order ID is **reused across legs**. Applying the
> orderbook key to the MLOB would merge a BUY leg with a SELL leg and corrupt PnL.

A compiled orderbook repeats rows — on 05-08-2026, **67% of rows** were part of an exact
duplicate pair.

**Key:** `User ID + Order ID + Order Time + Exchg Order ID + Exchange Time + Tag`

Six components, because a compiled orderbook that has been through Excel cannot be keyed on
Order ID alone. In the 11-08-2026 file **21% of Order IDs came back in scientific
notation** (`2.60811E+13`), which collapses thousands of distinct orders onto one value.

The exchange fields survive Excel intact, so including `Exchg Order ID` and both timestamps
keeps the key discriminating even when Order ID is destroyed. Every row can therefore be
keyed — there is no bypass path.

Lowest `SNO` wins. Measured on 11-08: the key collapses 2,28,704 rows, of which 2,28,719
are byte-identical duplicates — it removes duplicates and essentially nothing else. Exactly
**one** group in the file merged rows that differ, and those were four REJECTED rows at
price 0 with blank or non-index symbols, which the exchange filter discards anyway.

An Order ID in scientific notation is still **logged as a warning**: the parse survives it,
but the source file should be exported with that column as text.

**Why the MLOB key differs** — `(User ID, Date, Order ID, Symbol, Portfolio Name)`.
A multileg Order ID is reused across legs — on 05-08, 4,504 order ids spanned more than
one portfolio and 2,213 carried a BUY leg *and* a SELL leg. The orderbook key would merge
the two sides and corrupt PnL, which is `sell − buy`.

---

## 4. Lot sizes

`lots = |quantity| / lot size`, with the lot size resolved **by trade date** because the
exchanges revised them mid-history.

### NIFTY

| Period | Lot |
|---|---|
| up to 26-Dec-2024 | 25 |
| 27-Dec-2024 → 30-Dec-2025 | 75 |
| 29 & 30-Jan-2025 (exception inside that window) | 25 |
| from 31-Dec-2025 | 65 |

### SENSEX

| Period | Lot |
|---|---|
| up to 03-Jan-2025 | 10 |
| 30 & 31-Jan-2025 (exception) | 10 |
| otherwise | 20 |

### BANKNIFTY

| Period | Lot |
|---|---|
| up to 26-Feb-2025 | 15 |
| 27-Feb-2025 → 26-Jun-2025 | 30 |
| 27-Jun-2025 → 30-Dec-2025 | 35 |
| from 31-Dec-2025 | 30 |

An unparseable date falls back to the current lot size. **These tables are hardcoded and
must be extended when an exchange revises a lot size** — nothing detects that
automatically, and a wrong lot size scales lots and Lots per Cr proportionally.

---

## 5. Trade Value

Grouped per `(date, server, user, index)`:

```
order count = number of deduped COMPLETE orders
quantity    = Σ quantity                       (signed)
lots        = Σ |quantity| / lot size
trade value = Σ avg price × quantity           (signed, not absolute)
```

`trade value` keeps its sign, so buys and sells offset. It is a turnover figure, not a P&L.

### Normalise and Lots per Cr

```
normalise   = allocation / 100,000
lots per Cr = lots / normalise
```

**Why 1,00,000.** The stored allocation is written in hundreds — a displayed ₹1 Cr is
stored as 1,00,000. Dividing by that base makes 1 unit of normalise = 1 Cr of allocation,
so Lots per Cr is comparable across account sizes. Sub-crore accounts normalise
fractionally: 80,000 → 0.8, 20,000 → 0.2.

An account with no usable allocation has no normalise, no Lots per Cr, and no outlier flag.

---

## 6. Outlier detection

The unit judged is the **user**, not the report row. A row is per (user, index), so a user
trading two indexes would otherwise be judged twice on partial exposures. Their lots are
summed across servers and indexes inside the algo, then normalised by **total** allocation.

### The band

```
median = middle per-user lots per Cr within (date, algo, type)
MAD    = 1.4826 × median( |value − median| )
band   = [ median − k·MAD , median + k·MAD ]        k = "Outlier deviation", default 1
```

Below the band → *Below average range*; above → *Above average range*; else *In range*.

### Why MAD and not standard deviation

Mean and standard deviation are destroyed by the very outliers they are meant to find.
Measured on synthetic data:

| Contaminated | Mean | Std dev | Median | Scaled MAD |
|---|---|---|---|---|
| 0% | 100.0 | 9.4 | 100.7 | 9.1 |
| 5% | 5,094.9 | 21,772.7 | 100.8 | 10.2 |
| 40% | 40,060.1 | 48,940.7 | 109.4 | 27.6 |

Median and MAD have a **50% breakdown point** — up to half the data can be arbitrary
before they move. On real algo-8 data, inflating one user 50× widened the std-dev band by
**2,791%** and the MAD band by **0%**.

### Why 1.4826 — the short version

**2.54 converts inches to centimetres. 1.4826 converts MAD to standard deviation.**

It is a **unit conversion between two rulers for the same thing**, nothing more.

MAD is literally *"the middle distance from the middle"*: half the users are closer than
it, half are further. So **MAD covers 50% of the data by definition**. One standard
deviation covers **68%**. Same spread, two scales — so a MAD reading has to be scaled up
before it can be read as a std-dev.

For bell-shaped data that factor is always ~1.4826, whatever the units:

| Data | MAD | Std dev | ratio |
|---|---|---|---|
| bell-shaped, σ = 20 | 13.42 | 20.02 | **1.4915** |
| bell-shaped, σ = 3 | 2.03 | 3.00 | **1.4812** |

### What it changes in the report

Without the conversion the band is a third too narrow. On 05-08 at `k = 1`:

| | Users flagged |
|---|---|
| **Without** 1.4826 | **175** |
| **With** 1.4826 | **118** |

57 perfectly ordinary users would be called outliers (algo 7 alone goes from 11 to 27).
Put plainly: **`k = 1` would secretly mean 0.67**, covering 50% of users instead of 68%.
The constant makes the deviation input mean what a reader expects.

### The derivation, for the record

```
MAD is the m where P(|x − μ| ≤ m) = 0.50
2·Φ(m/σ) − 1 = 0.50
Φ(m/σ)       = 0.75
m/σ          = Φ⁻¹(0.75) = 0.6745
→ MAD = 0.6745 σ  →  σ = 1.4826 × MAD
```

Standard convention: R's `mad()` default is 1.4826 and scipy's
`median_abs_deviation(scale="normal")` divides by the same 0.6744897501960817.

**Derived in code, not typed:** `Decimal(str(1 / NormalDist().inv_cdf(0.75)))`.

### The assumption it carries — and where it does not hold

The 0.6745 relationship is a property of the **normal distribution only**. On data that is
not bell-shaped the ratio is different, so the conversion is approximate:

| Data | MAD | Std dev | ratio |
|---|---|---|---|
| flat / uniform `10,20,…,110` | 30.00 | 31.62 | **1.0541** ← not 1.48 |

This matters in practice. Where an algo's users are near-clones the distribution is nothing
like a bell curve, the band is legitimately tiny, and "std-dev equivalent" is a loose
description — read such a band as plain *distance from the median*. Observed band widths on
05-08:

| Algo | n | Median | Band | as % of median |
|---|---|---|---|---|
| 1 | 160 | 68.0 | ±21.7 | ±31.98% |
| 8 | 84 | 250.0 | ±1.9 | **±0.74%** |

One global `k` therefore gives algo 1 a ±32% tolerance and algo 8 a ±0.74% one. This is a
known limitation, not a defect.

### Per-index re-judging

Lot sizes differ per index, so Lots per Cr sits on a different scale in each. Bands are
recomputed **within each index**. Trade-off: a user trading two indexes appears in both
tables, and in each their Lots per Cr is that index's lots over their **whole** allocation
— a partial-exposure figure. Such users are flagged `⚠ also trades …`.

---

## 7. Strike parsing and the option chain

The same contract circulates under several symbol formats, all of which must normalise to
one key or a contract would count as two strikes.

**The index must be the symbol's prefix** — a substring test alone would let FINNIFTY and
MIDCPNIFTY through as NIFTY.

**The expiry is deliberately not parsed.** One orderbook covers one session, in which every
symbol of an index shares one expiry; that date is entered in the dashboard. This removes
the date ambiguity of compact forms (`NIFTY26JUL…` is July-2026 monthly, not 26-Jul).

**The strike is found by value.** In compact forms the tail digits before `CE`/`PE` are
tried at width 5, then 6, then 4, and accepted only if they fall in the index's band.

**The band is derived from the day's own fetched High / Low — nothing is hardcoded:**

```
band = [ day Low × 0.5 , day High × 1.5 ]
```

An index drifts, so a fixed rupee range silently starts rejecting real strikes as it does.
The day's own range is the correct reference because every strike traded that day sits near
it. Example bands on 11-08-2026:

| Index | Derived band |
|---|---|
| NIFTY | 12,215 – 36,865 |
| BANKNIFTY | 28,579 – 86,411 |
| SENSEX | 39,024 – 117,765 |

The window is deliberately generous: it exists only to reject a **mis-parse** (reading
`1124050` out of `NIFTY2681124050PE` instead of `24050`), not to judge whether a strike is
sensibly priced. A mis-parse is out by an order of magnitude, so nothing this wide lets one
through. Verified: **0 of 181 distinct symbols** on 11-08 parse differently from the old
fixed bands.

**Two fallbacks when market data is unavailable**, both verified to give identical results:

1. the median of **spaced** symbols in the same orderbook — those are unambiguous, since
   the strike is its own token
2. no band at all → take the candidate that is a multiple of **50**, because listed strikes
   sit on a round step while a mis-parse almost never does

**ATM anchor:** `day mid = (day High + day Low) / 2`, per index, from the fetched market
data. The chain centres on the traded strike nearest that mid.

---

## 8. Account classification

Classification comes from the **All User Details** sheet, tab `Main`. Six columns are read:
`userId`, `server`, `algo`, `max_loss`, `Running Type`, `Running Days`.

### Step 1 — drop on upload

A row is dropped if **any** of `server`, `Running Type`, `Running Days` reads
`DLR ACC` or `NOT RUNNING`. Dealer and stopped accounts are in no report, whatever the DTE.
On the reference file this removed 150 of 786 rows.

Deliberately not an all-three test: a row stopped in one column but stale in another is
still a row no report should classify.

### Step 2 — DTE scope

`Running Days` states the days an account **runs on**, so the scopes are **cumulative**:

| DTE | Running Days included |
|---|---|
| 0DTE | all — `0DTE` + `1DTE/0DTE` + `DAILY` |
| 1DTE | `1DTE/0DTE` **+ `DAILY`** |
| 4DTE | `DAILY` |

A `DAILY` account trades every day, including 1DTE days. Matching `1DTE` against the
literal string `1DTE/0DTE` alone was a bug: it discarded every DAILY account from a 1DTE
report. The 05-08 MTM settles it — all 374 matched accounts were `1DTE/0DTE` (181) or
`DAILY` (193), and none were `0DTE`-only.

### Step 3 — type

| Running Type | Type | Block label |
|---|---|---|
| `INT` | Intraday | **Int** |
| `POS` | Positional | **Pos+Int** |
| anything else / not in scope | Unclassified | own section |

"Pos+Int" is retained because a POS account runs intraday as well.

Matching is on the canonical **user id alone** — `userId` is unique in the sheet, so the
server adds nothing but a chance to mismatch.

### Unclassified

A compiled account the DTE scope does not cover is typed `Unclassified` and reported in its
own section at the **bottom** of the pivot. Consequence: **each Algo Total covers only that
algo's classified accounts**, and

```
Grand Total = Σ Algo Totals + Unclassified sub-total
```

The `unclassified` sheet shows `Not Found` for Running Type / Running Days when the row was
dropped on upload or is absent from the sheet — a dropped row is gone and is not resurrected.

---

## 9. MTM

```
MTM   = Realized P&L + Unrealized P&L        per account, ALWAYS computed
MTM % = MTM / ALLOCATION
```

**The compiled sheet's own `MTM` column is ignored.** It is only populated where Unrealized
is non-zero — on 30-07-2026, **577 of 633 rows** were left at 0 while Realized was not,
understating the book by ₹4.17 Cr. A different file disagreed on only 2 rows. Because the
failure mode varies, the column is overwritten rather than read, and a dashboard note
reports what was measured (rows differing, blank vs wrong value, and the rupee gap).

`MTM %` uses the same number as the money column beside it. In Excel it is a live formula
`=IF(G=0,0,J/G)` so it stays correct if the sheet is edited.

**Display convention:** the stored allocation is in hundreds, so the pivot and the report
show `Allocation × 100`. `MTM %` uses the stored basis, not the displayed one.

Percentiles P5/P95 and the Worst-10%ile sheet were removed — they appeared nowhere in the
HTML report.

---

## 10. Slippage

Both limits are plain **ratios of allocation** — the same unit convention as MTM %, *not*
multiplied by 100.

```
ML %          = MAX LOSS / ALLOCATION
Realized ML % = −Realized P&L / ALLOCATION           (positive = loss)

slippage  when  Realized ML % − ML % ≥ 0.1
```

So ML% 1.00 → Realized 1.09 is **not** slippage; 1.10 is. A float guard of `1e-9` makes a
difference of exactly 0.1 count.

**Measured on realized loss, not MTM**, because a max-loss stop is about what was actually
booked.

### Eligibility

Only accounts with `ALLOCATION > 0` **and** `MAX LOSS > 0` are judged. Accounts with an
allocation but no configured max loss would have `ML % = 0`, so any loss past the threshold
would read as slippage against a limit that never existed. They are listed in the
**`no_sl_Acc`** sheet instead. On 30-07 there were 4, and their aliases confirm the
intent — `CC_BHADADAR_FIX_1CR_0_SL_POS`, `MSV_VTPRATIM_0SL_POS`.

Account counts in the summary cover the eligible set only, so the denominator never
includes accounts the analysis could not judge.

### Major slippage

```
Avg Slippage (algo) = mean Realized ML % over that algo's slippage accounts
Major               = Realized ML % > that average
                      OR the algo has exactly ONE slippage account
```

**Why the special case.** With one account, that account *is* the average, so a strict `>`
can never fire and the algo could never report a major however badly it overshot. A lone
slippage account is its algo's worst by definition. With two accounts the mean sits between
them and exactly one qualifies; the test only becomes discriminating from three up.

The Excel highlight calls `major_slippages()` rather than re-implementing the test, so the
sheet and the report cannot disagree.

---

## 11. Orders Summary

Counts **orders**, not lots. Options only — futures are excluded, matching the strikes
section, so a rejected FUT order cannot invent an algo row in an index that algo never
traded.

| Column | Definition |
|---|---|
| Total Orders | every order, any outcome |
| Executed | `status == COMPLETE` |
| Failed/Cancelled/Rejected | not complete and not pending |
| Pending | still live at close — `OPEN`, `OPEN_PENDING`, `TRIGGER_PENDING`, `TRIGGER PENDING`, `AMO_SUBMITTED`, `PENDING` |
| Hedge | **executed** orders tagged `h_…` |
| VAR | **executed** orders tagged `v_…` |

**Status decides first, then the tag sub-divides only what executed.** Counting the tag
across all statuses reported cancelled hedges as hedge activity: on 11-08 that read
1,82,855 hedge against 3,09,760 executed. Hedge and VAR are slices of **Executed**, never
of Total Orders.

Pending is kept apart from failed — a live order is not a failure.

Deduplication matches Trade Value exactly: COMPLETE rows are deduped as their own set, the
rest as another, so Executed ties to the Trade Value order count.

Order tags: `h_…` hedge, `v_…` VAR, `s_…` square-off, anything else normal.

---

## 12. Intraday chart

One chart per index traded. Lines plotted on the **close** of each bucket.

### Series

| Series | Axis | Source |
|---|---|---|
| Index | left | fetched 1-minute closes |
| Premium | right | uploaded ATM premium file |
| Lots dots | lower panel | orderbook |

An index in the tens of thousands and a premium in the hundreds cannot share a scale, hence
two axes. Axis ticks snap to a round step — **index 50, premium 1** — and the range spans
exactly four steps so gridlines land *on* the numbers.

### Timeframe

Data is embedded at **1-minute** granularity and re-bucketed in the browser to
1 / 5 / 15 / 30 / 60 minutes, taking the **last** value in each bucket (the close). One
payload serves every timeframe, so switching needs no refetch.

### Lot dot categories

**Status decides first; the tag then sub-divides only what executed.**

| Series | Definition |
|---|---|
| **Completed** | every executed lot — **= Stoxxo + Hedge + VAR** |
| Stoxxo | executed, normal tag |
| Hedge | executed, `h_` tag |
| VAR | executed, `v_` tag |
| **Failed** | every cancelled / rejected lot, **any** tag |

These **overlap by design**: `Completed` is the total and the next three are its parts, so
the legend must not be summed.

Excluded: **pending** orders (live, neither executed nor failed) and **square-off** orders
(they close a position rather than place one).

**Why status-first.** An earlier version checked the tag first. On 11-08 that put failed
hedges into `hedge`, inflating it to 15,58,878 lots against a 5,59,259 executed book, while
`failed` showed 15,188 — **1.1% of the real 13,55,012**. One bug, two opposite distortions.

### The 15:29 gap

The index is the cash index and has no continuous ticks during the closing auction, so its
line ends at **15:29** while order-driven series run to **15:40**. The 15:15–15:40 band is
shaded and labelled rather than the gap being hidden or filled with invented data.

---

## 13. Portfolio analysis

> ### Where the Multileg Orders (MLOB) sheet is used
>
> **The MLOB feeds the Portfolio Analysis and nothing else.**
>
> | Calculation | Uses MLOB? |
> |---|---|
> | Portfolio summary — orders, PnL, algo → server → user drill-down | ✅ |
> | `Portfolio QS` Excel sheet | ✅ |
> | Any-pattern portfolio analysis inside DOR.html | ✅ |
> | Trade Value, lots, Lots per Cr, outliers | ❌ orderbook |
> | Strikes / option chain | ❌ orderbook |
> | Orders Summary | ❌ orderbook |
> | Intraday chart — index, premium, lot dots | ❌ orderbook + market data |
> | Segregation pivot, MTM, MTM % | ❌ User MTM + All User |
> | Slippage, `no_sl_Acc` | ❌ User MTM |
>
> Omitting the MLOB removes the Portfolio section entirely; every other figure in the
> report is unchanged.
>
> **Columns the analysis uses** (7): `User ID`, `Portfolio Name`, `Transaction`,
> `Avg Price`, `Filled Quantity`, `Status`, `Server`.
> **Read only to build the dedup key** (3): `Date`, `Order ID`, `Symbol`.
> Everything else in the file — `Leg ID`, `Quantity`, `Exchg Order ID`, `Tag`, `Remarks`,
> `Order Time`, … — is ignored.
>
> Note `Filled Quantity` is used, not `Quantity`: PnL must reflect what actually traded.
> Lots are **not** derived from the MLOB — that comes from the orderbook.

From the MLOB, **COMPLETE rows only**:

```
sell value = Σ Avg Price × Filled Quantity   over SELL rows
buy value  = Σ Avg Price × Filled Quantity   over BUY  rows
PnL        = sell value − buy value
```

Grouped per `(algo, server, user, portfolio)`. The MLOB has no algo column, so each
`(User ID, Server)` is matched against the User MTM exactly as the trade value rows are;
users with no MTM entry inherit their **server's most common algo**.

Default report covers every portfolio whose name contains `QS` (case-insensitive substring).

---

## 14. Index market data

Day High / Low per index for the report date, used for the chain's ATM anchor and the
chart's index line.

| Index | Fyers (primary) | Yahoo (fallback) |
|---|---|---|
| NIFTY | `NSE:NIFTY50-INDEX` | `^NSEI` |
| BANKNIFTY | `NSE:NIFTYBANK-INDEX` | `^NSEBANK` |
| SENSEX | `BSE:SENSEX-INDEX` | `^BSESN` |

**Only the indexes the orderbook actually traded are fetched.** `indexes_in_orderbook()`
reads the segments present and the fetch is restricted to those — BANKNIFTY is absent on
most days, and fetching it anyway costs a round trip and produces a spurious "no data"
warning.

**Fyers is the primary source, Yahoo the fallback.** Fyers is already the feed behind the
volume panel, so the levels and the volume come from one provider and one login. The two
sources were compared before the switch and agree **to the paisa** on all three indexes:

| 12-08-2026 | Fyers H / L / mid | Yahoo H / L / mid |
|---|---|---|
| NIFTY | 24,473.30 / 24,265.95 / 24,369.62 | identical |
| BANKNIFTY | 57,885.85 / 57,254.00 / 57,569.93 | identical |
| SENSEX | 78,263.33 / 77,497.93 / 77,880.63 | identical |

Fyers is also **fresher**: on 17-08-2026 Yahoo returned "no data — market holiday, weekend,
or the date has not settled yet" for all three, while Fyers had the full session. Yahoo is
kept only for days when no Fyers token has been issued.

**TradingView was rejected** as a source: it has no official public API for historical
OHLC — its only official offering is the Charting Library, where the caller supplies the
data. The unofficial websocket scrapers need a TradingView login and break when TV changes
its internals.

```
day mid = (High + Low) / 2
```

**The day's OHLC is derived from the 1-minute candles, not the daily bar.** The two were
verified identical to the paisa on every index and date tested, and the daily endpoint
returns `None` intermittently — so deriving costs nothing and removes a failure mode. It
also means one API call per index serves both the day High/Low and the chart's series.

Every fetch is best-effort: a market holiday, weekend, future date or dropped connection
yields an explanatory message and manual entry boxes, never an exception.

**Retention limit:** Yahoo serves 1-minute history for roughly 30 days and 5/15-minute for
about 60; Fyers serves 1-minute in requests of up to ~100 days and refuses wider ranges
with `Invalid input`. A report date outside the available window produces no intraday
series, and the chart section is omitted rather than half-drawn.

---

## 15. Constants — the complete list

| Constant | Value | Why |
|---|---|---|
| `NORMALISE_BASE` | 100,000 | allocation is stored in hundreds; makes 1 normalise unit = 1 Cr |
| `_MAD_SCALE` | 1.482602218505602 | `1/Φ⁻¹(0.75)` — converts MAD to a σ equivalent. Derived, not typed |
| `DEFAULT_STD_MULTIPLIER` | 1 | default `k`; a dashboard input |
| `SLIP_THRESHOLD` | 0.1 | slippage only past 10% of allocation beyond the limit |
| `_SLIP_EPS` | 1e-9 | float guard so exactly 0.1 counts |
| `FO_EXCHANGES` | `{BFO, NFO}` | F&O only; cash quantities are not lot multiples |
| `PENDING_STATUSES` | OPEN, OPEN_PENDING, TRIGGER_PENDING, TRIGGER PENDING, AMO_SUBMITTED, PENDING | live ≠ failed |
| `ORDER_CATEGORIES` | normal, hedge, var, sqoff | from the tag prefix `h_ v_ s_` |
| `LOT_CATEGORIES` | complete, stoxxo, hedge, var, failed | chart dots; complete overlaps its parts |
| `STRIKE_BAND_LOW` / `HIGH` | 0.5 / 1.5 | strike band = day Low × 0.5 … day High × 1.5, **derived, not hardcoded** |
| `_STRIKE_STEP` | 50 | fallback only — listed strikes sit on a round step |
| strike widths tried | 5, then 6, then 4 | strikes are 5 digits today |
| `ALL_USER_SHEET` | `Main` | the account master tab |
| `ALL_USER_DROP_VALUES` | `DLR ACC`, `NOT RUNNING` | dealer / stopped accounts |
| `DTE_RUNNING_DAYS` | 0DTE all · 1DTE `1DTE/0DTE`+`DAILY` · 4DTE `DAILY` | cumulative — DAILY trades every day |
| `RUNNING_TYPE_TO_TYPE` | INT→Intraday, POS→Positional | |
| `DEFAULT_PORTFOLIO_PATTERN` | `QS` | default portfolio report |
| axis steps | index 50, premium 1 | round chart ticks |
| chart shade band | 15:15 – 15:40 | CAS / extended derivatives window |

### Market-timing facts encoded

Following the **Closing Auction Session (CAS)**, live 3 August 2026:

- Cash continuous trading ends **15:15** for F&O stocks; CAS runs **15:15–15:35**
- Equity **derivatives now trade to 15:40** (was 15:30)
- The cash index has no continuous ticks during the auction, so index intraday ends 15:29
- Orderbooks therefore legitimately contain orders after 15:30 — **no time filter is
  applied anywhere**; the full book to 15:40 is read

---

## 16. Reconciliation identities

These hold by construction and are asserted in the code. If one breaks, something is wrong.

```
Trade Value
  Below + In + Above          = Total Users            (per algo/type; users with no
                                                        allocation carry no flag)

Orders Summary
  Executed + Failed + Pending = Total Orders           (per algo, per server)
  Hedge + VAR                ≤ Executed

Chart lots
  Stoxxo + Hedge + VAR        = Completed              (exact)
  Completed                   = Trade Value lots − square-off lots
  Completed + Failed + Pending = total lots fired

Segregation pivot
  Σ Algo Totals + Unclassified = Grand Total
  Realized + Unrealized        = MTM                   (per row and per block)

Portfolio
  Σ sell − Σ buy               = PnL
```

Verified on 11-08-2026 NIFTY: Completed 5,59,259 = Stoxxo 2,90,076 + Hedge 2,61,286 +
VAR 7,897, and Trade Value 5,59,307 − 48 square-off lots = 5,59,259.

---

## Known limitations

1. **Lot-size tables are hardcoded** and have no expiry. An exchange revision must be added
   manually; nothing detects it.
2. **One global `k`** for outliers. Band widths range from ±0.74% to ±32% of the median
   across algos, so a single multiplier treats very different populations identically.
3. **`_classify_symbol` is a substring test** for the segment field, so FINNIFTY and
   MIDCPNIFTY count as NIFTY in Trade Value. `parse_strike` (strikes, Orders Summary)
   requires the prefix and excludes them.
4. **Intraday index data ends 15:29**, so chart activity between 15:29 and 15:40 has no
   index line behind it.
5. **1-minute retention** — Yahoo ~30 days, Fyers ~100 days per request; older report dates get no chart.
