# Daily Operations Report (DOR)

A Streamlit app (plus standalone CLIs) that turns four raw inputs into the daily ops report:

1. **Trade Value analysis** — per-user lots / trade value from the orderbook, with statistical
   outlier flagging on lot pct.
2. **Intraday / Positional segregation** — a pivot of Realized/Unrealized P&L per algo → Int /
   Pos+Int → server, with max-loss addons applied to positional accounts.

Outputs: one Excel workbook `DOR_<date>.xlsx` (5 sheets: `tradevalue`, `summary`, `Segregation`,
`Raw_Data_Per_User`, `Worst 10%ile`) and one self-contained, client-shareable `DOR_<date>.html`.

## Files

| File | Role |
|---|---|
| `app.py` | Streamlit UI — upload the 4 inputs, set the outlier deviation, one **Process** click computes both reports. Run with `streamlit run app.py`. |
| `tradevalue.py` | Trade value engine (also a CLI: `python tradevalue.py orderbook.csv -s user_mtm.xlsx -d 1`). |
| `segregate_int_pos_mtm2.py` | Segregation engine (also an interactive CLI). |
| `dor.py` | Renders the DOR.html summary (inline CSS/SVG/JS, no external assets). |

## Inputs

| Input | Used for |
|---|---|
| Compiled Orderbook (CSV/Excel) | Trade value rows |
| Compiled User MTM (CSV/Excel) | Allocations + algo for trade value; the account universe for segregation |
| Combined Max Loss — 1DTE (required) | Positional classification + realized-P&L addon |
| Combined Max Loss — 4DTE (optional) | Extra addon for non-Noren positional accounts |
| Outlier deviation `k` (default 1.0) | Width of the lot-pct outlier band (± k standard deviations) |

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

## 1.7 Lot Pct and outlier flagging

```
Lot Pct = Lots / Allocation × 100        (blank when the user has no allocation)
```

Outliers are judged **per (trade date, algo) group**, using the **population standard deviation**
(variance divided by *n*, not *n−1*):

```
mean = Σ lot_pct / n
std  = sqrt( Σ (lot_pct − mean)² / n )
band = [ mean − k·std , mean + k·std ]      (k = the chosen deviation, default 1)
```

- `lot_pct < mean − k·std` → **"Below average range"**
- `lot_pct > mean + k·std` → **"Above average range"**
- otherwise → **"In range"**
- rows with no allocation or no algo can't be judged → blank flag

## 1.8 Algo summary & totals

Per (date, algo): row count, `Avg Lot Pct` (the mean above), `Std Dev`, the band
`low–high`, and how many rows fall Below / In / Above it.

Report KPIs: `Users` = distinct user ids, `Orders` = Σ order counts, `Total Lots` = Σ lots,
`Trade Value` = Σ trade values.

## 1.9 Box plot (dashboard + DOR.html)

Per algo, lot pct distribution: box = Q1–Q3 (quartiles via the *inclusive* method), line = median,
whiskers = the furthest actual values within **1.5 × IQR** of the box. The overlaid points are
**not** IQR outliers — they are the rows flagged by the ± k·σ rule above (blue = below,
orange = above).

---

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
Return %            = Realized / Allocation        (0 when allocation is 0)
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
accounts − positional, `Noren` = rows with User Type "Noren".

## 2.4 Excel formulas (Segregation sheet)

Aggregate cells are written as live formulas so the sheet stays auditable in Excel:
Sub-Totals use `=SUM(...)` over their data block, Algo/Grand Totals add the sub-total rows,
and every `Return %` cell is `=IF(G{row}=0,0,H{row}/G{row})` (Realized ÷ Allocation).
P5/P95 are written as computed values.

## 2.5 Raw_Data_Per_User sheet

One row per account with the full audit trail: Type, User Type, algo, server, id, alias, SL Hit
(0/1), MAX LOSS, ALLOCATION, Compiled Realized P&L, Addon 4DTE, Addon 1DTE, Addon Applied,
Realized P&L (Final), Unrealized P&L, and `MTM = Final Realized + Unrealized`.

## 2.6 Worst 10%ile sheet

Per **Algo × Int/Pos** group:

- Rank accounts by `UserReturn`; keep those **at or below the group's 10th percentile**.
  Only accounts with `ALLOCATION > 0` are ranked (zero-allocation returns are undefined).
- `Algo Avg Return %` shown next to each user = the **group-level** return
  `Σ Realized (Final) / Σ Allocation` — the same value as that section's sub-total Return %.
- `Reason` is auto-inferred (first match wins, editable afterwards in Excel):
  1. SL Hit == 1 → **"SL Hit"**
  2. allocation ≤ 0 → **"Zero/low allocation"**
  3. return < 0 → **"Negative return"**
  4. otherwise → **"Low relative return"**

## 2.7 Report date

Inferred from the first non-empty `Date` value in the compiled MTM sheet, formatted `dd-mm-YYYY`.

---

# Part 3 — DOR.html (`dor.py`)

Pure presentation — no new math. It embeds the already-computed numbers into a single
self-contained HTML file: KPI tiles, the per-algo trade-value outlier summary, the box plot as
inline SVG (same quartile/whisker/point rules as §1.9, with deterministic point jitter so stacked
outliers stay visible), and the segregation pivot as a click-to-drill-down table
(Algo → Int / Pos+Int → servers). Row-level data intentionally stays out of the HTML; it lives in
the Excel workbook.

---

## Notes

- `app.py` caches processing on the file contents + deviation + report schema, and warns when the
  uploaded inputs changed since the last **Process** click.
- Trade value math uses `Decimal` end to end; segregation uses pandas floats with money rounded
  to whole rupees in the outputs.
