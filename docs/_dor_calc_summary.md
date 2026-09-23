# DOR — Calculation Summary

Four calculations, their formulas, and the columns they read. One page each.

Source file for every figure below: **`Compiled_Orderbook_<date>.csv`**, except
Portfolio analysis, which reads **`Compiled_Multileg_Orders_<date>.xlsx`**.
Allocation and algo come from **`Compiled_User_MTM_<date>.xlsx`**.

---

## Rows that enter every calculation

```
Exchange        ∈ { NFO, BFO }                    strict — no column is an error
Exchange Time   present, and not year 0001        an order the exchange never saw
Symbol          resolves to NIFTY / BANKNIFTY / SENSEX
User ID         found in the User MTM, or in aliases.json
```

Orders stamped outside **09:15 – 15:40** are *kept* and reported separately as
**Out of market order**, whatever their status.

All timing uses **Exchange Time**, not Order Time.

Duplicates are removed first, on:

```
User ID + Order ID + Order Time + Exchg Order ID + Exchange Time + Tag
```

---

## 1. Trade Value

Per user, per server, per index — from **COMPLETE** rows only.

```
lot size     = per index, per date          NIFTY 65 · SENSEX 20 · BANKNIFTY 30
lots         = Σ |Quantity| ÷ lot size
trade value  = Σ (Avg Price × Quantity)
```

Outlier flagging normalises lots by the account's allocation:

```
normalise    = ALLOCATION ÷ 1,00,000
lots per Cr  = lots ÷ normalise

band         = median ± k × MAD × 1.4826       k = deviation input, default 1.0
               MAD = median(|x − median|)
```

Median and MAD, not mean and standard deviation: one blown-up account cannot
widen the band that is meant to catch it. `1.4826` rescales MAD to a standard
deviation on normal data — a unit conversion, like 2.54 cm per inch.

| Reads | From |
|---|---|
| `Quantity`, `Avg Price`, `Symbol`, `User ID`, `Server`, `Date` | orderbook |
| `ALLOCATION`, `ALGO`, `SERVER` | User MTM |

---

## 2. Orders Summary

Counts **orders**, not lots. Options only. Every status is read.

```
Total Orders   = every order that passed the row filters
Executed       = status COMPLETE
Rejected       = not complete, not pending      (rejected + cancelled)
Pending        = OPEN · OPEN_PENDING · TRIGGER_PENDING · AMO_SUBMITTED
Out of market  = Exchange Time outside 09:15–15:40, whatever the status
Hedge / VAR    = EXECUTED orders tagged h_… / v_…
```

The rejection split, on the `Status Message` text:

```
Margin Rejection = message contains any of
                   margin shortfall · margin exceeds · insufficient funds ·
                   red:sqroff shortfall · rrm: collateral · span limit ·
                   oems:[buy exposure limit for options]
                   …plus "saf:order is not open to cancel" when status = REJECTED

Others           = Rejected − Margin Rejection
```

Two identities hold on every row:

```
Total Orders = Executed + Rejected + Pending + Out of market
Rejected     = Margin Rejection + Others
```

Status decides first, the tag only subdivides what executed — so Hedge and VAR
are slices of Executed and can never exceed it.

---

## 3. MS volume

Our own traded quantity per minute — the teal line on the chart.

```
MS volume(index, minute) = Σ Quantity
                           over  status = COMPLETE
                           and   minute(Exchange Time) = minute
```

| Choice | Reason |
|---|---|
| Quantity, not lots | exchanges report volume in contracts, so the two are comparable |
| COMPLETE only | volume is what transacted; a cancelled order never touched the tape |
| Both sides counted | `Transaction` is not read, so a BUY and a SELL each add their quantity |

A strike-minute share above 100% therefore means internal crossing, not an error.

---

## 4. Portfolio analysis

From the Multileg Orders file, **COMPLETE** rows only, per portfolio and per user.

```
sell value = Σ (Avg Price × Filled Quantity)    over SELL rows
buy value  = Σ (Avg Price × Filled Quantity)    over BUY rows
PnL        = sell value − buy value
```

The MLOB carries no algo, so each `(User ID, Server)` is matched against the User
MTM exactly as the Trade Value rows are; a user with no MTM entry inherits their
server's algo.

Default report covers portfolios whose name contains **QS**; any other pattern can
be analysed in the HTML report itself.

---

## Reconciliation

```
Executed (Orders Summary)  =  order count behind Trade Value
Total Orders               =  Executed + Rejected + Pending + Out of market
Rejected                   =  Margin Rejection + Others
Completed lots             =  Stoxxo + Hedge + VAR
```

## What is excluded, and reported

| Excluded | Why | 20-08-2026 |
|---|---|---|
| Blank / non-F&O rows | not NFO or BFO | 8,636 |
| No exchange timestamp | the exchange never saw the order | 628 |
| Account in no reference file | carries no algo or allocation | 84 |

Each count is logged on every run, never absorbed silently.
