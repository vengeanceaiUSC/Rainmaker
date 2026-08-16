# Working capital refined — AR vs NWC + deferred revenue

## Your DSO → AR math (kept)
`6 / 59 ≈ 10.16%` DSO reduction → **~10% lower AR** if sales are steady.  
That frees cash trapped in **Accounts Receivable**. Modeled as **C13 = 10%** applied to **ΔAR only**.

## Asterisk: not “10% of total NWC”
NWC = Current assets − Current liabilities. In B2B SaaS, deferred revenue is large, so a 10% AR drop is **not** the same as a 10% change in the net NWC ratio.

## Equation now in the model
```
ΔAR_base = ΔRev × (59/365)     # ≈16.2% AR intensity (Fairview median DSO)
ΔAR_AI   = ΔRev × (59/365) × (1 − 10%)
ΔDeferred = ΔRev × 39.0%       # ZI FY24 unearned $473.8m / $1,214.3m
ΔNWC ≈ ΔAR − ΔDeferred
```

| Driver | Cell | Value | Role |
|---|---|---:|---|
| AR / receivables improve | C13 | **10%** | Cuts **ΔAR** only |
| AR intensity (DSO÷365) | C25 | **16.2%** | Gross AR leg |
| Deferred revenue % of rev | C26 | **39.0%** | Liability growth = cash source |

## FCF impact (Y1 at +5% rev)
- Base ΔNWC ≈ **−$13.9m** (cash source: deferred > AR on growth)
- AI ΔNWC ≈ **−$14.9m**
- AI FCF vs Base from lower ΔAR ≈ **+$1.0m** in Y1
- Illustrative stock: 10% × FY24 AR $246.1m ≈ **$24.6m** if the AR *balance* compresses (model uses **flow** ΔNWC, not a one-time write-down)

## Why 20–40% headlines stay out of the EQ
Those often measure overdue buckets or FTE hours, or start from 90+ day DSO. At a healthy ~59-day median, ~6 days / ~10% AR is the realistic cash-acceleration metric.
