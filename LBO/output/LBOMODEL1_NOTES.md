# vengeanceaiUSC-LBOMODEL1 — ZoomInfo AI-SDR LBO

## Thesis
Replace outbound human SDRs with AI agents → cut bloated S&M → **≥500 bps EBITDA margin expansion** → higher FCF → faster debt paydown → **higher LBO IRR**.

## Download
https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/vengeanceaiUSC_LBOMODEL1.xlsx

## FY2024 baseline (entry)
| Item | Value |
|---|---|
| Revenue | $1,214.3m |
| Gross Profit / GM | $1,024.5m / 84% |
| OpInc (GAAP) | $97.4m |
| D&A | $85.7m |
| Baseline EBITDA | **$183.1m (15.1% margin)** |
| S&M baseline | **$425.0m (35% of rev)** per strategy prompt |

## AI strategy rules
1. **Y1:** S&M −15%; add **$15m** AI compute/software cost  
2. **Y2–Y5:** S&M grows **2%/yr** only  
3. **Target:** ≥500 bps EBITDA margin expansion by Y2  
4. **Exit:** **13.0x** EV/EBITDA (mid of 12–14x)  
5. Higher FCF **sweeps debt** faster than base case  

## Results (in `Strategy_Summary`)
| | Base (no AI) | AI-SDR case |
|---|---|---|
| Y2 margin vs baseline | ~14 bps | **~566 bps** |
| MOIC | ~1.33x | **~2.18x** |
| IRR | ~5.8% | **~16.9%** |
| **IRR from AI S&M** | — | **+~11.1 percentage points** |

Open sheets: `Strategy_Summary`, `AI_Operating`, `Base_Operating`, `Debt_Sweep_AI`, `00_Assumptions_List`, `LBO`.

Rebuild: `python3 LBO/output/build_lbomodel1_ai_strategy.py`

Educational / research use only — not investment advice.

## Strategy Summary PDF
https://raw.githubusercontent.com/vengeanceaiUSC/Rainmaker/cursor/vengeanceaiusclbomodel1-44dc/LBO/output/LBOMODEL1_STRATEGY_SUMMARY.pdf

Includes metric commentary and Base vs AI results. Chart also lives on `Strategy_Summary` in the workbook.
