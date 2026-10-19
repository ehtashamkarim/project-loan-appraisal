# Construction-Finance Appraisal: G+5 Commercial Building, Mangaluru

A bank-style credit appraisal of a Rs 23.2 Cr lease-rental commercial project, built as a live Excel model (`Project_Loan_Appraisal_Model.xlsx`, 1,000+ formulas, zero hard-coded outputs). It combines a civil engineering cost build-up with a lender's view of cash flows and covenants.

## The deal
| Item | Value |
|---|---|
| Built-up area | 60,000 sq ft (G+5, RCC frame) |
| Hard cost | Rs 15.6 Cr (Rs 2,600/sq ft) + 10% soft cost |
| Land (promoter) | Rs 6.0 Cr |
| Total project cost | **Rs 23.2 Cr** |
| Bank loan asked | **Rs 12.0 Cr** (70% of construction + soft cost), 10.5%, 2-yr construction + 10-yr tenor |
| Revenue | Lease rent Rs 75/sq ft/month, 80% leasable, occupancy 60% → 80% → 90% |

## What the model does
- Year-wise rental income, opex, NOI (CFADS), loan drawdown, interest, principal, **DSCR**
- **Project IRR, equity IRR, LTV** at stabilisation (exit cap rate 8.5%)
- **6 stress scenarios**, each a full recalculating model block: rent −10%, cost overrun +15%, rate +100 bps, occupancy −10 pp, and combined
- Automatic covenant check (min DSCR ≥ 1.20x, average ≥ 1.40x)

## Key findings
1. **A standard front-loaded repayment fails.** Starting principal in the first operating year (5% step-up) gives a **minimum DSCR of 1.18x** in the lease-up year, below the 1.20x covenant, even though the average is 1.9x.
2. **Restructured to match the rent ramp:** a 1-year principal moratorium after completion, then step-up repayment (7% → 16%). Base case: **min DSCR 1.46x, avg 1.91x, project IRR 19.4%, equity IRR 21.9%, LTV 42%**.
3. **Passes all 4 single stresses** (worst: cost overrun +15% → min DSCR 1.27x).
4. **Fails only the combined stress** (rent −10% and cost +15% → 1.14x).

## Recommendation
**Sanction Rs 12.0 Cr** with the restructured repayment, subject to:
- a Debt Service Reserve Account (DSRA) of 2 quarters' debt service, which covers the combined-stress gap;
- promoter undertaking to fund cost overruns (cost overrun is the largest single risk);
- escrow of lease rentals (Lease Rental Discounting style), and pre-leasing of at least 40% before final disbursement.

## Interview talking points
- Why DSCR matters more than IRR to a lender: the bank's upside is capped at interest, so it cares about the weakest year, not the average.
- Why the lease-up year is the weakest: occupancy is lowest while interest is highest (full loan outstanding).
- Civil link: the cost build-up and overrun stress come from construction knowledge (G+5 STAAD/ETABS project).

## Rebuild the workbook
```
pip install openpyxl
python build_model.py
```
All figures are assumptions for a learning case study, not a real loan.
