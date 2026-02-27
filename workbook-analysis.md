# Deep Analysis: `Sample Goal Plan 1.xlsx`

## Workbook structure
- Total sheets: 6
- `Input Sheet`: primary manual entry sheet (118 non-empty rows, 5 formula cells)
- `Goal-Sheet`: goal projection + SIP strategy (138 formula cells)
- `Networth Statement`: net worth and allocation breakdown (69 formula cells)
- `Cash Flow`: yearly projection and retirement sustainability (318 formula cells)
- `Goal Sheet Breakup`: scheme-level mapping to goals (44 formula cells)
- `My Page`: assumptions and rate table (10 formula cells)

## Dependency map
- `Input Sheet` -> no external references; this is the base user-input layer.
- `My Page` -> references `Input Sheet` for goal labels.
- `Goal-Sheet` -> references both `Input Sheet` and `My Page`.
- `Networth Statement` -> references `Input Sheet`.
- `Cash Flow` -> references `Input Sheet` and `Goal-Sheet`.
- `Goal Sheet Breakup` -> references `Goal-Sheet`.

## What should be editable in UI
- Personal details, dependents, employment profile.
- Goal details (years left, required amount, current provision).
- Retirement ages / expectancy.
- Inflow / outflow line items.
- Assets, investments, liabilities.
- Optional assumptions (`My Page` rates) if you want advisor overrides.

## What should be automated (computed)
- Goal future value projections using inflation rates.
- Goal corpus, gap, and SIP required (PMT/RATE-based style logic).
- Net worth totals, allocation percentages, and liability impact.
- Cash flow yearly trajectory (opening balance, growth, goal outflow, closing balance).
- Retirement drawdown years from `Goal-Sheet` retirement corpus lines.

## Key formulas to preserve in app logic
- Age: `ROUNDDOWN(YEARFRAC(DOB, PlanDate),0)`
- Future value: `FV(inflation, years, , -currentCost)`
- Goal SIP: `PMT(monthlyRate, years*12, currentProvision, -targetCorpus, 1)`
- Net worth: `Total Assets - Total Liabilities`
- Cash flow close: `FV(growth,1,-cashIn,-opening) - goalOut`

## UI-first architecture recommendation
- Build only one visible page: `Input`.
- Keep computed JSON objects for:
  - `goalSheetModel`
  - `networthModel`
  - `cashflowModel`
  - `goalBreakupModel`
- Expose output tables/cards for advisor review; no manual edits there.

## Notes from sample file quality
- Some formulas are hard-coded to sample years (cashflow starts from 2018 in template), so dynamic year logic should be tied to plan date in the web app.
- A few lines have presentation artifacts in Excel (`+`, duplicate-style expressions, fixed constants). These should be normalized during implementation.
