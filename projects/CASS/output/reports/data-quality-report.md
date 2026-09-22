# Data Quality Report: CASS

## Scope
- Files inspected: `Academic_services_without_income.xlsx`, `Operational & Service Excellence by CAAS.xlsx`, and the three budget-control workbooks for fiscal years 2567–2569
- Dashboard financial source: budget-control registers (83 records)

## Findings
- **Pass**: Financial profit is derived per record as revenue minus total expense. Aggregate revenue 13,400,281.96 THB minus expense 7,654,152.68 THB equals profit 5,746,129.28 THB; reconciliation difference is -0.00 THB.
- **Medium**: Fiscal year 2569 may be incomplete or include activities spanning fiscal years; compare it with completed years cautiously.
- **Medium**: MU-ELT รหัส 67 is retained as a separately labelled register because it contains distinct placement-test and equipment activities.
- **Medium**: Community activity data has 33 rows but only 12 have project names and are included in the non-income drill-down. Blank source values remain blank/zero only where the source supplied no numeric value.

## Handling
No raw record was changed or removed. Tab 1 and Tab 3 use the same reconciled budget-register dataset; Tab 3 exposes its source grouping and detail records.
