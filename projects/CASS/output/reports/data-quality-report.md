# Data Quality Report: CASS

## Scope
- Files inspected: `Academic_services_without_income.xlsx` and `Operational & Service Excellence by CAAS.xlsx`
- Dashboard analytical source: `Service_Excellence_by_CAAS` (1,738 records)

## Findings
- **Critical**: Reported profit totals 21,128,837.51 THB; revenue minus expense is 23,558,726.27 THB. The reconciliation difference is -2,429,888.76 THB.
- **High**: The consolidated source contains 34 exact duplicate rows. Records are preserved and included in totals.
- **High**: The consolidated source has material missingness in participant, satisfaction, and segmentation fields. Missing numeric values are represented as 0 only in the analytical copy; this is not an imputation.
- **Medium**: In `income_generating_project`, satisfaction reaches 6.46, outside the apparent 0-5 scale.
- **Medium**: Community activity data has 33 rows but only 10 have project names and financial/participant fields.

## Handling
No raw record was changed or removed. The dashboard labels reported profit separately from derived profit and shows a NOT READY quality status.
