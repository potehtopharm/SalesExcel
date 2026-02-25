# Excel VBA Report Formatter (Arrival/Booking Export Cleanup)

A lightweight VBA toolkit that turns raw Excel exports into consistent, operationally usable reports.

![Use Case 1 Before/After](images/BeforeAndAfter.png)

## Use Case 1 — Report Formatting (Layout Cleanup)
Formats a raw export into an export-ready report:
- Resizes key columns and wraps long fields
- Removes unneeded columns
- Standardizes headers
- Adds row borders for readability

Screenshots:
- ![Use Case 1 Before](images/usecase-1-before.png)
- ![Use Case 1 After](images/usecase-1-after.png)

Output file:
- `examples/usecase-1/after-usecase-1.xlsx`

## Use Case 2 — Accounts of the Quarter Filter (AOQ)
Filters the same raw export to isolate guests under selected “Accounts of the Quarter” companies (anonymized for demo), then outputs a simplified sheet for operations use.

Screenshot:
- ![Use Case 2 After](images/usecase-2-after.png)

Output file:
- `examples/usecase-2/after-usecase-2.xlsx`

## Example files
Shared input (used by both use cases):
- `examples/input/before.xlsx`

## Repo structure
- `src/` — VBA source (`.bas`)
- `examples/` — anonymized input + outputs
- `images/` — screenshots

## How to run
1. Open the workbook containing the raw export (or paste the export into a sheet).
2. Open the VBA editor (**Alt + F11**) and import the `.bas` file(s) in `src/`.
3. Run the macro on the target sheet:
   - Use Case 1: `FormatAndPrepareSheet` (from `src/FormatAndPrepareSheet.bas`)
   - Use Case 2: `FilterAndFormatCompanyNames` (from `src/FilterAndFormat_AOQ.bas`)

## Notes
- Demo data is fully anonymized / placeholder.
- Goal: reduce manual reporting time and improve consistency for repeated operational workflows.
