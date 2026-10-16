# VAT invoice reconciliation: GDT and MISA

A Python project for reconciling VAT input and VAT output invoice records against VAT return schedules. Start with the notebook for the analysis, or run the module to regenerate both review workbooks and charts.

## Business situation and purpose

The group comprises two companies. One company handles approximately 800 VAT input invoices and 200 VAT output invoices per month. Most VAT input invoices are printed from MISA meInvoice, with other printed invoices submitted by staff for miscellaneous expenses. According to the company’s discussions with MISA, meInvoice imports data from GDT. The company has no workflow to check whether all VAT input invoices have been collected for VAT return preparation. This project reconciles GDT against the accounting schedules to identify records requiring follow-up.

The task is to automate routine comparisons and provide a traceable exception report for the VAT return preparer. The workflow matches records, checks amounts and separates items requiring investigation. The manual printing and checking process involves approximately 120 minutes of active work, spread across 1–2 elapsed days. The estimated workflow time of 15–20 minutes represents an approximately 83%–88% reduction when comparing 120 minutes to 15–20 minutes. This estimate awaits validation through a comparable timed review and is not a measured reduction in total VAT return preparation time. The project dataset contains 839 GDT VAT input invoices and 189 VAT output invoices, providing a workload close to the monthly business context. Document checks and input VAT deductibility decisions remain subject to the preparer's review.

## Public data scope

The five source workbooks retain the original headers, column order and worksheet layouts. Company names, addresses, tax codes and invoice identifiers have been replaced. Amounts are varied consistently across corresponding records, while additional sample records retain the matching and review patterns of the baseline. Only the input and output invoice schedules are populated for reconciliation; other MISA worksheets retain their template layouts. Original confidential workbooks, mapping files, workstation paths and tax-portal access components are excluded. Both notebooks use the same processing code and guidance, with separate source configuration.

[Read the rendered notebook](https://ntxhien25.github.io/VAT_Reconciliation_Demo.html) · [Portfolio case study](https://ntxhien25.github.io/vat-reconciliation.html)

## Run locally

```text
python -m pip install -r requirements.txt
python vat_reconciliation_demo.py
python -m unittest test_demo -v
```

Open `VAT_Reconciliation_Demo.ipynb` and run the cells in order for the guided analysis. No tax-portal account or private source file is needed.

## Matching rules

Inputs use invoice number, issue date and seller's tax code. Outputs use invoice number, issue date and purchaser's tax code. Leading zeros are ignored only in comparison keys; displayed identifiers retain their padding. Duplicate matching keys require review.

Net amounts, invoice totals and VAT amounts are checked independently within VND 1. MISA invoice total is the sum of the value excluding VAT and VAT in its schedule. Blank GDT net amounts remain missing and are labelled Commercial Invoices for review under the project rules. Blank GDT VAT and totals are treated as zero for comparison only. Review labels do not establish legal document classification or entitlement to input VAT credits.

## Results

| Source | Records | Matching amounts | Paired records for review | Source only |
| --- | ---: | ---: | ---: | ---: |
| Input / GDT | 839 | 715 | 33 | 91 |
| Input / MISA | 774 | 715 | 33 | 26 |
| Output / GDT | 189 | 182 | - | 7 |
| Output / MISA | 182 | 182 | - | - |

The 33 input pairs for review have missing GDT net amounts. The MISA-only records contain 20 customs declarations and six prior-period invoices. Each paired invoice is counted once per source; do not add the GDT and MISA counts together.

## Project files

- `vat_reconciliation.py`: reusable matching, checks, exports and charts.
- `vat_reconciliation_demo.py`: runs the workflow against the five masked source files.
- `VAT_Reconciliation_Demo.ipynb`: guided analysis with saved results.
- Four `Purchases_*.xlsx` / `Sales_*.xlsx` files and `MISA_VAT_Return_Schedules.xlsx`: complete masked records in the original source layouts.
- `VAT_Reconciliation_Demo.html`: readable notebook with code, charts and tables, suitable for static hosting.
- `reports/`: input and output Excel workbooks, each containing four reconciliation sheets.
- `assets/`: chart images generated when the demonstration is run.
- `test_demo.py`: 21 checks using fictional fixtures testing missing amounts, duplicates, leading zeros, source counts and Excel outputs.

Invoice statuses use New `#50A2A7`, Adjustment `#9F838C`, Replacement `#471323`, Cancelled `#334E58` and Other `#EBEBD3`. The reconciliation chart retains its approved red family. Bar opacity is 88%; labels use the same base colours at full opacity. Outer corners have a radius of 10% of bar width; internal stack boundaries remain flat.

GitHub profile: https://github.com/ntxhien25
