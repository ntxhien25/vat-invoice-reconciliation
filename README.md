# VAT invoice reconciliation: fictional demonstration

Work-informed method, aggregate-calibrated fictional records. No client source rows, amounts, tax identifiers, invoice lookup URLs or measured client results are reproduced.

The illustrative period is April 2025. Source volumes and exception breakdowns retain the earlier aggregate presentation. Tax calculations are illustrative, not tax advice or confirmation of deductibility.

## Run

Install `requirements.txt`, then run `python vat_reconciliation_demo.py` or the notebook from this folder. Run `python -m unittest test_demo` to verify matching controls. `python generate_fictional_data.py` recreates all five source workbooks from the explicit fictional specification.

The reusable matching logic compares invoice keys, net amounts, VAT and total amounts, flags duplicates and missing values, and retains source references. The four report sheets separate matches, differences and source-only records. Human review of underlying documents remains necessary.

No portal automation, credentials or lookup links are included. Do not replace the fictional inputs with client data in a public repository.

Legacy filenames containing 2026-08 are compatibility aliases of the fictional April 2025 inputs. They contain no original records. Use the 2025-04 files for the current demonstration.
