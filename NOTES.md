# Vista DCF — project notes and technical reference

[Project introduction](README.md) · [How to use the model](HOW_TO_USE.md)

This file preserves the technical notes, accounting choices, verification details and sources from the original README. The original workbook and command instructions are in [HOW_TO_USE.md](HOW_TO_USE.md).

## Additional project notes

Use this section for future observations, decisions and follow-up items. Add a date and enough context to explain what changed or what still needs review. The reference sections below retain the existing project documentation.

### Repository scope and local files

The Git repository covers only the `dcf_model` folder. It includes source code, documentation, tests and the reviewed public Vista source cache. Generated `outputs/` files, including workbooks, saved assumptions, input notes, previews and validation logs, stay local. Dependencies, environment files, credentials and editor settings are also excluded by `.gitignore`.

The cached Excel and PDF files are company-published source documents, not generated models or personal workbooks. The derived audited-text extract is excluded because it can be regenerated from the PDF. Review new files before committing: ignore rules do not detect every possible secret or private note.

## Command-line reference

The equivalent Python entry point is `python -m dcf_model.cli build`. Optional arguments include `--source-xlsx`, `--source-url`, `--from-workbook`, `--config`, `--output-dir` and `--no-previews`.

## Generated supporting files

These files are written to `dcf_model/outputs` by default. For workbook naming, version selection and saving inputs, see the [usage guide](HOW_TO_USE.md).

- `latest_workbook.json`: paths to the newly generated and previous workbooks.
- `historical_financial_data.json`: historical financial and operating data.
- `model_assumptions.json`: assumptions used for the latest build.
- `dcf_valuation_results.json`: valuation calculated by Python.
- `excel_workbook_definition.json`: Excel layout, formulas and input mappings.
- `formula_validation.json`, `formula_error_scan.ndjson`, `valuation_inspection.ndjson` and `native_excel_validation.json`: calculation checks and inspection results.
- `previews/`: images of workbook sheets.

## Architecture and dependencies

- `data_sources.py`: company-specific source mappings, annual-period selection, USD scaling, downloads, provenance and independent checks against audited PDFs.
- `model.py`: pure Python normalization, validation, upstream forecasting, terminal value, equity bridge and sensitivity calculations. `forecast(data, config, case)` returns annual cash flows, terminal calculations, valuation and warnings.
- `workbook_spec.py`: Excel cell references, formulas and editable-input mappings, plus expected values from the independent Python engine.
- `build_workbook.mjs`: workbook creation, calculation checks, visual previews and XLSX export using `@oai/artifact-tool`.
- `cli.py`: cached/online refresh, saved-input recovery and versioned outputs.

Python requires `openpyxl` and `pypdf` as listed in `requirements.txt`. `openpyxl` only reads source and saved workbooks. Excel authoring uses Node and the Codex bundled `@oai/artifact-tool` runtime. The supplied launcher detects the bundled Python; the exporter detects bundled Node or `DCF_NODE`. The local `node_modules` junction points to the bundled Node dependencies. These installed runtime paths are machine-specific and the junction should not be committed. The Python valuation can run independently on another machine; workbook generation also requires the Artifact Tool runtime.

The JSON model separates reported history, manual overrides, scenario policies, annual drivers and the equity bridge. Period keys are explicit calendar years. Input schema version is 1. Missing required values, invalid rates and unrecognized source layouts fail explicitly. V1 deliberately requires reviewed 2023–2025 sources and 2025 share metadata; it refuses a new fiscal-year window until the new audited filings and company-specific mappings have been reviewed.

To add an upstream issuer, implement a source adapter producing the same history/fact contract, then supply reviewed company defaults, normalization adjustments and security metadata. Other industries can retain the valuation/output layers while replacing the operating forecast engine.

## Accounting and valuation choices

- **USD consistency:** historical reported USD values are used directly and scaled to millions. No conversion using today's peso exchange rate occurs. FCFF and WACC are nominal USD.
- **Revenue:** realized oil prices are net of export duties. Duties are removed from both revenue and royalties; freight reimbursements are removed from both revenue and selling expense. Gas and NGL are modeled separately.
- **Core EBIT:** revenue less lifting costs, royalties, selling, overhead and DDA. The residual versus reported EBIT is shown explicitly and excludes other gains, impairments, exploration and other non-cash operating items. It is not the company's adjusted EBITDA definition.
- **Production:** average daily growth converts to annual volumes using actual calendar days. Sales-to-production factors are anchored to FY2025. Gas sales use a sourced million-MMBtu/produced-million-boe factor, not an assumed universal energy conversion.
- **Investment:** required new annual capacity equals forecast annual production less surviving prior production after natural decline, floored at zero. Investment equals new capacity times an editable cost per annual boe, plus infrastructure/intangible investment per boe and extra capex. A total-capex override is available, but a shortfall warning remains visible.
- **Leases:** capital investment includes owned and leased productive capacity; lease liabilities are deducted as financial debt. Lease payments are not deducted again in FCFF. This is a simplified capital-intensity model, not a detailed lease schedule.
- **Taxes:** `max(EBIT, 0) × cash-tax rate`. No automatic credit for operating losses, NOL carryforward or deferred-tax schedule.
- **Working capital:** current receivables plus inventories minus current trade payables, with an explicit adjustment for current financing payables. This is an editable proxy, not a full working-capital schedule.
- **Terminal value:** year 11 uses flat real production and user-selected nominal growth in prices/costs. Investment continues to replace declining production. Terminal capex is recalculated from replacement needs; an arbitrary year-10 override does not carry into perpetuity. The model assumes ongoing reserve replenishment rather than valuing only currently proved reserves.
- **Equity bridge:** borrowings plus lease liabilities, cash/investments, and explicit other-claim/asset inputs. Other claims default to deferred/farmout payables, provisions and income-tax liabilities at book value. Associates use a book-value proxy. These require valuation judgment and can be replaced.
- **Shares:** 104.299705 million period-end shares plus an editable 1.778830 million dilution allowance, based on the cap disclosed in audited Note 11.2. This is distinct from the weighted-average diluted EPS denominator. One ADS represents one Series A share.
- **Timing and scope:** year-end discounting from December 31, 2025. FY2025 includes a partial-year acquisition. No automatic 2026 acquisition, financing, live price, quarterly or ownership-perimeter updates.

## Verification

```powershell
python -m unittest discover -s dcf_model/tests -v
```

Use the bundled Python executable if the shell's Python lacks the reading dependencies. The suite tests audited reconciliation, missing data, period selection, normalization, WACC and debt direction, ADR conversion, leap years, loss taxes, production-linked capex, terminal reinvestment, sensitivity consistency, and saved-input preservation. Each workbook build independently compares 787 calculated cells to the Python engine and exercises case selection, terminal-growth blocking and later-year input changes.

Visual previews and formula calculations are checked with Artifact Tool. Native Microsoft Excel UI behavior is not automated by this project; keep Excel calculation mode on Automatic.

## Primary sources

- [Vista investor relations](https://www.vistaenergy.com/en/investors): historical operating/financial workbook, with per-field source cells recorded in the model.
- Audited 2025 and 2024 consolidated statements, linked in the workbook's Sources sheet and cached for reproducibility.
- [Vista 2024 Form 20-F, Item 12](https://vistaenergy.com/contenidos/VistaFY202420F.pdf): ADS ratio.

Source files are provided for reproducible personal research. The company notes that its historical workbook is unaudited; audited statements prevail. Five key statement totals per year are independently checked against the audited reports; that does not constitute a full audit of every imported metric.
