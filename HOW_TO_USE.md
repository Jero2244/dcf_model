# How to use the Vista DCF model

[Project introduction](README.md) · [Project notes and technical reference](NOTES.md)

This guide is for someone using the model in Excel. You do not need to edit Python code to change assumptions or compare scenarios in an existing workbook. The command-line steps later in this guide are for generating or refreshing a workbook.

## Find and open the workbook

If you cloned the GitHub repository, the clone folder is the `dcf_model` project. Generated workbooks and personal input notes are kept local and are not included in the repository, so start with **Generate or refresh a workbook** to create your first file. The command examples use the existing workspace's parent folder, `UNSAM_Python`; in another location, run them from the folder containing your `dcf_model` clone.

Open the project's `dcf_model/outputs` folder in File Explorer and look for an Excel file named `Vista_Energy_VIST_ADR_DCF_YYYY-MM-DD_HH-MM-SS_UTC.xlsx`. Choose the newest generation timestamp in the filename, or open `latest_workbook.json` in a text editor and use the path labeled `workbook`. Choose the file you want to continue if you have been working on an older version.

Open the `.xlsx` file in Microsoft Excel and read its **Instructions** tab. In Excel, set **Formulas → Calculation Options → Automatic** so results update when you edit inputs.

Keep the matching `.inputs.json` file beside the workbook. This companion file lets the project recover your saved assumptions when generating another version; you do not need to open or edit it. If you rename the workbook, rename the companion to the same base name: for example, `My_Vista_Model.xlsx` and `My_Vista_Model.inputs.json`. Use `-FromWorkbook` when rebuilding or refreshing a renamed copy, because automatic selection looks for generated filenames.

If there is no workbook yet, follow **Generate or refresh a workbook** below, then return here.

## Get familiar with the tabs

| Tab | What you use it for |
| --- | --- |
| Instructions | Read the workbook's overview and reminders. |
| Valuation | See the selected case's estimated business value and value per ADR. |
| Assumptions | Select a case and change its inputs. |
| Assumption Guide | Understand the original defaults and why they were chosen. |
| Input Notes | Find editable inputs and record why you changed them. |
| Forecast | Follow production, revenue, costs, investment and cash flow through the forecast years. |
| Scenarios | Compare Base, Bear and Bull and explore sensitivity to WACC, terminal growth and oil prices. |
| Historical | Read the imported historical figures. |
| Adjustments | Override historical figures while retaining the reported data. |
| Sources | Trace inputs to their sources and review reconciliation checks. |

Financial totals are generally in **USD millions**; the value per ADR is in **USD per ADR**. Read each row's unit before entering a number. An ADR represents shares traded in the US; this model uses one underlying share per ADS.

In this guide, **Assumptions E6** means cell E6 on the Assumptions tab. To jump to a cell, select the tab, type its address into Excel's Name Box (to the left of the formula bar), and press Enter.

Base, Bear and Bull are three separate sets of assumptions. Selecting a case changes the main valuation display; each case's own inputs remain editable.

## Use the workbook

1. Open the generated workbook and select a case in **Assumptions E6**.
2. Edit the pale-yellow input cells in the case sections. Base annual inputs start at row 52, Bear at row 94, and Bull at row 136. Every forecast year has its own inputs. The top section displays the selected case.
3. Review **Valuation** and **Scenarios**. All outputs use Excel formulas and update without Python. Growth, oil and gas prices, production costs, royalties, capex, taxes and WACC are editable.
4. Use **Adjustments E:G** for historical overrides and column N for explanations. A blank keeps the reported figure. Zero is a real override.
5. Save the workbook before refreshing. Keep its adjacent `.inputs.json` manifest with it, including when moving or renaming it.

**Assumption Guide** explains every starting assumption: its initial value or construction, historical source where applicable, why it was used, and where to edit it. The guide documents the original defaults, so it remains a baseline after edits. Growth paths, discount-rate inputs, 35% cash tax, 25% decline and capital-intensity parameters are illustrative choices, not management forecasts. History-based annual inputs are initially seeded from reported 2025 figures; changing history later does not automatically reset those annual inputs.

**Input Notes** has one row for every editable numeric input, including each scenario/year, WACC inputs, equity-bridge inputs, freight normalization and historical overrides. Filter the table by scenario, year or input. Its **Current input** column links to the model; change the number at the address shown in **Edit number at**, then write the reason in **Reason / source / date** (column I). For example: “2030 oil price changed to $75/bbl based on my revised price outlook, reviewed 2026-09-06.” Notes are optional and survive build/refresh alongside the saved numbers. Historical row-level notes in Adjustments N also remain supported.

All 18 annual drivers can be edited independently for all ten years in each case. The selected-case summary and valuation totals are calculated; edit their underlying yellow inputs. Use Adjustments E:G to replace any reported historical numeric metric without overwriting the source data.

All forecast defaults are illustrative starting points, not management guidance. Blue text identifies user inputs, green text identifies links to other sheets, and black text identifies ordinary calculations and imports. Negative residual equity indicates no modeled value left after claims; it is not a prediction of a negative stock price.

### Try one change

1. Select **Base** in **Assumptions E6** and look at its current value per ADR on **Valuation**.
2. On **Input Notes**, filter for Base, year 2030 and the oil-price input. Follow the address in **Edit number at**.
3. Change that yellow input to `75` for $75 per barrel. Enter a number directly; for percentage inputs, type a percentage such as `5%`.
4. Return to **Valuation** and **Scenarios** to see the effect. In **Forecast**, follow the change through revenue and cash flow.
5. Record the reason, source and date in **Input Notes**, column I. Save the workbook, or undo the example change if you were only experimenting.

This is a walkthrough, not a recommended oil-price assumption. Changing one year in Base does not change the same year's Bear or Bull input.

## Read the results

**Valuation** shows the selected case. Enterprise value is the modeled value of the operating business. The equity bridge then accounts for financial debt, cash, other claims and assets, and shares to arrive at value per ADR.

**Scenarios** compares all three cases and shows sensitivities. WACC is the discount rate applied to future cash flows. Terminal growth is the nominal long-run growth assumption after the explicit forecast. The oil sensitivity shifts every selected-case oil price; it is different from editing a single forecast year.

Review the production and investment figures in **Forecast**, along with any warnings. A capex shortfall means the investment you entered is below the model's estimate of what the production plan needs. WACC must exceed terminal growth for the model's terminal-value calculation to work.

The valuation date stays **December 31, 2025**, even if you generate or refresh a file later. For the assumptions behind taxes, leases, terminal investment and the equity bridge, read the [accounting and valuation choices](NOTES.md#accounting-and-valuation-choices).

## Save your work

Save in Excel before using any build or refresh command: the project reads the file saved on disk. Keep the workbook and its matching `.inputs.json` together when copying, moving or renaming them.

Build and refresh preserve the supported input values and explanatory notes in a newly generated workbook. Treat them as regeneration steps: do not rely on them to carry over custom formatting, added sheets or edits to calculated formulas. Keep your saved original if you make those kinds of changes.

## Generate or refresh a workbook

Use these steps when you need a new Excel file. Editing an existing workbook does not require them.

Open the `UNSAM_Python` folder in File Explorer, then open PowerShell in that folder (for example, type `powershell` in File Explorer's address bar and press Enter). Run one command at a time. The paths below assume you are in that parent folder, which contains `dcf_model`.

- **build** creates a new workbook from the included cached sources.
- **refresh** downloads source data and creates a new workbook.
- **value** prints the selected case's Python valuation without creating Excel.

If this is a different computer, first read [Architecture and dependencies](NOTES.md#architecture-and-dependencies): generating Excel requires the runtime described there, in addition to Python.

### Commands

From the `UNSAM_Python` directory in PowerShell:

```powershell
# Rebuild offline using the included official source cache.
.\dcf_model\run.ps1 build

# Download the latest Vista historical workbook, verify audited totals,
# preserve inputs from the latest generated model, and create a new version.
.\dcf_model\run.ps1 refresh

# Explicitly preserve a saved workbook's assumptions.
.\dcf_model\run.ps1 refresh -FromWorkbook 'C:\path\to\Vista_Energy_VIST_ADR_DCF_2026-09-06_23-06-17_UTC.xlsx'

# Use an assumption JSON file instead of a saved workbook.
.\dcf_model\run.ps1 build -Config 'C:\path\to\model_assumptions.json'

# Print the Python valuation without creating Excel.
.\dcf_model\run.ps1 value
```

Network access is required for `refresh`. `build` and `value` work offline with the included source cache. Without an explicit `-Config` or `-FromWorkbook`, the most recent generated workbook supplies assumptions. A refresh produces a timestamped version and leaves the prior workbook unchanged. Numeric input values and explanatory notes are preserved. Formulas manually inserted into input cells are rejected during refresh with a location-specific message; they are not silently discarded.

The default output directory is `dcf_model/outputs`, regardless of the working directory. Workbooks are named `Vista_Energy_VIST_ADR_DCF_YYYY-MM-DD_HH-MM-SS_UTC.xlsx`; the date and time indicate when the file was generated, not the valuation date. Builds within the same second receive a numbered suffix to preserve earlier files. Keep each workbook's matching `.inputs.json` file beside it so refresh can recover saved inputs. Older `vista_dcf_*.xlsx` filenames are also supported.

When generation succeeds, the command prints `Created` followed by the new workbook's path. Open that file to continue. To rebuild a specific saved workbook offline, use `build -FromWorkbook` with its path, just as in the refresh example above. Supply either `-Config` or `-FromWorkbook`, not both.

Automatic selection uses the newest **generation timestamp**, not the file you edited most recently. Use `-FromWorkbook` whenever you want to be certain which saved model is carried forward. With no previous workbook and no explicit configuration, the project uses its starting defaults.

## If something does not work

| What you see | What to do |
| --- | --- |
| Results do not change after editing an input | Check Excel's calculation mode is Automatic, confirm the selected case, and check that you edited that case's yellow input. |
| A missing input manifest message | Restore the matching `.inputs.json` beside the workbook with the matching base name, or use an existing assumption JSON with `-Config`. |
| A message naming a formula in an input cell | Replace that input formula with its intended numeric value, save the workbook and retry. |
| A refresh fails while downloading sources | Check your internet connection. You can use `build` offline if the required cache is available. |
| A source-layout or fiscal-year validation error | The source needs technical review. This version supports reviewed FY2023–2025 history; refreshing does not automatically roll the model into a new fiscal-year window. |
| A missing Python package, Node or Artifact Tool error | Check the required runtime and dependencies in [Project notes](NOTES.md#architecture-and-dependencies). |

For output-file descriptions, calculation checks and model limitations, see [Project notes and technical reference](NOTES.md).
