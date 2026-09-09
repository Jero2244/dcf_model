# Vista DCF

A modular Python valuation engine and a formula-driven Excel model for Vista Energy (NYSE: VIST). The initial company universe includes Argentine operating exposure; Vista is incorporated in Mexico.

The workbook contains actual FY2023–2025 data, ten annual forecast periods, three editable scenarios, WACC/terminal-growth sensitivity, oil-price sensitivity, and a bridge from enterprise value to USD per ADR. It values the business at **December 31, 2025**, using reports published later to verify historical data. It is not a point-in-time backtest or a current market price target.

## What this project does

Vista DCF helps you explore how assumptions about Vista Energy's production, commodity prices, costs and investment affect its estimated value. DCF means **discounted cash flow**: the model forecasts the cash the business can generate, discounts those future cash flows to the valuation date, and accounts for debt, cash and other claims to estimate value per ADR.

The project combines a Python engine with an Excel workbook you can use directly. Once a workbook has been generated, you can change assumptions and compare results in Excel without running Python.

## What is included

- Reported financial and operating history, with source references and checks against audited statements.
- Annual forecasts for FY2026–2035 and editable Base, Bear and Bull cases.
- Valuation and sensitivity views showing how assumptions affect the result.
- Explanations of the starting assumptions and space to record your own changes.
- Commands to rebuild the workbook or refresh its source data while preserving saved inputs and notes.

## Where to start

Read **[How to use the model](HOW_TO_USE.md)** for a walkthrough of opening the workbook, changing assumptions, understanding the results and saving or refreshing your work.

Read **[Project notes and technical reference](NOTES.md)** for accounting choices, limitations, architecture, dependencies, verification and sources. The detailed reference material from the original README is preserved there; its operating instructions are preserved in the usage guide.
