"""Documentation of the existing default construction, not new forecasts."""

# Key, starting rule, explanation. Sources refer to the workbook's source register.
DRIVER_GUIDE = {
    'oil_growth': ('Base: 15%, 12%, 10%, 8%, 5%, 4%, 3%, 2%, 1%, 0% for 2026–2035. Bear: 5%, 4%, 3%, 2%, then 0%. Bull: 22%, 18%, 15%, 12%, 8%, 6%, 4.5%, 3%, 1.5%, 0%.', 'Illustrative expansion followed by a fade to flat production. These paths were chosen for scenario analysis, not estimated from reserves or management guidance.'),
    'gas_growth': ('Initially the same annual path as oil growth in each scenario.', 'A simplifying starting point. Edit gas independently if its production outlook differs from oil.'),
    'ngl_growth': ('Initially the same annual path as oil growth in each scenario.', 'A simplifying starting point. Edit NGL independently if its production outlook differs from oil.'),
    'oil_price': ('2026–2030: Base uses reported 2025 oil price; Bear uses $55/bbl; Bull uses $75/bbl. In 2031–2035 prices move in five equal steps to $70 / $60 / $80, respectively.', 'Base starts from the realized-price history in S1. Scenario prices and endpoints are illustrative choices, not a forward curve. Prices are net of export duties.'),
    'gas_price': ('Reported 2025 gas price × 1.02^(forecast year − 2026).', 'Anchored to S1, with an illustrative 2% annual nominal escalation. Each annual price is editable; the escalation is only a rule for generating initial values.'),
    'ngl_price': ('2025 NGL revenue / (2025 daily NGL production × 365 / 1,000,000), then 2% annual escalation from 2026.', 'S1 revenue divided by produced boe is a revenue-per-volume proxy, not an independently quoted NGL price.'),
    'lifting': ('2025 operating expenses / annual 2025 production in million boe, then 2% annual escalation from 2026.', 'S1 provides the historical cost intensity. The 2% escalation is illustrative and does not model a separate efficiency gain.'),
    'royalty_rate': ('(2025 royalties − export duties) / (2025 revenue − export duties − $29.8m freight). Held flat initially.', 'Uses normalized S1 history to avoid counting export duties twice. A historical blended ratio, not a statutory royalty rate.'),
    'selling_rate': ('(2025 selling expense − $29.8m freight) / (2025 revenue − export duties − $29.8m freight). Held flat initially.', 'Removes matching freight from both sides. The resulting historical ratio is a starting proxy for future selling costs.'),
    'gna_rate': ('2025 general and administrative expense / normalized 2025 revenue. Held flat initially.', 'Anchored to S1 historical overhead intensity. No separate operating-leverage improvement is assumed.'),
    'dda': ('2025 depreciation and depletion / annual 2025 production in million boe, then 2% annual escalation from 2026.', 'S1 supplies the historical intensity. This is a simple production-based estimate, not a fixed-asset or reserve depletion schedule.'),
    'tax_rate': ('35% of positive forecast EBIT; no tax credit on losses.', 'Illustrative cash-tax assumption. It is not a jurisdiction-by-jurisdiction effective tax forecast and does not model tax losses or deferred taxes.'),
    'nwc_rate': ('(2025 receivables + inventory − current trade payables + $19.236m financing-payable adjustment) / normalized 2025 revenue.', 'Uses S1 balances and the audited current farmout-payable amount. The initial ratio can be negative because suppliers fund part of operations. It is a working-capital proxy.'),
    'decline': ('25% per year in all scenarios.', 'Illustrative natural decline of existing production used to estimate replacement capacity. No field-level engineering calibration supports this specific default.'),
    'replacement_cost': ('$60 per new annual boe of capacity in 2026, then 2% annual escalation.', 'Illustrative capital intensity used for both replacement and expansion. It is not cost per reserve boe or cost per flowing daily boe, and is not calibrated to a drilling plan.'),
    'infrastructure': ('$3 per produced boe in 2026, then 2% annual escalation.', 'Illustrative allowance for infrastructure and intangible investment beyond new production capacity. Review against a company investment plan.'),
    'extra_capex': ('$0m in each year.', 'No additional discrete project is assumed. Enter project-specific investment here if it is not already included in the production-linked calculation.'),
    'capex_override': ('Blank: use production-linked investment plus infrastructure and extra capex. Zero: explicitly use zero total capex.', 'An optional total-investment replacement for a reviewed external capex estimate. A value below modeled investment needs produces a shortfall warning.'),
}

POLICY_GUIDE = {
    'risk_free': ('4.0% in every scenario.', 'Illustrative USD risk-free input. It is not a retrieved Treasury yield for the valuation date.'),
    'beta': ('1.10 in every scenario.', 'Illustrative equity sensitivity to market risk. It was not estimated by a stock-return regression or relevered from a peer group.'),
    'equity_premium': ('5.0% in every scenario.', 'Illustrative mature-market equity risk premium. The model does not derive it from a market-implied premium estimate.'),
    'country_premium': ('Base 6.0%; Bear 9.0%; Bull 4.0%.', 'Illustrative additional country-risk compensation. The Bear case assumes more risk and the Bull case less. These are not observed sovereign spreads.'),
    'debt_cost': ('9.0% before tax in every scenario.', 'Illustrative marginal funding cost, not a retrieved bond yield or the historical accounting interest rate.'),
    'debt_weight': ('25.0% of enterprise capital in every scenario.', 'Illustrative target financing mix. It is not calculated from current equity market value and outstanding debt.'),
    'debt_tax': ('35.0% in every scenario.', 'Illustrative marginal tax shield on debt interest. Review deductibility and effective tax capacity separately from the operating cash-tax assumption.'),
    'wacc_override': ('Blank: calculate WACC from the seven inputs above. Enter a rate to replace that calculation.', 'Cost of equity = risk-free + beta × equity premium + country premium. WACC weights equity cost and after-tax debt cost. This rate must exceed terminal growth.'),
    'terminal_growth': ('2.0% nominal USD growth in every scenario.', 'Illustrative perpetual price/cost growth, with flat real production in year 11. It is not an independently sourced long-term forecast. Replacement investment continues.'),
}

BRIDGE_GUIDE = {
    'shares': ('104.299705 million period-end ordinary shares.', 'Audited 2025 Note 11.2: 104,299,705 shares. Uses period-end shares rather than the weighted-average EPS denominator.'),
    'dilution': ('1.778830 million additional shares.', 'Audited Note 11.2 cap of 106,078,535 less 104,299,705 outstanding shares. A dilution allowance based on the cap, not a forecast of actual issuance.'),
    'adr_ratio': ('1 ordinary share per ADS.', 'Vista 2024 Form 20-F, Item 12 (S4). USD per ordinary share is multiplied by this ratio to obtain USD per ADS.'),
    'cash_adjustment': ('$0m.', 'No extra change to historical cash and short-term investments. Use a negative adjustment for restricted/unavailable cash or a positive adjustment for additional cash.'),
    'debt_adjustment': ('$0m.', 'No extra change to historical borrowings plus lease debt. Positive values increase the debt deduction; negative values reduce it.'),
    'other_claims': ('$508.659m = $311.472m deferred/farmout payables + $62.313m provisions + $134.874m income-tax liabilities.', 'Uses audited 2025 balance-sheet amounts as debt-like book-value proxies. This classification and their economic value are modeling choices, not an independent fair-value assessment.'),
    'nonoperating_assets': ('Reported 2025 investments in associates from S1.', 'Book value is the initial proxy for nonoperating assets. Replace it with an independently reviewed value when available.'),
    'minority_interest': ('$0m.', 'No additional minority-interest deduction is assumed. Review whether a noncontrolling economic claim needs to be valued.'),
    'opening_nwc_adjustment': ('$19.236m.', 'Audited 2025 current farmout payables are removed from operating NWC by adding them back. This keeps the financing claim out of operating working capital.'),
}

FREIGHT_GUIDE = ('2023 and 2024: $0m. 2025: $29.8m.', 'The model uses a rounded 2025 freight normalization. Audited 2025 sea freight is $29.840m; earlier zeros mean no separate normalization was applied. The assumption removes equal amounts from revenue and selling expense. Review and edit by year as needed.')
