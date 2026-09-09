"""Pure-Python upstream FCFF engine. No I/O or spreadsheet dependencies."""
from __future__ import annotations

import calendar
import copy
import math

DRIVERS = [
    ("oil_growth", "Oil production growth", "%"),
    ("gas_growth", "Gas production growth", "%"),
    ("ngl_growth", "NGL production growth", "%"),
    ("oil_price", "Oil price, net of export duties", "USD/bbl"),
    ("gas_price", "Gas price", "USD/MMBtu"),
    ("ngl_price", "NGL revenue per produced boe", "USD/boe"),
    ("lifting", "Lifting cost", "USD/boe"),
    ("royalty_rate", "Royalties excluding export duties", "% revenue"),
    ("selling_rate", "Selling expenses excluding freight", "% revenue"),
    ("gna_rate", "General and administrative expenses", "% revenue"),
    ("dda", "Depreciation and depletion", "USD/boe"),
    ("tax_rate", "Unlevered cash-tax rate", "% positive EBIT"),
    ("nwc_rate", "Operating working capital", "% revenue"),
    ("decline", "Annual natural production decline", "%"),
    ("replacement_cost", "Investment per new annual boe capacity", "USD/annual boe"),
    ("infrastructure", "Infrastructure and intangible investment", "USD/boe"),
    ("extra_capex", "Additional capital investment", "USD m"),
    ("capex_override", "Total capex override (blank = calculated)", "USD m"),
]
POLICIES = [
    ("risk_free", "USD risk-free rate", "%"),
    ("beta", "Levered equity beta", "x"),
    ("equity_premium", "Mature-market equity risk premium", "%"),
    ("country_premium", "Additional country risk premium", "%"),
    ("debt_cost", "Pre-tax cost of debt", "%"),
    ("debt_weight", "Target debt / enterprise capital", "%"),
    ("debt_tax", "Marginal tax rate for debt shield", "%"),
    ("wacc_override", "WACC override (blank = calculated)", "%"),
    ("terminal_growth", "Terminal nominal USD growth", "%"),
]
BRIDGE = [
    ("shares", "Period-end ordinary shares", "million shares"),
    ("dilution", "Additional dilution allowance", "million shares"),
    ("adr_ratio", "Ordinary shares per ADR / ADS", "shares/ADS"),
    ("cash_adjustment", "Adjustment to cash and short-term investments", "USD m"),
    ("debt_adjustment", "Adjustment to financial debt", "USD m"),
    ("other_claims", "Other debt-like claims", "USD m"),
    ("nonoperating_assets", "Additional nonoperating assets", "USD m"),
    ("minority_interest", "Minority interest value", "USD m"),
    ("opening_nwc_adjustment", "Remove current financing payables from NWC", "USD m"),
]


def normalized(h: dict, freight: float) -> dict:
    revenue = h["revenue"] - h["export_duties"] - freight
    royalties = h["royalties"] - h["export_duties"]
    selling = h["selling"] - freight
    ebit = revenue - h["opex"] - royalties - selling - h["gna"] - h["dda"]
    return dict(revenue=revenue, royalties=royalties, selling=selling, ebit=ebit,
                excluded_ebit=h["ebit"]-ebit,
                nwc=h["receivables"]+h["inventory"]-h["payables"])


def default_assumptions(data: dict) -> dict:
    h = data["history"]["2025"]
    n = normalized(h, 29.8)
    annual_boe = h["production"]*365/1e6
    years = list(range(2026, 2036))
    scenarios = {}
    for case, growth, oil, final_oil, country in [
        ("Base", [.15,.12,.10,.08,.05,.04,.03,.02,.01,0], h["oil_price"],70,.06),
        ("Bear", [.05,.04,.03,.02,0,0,0,0,0,0],55,60,.09),
        ("Bull", [.22,.18,.15,.12,.08,.06,.045,.03,.015,0],75,80,.04),
    ]:
        drivers = {}
        for i, year in enumerate(years):
            transition = max(0, i-4)/5
            inflation = 1.02**i
            drivers[str(year)] = dict(
                oil_growth=growth[i], gas_growth=growth[i], ngl_growth=growth[i],
                oil_price=oil+(final_oil-oil)*transition,
                gas_price=h["gas_price"]*inflation,
                ngl_price=h["ngl_revenue"]/(h["ngl_production"]*365/1e6)*inflation,
                lifting=h["opex"]/annual_boe*inflation,
                royalty_rate=n["royalties"]/n["revenue"], selling_rate=n["selling"]/n["revenue"],
                gna_rate=h["gna"]/n["revenue"], dda=h["dda"]/annual_boe*inflation,
                tax_rate=.35, nwc_rate=(n["nwc"]+19.236)/n["revenue"], decline=.25,
                replacement_cost=60*inflation, infrastructure=3*inflation,
                extra_capex=0, capex_override=None)
        scenarios[case] = dict(drivers=drivers, policies=dict(risk_free=.04, beta=1.1,
            equity_premium=.05,country_premium=country,debt_cost=.09,debt_weight=.25,
            debt_tax=.35,wacc_override=None,terminal_growth=.02))
    return dict(schema_version=1, selected_case="Base", years=years, scenarios=scenarios,
                bridge=dict(shares=104.299705, dilution=1.778830, adr_ratio=1,
                            cash_adjustment=0, debt_adjustment=0,
                            other_claims=311.472+62.313+134.874,
                            nonoperating_assets=h["associates"], minority_interest=0,
                            opening_nwc_adjustment=19.236),
                freight={"2023":0,"2024":0,"2025":29.8}, overrides={}, override_notes={}, input_notes={})


def adjusted_history(data: dict, config: dict) -> dict:
    result = copy.deepcopy(data["history"])
    for year, fields in config.get("overrides", {}).items():
        if year not in result:
            raise ValueError(f"Unknown override year: {year}")
        for key, value in fields.items():
            if key not in result[year]:
                raise ValueError(f"Unknown historical field: {key}")
            if value is not None:
                number(value, f"history {year} {key}")
                result[year][key] = value
    return result


def number(value, label):
    if isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value):
        raise ValueError(f"Missing or nonnumeric input: {label}")
    return value


def wacc(p: dict) -> float:
    for key, _, _ in POLICIES:
        if key != "wacc_override":
            number(p.get(key), key)
    if not 0 <= p["debt_weight"] <= 1 or not 0 <= p["debt_tax"] <= 1:
        raise ValueError("Debt weight and marginal tax rate must be between zero and one")
    calculated = ((p["risk_free"]+p["beta"]*p["equity_premium"]+p["country_premium"])
                  *(1-p["debt_weight"])+p["debt_cost"]*(1-p["debt_tax"])*p["debt_weight"])
    value = p.get("wacc_override")
    return calculated if value is None else number(value, "WACC override")


def validate(config: dict):
    if config.get("schema_version") != 1:
        raise ValueError("Unsupported assumption schema version")
    if config["years"] != list(range(2026,2036)):
        raise ValueError("The model requires a continuous FY2026–2035 forecast")
    if config["selected_case"] not in ("Base","Bear","Bull"):
        raise ValueError("Select Base, Bear or Bull")
    b = config["bridge"]
    for key, _, _ in BRIDGE:
        number(b.get(key), key)
    if b["shares"] <= 0 or b["dilution"] < 0 or b["adr_ratio"] <= 0:
        raise ValueError("Shares and ADR ratio must be positive; dilution cannot be negative")
    for year in ("2023","2024","2025"):
        if number(config["freight"].get(year), f"freight {year}") < 0:
            raise ValueError("Freight reimbursements cannot be negative")
    for case in ("Base","Bear","Bull"):
        c = config["scenarios"][case]
        p = c["policies"]
        rate = wacc(p)
        if rate <= p["terminal_growth"] or rate <= -1 or p["terminal_growth"] <= -1:
            raise ValueError(f"{case}: WACC must exceed terminal growth and both must exceed -100%")
        for year in config["years"]:
            d = c["drivers"][str(year)]
            for key,_,_ in DRIVERS:
                if key == "capex_override" and d.get(key) is None:
                    continue
                value = number(d.get(key),f"{case} {year} {key}")
                if key.endswith("growth") and value <= -1:
                    raise ValueError("Production growth must exceed -100%")
                if key in ("decline","tax_rate","royalty_rate","selling_rate","gna_rate") and not 0 <= value <= 1:
                    raise ValueError(f"{key} must be between zero and one")
                if key not in ("oil_growth","gas_growth","ngl_growth","nwc_rate") and value < 0:
                    raise ValueError(f"{key} cannot be negative")


def forecast(data: dict, config: dict, case: str) -> dict:
    validate(config)
    history = adjusted_history(data, config)
    h = history["2025"]
    for key in ("oil_production","gas_production","ngl_production","production","oil_sales","gas_sales"):
        if h[key] <= 0:
            raise ValueError(f"Historical {key} must be positive")
    n = normalized(h, config["freight"]["2025"])
    b, c = config["bridge"], config["scenarios"][case]
    rate, g = wacc(c["policies"]), c["policies"]["terminal_growth"]
    previous = dict(oil=h["oil_production"]*365/1e6,
                    gas=h["gas_production"]*365/1e6,
                    ngl=h["ngl_production"]*365/1e6,
                    nwc=n["nwc"]+b["opening_nwc_adjustment"])
    oil_factor = h["oil_sales"]/previous["oil"]
    gas_factor = h["gas_sales"]/previous["gas"]
    previous_days = 365
    rows, warnings = [], []
    for index, year in enumerate(config["years"]):
        d = c["drivers"][str(year)]
        days = 366 if calendar.isleap(year) else 365
        oil = previous["oil"]*(1+d["oil_growth"])*days/previous_days
        gas = previous["gas"]*(1+d["gas_growth"])*days/previous_days
        ngl = previous["ngl"]*(1+d["ngl_growth"])*days/previous_days
        total = oil+gas+ngl
        oil_revenue, gas_revenue, ngl_revenue = oil*oil_factor*d["oil_price"],gas*gas_factor*d["gas_price"],ngl*d["ngl_price"]
        revenue = oil_revenue+gas_revenue+ngl_revenue
        lifting, royalties, selling, gna, dda = total*d["lifting"],revenue*d["royalty_rate"],revenue*d["selling_rate"],revenue*d["gna_rate"],total*d["dda"]
        ebitda = revenue-lifting-royalties-selling-gna
        ebit = ebitda-dda
        tax = max(ebit,0)*d["tax_rate"]
        previous_total = (previous["oil"]+previous["gas"]+previous["ngl"])*days/previous_days
        new_capacity = max(0,total-previous_total*(1-d["decline"]))
        required_capex = new_capacity*d["replacement_cost"]+total*d["infrastructure"]
        capex = required_capex+d["extra_capex"] if d["capex_override"] is None else d["capex_override"]
        nwc = revenue*d["nwc_rate"]
        delta_nwc = nwc-previous["nwc"]
        fcf = ebit-tax+dda-capex-delta_nwc
        row = dict(year=year,days=days,oil=oil,gas=gas,ngl=ngl,total=total,
                   oil_revenue=oil_revenue,gas_revenue=gas_revenue,ngl_revenue=ngl_revenue,revenue=revenue,
                   lifting=lifting,royalties=royalties,selling=selling,gna=gna,ebitda=ebitda,
                   dda=dda,ebit=ebit,tax=tax,new_capacity=new_capacity,required_capex=required_capex,
                   capex=capex,nwc=nwc,delta_nwc=delta_nwc,fcf=fcf,discount_factor=(1+rate)**-(index+1),
                   pv=fcf/(1+rate)**(index+1),capex_gap=capex-required_capex)
        rows.append(row)
        if capex < required_capex-1e-8 or (new_capacity > 0 and d["replacement_cost"] == 0):
            warnings.append(f"{year}: investment is below the modeled production requirement")
        previous, previous_days = row, days
    last = rows[-1]
    d = c["drivers"][str(config["years"][-1])]
    # Terminal: no real volume growth; nominal revenues/costs grow by g.
    # Normalize reinvestment to replacement of declining production, not year-10 growth.
    terminal_capex = (last["total"]*d["decline"]*d["replacement_cost"]
                      +last["total"]*d["infrastructure"]+d["extra_capex"])*(1+g)
    terminal = dict(revenue=last["revenue"]*(1+g),ebit=last["ebit"]*(1+g),
                    dda=last["dda"]*(1+g),tax=max(last["ebit"]*(1+g),0)*d["tax_rate"],
                    capex=terminal_capex,nwc=last["nwc"]*(1+g),delta_nwc=last["nwc"]*g)
    terminal["fcf"] = terminal["ebit"]-terminal["tax"]+terminal["dda"]-terminal["capex"]-terminal["delta_nwc"]
    if terminal["fcf"] <= 0:
        warnings.append("Terminal cash flow is nonpositive; review sustainable operations")
    tv = terminal["fcf"]/(rate-g)
    pv_terminal = tv/(1+rate)**10
    ev = sum(r["pv"] for r in rows)+pv_terminal
    debt = h["debt_nc"]+h["debt_current"]+h["lease_nc"]+h["lease_current"]+b["debt_adjustment"]
    cash = h["cash"]+b["cash_adjustment"]
    equity = ev+cash+b["nonoperating_assets"]-debt-b["other_claims"]-b["minority_interest"]
    return dict(case=case,rows=rows,terminal=terminal,wacc=rate,terminal_growth=g,
                terminal_value=tv,pv_terminal=pv_terminal,enterprise_value=ev,debt=debt,cash=cash,
                equity_value=equity,value_per_adr=equity/(b["shares"]+b["dilution"])*b["adr_ratio"],
                terminal_share=pv_terminal/ev if ev else None,warnings=warnings)


def sensitivities(data: dict, config: dict) -> dict:
    case = config["selected_case"]
    base_rate = wacc(config["scenarios"][case]["policies"])
    base_growth = config["scenarios"][case]["policies"]["terminal_growth"]
    rates = [base_rate+x for x in (-.02,-.01,0,.01,.02)]
    growths = [base_growth+x for x in (-.01,-.005,0,.005,.01)]
    grid = []
    for rate in rates:
        row = []
        for growth in growths:
            changed = copy.deepcopy(config)
            changed["scenarios"][case]["policies"].update(wacc_override=rate,terminal_growth=growth)
            row.append(forecast(data,changed,case)["value_per_adr"] if rate > growth and min(rate,growth)>-1 else None)
        grid.append(row)
    return dict(rates=rates,growths=growths,values=grid)
