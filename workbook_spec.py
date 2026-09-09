"""Build an auditable cell specification for the Artifact Tool XLSX author."""
from __future__ import annotations

from .data_sources import METRICS, AUDIT_URLS, ADR_URL
from .model import BRIDGE, DRIVERS, POLICIES, forecast, sensitivities
from .assumption_explanations import DRIVER_GUIDE, POLICY_GUIDE, BRIDGE_GUIDE, FREIGHT_GUIDE

CASES = ["Base", "Bear", "Bull"]
CASE_ROWS = {"Base":52,"Bear":94,"Bull":136}
FORECAST_ROWS = {"Selected":10,"Base":44,"Bear":78,"Bull":112}
HROWS = {m[0]:8+i for i,m in enumerate(METRICS)}
DROWS = {m[0]:i for i,m in enumerate(DRIVERS)}
PROWS = {m[0]:i for i,m in enumerate(POLICIES)}
BROWS = {m[0]:176+i for i,m in enumerate(BRIDGE)}
FIELDS = [
    ("oil","Oil production","million bbl"),("gas","Gas production","million boe"),
    ("ngl","NGL production","million boe"),("total","Total production","million boe"),
    ("oil_sales","Oil sales volume","million bbl"),("gas_sales","Gas sales volume","million MMBtu"),
    ("oil_revenue","Oil revenue, net of export duties","USD m"),("gas_revenue","Gas revenue","USD m"),
    ("ngl_revenue","NGL revenue","USD m"),("revenue","Total modeled revenue","USD m"),
    ("lifting","Lifting costs","USD m"),("royalties","Royalties excluding export duties","USD m"),
    ("selling","Selling expenses excluding freight","USD m"),("gna","General and administrative expenses","USD m"),
    ("ebitda","Core EBITDA","USD m"),("dda","Depreciation and depletion","USD m"),
    ("ebit","Core EBIT","USD m"),("tax","Unlevered cash taxes","USD m"),
    ("new_capacity","Required new annual production capacity","million boe"),
    ("required_capex","Investment required by production plan","USD m"),("capex","Total capital investment","USD m"),
    ("nwc","Operating working capital","USD m"),("delta_nwc","Change in working capital","USD m"),
    ("fcf","Unlevered free cash flow","USD m"),("discount_factor","Discount factor","x"),
    ("pv","Present value of free cash flow","USD m"),("capex_gap","Investment less modeled requirement","USD m"),
    ("investment_warning","Investment shortfall indicator","0 / 1"),
]
FROWS = {k:i for i,(k,_,_) in enumerate(FIELDS)}
NUMBER = '#,##0.0;(#,##0.0);"-"'
PERCENT = '0.0%;(0.0%);"-"'
PRICE = '$0.00;($0.00);"-"'


def column(index: int) -> str:
    result = ""
    while index:
        index, r = divmod(index-1,26)
        result = chr(65+r)+result
    return result


def build_spec(data: dict, config: dict) -> dict:
    sheets = []
    inputs = []
    comparisons = []
    by_name = {}

    def sheet(name,title,subtitle,last_row=70,last_col="N"):
        s=dict(name=name,title=title,subtitle=subtitle,last_row=last_row,last_col=last_col,
               cells=[],styles=[],widths={},renders=[])
        sheets.append(s); by_name[name]=s
        return s

    def put(s,addr,value,role="body",fmt=None):
        cell=dict(address=addr,role=role)
        cell["formula" if isinstance(value,str) and value.startswith("=") else "value"]=value
        if fmt: cell["number_format"]=fmt
        s["cells"].append(cell)

    def inp(s,addr,value,path,fmt=NUMBER):
        put(s,addr,value,"input",fmt)
        inputs.append(dict(sheet=s["name"],cell=addr,path=path))

    def header(s,row,labels,start=3):
        for i,label in enumerate(labels,start): put(s,f"{column(i)}{row}",label,"header")

    def label(s,row,text,unit=None):
        put(s,f"C{row}",text)
        if unit: put(s,f"D{row}",unit,"note")

    def h(key,col="L"):
        return f"'Adjustments'!{col}{HROWS[key]}"

    def dr(case,key,col): return f"'Assumptions'!{col}{CASE_ROWS[case]+DROWS[key]}"
    def pr(case,key): return f"'Assumptions'!$E${CASE_ROWS[case]+21+PROWS[key]}"
    def fr(case,key,col): return f"'Forecast'!{col}{FORECAST_ROWS[case]+FROWS[key]}"
    def br(key): return f"'Assumptions'!$E${BROWS[key]}"
    def polw(case): return f"'Assumptions'!$E${CASE_ROWS[case]+32}"

    s=sheet("Instructions","Vista DCF","FY2023–2025 actuals and FY2026–2035 estimates. Nominal USD.",45,"H")
    s["widths"]={"C":38,"D":105,"E":2,"F":2,"G":2,"H":2}
    instructions=[
        (6,"Model date","December 31, 2025. Later-published reports verify history. This is not a point-in-time backtest."),
        (9,"1. Choose assumptions","Select Base, Bear or Bull in Assumptions E6. Blue text on pale yellow marks editable inputs."),
        (12,"2. Edit the annual drivers","Change production growth, realized prices, costs, cash taxes and capex assumptions by year. Defaults are illustrative, not management guidance."),
        (15,"3. Review value and scenarios","Valuation shows the selected case. Scenarios compares all cases and recalculating sensitivities."),
        (18,"4. Review historical adjustments","Historical contains imported values. Enter overrides in Adjustments E:G and explain them in N. Blank means use the reported figure; zero is a valid override."),
        (21,"5. Save before refreshing","Save Excel edits first. The refresh command reads the saved workbook, preserves numeric assumptions and notes, and creates a new version."),
        (24,"Production and capital investment","Annual volumes follow average daily production growth. Required investment replaces natural decline and adds new capacity. An override below that requirement displays a warning."),
        (27,"Terminal operations","Year 11 assumes flat real production and nominal USD price/cost growth. Replacement investment continues indefinitely. This is a going-concern DCF, not a reserve-depletion NAV."),
        (30,"Important accounting choices","Realized oil prices are net of export duties. Freight and export duties are removed from both revenue and matching expense. Other operating gains and impairments are excluded from core EBIT."),
        (33,"Valuation boundaries","No 2026 transactions, live share price, FX forecast or tax-loss schedule. Review equity-bridge proxies and 2025 acquisition effects before investment use."),
        (36,"Reading the model","Blue = editable inputs; green = links to another sheet; black = calculations/imports. Financial totals are USD millions; ADR values are USD per ADS."),
        (39,"Why these numbers?","Assumption Guide explains the original defaults and separates historical anchors from illustrative choices. Current numbers may differ after your edits."),
        (42,"Record why you changed an input","In Input Notes, filter by scenario, year or input. Edit the number at the listed cell, then write your reason in column I. Include your source and date. Save before refreshing."),
    ]
    for row,text,note in instructions:
        label(s,row,text); put(s,f"D{row}",note,"wide_note")
    s["renders"]=["C2:D20","C21:D44"]

    s=sheet("Valuation","Vista Valuation","As of December 31, 2025. Year-end discounting. USD millions except ADR values.",49,"H")
    s["renders"]=["C2:H29","C30:H47"]

    a=sheet("Assumptions","Vista Assumptions","Edit yellow scenario inputs below. See Assumption Guide for defaults and Input Notes to record reasons.",190,"N")
    put(a,"G6","Selected-case summary: edit the yellow cells in each scenario below.","note")
    a["freeze_rows"]=8; a["freeze_columns"]=4
    inp(a,"E6",config["selected_case"],["selected_case"],None)
    label(a,6,"Selected Case")
    a["validations"]=[dict(range="E6",values=CASES)]
    put(a,"D8","Forecast year","header")
    for i,year in enumerate(config["years"],5):
        put(a,f"{column(i)}8",f"=DATE({year},12,31)","header",'"FY "yyyy')
    for i,(key,text,unit) in enumerate(DRIVERS):
        row=10+i; label(a,row,text,unit)
        for col in map(column,range(5,15)):
            # Optional blank capex overrides must remain blank through selection.
            refs=[f'{col}{CASE_ROWS[c]+i}' for c in CASES]
            options=[f'IF({r}="","",{r})' if key=="capex_override" else r for r in refs]
            put(a,f"{col}{row}",f'=IF($E$6="Base",{options[0]},IF($E$6="Bear",{options[1]},{options[2]}))',fmt=PERCENT if unit.startswith("%") else NUMBER)
    for i,(key,text,unit) in enumerate(POLICIES):
        row=32+i; label(a,row,text,unit)
        refs=[f'E{CASE_ROWS[c]+21+i}' for c in CASES]
        options=[f'IF({r}="","",{r})' if key=="wacc_override" else r for r in refs]
        put(a,f"E{row}",f'=IF($E$6="Base",{options[0]},IF($E$6="Bear",{options[1]},{options[2]}))',fmt=NUMBER if unit=="x" else PERCENT)
    label(a,41,"Cost of equity", "%"); put(a,"E41","=E32+E33*E34+E35",fmt=PERCENT)
    label(a,42,"Selected WACC", "%"); put(a,"E42",'=IF(E39="",E41*(1-E37)+E36*(1-E38)*E37,E39)',fmt=PERCENT)
    for case,start in CASE_ROWS.items():
        header(a,start-4,[f"{case} assumptions"])
        for i,year in enumerate(config["years"],5):
            col=column(i); put(a,f"{col}{start-2}",f"={col}$8","header",'"FY "yyyy')
        for j,(key,text,unit) in enumerate(DRIVERS):
            row=start+j; label(a,row,text,unit)
            for i,year in enumerate(config["years"],5):
                inp(a,f"{column(i)}{row}",config["scenarios"][case]["drivers"][str(year)][key],
                    ["scenarios",case,"drivers",str(year),key],PERCENT if unit.startswith("%") else NUMBER)
        for j,(key,text,unit) in enumerate(POLICIES):
            row=start+21+j; label(a,row,text,unit)
            inp(a,f"E{row}",config["scenarios"][case]["policies"][key],
                ["scenarios",case,"policies",key],NUMBER if unit=="x" else PERCENT)
        label(a,start+31,"Cost of equity", "%")
        put(a,f"E{start+31}",f"=E{start+21}+E{start+22}*E{start+23}+E{start+24}",fmt=PERCENT)
        label(a,start+32,"Calculated or overridden WACC", "%")
        put(a,f"E{start+32}",f'=IF(E{start+28}="",E{start+31}*(1-E{start+26})+E{start+25}*(1-E{start+27})*E{start+26},E{start+28})',fmt=PERCENT)
    header(a,174,["Equity bridge inputs"])
    for i,(key,text,unit) in enumerate(BRIDGE):
        label(a,176+i,text,unit)
        inp(a,f"E{176+i}",config["bridge"][key],["bridge",key],NUMBER)
    put(a,"G176","Period-end shares and dilution allowance: audited Note 11.2. Not the weighted-average EPS denominator.","side_note")
    put(a,"G180","Other claims = deferred/farmout payables 311.472 + provisions 62.313 + income tax payable 134.874. Book-value proxies.","side_note")
    put(a,"G184","Capital investment includes owned and leased productive assets. Lease debt is deducted in the equity bridge; lease principal is not deducted again in FCFF.","side_note")
    a["renders"]=["C2:J27","C30:J43","C48:J70","C73:J86","C90:J112","C115:J128","C132:J154","C157:J169","C174:J188"]

    hist=sheet("Historical","Vista Historical Data","Company-reported data. Financial totals converted from USD thousands where applicable.",67,"H")
    adj=sheet("Adjustments","Vista Historical Adjustments","Blank overrides retain reported values. Enter a reason for each manual override.",80,"N")
    hist["freeze_rows"]=7; adj["freeze_rows"]=7
    header(hist,7,["Metric","Units",2023,2024,2025,"Source"])
    header(adj,7,["Metric","Units","2023 override","2024 override","2025 override","","","2023 applied","2024 applied","2025 applied","","Reason"])
    adj["widths"]={"N":65}
    for i,(key,text,unit,*_) in enumerate(METRICS):
        row=8+i; label(hist,row,text,unit); label(adj,row,text,unit)
        put(hist,f"H{row}","S1","note")
        for j,year in enumerate(("2023","2024","2025"),5):
            col=column(j); outcol=column(j+5)
            put(hist,f"{col}{row}",data["history"][year][key],fmt=NUMBER)
            inp(adj,f"{col}{row}",config.get("overrides",{}).get(year,{}).get(key),["overrides",year,key])
            put(adj,f"{outcol}{row}",f'=IF({col}{row}="",\'Historical\'!{col}{row},{col}{row})',"link",NUMBER)
        inp(adj,f"N{row}",config.get("override_notes",{}).get(key,""),["override_notes",key],None)
    label(adj,58,"Freight revenue and matching expense", "USD m")
    for j,year in enumerate(("2023","2024","2025"),5):
        inp(adj,f"{column(j)}58",config["freight"][year],["freight",year])
    normal_labels={60:"Revenue excluding duties and freight",61:"Royalties excluding export duties",62:"Selling expenses excluding freight",63:"Core EBITDA",64:"Depreciation and depletion",65:"Core EBIT",66:"Other items excluded from reported EBIT",67:"Current operating NWC proxy",68:"Oil sales / production factor",69:"Gas sales MMBtu / produced boe"}
    for row,text in normal_labels.items(): label(adj,row,text,"x" if row>=68 else "USD m")
    for j,year in enumerate((2023,2024,2025),5):
        col=column(j); src=column(j+5); days=366 if year==2024 else 365
        ref=lambda k:f"{src}{HROWS[k]}"
        formulas={60:f"{ref('revenue')}-{ref('export_duties')}-{col}58",
                  61:f"{ref('royalties')}-{ref('export_duties')}",62:f"{ref('selling')}-{col}58",
                  63:f"{col}60-{ref('opex')}-{col}61-{col}62-{ref('gna')}",64:ref('dda'),
                  65:f"{col}63-{col}64",66:f"{ref('ebit')}-{col}65",
                  67:f"{ref('receivables')}+{ref('inventory')}-{ref('payables')}",
                  68:f"{ref('oil_sales')}/({ref('oil_production')}*{days}/1000000)",
                  69:f"{ref('gas_sales')}/({ref('gas_production')}*{days}/1000000)"}
        for row,formula in formulas.items(): put(adj,f"{col}{row}","="+formula,fmt=NUMBER)
        for row,text in [(60,"Balance sheet difference"),(61,"Revenue component difference"),(62,"Cash reconciliation difference"),(63,"Production rounding difference (boe/day)")]:
            label(hist,row,text)
        hr=lambda k:f"{col}{HROWS[k]}"
        checks={60:f"{hr('assets')}-{hr('liabilities')}-{hr('equity')}",61:f"{hr('revenue')}-{hr('oil_revenue')}-{hr('gas_revenue')}-{hr('ngl_revenue')}",
                62:f"{hr('closing_cash')}-{hr('opening_cash')}-{hr('cfo')}-{hr('cfi')}-{hr('cff')}-{hr('fx_cash')}",
                63:f"{hr('production')}-{hr('oil_production')}-{hr('gas_production')}-{hr('ngl_production')}"}
        for row,formula in checks.items(): put(hist,f"{col}{row}","="+formula,fmt=NUMBER)
    put(adj,"C73","Normalization notes","header")
    put(adj,"C75","Core EBIT excludes the residual of other income, impairments, exploration and other non-cash operating items. It is a model subtotal, not Vista's adjusted EBITDA.","long_note")
    put(adj,"C78","Historical NWC is current receivables + inventories - current trade payables. The opening financing-payable adjustment is in Assumptions. Capex uses a production-linked forecast, not historical acquisition cash outflows.","long_note")
    hist["renders"]=["C2:H29","C30:H54","C58:H65"]
    adj["renders"]=["C2:L29","C30:L54","C56:L79"]

    fs=sheet("Forecast","Vista Production and Cash Flow","Five detailed years, five transition years. All scenarios calculate independently.",145,"N")
    fs["freeze_rows"]=8; fs["freeze_columns"]=4
    for case,start in FORECAST_ROWS.items():
        header(fs,start-3,[f"{case} forecast"])
        for j,year in enumerate(config["years"],5):
            col=column(j); put(fs,f"{col}{start-2}",f"='Assumptions'!{col}$8","header",'"FY "yyyy')
        for k,(key,text,unit) in enumerate(FIELDS): label(fs,start+k,text,unit)
        for j,year in enumerate(config["years"],5):
            col=column(j); prev=column(j-1); i=j-5
            loc=lambda key:f"{col}{start+FROWS[key]}"
            old=lambda key:f"{prev}{start+FROWS[key]}"
            if case=="Selected":
                for key,_,unit in FIELDS:
                    put(fs,loc(key),f'=IF(\'Assumptions\'!$E$6="Base",{fr("Base",key,col)},IF(\'Assumptions\'!$E$6="Bear",{fr("Bear",key,col)},{fr("Bull",key,col)}))',"link",NUMBER)
                continue
            d=lambda key:dr(case,key,col)
            days=366 if year%4==0 else 365
            lastdays=366 if (year-1)%4==0 else 365
            prior={k:(f"{h(k+'_production')}*365/1000000" if i==0 else old(k)) for k in ('oil','gas','ngl')}
            prevtotal=f"({prior['oil']}+{prior['gas']}+{prior['ngl']})*{days}/{lastdays}"
            pnwc=f"'Adjustments'!G67+{br('opening_nwc_adjustment')}" if i==0 else old('nwc')
            formulas={
                **{k:f"({prior[k]})*(1+{d(k+'_growth')})*{days}/{lastdays}" for k in ('oil','gas','ngl')},
                'total':f"SUM({loc('oil')}:{loc('ngl')})",'oil_sales':f"{loc('oil')}*'Adjustments'!$G$68",
                'gas_sales':f"{loc('gas')}*'Adjustments'!$G$69",'oil_revenue':f"{loc('oil_sales')}*{d('oil_price')}",
                'gas_revenue':f"{loc('gas_sales')}*{d('gas_price')}",'ngl_revenue':f"{loc('ngl')}*{d('ngl_price')}",
                'revenue':f"SUM({loc('oil_revenue')}:{loc('ngl_revenue')})",'lifting':f"{loc('total')}*{d('lifting')}",
                **{k:f"{loc('revenue')}*{d(driver)}" for k,driver in [('royalties','royalty_rate'),('selling','selling_rate'),('gna','gna_rate')]},
                'ebitda':f"{loc('revenue')}-SUM({loc('lifting')}:{loc('gna')})",'dda':f"{loc('total')}*{d('dda')}",
                'ebit':f"{loc('ebitda')}-{loc('dda')}",'tax':f"MAX(0,{loc('ebit')})*{d('tax_rate')}",
                'new_capacity':f"MAX(0,{loc('total')}-({prevtotal})*(1-{d('decline')}))",
                'required_capex':f"{loc('new_capacity')}*{d('replacement_cost')}+{loc('total')}*{d('infrastructure')}",
                'capex':f"IF({d('capex_override')}=\"\",{loc('required_capex')}+{d('extra_capex')},{d('capex_override')})",
                'nwc':f"{loc('revenue')}*{d('nwc_rate')}",'delta_nwc':f"{loc('nwc')}-({pnwc})",
                'fcf':f"{loc('ebit')}-{loc('tax')}+{loc('dda')}-{loc('capex')}-{loc('delta_nwc')}",
                'discount_factor':f"1/(1+{polw(case)})^{i+1}",'pv':f"{loc('fcf')}*{loc('discount_factor')}",
                'capex_gap':f"{loc('capex')}-{loc('required_capex')}",
                'investment_warning':f"IF(OR({loc('capex_gap')}<-0.00000001,AND({loc('new_capacity')}>0,{d('replacement_cost')}=0)),1,0)",
            }
            for key,formula in formulas.items(): put(fs,loc(key),'='+formula,'link' if "!" in formula else "body",NUMBER)
    fs["renders"]=["C2:J37","C41:J71","C75:J105","C109:J139"]

    def terminal_capex(case):
        return f"({fr(case,'total','N')}*({dr(case,'decline','N')}*{dr(case,'replacement_cost','N')}+{dr(case,'infrastructure','N')})+{dr(case,'extra_capex','N')})"
    def terminal_fcf(case,g):
        return f"({fr(case,'ebit','N')}-{fr(case,'tax','N')}+{fr(case,'dda','N')}-{terminal_capex(case)})*(1+{g})-{fr(case,'nwc','N')}*{g}"
    bridge_delta=f"({h('cash')}+{br('cash_adjustment')}+{br('nonoperating_assets')}-{h('debt_nc')}-{h('debt_current')}-{h('lease_nc')}-{h('lease_current')}-{br('debt_adjustment')}-{br('other_claims')}-{br('minority_interest')})"
    denom=f"({br('shares')}+{br('dilution')})"

    sc=sheet("Scenarios","Vista Scenarios and Sensitivities","USD per ADR. Oil sensitivity shifts every selected-case oil price, with taxes and NWC recalculated.",122,"N")
    header(sc,6,["Scenario comparison","Units","Base","Bear","Bull"])
    for row,text in [(8,"Value per ADR"),(9,"Enterprise value"),(10,"Equity value"),(11,"WACC"),(12,"Terminal value / enterprise value")]: label(sc,row,text)
    for col,case in zip('EFG',CASES):
        g=pr(case,'terminal_growth'); w=polw(case)
        pvf=f"SUM({fr(case,'pv','E')}:{fr(case,'pv','N')})"
        pvt=f"(({terminal_fcf(case,g)})/({w}-{g}))/(1+{w})^10"
        case_check=f"'Sources'!{'HIJ'[CASES.index(case)]}6"
        valid=f"AND({w}>{g},{w}>-1,{g}>-1,{case_check}=0,COUNT('Assumptions'!E176:E184)=9,{br('shares')}>0,{br('dilution')}>=0,{br('adr_ratio')}>0)"
        put(sc,f"{col}9",f'=IF({valid},{pvf}+{pvt},"Review inputs")',"link",NUMBER)
        put(sc,f"{col}10",f'=IF(ISNUMBER({col}9),{col}9+{bridge_delta},"Review inputs")',"link",NUMBER)
        put(sc,f"{col}8",f'=IF(ISNUMBER({col}10),{col}10/{denom}*{br("adr_ratio")},"Review inputs")',"link",PRICE)
        put(sc,f"{col}11",f"={w}","link",PERCENT)
        put(sc,f"{col}12",f'=IF(AND(ISNUMBER({col}9),{col}9<>0),({pvt})/{col}9,"")',"link",PERCENT)

    v=by_name['Valuation']
    label(v,6,"Selected case"); put(v,"E6","='Assumptions'!E6","link")
    vlabs={8:"WACC",9:"Terminal nominal USD growth",10:"PV of forecast cash flows",11:"Terminal-year free cash flow",12:"Terminal value",13:"PV of terminal value",14:"Enterprise value",15:"Cash and short-term investments",16:"Other nonoperating assets",17:"Borrowings and lease debt",18:"Other debt-like claims",19:"Minority interest",20:"Equity value",21:"Diluted ordinary shares (millions)",22:"Ordinary shares per ADR",23:"Value per ADR (USD)",24:"Terminal value / enterprise value"}
    for row,text in vlabs.items(): label(v,row,text)
    for row,formula in {
        8:"'Assumptions'!E42",9:"'Assumptions'!E40",10:f"SUM({fr('Selected','pv','E')}:{fr('Selected','pv','N')})",
        11:"E41",12:'IF(E28="",E11/(E8-E9),"Review inputs")',13:'IF(E28="",E12/(1+E8)^10,"Review inputs")',
        14:'IF(E28="",SUM(E10,E13),"Review inputs")',15:f"{h('cash')}+{br('cash_adjustment')}",16:br('nonoperating_assets'),
        17:f"{h('debt_nc')}+{h('debt_current')}+{h('lease_nc')}+{h('lease_current')}+{br('debt_adjustment')}",
        18:br('other_claims'),19:br('minority_interest'),20:'IF(E28="",E14+SUM(E15:E16)-SUM(E17:E19),"Review inputs")',
        21:denom,22:br('adr_ratio'),23:'IF(E28="",E20/E21*E22,"Review inputs")',24:'IF(AND(E28="",E14<>0),E13/E14,"")',
    }.items(): put(v,f"E{row}",'='+formula,'link' if '!' in formula else 'body',PERCENT if row in (8,9,24) else PRICE if row==23 else NUMBER)
    label(v,28,"Input issue")
    put(v,"E28",'=IF(OR(COUNT(E8:E9)<2,E8<=E9,E8<=-1,E9<=-1),"WACC must exceed terminal growth",IF(OR(COUNT(\'Assumptions\'!E176:E184)<9,\'Assumptions\'!E176<=0,\'Assumptions\'!E177<0,E22<=0),"Review equity bridge inputs",IF(\'Sources\'!E6>0,"Missing or invalid annual inputs","")))')
    label(v,29,"Production investment")
    put(v,"E29",f'=IF(SUM({fr("Selected","investment_warning","E")}:{fr("Selected","investment_warning","N")})>0,"Capex below production requirement","")')
    label(v,30,"Terminal operations")
    put(v,"E30",'=IF(E41<=0,"Terminal cash flow is nonpositive","")')
    label(v,31,"Residual equity")
    put(v,"E31",'=IF(AND(ISNUMBER(E20),E20<0),"Negative equity: no residual for shareholders","")')
    header(v,33,["Terminal year calculation"])
    for row,text in [(35,"Revenue"),(36,"EBIT"),(37,"Depreciation and depletion"),(38,"Cash taxes"),(39,"Sustaining capital investment"),(40,"Change in working capital"),(41,"Terminal-year free cash flow")]: label(v,row,text,"USD m")
    for row,key in [(35,'revenue'),(36,'ebit'),(37,'dda')]: put(v,f"E{row}",f"={fr('Selected',key,'N')}*(1+E9)","link",NUMBER)
    put(v,"E38",'=MAX(0,E36)*\'Assumptions\'!N21',"link",NUMBER)
    put(v,"E39",f"=({fr('Selected','total','N')}*('Assumptions'!N23*'Assumptions'!N24+'Assumptions'!N25)+'Assumptions'!N26)*(1+E9)","link",NUMBER)
    put(v,"E40",f"={fr('Selected','nwc','N')}*E9","link",NUMBER)
    put(v,"E41","=E36-E38+E37-E39-E40",fmt=NUMBER)
    put(v,"C44","Terminal reinvestment replaces natural decline. Final-year capex overrides do not carry into perpetuity. Tax losses are not carried forward; taxes apply only to positive EBIT.","long_note")
    put(v,"C47","Equity bridge uses book-value proxies for other claims and associates. Review the explicit amounts in Assumptions and their source notes before relying on the result.","long_note")

    header(sc,15,["WACC versus terminal growth"])
    put(sc,"D17","WACC / g","header")
    for j,delta in enumerate((-.01,-.005,0,.005,.01),5): put(sc,f"{column(j)}17",f"='Valuation'!$E$9+({delta})","header",PERCENT)
    for row,delta in enumerate((-.02,-.01,0,.01,.02),18):
        put(sc,f"D{row}",f"='Valuation'!$E$8+({delta})","header",PERCENT)
        for j in range(5,10):
            col=column(j); rate=f"$D{row}"; g=f"{col}$17"
            pv='+'.join(f"{fr('Selected','fcf',column(k))}/(1+{rate})^{k-4}" for k in range(5,15))
            # Terminal no-growth operating FCF is recovered from selected helper cells.
            tf=f"('Valuation'!$E$36-'Valuation'!$E$38+'Valuation'!$E$37-'Valuation'!$E$39)/(1+'Valuation'!$E$9)*(1+{g})-{fr('Selected','nwc','N')}*{g}"
            put(sc,f"{col}{row}",f'=IF(AND({rate}>{g},{rate}>-1,{g}>-1,\'Valuation\'!$E$28=""),({pv}+({tf})/({rate}-{g})/(1+{rate})^10+{bridge_delta})/{denom}*{br("adr_ratio")},"Invalid")',"link",PRICE)
    header(sc,27,["Oil-price sensitivity"])
    for j,shock in enumerate((-.2,-.1,0,.1,.2),5):
        col=column(j); put(sc,f"{col}30",shock,"header",PERCENT)
        start=40+(j-5)*16
        header(sc,start-2,[f"Oil-price change {shock:+.0%}"])
        for k in range(5,15):
            ccol=column(k); prev=column(k-1)
            revenue=f"{fr('Selected','revenue',ccol)}+{fr('Selected','oil_revenue',ccol)}*${col}$30"
            ebit=f"{ccol}{start}*(1-'Assumptions'!{ccol}17-'Assumptions'!{ccol}18-'Assumptions'!{ccol}19)-{fr('Selected','lifting',ccol)}-{fr('Selected','dda',ccol)}"
            pnwc=f"'Adjustments'!G67+{br('opening_nwc_adjustment')}" if k==5 else f"{prev}{start+3}"
            formulas=[revenue,ebit,f"MAX(0,{ccol}{start+1})*'Assumptions'!{ccol}21",
                      f"{ccol}{start}*'Assumptions'!{ccol}22",f"{ccol}{start+3}-({pnwc})",
                      f"{ccol}{start+1}-{ccol}{start+2}+{fr('Selected','dda',ccol)}-{fr('Selected','capex',ccol)}-{ccol}{start+4}",
                      f"{ccol}{start+5}/(1+'Valuation'!$E$8)^{k-4}"]
            for offset,formula in enumerate(formulas): put(sc,f"{ccol}{start+offset}",'='+formula,'link',NUMBER)
        for off,text in enumerate(["Revenue","EBIT","Cash taxes","Working capital","Change in working capital","Free cash flow","PV of free cash flow"]): label(sc,start+off,text,"USD m")
        tf=f"(N{start+1}-N{start+2})*(1+'Valuation'!$E$9)+'Valuation'!$E$37-'Valuation'!$E$39-N{start+3}*'Valuation'!$E$9"
        formula=f'(SUM(E{start+6}:N{start+6})+({tf})/(\'Valuation\'!$E$8-\'Valuation\'!$E$9)/(1+\'Valuation\'!$E$8)^10+{bridge_delta})/{denom}*{br("adr_ratio")}'
        put(sc,f"{col}31",f'=IF(\'Valuation\'!$E$28="",{formula},"Invalid")',"link",PRICE)
    label(sc,30,"Change in all oil-price assumptions")
    label(sc,31,"Value per ADR (USD)")
    put(sc,"C34","Negative residual values indicate that modeled enterprise value does not cover claims. They are not negative share-price predictions.","long_note")
    sc["renders"]=["C2:J36","C38:J50"]

    src=sheet("Sources","Vista Sources and Checks","All source values retain their period, units, location and import conversion.",len(data['facts'])+53,"L")
    src["widths"]={"C":38,"D":16,"E":20,"F":18,"G":26,"H":38,"I":65,"J":20,"K":18,"L":18}
    label(src,6,"Invalid selected annual inputs")
    # Explicit blanks/invalids; capex and WACC overrides are the only optional numeric inputs.
    terms=["IF(COUNT('Assumptions'!E10:N26)<170,1,0)",
           "IF(MIN('Assumptions'!E10:N12)<=-1,1,0)",
           "IF(MIN('Assumptions'!E13:N21)<0,1,0)",
           "IF(MAX('Assumptions'!E17:N19)>1,1,0)",
           "IF(MAX('Assumptions'!E21:N21)>1,1,0)",
           "IF(OR(MIN('Assumptions'!E23:N26)<0,MAX('Assumptions'!E23:N23)>1),1,0)",
           "IF(COUNT('Assumptions'!E32:E38)<7,1,0)",
           "IF(COUNT('Assumptions'!E40)<1,1,0)",
           "IF('Assumptions'!E39=\"\",0,IF(ISNUMBER('Assumptions'!E39),0,1))",
           "IF(OR('Assumptions'!E37<0,'Assumptions'!E37>1,'Assumptions'!E38<0,'Assumptions'!E38>1),1,0)",
           "IF(AND('Assumptions'!E6<>\"Base\",'Assumptions'!E6<>\"Bear\",'Assumptions'!E6<>\"Bull\"),1,0)"]
    for col in map(column,range(5,15)):
        terms.append(f'IF({chr(39)}Assumptions{chr(39)}!{col}27="",0,IF(ISNUMBER(\'Assumptions\'!{col}27),IF(\'Assumptions\'!{col}27<0,1,0),1))')
    # Apply the same check directly to each case's editable cells, so blanks do
    # not turn into valid zeroes through the selected-case link formulas.
    import re
    for case,check_col in zip(CASES,'HIJ'):
        offset=CASE_ROWS[case]-10
        def relocate(match):
            col,row=match.group(1),int(match.group(2))
            if 10<=row<=27: row+=offset
            elif 32<=row<=40: row+=offset-1
            return f"'Assumptions'!{col}{row}"
        case_terms=[re.sub(r"'Assumptions'!([A-Z]+)(\d+)",relocate,t) for t in terms]
        # Range end references also need their row adjusted.
        case_terms=[re.sub(r":([A-Z]+)(\d+)",lambda m:f":{m.group(1)}{int(m.group(2))+offset-(1 if 32<=int(m.group(2))<=40 else 0)}",t) for t in case_terms]
        put(src,f"{check_col}5",case,"note")
        put(src,f"{check_col}6",'='+ '+'.join(case_terms),"link",'0')
    put(src,"E6",'=IF(\'Assumptions\'!E6="Base",H6,IF(\'Assumptions\'!E6="Bear",I6,J6))',"link",'0')
    header(src,9,["Source","Document","","Period","URL"])
    for row,code,title,period,url in [(10,"S1","Official historical workbook","2023–2025",data['source_url']),
        (12,"S2","Audited consolidated financial statements","2025 / 2024",AUDIT_URLS[2025]),
        (14,"S3","Audited consolidated financial statements","2024 / 2023",AUDIT_URLS[2024]),
        (16,"S4","Form 20-F, Item 12: ADS ratio","2024",ADR_URL)]:
        put(src,f"C{row}",code); put(src,f"D{row}",title,"side_note");put(src,f"F{row}",period);put(src,f"G{row}",url,"url")
    notes=[
        "Shares: 104,299,705 outstanding; allowance 1,778,830 to the 106,078,535 cap in audited 2025 Note 11.2. Dilution is an editable modeling allowance, not an EPS weighted average.",
        "Other claims: 292.236 noncurrent payables + 19.236 current farmout payable (Note 26), 62.313 provisions and 134.874 income-tax payable (balance sheet). Associate asset proxy 54.542. Amounts in USD m.",
        "Revenue/royalty normalization uses net-of-duty oil prices to avoid double counting export duties. Freight reimbursement 29.8 in 2025 is removed from revenue and selling expense (company workbook notes).",
        "WACC risk-free, beta, risk premiums, debt cost/weight, 35% cash-tax rate, 25% decline and USD60 per new annual boe capacity are illustrative user-editable assumptions, not estimates sourced from management.",
        "FY2025 production includes only part-year acquired assets. Defaults start from annual average production, not Q4 exit production. The 2026 transaction perimeter is excluded. Terminal value assumes continued reserve replacement.",
        "Leases are financing: lease liabilities are included in debt. Forecast replacement/infrastructure investment covers owned and leased capacity; do not subtract lease principal again. DDA is a per-boe tax proxy; no tax-loss or deferred-tax schedule.",
    ]
    for i,note in enumerate(notes): put(src,f"C{19+i*3}",note,"long_note")
    header(src,38,["Audited comparison","Year","Imported USD m","Audited USD m","Difference USD m"])
    for i,check in enumerate(data['audit_checks']):
        row=39+i
        put(src,f"C{row}",check['key']);put(src,f"D{row}",check['year'],fmt='0')
        ycol=column(5+check['year']-2023)
        put(src,f"E{row}",f"='Historical'!{ycol}{HROWS[check['key']]}","link",NUMBER)
        put(src,f"F{row}",check['audited'],fmt=NUMBER)
        put(src,f"G{row}",f"=E{row}-F{row}",fmt=NUMBER)
    fact_start=57
    src['last_row']=fact_start+len(data['facts'])+4
    header(src,56,["Metric","Year","Value","Units","Source location","Conversion multiplier","URL"])
    for i,fact in enumerate(data['facts']):
        row=fact_start+i
        for col,key in zip('CDEFGHI',['label','year','value','unit','location','scale','url']):
            put(src,f"{col}{row}",fact[key],"url" if col=='I' else 'body',NUMBER if col=='E' else '0.000' if col=='H' else '0' if col=='D' else None)
    src['renders']=["C2:J17","C19:J36","C38:H54","C56:I64"]

    guide=sheet('Assumption Guide','Why These Assumptions?','Original default rules, not current forecasts. Historical anchors use the Sources sheet; other choices are illustrative.',49,'F')
    guide['widths']={'C':43,'D':66,'E':83,'F':30}
    guide['freeze_rows']=7
    header(guide,7,['Input','Original default construction','Why this starting point / source','Edit values at'])
    guide_row=8
    for items, explanations, locations in [
        (DRIVERS,DRIVER_GUIDE,lambda key: 'Assumptions: rows '+', '.join(str(CASE_ROWS[c]+DROWS[key]) for c in CASES)+' (E:N)'),
        (POLICIES,POLICY_GUIDE,lambda key: 'Assumptions: '+', '.join('E'+str(CASE_ROWS[c]+21+PROWS[key]) for c in CASES)),
        (BRIDGE,BRIDGE_GUIDE,lambda key: 'Assumptions!E'+str(BROWS[key])),
    ]:
        for key,text,unit in items:
            rule,reason=explanations[key]
            for col,value in zip('CDEF',[text+' ('+unit+')',rule,reason,locations(key)]):
                put(guide,f'{col}{guide_row}',value,'guide_text')
            guide_row+=1
    for text,rule,reason,location in [
        ('Freight normalization (USD m)',*FREIGHT_GUIDE,'Adjustments!E58:G58'),
        ('Historical financial and operating inputs','Reported values remain in Historical. Blank overrides use reported data; zero replaces a value with zero.','Every imported numeric metric can be overridden by year in Adjustments. Source locations and units are listed in Sources. Input Notes provides a separate reason for each year; existing row notes remain in Adjustments N.','Adjustments!E:G'),
        ('Changing assumptions after a build','Annual inputs are stored independently. Editing historical values changes linked calculations, but does not regenerate the originally seeded annual assumptions.','Review the forecast assumptions explicitly after changing history. Use Input Notes to record the input, rationale, source and date. Calculated totals and valuation outputs update from the editable inputs.','Assumptions and Input Notes'),
    ]:
        for col,value in zip('CDEF',[text,rule,reason,location]): put(guide,f'{col}{guide_row}',value,'guide_text')
        guide_row+=1
    guide['last_row']=guide_row+1
    guide['renders']=['C2:F17','C18:F25','C26:F34',f'C35:F{guide_row-1}']

    # One persistent reason per numeric input, including each scenario/year and historical override.
    numeric_inputs=[entry for entry in inputs if entry['path'][0] not in ('selected_case','override_notes')]
    notes=sheet('Input Notes','Reasons for Input Changes','Filter by scenario, year or input. Change values at the listed cell; write your reason, source and date in column I.',len(numeric_inputs)+9,'I')
    notes['widths']={'C':18,'D':12,'E':43,'F':19,'G':17,'H':26,'I':66}
    notes['freeze_rows']=7
    header(notes,7,['Scenario / group','Year','Input','Units','Current input','Edit number at','Reason / source / date'])
    labels={key:(text,unit) for key,text,unit in DRIVERS+POLICIES+BRIDGE}
    labels.update({m[0]:(m[1],m[2]) for m in METRICS})
    for row,entry in enumerate(numeric_inputs,8):
        p=entry['path']; key=p[-1]
        if p[0]=='scenarios':
            group=p[1]; year=int(p[3]) if p[2]=='drivers' else 'All years'
        elif p[0]=='overrides': group='Historical'; year=int(p[1])
        elif p[0]=='freight': group='Freight'; year=int(p[1])
        else: group='Equity bridge'; year='Opening'
        text,unit=('Freight normalization','USD m') if p[0]=='freight' else labels[key]
        for col,value in zip('CDEF',[group,year,text,unit]): put(notes,f'{col}{row}',value,'notes_text','0' if col=='D' and isinstance(year,int) else None)
        ref=f"'{entry['sheet']}'!{entry['cell']}"
        fmt=PERCENT if unit.startswith('%') else NUMBER
        put(notes,f'G{row}',f'=IF(ISBLANK({ref}),"",{ref})','link',fmt)
        put(notes,f'H{row}',f"{entry['sheet']}!{entry['cell']}",'notes_text')
        note_key='/'.join(p)
        inp(notes,f'I{row}',config.get('input_notes',{}).get(note_key,''),['input_notes',note_key],None)
        notes['cells'][-1]['role']='input_note'
    notes['tables']=[dict(name='InputChangeReasons',range=f'C7:I{len(numeric_inputs)+7}')]
    notes['renders']=['C2:I13','C185:I191',f'C{len(numeric_inputs)+2}:I{len(numeric_inputs)+7}']

    # Independent Python expected values for every core case/year cell and sensitivities.
    for case in CASES:
        result=forecast(data,config,case)
        for j,row in enumerate(result['rows'],5):
            for key,_,_ in FIELDS:
                if key in row:
                    comparisons.append(dict(sheet='Forecast',cell=f'{column(j)}{FORECAST_ROWS[case]+FROWS[key]}',expected=row[key]))
        comparisons.append(dict(sheet='Scenarios',cell=f'{"EFG"[CASES.index(case)]}8',expected=result['value_per_adr']))
    selected=forecast(data,config,config['selected_case'])
    for cell,key in [('E14','enterprise_value'),('E20','equity_value'),('E23','value_per_adr'),('E24','terminal_share')]:
        comparisons.append(dict(sheet='Valuation',cell=cell,expected=selected[key]))
    sens=sensitivities(data,config)
    for i,row in enumerate(sens['values'],18):
        for j,value in enumerate(row,5):
            if value is not None: comparisons.append(dict(sheet='Scenarios',cell=f'{column(j)}{i}',expected=value))
    import copy
    for j,shock in enumerate((-.2,-.1,0,.1,.2),5):
        changed=copy.deepcopy(config)
        for d in changed['scenarios'][config['selected_case']]['drivers'].values(): d['oil_price']*=1+shock
        comparisons.append(dict(sheet='Scenarios',cell=f'{column(j)}31',expected=forecast(data,changed,config['selected_case'])['value_per_adr']))
    # Reader order: outputs and inputs first, followed by supporting schedules.
    order=['Instructions','Valuation','Assumptions','Assumption Guide','Input Notes','Forecast','Scenarios','Historical','Adjustments','Sources']
    return dict(schema_version=1,sheets=[by_name[n] for n in order],input_map=inputs,
                comparisons=comparisons,config=config,data=data,summary=selected)
