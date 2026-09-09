from __future__ import annotations

import copy
import io
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from dcf_model.data_sources import load_vista, audit_history, annual_column, read_number, discover_history_url, download
from dcf_model.model import default_assumptions, adjusted_history, normalized, forecast, sensitivities, validate
from dcf_model.cli import saved_assumptions
from dcf_model.workbook_spec import build_spec

RAW=Path(__file__).resolve().parents[1]/'data/raw'


class DCFTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data=load_vista(RAW/'vista_historical.xlsx')
        cls.data['audit_checks']=audit_history(cls.data,RAW)

    def setUp(self):
        self.config=default_assumptions(self.data)

    def result(self,case='Base'):
        return forecast(self.data,self.config,case)

    def test_audited_reconciliations(self):
        self.assertEqual(len(self.data['audit_checks']),15)
        self.assertTrue(all(abs(c['difference'])<.002 for c in self.data['audit_checks']))
        for h in self.data['history'].values():
            self.assertAlmostEqual(h['assets'],h['liabilities']+h['equity'],places=3)
            self.assertAlmostEqual(h['revenue'],h['oil_revenue']+h['gas_revenue']+h['ngl_revenue'],places=3)
            self.assertAlmostEqual(h['closing_cash'],h['opening_cash']+h['cfo']+h['cfi']+h['cff']+h['fx_cash'],places=3)

    def test_normalization_does_not_double_count_export_duties_or_freight(self):
        h=self.data['history']['2025']
        n=normalized(h,29.8)
        # Equal expense/revenue removals do not manufacture operating profit.
        expected=h['revenue']-h['opex']-h['royalties']-h['selling']-h['gna']-h['dda']
        self.assertAlmostEqual(n['ebit'],expected)
        self.assertAlmostEqual(n['revenue'],2370.697,places=3)
        self.assertGreater(h['ebit']-n['ebit'],400)

    def test_higher_wacc_reduces_positive_cash_flow_value(self):
        base=self.result()
        self.assertTrue(all(r['fcf']>0 for r in base['rows']))
        self.config['scenarios']['Base']['policies']['wacc_override']=base['wacc']+.02
        self.assertLess(self.result()['enterprise_value'],base['enterprise_value'])

    def test_more_debt_reduces_equity_dollar_for_dollar(self):
        base=self.result()
        self.config['bridge']['debt_adjustment']=100
        changed=self.result()
        self.assertAlmostEqual(changed['enterprise_value'],base['enterprise_value'])
        self.assertAlmostEqual(changed['equity_value'],base['equity_value']-100)

    def test_adr_ratio_direction(self):
        base=self.result()['value_per_adr']
        self.config['bridge']['adr_ratio']=10
        self.assertAlmostEqual(self.result()['value_per_adr'],base*10)

    def test_later_year_input_leaves_prior_years_unchanged(self):
        base=self.result()
        self.config['scenarios']['Base']['drivers']['2030']['oil_price']+=10
        changed=self.result()
        self.assertEqual(changed['rows'][:4],base['rows'][:4])
        self.assertGreater(changed['rows'][4]['fcf'],base['rows'][4]['fcf'])

    def test_growth_requires_more_investment(self):
        base=self.result()
        self.config['scenarios']['Base']['drivers']['2026']['oil_growth']+=.10
        self.assertGreater(self.result()['rows'][0]['required_capex'],base['rows'][0]['required_capex'])

    def test_override_zero_is_not_missing_and_shortfall_is_visible(self):
        self.config['scenarios']['Base']['drivers']['2026']['capex_override']=0
        result=self.result()
        self.assertEqual(result['rows'][0]['capex'],0)
        self.assertLess(result['rows'][0]['capex_gap'],0)
        self.assertTrue(result['warnings'])

    def test_zero_capital_productivity_is_flagged(self):
        self.config['scenarios']['Base']['drivers']['2026']['replacement_cost']=0
        self.assertTrue(self.result()['warnings'])

    def test_no_tax_credit_on_operating_losses(self):
        for d in self.config['scenarios']['Bear']['drivers'].values():
            d['oil_price']=0
        result=self.result('Bear')
        self.assertTrue(all(r['tax']==0 for r in result['rows']))
        self.assertEqual(result['terminal']['tax'],0)

    def test_leap_year_volume_follows_daily_production(self):
        drivers=self.config['scenarios']['Base']['drivers']
        for d in drivers.values():
            d['oil_growth']=0
        rows=self.result()['rows']
        self.assertAlmostEqual(rows[2]['oil']/rows[1]['oil'],366/365)
        self.assertAlmostEqual(rows[3]['oil']/rows[2]['oil'],365/366)

    def test_terminal_reinvestment_ignores_year_ten_override(self):
        base=self.result()
        self.config['scenarios']['Base']['drivers']['2035']['capex_override']=0
        changed=self.result()
        self.assertEqual(changed['terminal']['capex'],base['terminal']['capex'])
        self.assertGreater(changed['terminal']['capex'],0)

    def test_terminal_growth_boundary_blocked(self):
        p=self.config['scenarios']['Base']['policies']
        p['wacc_override']=p['terminal_growth']
        with self.assertRaisesRegex(ValueError,'WACC must exceed'):
            self.result()

    def test_missing_and_invalid_inputs_fail_clearly(self):
        for bad in (None,'',float('nan'),float('inf'),True):
            with self.subTest(bad=bad),self.assertRaises(ValueError):
                read_number(bad,'test cell')
        self.config['scenarios']['Base']['drivers']['2026']['oil_price']=None
        with self.assertRaisesRegex(ValueError,'nonnumeric'):
            self.result()

    def test_bad_share_count_and_tax_rate_rejected(self):
        self.config['bridge']['shares']=0
        with self.assertRaises(ValueError): validate(self.config)
        self.config=default_assumptions(self.data)
        self.config['scenarios']['Base']['drivers']['2026']['tax_rate']=1.1
        with self.assertRaises(ValueError): validate(self.config)

    def test_historical_override_zero_retained(self):
        self.config['overrides']={'2025':{'cash':0}}
        changed=adjusted_history(self.data,self.config)
        self.assertEqual(changed['2025']['cash'],0)
        self.assertNotEqual(self.data['history']['2025']['cash'],0)

    def test_sensitivity_center_matches_selected_case(self):
        self.assertAlmostEqual(sensitivities(self.data,self.config)['values'][2][2],self.result()['value_per_adr'])

    def test_new_fiscal_year_fails_before_stale_metadata_is_used(self):
        with self.assertRaisesRegex(ValueError,'reviewed audit'):
            load_vista(RAW/'vista_historical.xlsx',2026)

    def test_annual_column_rejects_quarters_and_duplicate_years(self):
        fake=SimpleNamespace(title='Income Statement')
        class Sheet:
            title='Income Statement'
            def __getitem__(self,key):
                return [SimpleNamespace(column=1,value='Q4 2025'),SimpleNamespace(column=2,value=2025),SimpleNamespace(column=3,value=2024)]
        self.assertEqual(annual_column(Sheet(),2025),2)
        with self.assertRaises(ValueError): annual_column(Sheet(),2023)

    def test_discovery_and_failed_download_preserve_cache(self):
        content=b'<a href="https://vista.example/HistoricalFinancialData.xlsx">Historical data</a>'
        with patch('dcf_model.data_sources.urlopen',return_value=io.BytesIO(content)):
            self.assertEqual(discover_history_url(),'https://vista.example/HistoricalFinancialData.xlsx')
        with tempfile.TemporaryDirectory() as temp:
            target=Path(temp)/'cached.xlsx';target.write_bytes(b'good cache')
            with patch('dcf_model.data_sources.urlopen',side_effect=OSError('network unavailable')):
                with self.assertRaises(OSError): download('https://vista.example/test',target)
            self.assertEqual(target.read_bytes(),b'good cache')

    def test_refresh_preserves_saved_numeric_inputs_zero_overrides_and_notes(self):
        entries=[dict(sheet='Assumptions',cell='E6',path=['selected_case']),
                 dict(sheet='Assumptions',cell='I55',path=['scenarios','Base','drivers','2030','oil_price']),
                 dict(sheet='Adjustments',cell='G32',path=['overrides','2025','cash']),
                 dict(sheet='Adjustments',cell='N32',path=['override_notes','cash'])]
        values={'E6':'Bull','I55':81,'G32':0,'N32':'Restricted cash excluded'}
        class FakeSheet:
            def __getitem__(self,key): return SimpleNamespace(value=values[key])
        class FakeWorkbook:
            def __getitem__(self,key): return FakeSheet()
            def close(self): pass
        with tempfile.TemporaryDirectory() as temp:
            book=Path(temp)/'saved.xlsx'
            book.with_suffix('.inputs.json').write_text(json.dumps(dict(input_map=entries,config=self.config)))
            with patch('dcf_model.cli.openpyxl.load_workbook',return_value=FakeWorkbook()):
                restored=saved_assumptions(book)
            self.assertEqual(restored['selected_case'],'Bull')
            self.assertEqual(restored['scenarios']['Base']['drivers']['2030']['oil_price'],81)
            self.assertEqual(restored['overrides']['2025']['cash'],0)
            self.assertEqual(restored['override_notes']['cash'],'Restricted cash excluded')
            self.assertEqual(restored['scenarios']['Bear'],self.config['scenarios']['Bear'])

    def test_every_editable_cell_has_unique_refresh_mapping(self):
        spec=build_spec(self.data,self.config)
        keys=[(entry['sheet'],entry['cell']) for entry in spec['input_map']]
        self.assertEqual(len(keys),len(set(keys)))
        paths=[tuple(entry['path']) for entry in spec['input_map']]
        self.assertEqual(len(paths),len(set(paths)))

    def test_every_numeric_input_has_a_persistent_reason_and_live_value(self):
        spec=build_spec(self.data,self.config)
        numeric=[e for e in spec['input_map'] if e['path'][0] not in ('selected_case','override_notes','input_notes')]
        reasons={e['path'][1]:e for e in spec['input_map'] if e['path'][0]=='input_notes'}
        self.assertEqual(set(reasons),{'/'.join(e['path']) for e in numeric})
        notes=next(s for s in spec['sheets'] if s['name']=='Input Notes')
        cells={c['address']:c for c in notes['cells']}
        for entry in numeric:
            reason=reasons['/'.join(entry['path'])]
            value=cells['G'+reason['cell'][1:]]
            self.assertIn(f"'{entry['sheet']}'!{entry['cell']}",value['formula'])

    def test_saved_annual_policy_bridge_and_freight_notes_survive_refresh(self):
        spec=build_spec(self.data,self.config)
        values={(s['name'],c['address']):c.get('value') for s in spec['sheets'] for c in s['cells']}
        reasons={'scenarios/Base/drivers/2030/oil_price':'Revised 2030 price outlook',
                 'scenarios/Bear/policies/risk_free':'Updated USD rate reference',
                 'bridge/shares':'Updated period-end share count',
                 'freight/2025':'Reviewed reimbursement amount',
                 'overrides/2025/cash':'Exclude restricted cash in 2025'}
        for entry in spec['input_map']:
            if entry['path'][0]=='input_notes' and entry['path'][1] in reasons:
                values[(entry['sheet'],entry['cell'])]=reasons[entry['path'][1]]
        values[('Assumptions','I55')]=75
        class FakeWorkbook:
            def __getitem__(self,sheet):
                class FakeSheet:
                    def __getitem__(self,cell): return SimpleNamespace(value=values[(sheet,cell)])
                return FakeSheet()
            def close(self): pass
        with tempfile.TemporaryDirectory() as temp:
            book=Path(temp)/'saved.xlsx'
            book.with_suffix('.inputs.json').write_text(json.dumps(dict(input_map=spec['input_map'],config=self.config)))
            with patch('dcf_model.cli.openpyxl.load_workbook',return_value=FakeWorkbook()):
                restored=saved_assumptions(book)
        for key,note in reasons.items(): self.assertEqual(restored['input_notes'][key],note)
        self.assertEqual(restored['scenarios']['Base']['drivers']['2030']['oil_price'],75)
        rebuilt=build_spec(self.data,restored)
        note_values={c.get('value') for s in rebuilt['sheets'] if s['name']=='Input Notes' for c in s['cells'] if c['role']=='input_note'}
        self.assertTrue(set(reasons.values()).issubset(note_values))


if __name__=='__main__': unittest.main()
