"""Refresh inputs and produce a versioned, editable Excel DCF."""
from __future__ import annotations

import argparse
import copy
import json
import os
from pathlib import Path
import shutil
import subprocess
from datetime import datetime, timezone

import openpyxl

from .data_sources import load_vista, audit_history, discover_history_url, download, AUDIT_URLS, HISTORY_URL
from .model import default_assumptions, validate, forecast
from .workbook_spec import build_spec

PACKAGE = Path(__file__).resolve().parent
DEFAULT_OUTPUT = PACKAGE / 'outputs'


def write_json(path: Path, value):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(value,indent=2,allow_nan=False),encoding='utf-8')


def workbook_version(path: Path):
    """Order generated versions by build time, even after editing older files."""
    if path.stem.startswith('vista_dcf_'):
        stamp=datetime.strptime(path.stem.removeprefix('vista_dcf_'),'%Y%m%d_%H%M%S_%f')
        return stamp, 1
    name=path.stem.removeprefix('Vista_Energy_VIST_ADR_DCF_')
    stamp=datetime.strptime(name[:19],'%Y-%m-%d_%H-%M-%S')
    suffix=name[19:].removeprefix('_UTC')
    return stamp, int(suffix.removeprefix('_')) if suffix else 1


def saved_assumptions(workbook_path: Path) -> dict:
    manifest_path=workbook_path.with_suffix('.inputs.json')
    if not manifest_path.exists():
        raise ValueError(f'Input manifest missing: {manifest_path}. Keep this file beside the workbook, or use --config.')
    manifest=json.loads(manifest_path.read_text(encoding='utf-8'))
    config=copy.deepcopy(manifest['config'])
    workbook=openpyxl.load_workbook(workbook_path,data_only=False)
    try:
        for entry in manifest['input_map']:
            value=workbook[entry['sheet']][entry['cell']].value
            if isinstance(value,str) and value.startswith('='):
                raise ValueError(f"Refresh expects numeric input cells, not formulas: {entry['sheet']}!{entry['cell']}. Original file is unchanged.")
            branch=config
            for key in entry['path'][:-1]: branch=branch.setdefault(key,{})
            branch[entry['path'][-1]]=value
    finally:
        workbook.close()
    validate(config)
    return config


def node_executable() -> str:
    override=os.environ.get('DCF_NODE')
    bundled=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe'
    executable=override or (str(bundled) if bundled.exists() else shutil.which('node'))
    if not executable:
        raise RuntimeError('Node.js is required for workbook export. Set DCF_NODE to the executable.')
    return executable


def main(argv=None):
    parser=argparse.ArgumentParser(description='Vista annual USD DCF. Offline build uses cached official sources.')
    parser.add_argument('command',choices=['build','refresh','value'])
    parser.add_argument('--source-xlsx',type=Path)
    parser.add_argument('--source-url',help='Provenance URL for an explicitly supplied source workbook')
    parser.add_argument('--from-workbook',type=Path,help='Preserve saved numeric assumptions and overrides')
    parser.add_argument('--config',type=Path,help='Explicit assumption JSON instead of previous workbook inputs')
    parser.add_argument('--output-dir',type=Path,default=DEFAULT_OUTPUT)
    parser.add_argument('--no-previews',action='store_true')
    args=parser.parse_args(argv)
    if args.config and args.from_workbook:
        parser.error('Use either --config or --from-workbook')
    raw=PACKAGE/'data/raw'
    source=args.source_xlsx or raw/'vista_historical.xlsx'
    source_url=args.source_url or HISTORY_URL
    metadata=raw/'download_metadata.json'
    if metadata.exists() and not args.source_xlsx:
        source_url=json.loads(metadata.read_text())['url']
    if args.command=='refresh':
        if not args.source_xlsx:
            source_url=discover_history_url()
            download(source_url,source)
            write_json(metadata,dict(url=source_url,retrieved_at=datetime.now(timezone.utc).isoformat()))
        for year,url in AUDIT_URLS.items(): download(url,raw/f'vista_{year}_audited.pdf')
    if not source.exists():
        raise ValueError('Historical source missing. Run refresh with network access, or supply --source-xlsx.')
    data=load_vista(source,source_url=source_url if not args.source_xlsx or args.source_url else f'User-supplied file: {source.name}')
    data['audit_checks']=audit_history(data,raw)
    previous=args.from_workbook
    if previous is None and args.config is None:
        existing=list(args.output_dir.glob('Vista_Energy_VIST_ADR_DCF_*.xlsx'))
        existing.extend(args.output_dir.glob('vista_dcf_*.xlsx'))
        if existing: previous=max(existing,key=workbook_version)
    config=(json.loads(args.config.read_text(encoding='utf-8')) if args.config else saved_assumptions(previous) if previous else default_assumptions(data))
    validate(config)
    result=forecast(data,config,config['selected_case'])
    if args.command=='value':
        print(json.dumps(result,indent=2)); return 0
    stamp=datetime.now(timezone.utc).strftime('%Y-%m-%d_%H-%M-%S')
    output=args.output_dir/f'Vista_Energy_VIST_ADR_DCF_{stamp}_UTC.xlsx'
    version=2
    while output.exists():
        output=args.output_dir/f'Vista_Energy_VIST_ADR_DCF_{stamp}_UTC_{version}.xlsx'
        version+=1
    spec_path=args.output_dir/'excel_workbook_definition.json'
    spec=build_spec(data,config)
    write_json(spec_path,spec)
    write_json(args.output_dir/'historical_financial_data.json',data)
    write_json(args.output_dir/'model_assumptions.json',config)
    write_json(args.output_dir/'dcf_valuation_results.json',result)
    command=[node_executable(),str(PACKAGE/'build_workbook.mjs'),str(spec_path),str(output)]
    if args.no_previews: command.append('--no-previews')
    subprocess.run(command,check=True,cwd=PACKAGE)
    write_json(args.output_dir/'latest_workbook.json',dict(workbook=str(output.resolve()),previous=str(previous.resolve()) if previous else None))
    print(f'Created {output}')
    return 0


if __name__=='__main__':
    raise SystemExit(main())
