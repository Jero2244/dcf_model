import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const [specPath, outputPath, ...flags] = process.argv.slice(2);
if (!specPath || !outputPath) throw new Error('Usage: node build_workbook.mjs spec.json output.xlsx [--no-previews]');
const spec = JSON.parse(await fs.readFile(specPath, 'utf8'));
const wb = Workbook.create();
const outDir = path.dirname(outputPath);
await fs.mkdir(outDir, { recursive: true });
const font = 'Arial';
const sheets = new Map();
for (const def of spec.sheets) sheets.set(def.name, wb.worksheets.add(def.name));

function splitText(text, max = 150) {
  const lines = []; let line = '';
  for (const word of String(text).split(/\s+/)) {
    if (line.length + word.length > max) { lines.push(line); line = word; }
    else line += (line ? ' ' : '') + word;
  }
  if (line) lines.push(line);
  return lines;
}

for (const def of spec.sheets) {
  const sh = sheets.get(def.name);
  sh.showGridLines = false;
  const used = sh.getRange(`A1:${def.last_col}${def.last_row}`);
  used.format.font = { name: font, size: 10, color: '#202A35' };
  used.format.rowHeight = 22;
  used.format.columnWidth = 16;
  used.format.verticalAlignment = 'center';
  sh.getRange('A:B').format.columnWidth = 2;
  sh.getRange('C:C').format.columnWidth = 49;
  sh.getRange('D:D').format.columnWidth = 18;
  for (const [col, width] of Object.entries(def.widths)) sh.getRange(`${col}:${col}`).format.columnWidth = width;
  sh.getRange(`C2:${def.last_col}2`).format.borders = { bottom: { style: 'thin', color: '#A4ACB4' } };
  sh.getRange('C2').values = [[def.title]];
  sh.getRange('C2').format.font = { name:font, size:16, bold:true, color:'#192D45' };
  sh.getRange('C2').format.rowHeight = 29;
  sh.getRange('C3').values = [[def.subtitle]];
  sh.getRange('C3').format.font = { name:font, size:10, italic:true, color:'#576270' };
  if (def.freeze_rows) sh.freezePanes.freezeRows(def.freeze_rows);
  if (def.freeze_columns) sh.freezePanes.freezeColumns(def.freeze_columns);
  if (def.name === 'Assumptions') sh.tabColor = '#E8D6B8';

  const unique = new Map(def.cells.map(c => [c.address, c]));
  for (const cell of unique.values()) {
    const range = sh.getRange(cell.address);
    if ('formula' in cell) range.formulas = [[cell.formula]];
    else if (cell.value !== null && cell.value !== undefined) range.values = [[cell.value]];
    if (cell.number_format) range.setNumberFormat(cell.number_format);
    const row = Number(cell.address.match(/\d+/)[0]);
    if (cell.role === 'input' || cell.role === 'input_note') {
      range.format.fill = '#FFF3CF';
      range.format.font.color = '#0000FF';
    } else if (cell.role === 'link' || (cell.formula || '').includes('!')) {
      range.format.font.color = '#008000';
    }
    if (cell.role === 'header') {
      range.format.fill = '#263E58';
      range.format.font = { name:font, size:10, bold:true, color:'#FFFFFF' };
      range.format.rowHeight = 29;
      range.format.wrapText = true;
      if (cell.address.startsWith('C')) {
        sh.getRange(`C${row}:${def.last_col}${row}`).format.fill = '#263E58';
      }
    }
    if (cell.role === 'note') range.format.font.color = '#657181';
    if (cell.role === 'wide_note') {
      range.format.wrapText = true;
      range.format.rowHeight = 66;
      range.format.verticalAlignment = 'top';
    }
    if (cell.role === 'side_note') {
      const col = cell.address.match(/[A-Z]+/)[0];
      // Short notes in side areas flow over otherwise empty cells on separate lines.
      const max = def.name === 'Assumptions' ? 85 : 38;
      splitText(cell.value,max).forEach((line,i) => {
        const target=sh.getRange(`${col}${row+i}`);
        target.values=[[line]];
        target.format.font.color='#596575';
      });
    }
    if (cell.role === 'long_note') {
      const max = def.name === 'Valuation' ? 100 : 145;
      splitText(cell.value,max).forEach((line,i) => {
        const target=sh.getRange(`C${row+i}`);
        target.values=[[line]];
        target.format.font.color='#596575';
      });
    }
    if (cell.role === 'url') {
      range.format.wrapText = true;
      range.format.font.color = '#52657A';
      range.format.rowHeight = def.name==='Sources' && row>=57 ? 48 : 90;
    }
    if (typeof cell.value==='number' || cell.formula) range.format.horizontalAlignment='right';
    if (cell.role==='header') {
      range.format.horizontalAlignment='center';
      range.format.borders={right:{style:'thin',color:'#FFFFFF'}};
    }
  }
  for (const validation of def.validations || []) {
    sh.getRange(validation.range).dataValidation = { rule: { type:'list', values:validation.values } };
  }
  for (const table of def.tables || []) {
    const added=sh.tables.add(table.range,true,table.name);
    added.style='TableStyleLight1';
    added.showFilterButton=true;
  }
  if (def.name==='Assumption Guide') {
    sh.getRange(`C8:F${def.last_row}`).format.wrapText=true;
    sh.getRange(`C8:F${def.last_row}`).format.verticalAlignment='top';
    sh.getRange(`C8:F${def.last_row}`).format.rowHeight=60;
  }
  if (def.name==='Input Notes') {
    sh.getRange(`C8:I${def.last_row}`).format.wrapText=true;
    sh.getRange(`C8:I${def.last_row}`).format.verticalAlignment='top';
    sh.getRange(`C8:I${def.last_row}`).format.rowHeight=54;
    sh.getRange(`I8:I${def.last_row-2}`).format.fill='#FFF3CF';
    sh.getRange(`I8:I${def.last_row-2}`).format.font.color='#0000FF';
  }
  if (def.name==='Historical' || def.name==='Adjustments' || def.name==='Forecast' || def.name==='Assumptions') {
    sh.getRange(`C8:D${def.last_row}`).format.wrapText = true;
    sh.getRange(`C8:D${def.last_row}`).format.rowHeight = 29;
  }
  // Keep prose flowing across blank cells after applying working-row wrapping.
  for (const cell of unique.values()) {
    if (cell.role==='long_note') {
      const row=Number(cell.address.match(/\d+/)[0]);
      sh.getRange(`C${row}:C${row+2}`).format.wrapText=false;
    }
  }
  if (def.name==='Valuation') {
    sh.getRange('C:C').format.columnWidth=51;
    sh.getRange('D:D').format.columnWidth=13;
    sh.getRange('E:E').format.columnWidth=22;
    for (const row of [14,20,23,41]) {
      sh.getRange(`C${row}:E${row}`).format.borders={top:{style:'thin',color:'#9AA7B6'}};
      sh.getRange(`C${row}:E${row}`).format.font.bold=true;
    }
    sh.getRange('C23:E23').format.fill='#E8EFF6';
    sh.getRange('E23').format.font.size=14;
    sh.getRange('E28:E31').format.horizontalAlignment='left';
  }
  if (def.name==='Scenarios') {
    sh.getRange('E18:I22').conditionalFormats.add('colorScale',{
      colors:['#F1D8C9','#F5F6F8','#C8DDD4'], thresholds:['min',{type:'percentile',value:50},'max']
    });
  }
}

// Compare calculated workbook values to the separate Python engine before export.
const errors=[];
for (const check of spec.comparisons) {
  const actual=sheets.get(check.sheet).getRange(check.cell).values[0][0];
  const tolerance=1e-7*Math.max(1,Math.abs(check.expected));
  if (typeof actual!=='number' || Math.abs(actual-check.expected)>tolerance) errors.push({...check,actual});
}
await fs.writeFile(path.join(outDir,'formula_validation.json'),JSON.stringify({comparisons:spec.comparisons.length,errors},null,2));
if (errors.length) throw new Error(`Python/Excel mismatch (${errors.length}): ${JSON.stringify(errors.slice(0,6))}`);
// Note-sheet mirrors must distinguish an explicit zero from a blank override.
const numericInputs=new Map(spec.input_map.filter(e=>!['input_notes','override_notes','selected_case'].includes(e.path[0])).map(e=>[e.path.join('/'),e]));
for (const entry of spec.input_map.filter(e=>e.path[0]==='input_notes')) {
  const source=numericInputs.get(entry.path[1]);
  const expected=sheets.get(source.sheet).getRange(source.cell).values[0][0];
  const actual=sheets.get('Input Notes').getRange(entry.cell.replace(/^I/,'G')).values[0][0];
  if ((expected ?? '') !== (actual ?? '')) throw new Error(`Input Notes mirror differs from ${source.sheet}!${source.cell}: expected ${expected}, got ${actual}`);
}

// Behavioral checks on the editable workbook, restoring every input afterward.
const a=sheets.get('Assumptions'), v=sheets.get('Valuation');
const originalCase=a.getRange('E6').values[0][0];
for (const [index,name] of ['Base','Bear','Bull'].entries()) {
  a.getRange('E6').values=[[name]];
  const actual=v.getRange('E23').values[0][0];
  const expected=sheets.get('Scenarios').getRange(`${'EFG'[index]}8`).values[0][0];
  if (typeof actual!=='number' || Math.abs(actual-expected)>1e-7) throw new Error(`Case selector failed: ${name}`);
}
a.getRange('E6').values=[[originalCase]];
const caseStart={Base:52,Bear:94,Bull:136}[originalCase];
const waccCell=`E${caseStart+28}`,growthCell=`E${caseStart+29}`;
const oldWacc=a.getRange(waccCell).values[0][0];
a.getRange(waccCell).values=[[a.getRange(growthCell).values[0][0]]];
if (typeof v.getRange('E23').values[0][0] === 'number') throw new Error('Invalid terminal growth did not block valuation');
a.getRange(waccCell).values=[[oldWacc ?? null]];
const laterCell=`I${caseStart+3}`, oldLater=a.getRange(laterCell).values[0][0];
const earlier=sheets.get('Forecast').getRange('H33').values[0][0];
const baseline=v.getRange('E23').values[0][0];
a.getRange(laterCell).values=[[oldLater+10]];
if (v.getRange('E23').values[0][0]===baseline || sheets.get('Forecast').getRange('H33').values[0][0]!==earlier) throw new Error('Later-year price input did not propagate correctly');
a.getRange(laterCell).values=[[oldLater]];
const firstPrice=`E${caseStart+3}`,oldPrice=a.getRange(firstPrice).values[0][0];
a.getRange(firstPrice).values=[[null]];
if (typeof v.getRange('E23').values[0][0]==='number') throw new Error('Missing price did not block valuation');
a.getRange(firstPrice).values=[[oldPrice]];

const scan=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:100},summary:'Formula error scan'});
await fs.writeFile(path.join(outDir,'formula_error_scan.ndjson'),scan.ndjson);
for (const def of spec.sheets) {
  for (const cell of def.cells) {
    if (!cell.formula) continue;
    const value=sheets.get(def.name).getRange(cell.address).values[0][0];
    if (typeof value==='string' && /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)/.test(value)) {
      throw new Error(`Formula error at ${def.name}!${cell.address}: ${value}`);
    }
  }
}
const table=await wb.inspect({kind:'table',range:'Valuation!C6:E30',include:'values,formulas',tableMaxRows:25,tableMaxCols:3,maxChars:5000});
await fs.writeFile(path.join(outDir,'valuation_inspection.ndjson'),table.ndjson);
const inputNotes=await wb.inspect({kind:'table',range:"'Input Notes'!C7:I13",include:'values,formulas',tableMaxRows:7,tableMaxCols:7,maxChars:5000});
await fs.writeFile(path.join(outDir,'input_notes_inspection.ndjson'),inputNotes.ndjson);

if (!flags.includes('--no-previews')) {
  const previewDir=path.join(outDir,'previews');
  await fs.mkdir(previewDir,{recursive:true});
  for (const def of spec.sheets) {
    for (const [i,range] of def.renders.entries()) {
      const blob=await wb.render({sheetName:def.name,range,scale:1,format:'png'});
      await fs.writeFile(path.join(previewDir,`${def.name.toLowerCase()}_${i+1}.png`),new Uint8Array(await blob.arrayBuffer()));
    }
  }
}
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);
await fs.writeFile(outputPath.replace(/\.xlsx$/i,'.inputs.json'),JSON.stringify({schema_version:1,input_map:spec.input_map,config:spec.config},null,2));
console.log(JSON.stringify({output:outputPath,formula_comparisons:spec.comparisons.length,behavioral_checks:5,value_per_adr:v.getRange('E23').values[0][0]}));
