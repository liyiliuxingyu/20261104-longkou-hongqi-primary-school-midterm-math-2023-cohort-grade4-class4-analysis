/**
 * Current continuation path: import the completed seven-sheet baseline,
 * recalculate its editable formulas, export, and preserve native metadata.
 * Unlike the archived five-sheet generator, this does not recreate or discard
 * existing worksheets, manual audit responses, names, or score inputs.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

async function main() {
  const root=path.resolve(process.env.GRADE4_PROJECT_ROOT ?? fileURLToPath(new URL('..',import.meta.url)));
  const config=JSON.parse(await fs.readFile(path.join(root,'project.json'),'utf8'));
  const input=path.resolve(process.env.GRADE4_INPUT_XLSX ?? path.join(root,config.baseline));
  const outputDir=path.resolve(process.env.GRADE4_OUTPUT_DIR ?? path.join(root,'outputs','rebuild'));
  const python=process.env.GRADE4_PYTHON ?? process.env.CODEX_PRIMARY_RUNTIME_PYTHON ?? 'python3';
  const stats=JSON.parse(await fs.readFile(path.join(outputDir,'analysis','analysis.json'),'utf8'));
  const workbook=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
  const names=[];
  for(let i=0;i<7;i++) names.push(workbook.worksheets.getItemAt(i).name);
  for(const name of ['一年级','二年级','三年级','四年级','五年级','四年级分析','数据核查']) {
    if(!names.includes(name)) throw new Error(`Required baseline worksheet is missing: ${name}`);
  }
  const sheet=workbook.worksheets.getItem('四年级');
  workbook.recalculate();
  const value=(address)=>sheet.getRange(address).values[0][0];
  const checks=[];
  function equal(address,expected) {
    const actual=value(address);
    if(typeof actual!=='number'||Math.abs(actual-expected)>1e-8) {
      throw new Error(`Recalculation mismatch at 四年级!${address}: ${JSON.stringify(actual)} / ${expected}`);
    }
    checks.push({cell:`四年级!${address}`,expected,actual});
  }
  equal('G50',stats.all.names);
  equal('G51',stats.all.base_count);
  equal('G54',stats.all.base_sum);
  equal('G65',stats.all.total_sum);
  for(let i=0;i<5;i++) {
    const c=String.fromCharCode(66+i), expected=stats.classes[i];
    equal(`${c}50`,expected.names); equal(`${c}51`,expected.base_count);
    equal(`${c}54`,expected.base_sum); equal(`${c}55`,expected.base_mean);
    equal(`${c}65`,expected.sum);
  }
  const errors=await workbook.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},maxChars:3000});
  await fs.mkdir(outputDir,{recursive:true});
  const target=path.join(outputDir,config.rebuilt_filename);
  if(path.resolve(target)===input) throw new Error('Refusing to overwrite the baseline input.');
  const output=await SpreadsheetFile.exportXlsx(workbook);
  await output.save(target);
  const metadataOutput=execFileSync(python,[path.join(root,'src','preserve_native_metadata.py'),input,target],{encoding:'utf8'});
  let metadata;
  try {metadata=JSON.parse(metadataOutput);} catch {metadata={stdout:metadataOutput};}
  const report={status:'passed',baseline_name:path.basename(input),output_name:path.basename(target),
    checks,formula_error_scan:errors.ndjson,native_metadata:metadata,
    authored_changes:'No score/name edits. Recalculated existing formulas and re-exported the seven-sheet baseline.',
    narrative_policy:'Dated 2026-11-06 commentary is a historical snapshot; review it when input scores change.'};
  await fs.writeFile(path.join(outputDir,'engine_selfcheck.json'),JSON.stringify(report,null,2)+'\n');
  if(process.env.GRADE4_RENDER==='1') {
    for(const [sheetName,range,name] of [['四年级分析','A1:L24','analysis'],['四年级分析','A26:L42','charts'],['数据核查','A1:G14','audit']]) {
      const preview=await workbook.render({sheetName,range,scale:1.1,format:'png'});
      await fs.writeFile(path.join(outputDir,`${name}.png`),new Uint8Array(await preview.arrayBuffer()));
    }
  }
  console.log(JSON.stringify({status:'passed',output:target,checked_cells:checks.length},null,2));
}

main().catch(error=>{console.error(error.name+': '+error.message);process.exitCode=1;});
