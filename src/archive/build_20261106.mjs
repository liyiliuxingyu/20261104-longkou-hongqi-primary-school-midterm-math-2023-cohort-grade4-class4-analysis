/*
 * Historical final workbook builder, restored from the visible conversation
 * on 2026-11-20, with all visible bug-fix patches applied.
 * This is a reconstruction of source text, not a recovered byte-identical file.
 * Path constants are now configurable. It requires the ORIGINAL five-sheet
 * workbook and historical/recreated audit/analysis.json. It intentionally
 * refuses a finished seven-sheet workbook so it cannot silently overwrite
 * subsequent user edits. Use the package's current continuation entry instead.
 *
 * GRADE4_PROJECT_ROOT     project root (default: two levels above this script)
 * GRADE4_ORIGINAL_WORKBOOK original uploaded five-sheet workbook
 * GRADE4_ANALYSIS_JSON    historical or explicitly reconstructed analysis JSON
 * GRADE4_OUTPUT_DIR       isolated output directory (default: outputs/archive_rebuild)
 * GRADE4_PYTHON           Python interpreter for the metadata pass
 * Requires @oai/artifact-tool through the primary runtime, not bundled here.
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const scriptDir=path.dirname(fileURLToPath(import.meta.url));
const root=path.resolve(process.env.GRADE4_PROJECT_ROOT??path.join(scriptDir,'../..'));
const outputDir=path.resolve(process.env.GRADE4_OUTPUT_DIR??path.join(root,'outputs/archive_rebuild'));
const previewDir=path.join(root,'audit/previews');
const input=path.resolve(process.env.GRADE4_ORIGINAL_WORKBOOK??path.join(root,'inputs/original_upload/2026.11期中考试四年级名单.xlsx'));
const analysisPath=path.resolve(process.env.GRADE4_ANALYSIS_JSON??path.join(root,'audit/analysis.json'));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const sheetInventory=(await wb.inspect({kind:'sheet',include:'id,name',maxChars:10000})).ndjson;
if(sheetInventory.split('\n').some(line=>{try{return ['四年级分析','数据核查'].includes(JSON.parse(line).name);}catch{return false;}})){
  throw new Error('此历史生成器需要原上传的五工作表版本，不能直接用于已含“四年级分析/数据核查”的成品。请使用工程包的当前续作入口，或显式提供真正的原上传文件。');
}
const data=JSON.parse(await fs.readFile(analysisPath,'utf8'));
await fs.mkdir(previewDir,{recursive:true});
const source=wb.worksheets.getItem('四年级');
const summary=wb.worksheets.add('四年级分析');
const audit=wb.worksheets.add('数据核查');
const font='Arial';
const navy='#263F60', blue='#426EA2', pale='#EEF3F9', amber='#FFF2CC', red='#9C3036';
const groups=['ABCDE','FGHIJ','KLMNO','PQRST','UVWXY'];
const markedRows=[[3,13,44,45],[12,18,43,45],[7,9,29,32,42],[32,37],[13,21,36,37]];
const col=i=>String.fromCharCode(65+i);
const ref=(c,r)=>`'四年级'!${c}${r}`;
const baseRanges=groups.map((g,i)=>`${ref(g[2],2)}:${g[2]}${data.classes[i].names+1}`);
const totalRanges=groups.map((g,i)=>`${ref(g[4],2)}:${g[4]}${data.classes[i].names+1}`);
const cellValue=(s,a)=>s.getRange(a).values[0][0];
const formula=(s,a,f)=>{s.getRange(a).formulas=[[f]];};
function header(s,r){s.getRange(r).format={fill:navy,font:{name:font,size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true};}
function section(s,a,text){s.getRange(a).values=[[text]];s.getRange(a).format.font={name:font,size:12,bold:true,color:navy};}
function baseline(s,r){s.showGridLines=false;s.getRange(r).format={font:{name:font,size:11,color:'#253246'},verticalAlignment:'center',rowHeight:23};}
function stripe(s,start,end,lastCol){for(let r=start;r<=end;r++){if((r-start)%2===1)s.getRange(`A${r}:${lastCol}${r}`).format.fill='#F4F6F9';}}
function tableEnd(s,r){s.getRange(r).format={fill:pale,font:{name:font,size:11,bold:true,color:navy},borders:{top:{style:'thin',color:navy}}};}

// Keep every original name, sex, score input and name color. Row coordinates stay fixed.
source.freezePanes.freezeRows(1);
source.getRange('A1:Y1').format.rowHeight=27;
source.getRange('A2:Y45').format.rowHeight=21;
for(let i=0;i<groups.length;i++){
  const g=groups[i],last=data.classes[i].names+1;
  source.getRange(`${g[0]}1:${g[4]}1`).values=[[`${i+1}班姓名`,'性别','基础分','附加分','总分']];
  header(source,`${g[0]}1:${g[4]}1`);
  source.getRange(`${g[0]}1:${g[0]}70`).format.columnWidth=18;
  source.getRange(`${g[1]}1:${g[1]}70`).format.columnWidth=9;
  source.getRange(`${g[2]}1:${g[3]}70`).format.columnWidth=10;
  source.getRange(`${g[4]}1:${g[4]}70`).format.columnWidth=11;
  source.getRange(`${g[1]}2:${g[1]}${last}`).format.horizontalAlignment='center';
  source.getRange(`${g[2]}2:${g[4]}${last}`).setNumberFormat('General');
  source.getRange(`${g[2]}2:${g[4]}${last}`).format.horizontalAlignment='right';
  const fs=[];
  for(let r=2;r<=last;r++)fs.push([`=IF(ISNUMBER(${g[2]}${r}),SUM(${g[2]}${r}:${g[3]}${r}),"")`]);
  source.getRange(`${g[4]}2:${g[4]}${last}`).formulas=fs;
  source.getRange(`${g[1]}2:${g[1]}${last}`).dataValidation={rule:{type:'list',values:['男','女']}};
  source.getRange(`${g[2]}2:${g[3]}${last}`).conditionalFormats.addCustom(`AND($${g[0]}2<>"",ISBLANK(${g[2]}2))`,{fill:amber});
  source.getRange(`${g[4]}2:${g[4]}${last}`).conditionalFormats.addCustom(`AND($${g[0]}2<>"",OR(ISBLANK($${g[2]}2),ISBLANK($${g[3]}2)))`,{fill:amber,font:{color:'#875400'}});
}
// Remove former totals from student-adjacent rows; preserve originals on the audit sheet.
for(const a of ['O44:O45','T44:T45','Y44:Y45'])source.getRange(a).clear({applyTo:'contents'});
source.getRange('A46:Y53').clear({applyTo:'all'});
source.getRange('A46').values=[['浅黄色为缺录或待核分数；空基础分不计作0分；附加分空白的总分仅供参考。']];
source.getRange('A46:Y46').format={font:{name:font,size:11,color:'#875400'},rowHeight:25};
for(let i=0;i<groups.length;i++){
  const g=groups[i];source.getRange(`${g[0]}47`).values=[['现录总分合计']];
  formula(source,`${g[4]}47`,`=SUM(${g[4]}2:${g[4]}${data.classes[i].names+1})`);
  source.getRange(`${g[0]}47:${g[4]}47`).format={fill:pale,font:{bold:true,color:navy,size:11},rowHeight:26};
}
source.getRange('A48').values=[['统计与图表见“四年级分析”；原汇总及待核实名单见“数据核查”。']];
source.getRange('A48:Y48').format={font:{name:font,size:11,color:'#536274'},rowHeight:25};
source.getRange('A49:G49').values=[['统计项目','1班','2班','3班','4班','5班','全年级']];
header(source,'A49:G49');
const metricLabels={50:'名单人数',51:'基础分已录人数',52:'基础分缺录人数',53:'仅缺附加分人数',54:'基础分合计',55:'基础分均分',56:'基础分中位数',57:'基础分标准差',58:'基础分＜60人数',59:'基础分60—＜70',60:'基础分70—＜80',61:'基础分80—＜90',62:'基础分≥90人数',63:'基础分≥60占比',64:'基础分≥90占比',65:'现录总分合计',66:'现录总分均分',67:'基础分≥80占比',68:'两项成绩齐全人数',69:'男生名单人数',70:'女生名单人数'};
source.getRange('A50:G70').format={font:{name:font,size:10,color:'#253246'},rowHeight:24,verticalAlignment:'center'};
source.getRange('B50:G70').format.horizontalAlignment='right';
source.getRange('B50:G70').setNumberFormat('0');
for(const [r,label] of Object.entries(metricLabels))source.getRange(`A${r}`).values=[[label]];
for(let i=0;i<groups.length;i++){
  const g=groups[i],end=data.classes[i].names+1,cc=col(i+1);
  const nr=`${g[0]}2:${g[0]}${end}`,br=`${g[2]}2:${g[2]}${end}`,ar=`${g[3]}2:${g[3]}${end}`,sr=`${g[1]}2:${g[1]}${end}`;
  const expressions={50:`COUNTA(${nr})`,51:`COUNT(${br})`,52:`${cc}50-${cc}51`,53:`${cc}51-COUNTIFS(${br},"<>",${ar},"<>")`,54:`SUM(${br})`,55:`IF(${cc}51=0,"",${cc}54/${cc}51)`,56:`IF(${cc}51=0,"",MEDIAN(${br}))`,57:`IF(${cc}51=0,"",STDEVP(${br}))`,58:`${cc}51-COUNTIFS(${br},">=60")`,59:`COUNTIFS(${br},">=60",${br},"<70")`,60:`COUNTIFS(${br},">=70",${br},"<80")`,61:`COUNTIFS(${br},">=80",${br},"<90")`,62:`COUNTIFS(${br},">=90")`,63:`IF(${cc}51=0,"",1-${cc}58/${cc}51)`,64:`IF(${cc}51=0,"",${cc}62/${cc}51)`,65:`${g[4]}47`,66:`IF(${cc}51=0,"",${cc}65/${cc}51)`,67:`IF(${cc}51=0,"",SUM(${cc}61:${cc}62)/${cc}51)`,68:`${cc}51-${cc}53`,69:`COUNTIFS(${sr},"男")`,70:`COUNTIFS(${sr},"女")`};
  for(const [r,e] of Object.entries(expressions))formula(source,`${cc}${r}`,`=${e}`);
}
for(const r of [50,51,52,53,54,58,59,60,61,62,65,68,69,70])formula(source,`G${r}`,`=SUM(B${r}:F${r})`);
for(const [r,e] of Object.entries({55:'IF(G51=0,"",G54/G51)',56:`IF(G51=0,"",MEDIAN(${baseRanges.join(',')}))`,57:`IF(G51=0,"",STDEVP(${baseRanges.join(',')}))`,63:'IF(G51=0,"",1-G58/G51)',64:'IF(G51=0,"",G62/G51)',66:'IF(G51=0,"",G65/G51)',67:'IF(G51=0,"",SUM(G61:G62)/G51)'}))formula(source,`G${r}`,`=${e}`);
for(const r of [55,56,57,66])source.getRange(`B${r}:G${r}`).setNumberFormat('0.00');
for(const r of [63,64,67])source.getRange(`B${r}:G${r}`).setNumberFormat('0.00%');
source.getRange('G50:G70').format.fill=pale;
source.getRange('A72').values=[['统计按已录基础分进行，0分保留。缺附加分人数仅指已有基础分者。标准差采用总体公式。']];
source.getRange('A73').values=[['60和90仅为分段界线。试卷满分、原姓名颜色含义及学校正式统计范围仍待确认。']];

// Main analysis: linked to the calculations above, with weighted grade averages.
baseline(summary,'A1:L51');summary.tabColor=navy;
const summaryWidths=[10,12,12,12,12,12,12,12,14,14,16,16];
summaryWidths.forEach((w,i)=>{summary.getRange(`${col(i)}1:${col(i)}51`).format.columnWidth=w;});
summary.getRange('A2').values=[['四年级期中成绩分析']];summary.getRange('A2').format.font={name:font,size:16,bold:true,color:navy};
summary.getRange('A3').values=[['2026年11月。本页以已录基础分的学生为统计对象，原已录0分保留，空白不作0分。']];
summary.getRange('A4').values=[['暂不按原姓名颜色剔除学生。满分与学校等级标准未注明，60、90仅作分段界线。']];
summary.getRange('A3:L4').format.font={name:font,size:10,color:'#536274'};
summary.getRange('A5:L5').format.borders={bottom:{style:'thin',color:'#CBD5E1'}};
summary.getRange('A6:L6').values=[['班级','名单人数','基础分人数','缺基础分','仅缺附加分','基础均分','中位数','标准差','基础≥60占比','基础≥90占比','现录总分合计','现录总分均分']];header(summary,'A6:L6');summary.getRange('A6:L6').format.rowHeight=38;
const summaryMetrics=[50,51,52,53,55,56,57,63,64,65,66];
for(let i=0;i<6;i++){
  const r=7+i,sc=col(i+1);summary.getRange(`A${r}`).values=[[i===5?'全年级':`${i+1}班`]];
  summary.getRange(`B${r}:L${r}`).formulas=[summaryMetrics.map(m=>[55,66].includes(m)?`=IF('四年级'!${sc}51=0,"",ROUND('四年级'!${sc}${m===55?54:65}*100/'四年级'!${sc}51,0)/100)`:[56,57].includes(m)?`=ROUND('四年级'!${sc}${m},2)`:`='四年级'!${sc}${m}`)];
}
stripe(summary,7,11,'L');tableEnd(summary,'A12:L12');
summary.getRange('B7:L12').format.horizontalAlignment='right';summary.getRange('B7:E12').setNumberFormat('0');summary.getRange('F7:H12').setNumberFormat('0.00');summary.getRange('I7:J12').setNumberFormat('0.00%');summary.getRange('K7:K12').setNumberFormat('0');summary.getRange('L7:L12').setNumberFormat('0.00');
summary.getRange('D7:E12').conditionalFormats.add('cellIs',{operator:'greaterThan',formula:0,format:{fill:amber,font:{color:'#875400'}}});
summary.getRange('A14').values=[['总分为参考口径：有基础分者沿用原表加分算法，附加分空白时暂不加分，不能认定其附加分已确认是0。']];
summary.getRange('A14:L14').format={font:{name:font,size:10,color:'#875400'},rowHeight:25};
section(summary,'A16','基础分分段人数');section(summary,'I16','全年级分段分布');
summary.getRange('A17:G17').values=[['班级','＜60','60—＜70','70—＜80','80—＜90','≥90','人数']];header(summary,'A17:G17');
for(let i=0;i<6;i++){
  const r=18+i,sc=col(i+1);summary.getRange(`A${r}`).values=[[i===5?'全年级':`${i+1}班`]];
  summary.getRange(`B${r}:G${r}`).formulas=[[58,59,60,61,62,51].map(m=>`='四年级'!${sc}${m}`)];
}
stripe(summary,18,22,'G');tableEnd(summary,'A23:G23');summary.getRange('B18:G23').setNumberFormat('0');summary.getRange('B18:G23').format.horizontalAlignment='right';
summary.getRange('I17:K17').values=[['分数段','人数','占比']];header(summary,'I17:K17');
const bands=['＜60','60—＜70','70—＜80','80—＜90','≥90'];
for(let i=0;i<5;i++){
  const r=18+i;summary.getRange(`I${r}`).values=[[bands[i]]];formula(summary,`J${r}`,`=${col(1+i)}23`);formula(summary,`K${r}`,`=IF($G$23=0,"",J${r}/$G$23)`);
}
summary.getRange('J18:J22').setNumberFormat('0');summary.getRange('K18:K22').setNumberFormat('0.00%');
formula(summary,'I24','="名单：男生 "&\'四年级\'!G69&" 人，女生 "&\'四年级\'!G70&" 人"');summary.getRange('I24:L24').format.font={name:font,size:10,color:'#536274'};
const ch1=summary.charts.add('bar',[summary.getRange('A6:A11'),summary.getRange('F6:F11'),summary.getRange('G6:G11')]);
ch1.title='各班基础分均值与中位数（分）';ch1.setPosition('A26','F41');ch1.titleTextStyle.fontSize=15;ch1.titleTextStyle.typeface=font;ch1.legend={position:'top',textStyle:{typeface:font,fontSize:12}};
ch1.xAxis={axisType:'textAxis',textStyle:{typeface:font,fontSize:12}};ch1.yAxis={min:0,max:100,numberFormatCode:'0',numberFormatSourceLinked:false,textStyle:{typeface:font,fontSize:12}};
ch1.series.items[0].fill=blue;ch1.series.items[1].fill='#AEBFD4';
const ch2=summary.charts.add('bar',summary.getRange('I17:J22'));
ch2.title='全年级基础分分布（人）';ch2.setPosition('H26','L41');ch2.hasLegend=false;ch2.titleTextStyle.fontSize=15;ch2.titleTextStyle.typeface=font;ch2.xAxis={axisType:'textAxis',textStyle:{typeface:font,fontSize:12}};ch2.yAxis={min:0,max:80,numberFormatCode:'0',numberFormatSourceLinked:false,textStyle:{typeface:font,fontSize:12}};ch2.series.items[0].fill=blue;
section(summary,'A44','本次核查的成绩结构（2026-11-06）');
const findings=[
 ['1班','基础均分最高，成绩分散程度较小。仍有2名基础分低于60的学生需要结合答卷分析。'],
 ['2班','基础分低于60的有8人，均低于50。均值明显低于中位数，应先核对记录，再分析个别困难。'],
 ['3班','基础分≥90的比例最高。中位数与1班相同，同时有5人低于60，需要兼顾不同分段。'],
 ['4班','基础均分与3班只差约0.17分。现录总分的差距主要来自附加部分，暂不夸大排序差异。'],
 ['5班','中位数79，60—79分有15人。后续讲评需要兼顾中间分段；原汇总另有3分差异待核。']
];
findings.forEach((v,i)=>summary.getRange(`A${45+i}:B${45+i}`).values=[v]);summary.getRange('A45:L49').format.font={name:font,size:10,color:'#253246'};
summary.getRange('A51').values=[['本表只有基础分与附加分，无法确定逐题错误、具体错因或进退步。需要原卷、逐题得分及前次成绩进一步分析。']];summary.getRange('A51:L51').format.font={name:font,size:10,color:'#536274'};

// Terminal audit: original manual figures and a labelled conditional comparison.
baseline(audit,'A1:G63');audit.tabColor='#A76E27';
[20,14,16,20,18,24,56].forEach((w,i)=>audit.getRange(`${col(i)}1:${col(i)}63`).format.columnWidth=w);
audit.getRange('A2').values=[['四年级数据核查']];audit.getRange('A2').format.font={name:font,size:16,bold:true,color:navy};
audit.getRange('A3').values=[['来源：2026.11期中考试四年级名单.xlsx“四年级”，本次上传原值。核查日期：2026-11-06。']];audit.getRange('A3:G3').format.font={name:font,size:10,color:'#536274'};
audit.getRange('A5:F5').values=[['原汇总与条件复算','1班','2班','3班','4班','5班']];header(audit,'A5:F5');
const auditLabels=['原手填总分','原手填人数','原手填均分','原标色姓名人数','条件复算人数','条件复算总分','条件复算均分','原总分减条件总分'];
auditLabels.forEach((v,i)=>audit.getRange(`A${6+i}`).values=[[v]]);
for(let i=0;i<5;i++){
  const cc=col(i+1),g=groups[i],cl=data.classes[i],end=cl.names+1;
  audit.getRange(`${cc}6:${cc}9`).values=[[cl.old_sum],[cl.old_count],[cl.old_mean],[markedRows[i].length]];
  formula(audit,`${cc}10`,`=COUNTA(${ref(g[0],2)}:${g[0]}${end})-${cc}9`);
  formula(audit,`${cc}11`,`=SUM(${totalRanges[i]})-SUM(${markedRows[i].map(r=>ref(g[4],r)).join(',')})`);
  formula(audit,`${cc}12`,`=IF(${cc}10=0,"",${cc}11/${cc}10)`);
  formula(audit,`${cc}13`,`=${cc}6-${cc}11`);
}
audit.getRange('B6:F13').format.horizontalAlignment='right';audit.getRange('B6:F13').setNumberFormat('0');for(const r of [8,12])audit.getRange(`B${r}:F${r}`).setNumberFormat('0.000');
audit.getRange('B13:F13').conditionalFormats.add('cellIs',{operator:'notEqual',formula:0,format:{fill:'#FCE5E7',font:{color:red,bold:true}}});
audit.getRange('G5:G13').format={wrapText:true,verticalAlignment:'center',font:{name:font,size:10,color:'#536274'}};
audit.getRange('G5').values=[['条件复算仅用于解释原汇总']];audit.getRange('G5').format.font={bold:true,color:navy};
audit.getRange('G6').values=[['假设所有原有色姓名不纳入统计。颜色含义未经确认，主分析仍保留所有已录基础分。']];audit.getRange('G6:G7').format.rowHeight=40;
audit.getRange('G8').values=[['前四班人数与总分在该假设下均相符，不能由此认定排除原因或正式统计口径。']];audit.getRange('G8').format.rowHeight=42;
audit.getRange('G10').values=[['5班3149－9－8＝3132，3132÷38≈82.421。原填3129与明细相差3分。']];audit.getRange('G10').format.rowHeight=45;
audit.getRange('G12').values=[['5班任何现有成绩子集均不能合计3129；需核查原汇总或是否有未同步到明细的改分。']];audit.getRange('G12').format.rowHeight=45;
section(audit,'A16','原表有色姓名（仅作口径对照，未据此删除成绩）');
audit.getRange('A17:F17').values=[['班级','姓名','原姓名位置','基础分','附加分','现录总分']];header(audit,'A17:F17');
let markRow=18;
for(let i=0;i<5;i++)for(const r of markedRows[i]){
  const g=groups[i],rec=data.students.find(x=>x.class===i+1&&x.row===r);
  audit.getRange(`A${markRow}:C${markRow}`).values=[[`${i+1}班`,rec.name,`${g[0]}${r}`]];
  for(let j=0;j<3;j++)formula(audit,`${col(3+j)}${markRow}`,`=IF(ISNUMBER(${ref(g[2+j],r)}),${ref(g[2+j],r)},"")`);
  markRow++;
}
stripe(audit,18,36,'F');audit.getRange('D18:F36').setNumberFormat('General');audit.getRange('D18:F36').format.horizontalAlignment='right';
section(audit,'A39','缺录与其他核查项');
audit.getRange('A40:G40').values=[['班级','姓名','原位置','核查事项','核实结果（可填）','当前处理','说明']];header(audit,'A40:G40');audit.getRange('A40:G40').format.rowHeight=30;
const issueRows=[];
for(const rec of data.students.filter(x=>x.base===null)){
  const g=groups[rec.class-1];
  issueRows.push([`${rec.class}班`,rec.name,`${g[2]}${rec.row}:${g[3]}${rec.row}`,'基础、附加分均空','待核实','留空，不当0分',rec.class===2&&rec.row===45?'原J45为全班合计3338，已移至J47。个人成绩仍待补录。':'原总分因空白显示0，现留空。未录入原因待核实。']);
}
for(const rec of data.students.filter(x=>x.base!==null&&x.bonus===null)){
  const g=groups[rec.class-1];
  issueRows.push([`${rec.class}班`,rec.name,`${g[3]}${rec.row}`,'附加分空白','待核实','总分仅供参考',`基础分已录${rec.base}。请核实附加分是否应为0，或需要补录。`]);
}
issueRows.push(['2班／5班','刘浩宇','F11／U13','跨班同名','待核实','分别保留','不同班级同名不能据此认定重复录入，可对照学籍信息核实。']);
issueRows.push(['2班','车国霆','I3','附加分为5','可复核','保留原值','这是全表唯一附加5分，仅因出现次数少不能判错，可对照原卷。']);
issueRows.push(['2班','赵家辉','H43','基础分明确录为0','待核实','保留0分并纳入','核实是否实得0分，不能自动改为空白或排除。']);
issueRows.push(['3班','胡语诺','M9／N9','两项明确录为0','待核实','保留0分并纳入','核实是否实得0分，不能自动改为空白或排除。']);
audit.getRange(`A41:G${40+issueRows.length}`).values=issueRows;
stripe(audit,41,40+issueRows.length,'G');
audit.getRange(`A41:G${40+issueRows.length}`).format.rowHeight=42;audit.getRange(`C41:G${40+issueRows.length}`).format.wrapText=true;audit.getRange(`A41:G${40+issueRows.length}`).format.font={name:font,size:10,color:'#253246'};
audit.getRange(`E41:E${40+issueRows.length}`).format.fill=amber;audit.getRange(`E41:E${40+issueRows.length}`).dataValidation={rule:{type:'list',values:['待核实','可复核','已确认','已补录']}};
audit.getRange('A62').values=[['核查结果：214条姓名和性别完整，班内未发现同名或姓名首尾空格。199条两项分数齐全记录的个人加法均正确。']];
audit.getRange('A63').values=[['原公式缓存与逐项复算一致。原手填平均分与对应手填总分÷人数相符，主要问题在统计口径及5班汇总。']];audit.getRange('A62:G63').format.font={name:font,size:10,color:'#536274'};
audit.freezePanes.freezeRows(5);

// Historical meaningful recalculation checks; every temporary input is restored.
function eq(actual,expected,label,tolerance=1e-8){if(typeof expected==='number'){if(typeof actual!=='number'||Math.abs(actual-expected)>tolerance)throw new Error(`${label}: ${JSON.stringify(actual)} != ${expected}`);}else if(actual!==expected)throw new Error(`${label}: ${actual} != ${expected}`);}
eq(cellValue(source,'G50'),214,'roster');eq(cellValue(source,'G51'),206,'recorded bases');eq(cellValue(source,'G53'),7,'missing bonuses');eq(cellValue(source,'G65'),17053,'recorded total');
eq(cellValue(source,'J45'),'','summary removed from student row');
eq(cellValue(source,'J43'),0,'existing base zero preserved');eq(cellValue(source,'O9'),0,'existing two zero scores preserved');eq(cellValue(audit,'D24'),0,'audit numeric zero preserved');eq(cellValue(audit,'E27'),0,'audit bonus zero preserved');
source.getRange('C44').values=[[0]];
eq(cellValue(source,'E44'),0,'real zero preserved');eq(cellValue(source,'G51'),207,'blank to zero counted');eq(cellValue(source,'G52'),7,'blank to zero missing count');
source.getRange('C44').values=[[null]];
eq(cellValue(source,'E44'),'','restored missing remains blank');eq(cellValue(source,'G51'),206,'roster restored');
source.getRange('I18').values=[[4]];
eq(cellValue(source,'J18'),14,'bonus update total');eq(cellValue(summary,'E8'),3,'missing bonus count updates');eq(cellValue(summary,'K12'),17057,'grade total update');
source.getRange('I18').values=[[null]];
wb.recalculate();
for(let i=0;i<5;i++){
  const sc=col(i+1),st=data.classes[i];
  eq(cellValue(source,`${sc}50`),st.names,`class ${i+1} names`);
  eq(cellValue(source,`${sc}51`),st.base_count,`class ${i+1} count`);
  eq(cellValue(source,`${sc}55`),st.base_mean,`class ${i+1} base mean`);
  eq(cellValue(source,`${sc}56`),st.base_median,`class ${i+1} median`);
  eq(cellValue(source,`${sc}57`),st.base_sd,`class ${i+1} sd`);
  eq(cellValue(source,`${sc}65`),st.sum,`class ${i+1} total`);
  st.base_bands.forEach((v,k)=>eq(cellValue(source,`${sc}${58+k}`),v,`class ${i+1} band ${k}`));
}
eq(cellValue(audit,'F11'),3132,'class 5 conditional sum');eq(cellValue(audit,'F13'),-3,'class 5 discrepancy');
eq(cellValue(summary,'L11'),78.73,'class 5 displayed mean rounded correctly');
console.log((await wb.inspect({kind:'table',range:'四年级分析!A6:L12',include:'values,formulas',tableMaxRows:7,tableMaxCols:12,maxChars:7000})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},summary:'final formula error scan',maxChars:1500})).ndjson);
await fs.mkdir(outputDir,{recursive:true});
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(`${outputDir}/2026.11期中考试四年级名单.xlsx`);
const python=process.env.GRADE4_PYTHON??process.env.CODEX_PRIMARY_RUNTIME_PYTHON??'python3';
console.log(execFileSync(python,[path.join(scriptDir,'preserve_native_metadata_20261106.py'),input,`${outputDir}/2026.11期中考试四年级名单.xlsx`],{encoding:'utf8'}));
const rendered=[['四年级分析','A1:L24','summary_table'],['四年级分析','A26:M42','summary_charts'],['四年级分析','A44:L51','summary_findings'],['四年级','A1:J15','updated_source_top'],['四年级','A43:J55','updated_source_bottom'],['数据核查','A1:G14','audit_comparison'],['数据核查','A16:F36','audit_marked'],['数据核查','A39:G59','audit_issues']];
for(const [sheetName,range,name] of rendered){
  const p=await wb.render({sheetName,range,scale:1.15,format:'png'});
  await fs.writeFile(`${previewDir}/${name}.png`,new Uint8Array(await p.arrayBuffer()));
  console.log(`Rendered ${name}`);
}
await fs.writeFile(path.join(root,'audit/validation.json'),JSON.stringify({allChecksPassed:true,recordedBaseCount:cellValue(source,'G51'),missingBase:cellValue(source,'G52'),missingBonus:cellValue(source,'G53'),sourceSum:cellValue(source,'G65'),gradeBaseMean:cellValue(source,'G55'),conditionalClass5Sum:cellValue(audit,'F11'),charts:summary.charts.items.map(c=>({title:c.title.text,series:c.series.items.map(s=>({formula:s.formula,categoryFormula:s.categoryFormula}))}))},null,2));
console.log(`Saved ${outputDir}/2026.11期中考试四年级名单.xlsx`);
