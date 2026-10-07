import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';

const root='/Users/kwu/PHVS';
const out=root+'/outputs/01a111f7-26b0-7fc2-b98f-41db76412d40';
const qa='/tmp/phvs_reverse_dcf_qa';
await fs.mkdir(qa,{recursive:true});
const wb=Workbook.create();
const names=['Reverse SOTP','Inputs','DCF','Launch comparison','Street evidence'];
const sheets=Object.fromEntries(names.map(n=>[n,wb.worksheets.add(n)]));
const navy='#17324D',blue='#0000FF',green='#008000',amber='#FFF2CC',grey='#667085';
const num='#,##0.0;(#,##0.0);"-"',price='$0.00;($0.00);"-"',pct='0.0%;(0.0%);"-"';
const col=n=>{let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;};
const val=(s,a,v)=>s.getRange(a).values=[[v]];
const formula=(s,a,v)=>{s.getRange(a).formulas=[[v]];s.getRange(a).format.font.color=v.includes('!')?green:'#111111';};
function title(s,t,last='O',rows=105){
 s.showGridLines=false;s.getRange(`A1:${last}${rows}`).format.font={name:'Arial',size:10,color:'#172B4D'};
 s.getRange(`A1:B${rows}`).format.columnWidth=2.5;
 s.getRange(`C1:C${rows}`).format.columnWidth=42;
 s.getRange(`D1:D${rows}`).format.columnWidth=3;
 s.getRange(`E1:${last}${rows}`).format.columnWidth=14;
 s.getRange(`C1:${last}${rows}`).format.rowHeight=21;
 s.getRange(`C1:${last}${rows}`).format.verticalAlignment='center';
 val(s,'C2',t);s.getRange('C2').format.font={name:'Arial',size:14,bold:true,color:navy};
 s.getRange(`C3:${last}3`).format.borders={bottom:{style:'thin',color:navy}};
 s.getRange(`E1:${last}${rows}`).setNumberFormat(num);
}
function band(s,row,text,last='O'){val(s,`C${row}`,text);s.getRange(`C${row}:${last}${row}`).format={fill:navy,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},rowHeight:24};}
function note(s,cell,text){val(s,cell,text);s.getRange(cell).format.font.color=grey;}
function input(s,cell,value,format=num){val(s,cell,value);s.getRange(cell).setNumberFormat(format);s.getRange(cell).format.font.color=blue;s.getRange(cell).format.fill=amber;}
function setrow(s,row,label,values){val(s,`C${row}`,label);s.getRange(`E${row}:${col(4+values.length)}${row}`).values=[values];}
function linkrow(s,row,label,forms){val(s,`C${row}`,label);s.getRange(`E${row}:${col(4+forms.length)}${row}`).formulas=[forms];s.getRange(`E${row}:${col(4+forms.length)}${row}`).format.font.color=forms.some(f=>f.includes('!'))?green:'#111111';}
function cfError(s,address,formulaText){s.getRange(address).conditionalFormats.addCustom(formulaText,{fill:'#FEE4E2',font:{color:'#B42318',bold:true}});}
for(const [name,s] of Object.entries(sheets))title(s,name);
sheets['Reverse SOTP'].tabColor=navy;sheets.Inputs.tabColor='#41658A';
const e=sheets['Street evidence'];
note(e,'C4','User-supplied September 2026 notes. A missing input is not zero. PTs are sell-side opinions, not buy-side consensus.');
const brokers=[
 ['Wolfe','2026-09-08',42,null,null,null,'Narrative / multiple framework','124292892 p1,3','No numerical DCF or PoS in supplied note'],
 ['BofA','2026-09-08',44,null,.70,.10,'DCF; terminal not disclosed','124292329 p1,3,4','XR $21/share; IR $14 derived; HAE total $35; $917m nominal XR peak'],
 ['Guggenheim','2026-09-09',44,.90,.90,.125,'2036; terminal g -5%','124306319 p1–3','Detailed ramps; printed tax / PV arithmetic requires normalization'],
 ['KBC','2026-09-08',49,null,.90,null,'6-month target convention','124291310 p1,3,4','Insufficient cash-flow detail for complete reconstruction'],
 ['Wedbush','2026-09-08',50,null,null,.15,'SOTP DCF; terminal undisclosed','124292125 p4,5,7','IR $24, XR $21, cash $5; PoS absent; table unit-heading inconsistency'],
 ['H.C. Wainwright','2026-09-08',60,.95,.90,.12,'Terminal multiple 6x','124292122 p1','AAE PoS 50%; multiple base not specified in supplied paragraph'],
 ['RBC','2026-09-09',62,.90,null,.105,'DCF g1.5% plus 30x 2034 EPS','124298222 p3,6','Blended method; AAE PoS80%; do not treat as pure DCF'],
 ['Morgan Stanley','2026-09-09',62,.90,.90,.125,'Terminal g0%','124295137 p2,3','XR90% update supersedes stale70% sidebar; AAE70%'],
 ['Wells Fargo','2026-09-09',65,null,.90,.10,'No terminal growth (not no TV)','124306208 p2','90% stated with liver-signal context; product mapping not fully specified'],
 ['Jones','2026-09-09',65,null,.90,.12,'2035; terminal g0%','124307375 p3–5','Total revenue path available; IR/XR split not supplied'],
 ['Leerink','2026-09-08',65,.95,.95,.13,'2040; terminal g2%','124293821 p1,4–6','Updated note supersedes same-day $50 / XR60% follow-up appendix'],
 ['Cantor','2026-09-14',65,null,null,.12,'Beyond2040; terminal g4%','124381519 p3','PoS absent; terminal growth matters independently of adoption']
];
e.getRange('C6:K6').values=[['Broker','Note date','PT USD','IR PoS','XR PoS','Discount rate','Method','Source ID / PDF page','Limitations']];
e.getRange('C7:K18').values=brokers.map(r=>r.map(v=>v===null?'n.a.':v));
e.getRange('C6:K6').format={fill:navy,font:{color:'#FFFFFF',bold:true},rowHeight:30};
e.getRange('C7:C18').format.columnWidth=23;e.getRange('D6:D18').format.columnWidth=14;
e.getRange('F7:H18').setNumberFormat(pct);e.getRange('E7:E18').setNumberFormat(price);
e.getRange('I6:I18').format.columnWidth=36;e.getRange('J6:J18').format.columnWidth=27;e.getRange('K6:K18').format.columnWidth=78;
note(e,'C20','Closest supplied PTs to spot: Wolfe $42, BofA/Guggenheim $44, KBC $49, Wedbush $50. Guggenheim has the fullest lower-PT revenue path.');
note(e,'C21','Preserve all brokers. Do not average parameters across different valuation methods or count two Leerink notes as independent views.');
const years=Array.from({length:11},(_,i)=>2026+i);
band(e,23,'Guggenheim printed forecasts (EUR millions)');setrow(e,24,'Calendar year',years);
const ir=[0,54.385,109.336,219.809,303.809,416.439,502.326,589.094,676.751,765.304,854.760];
const xr=[0,0,64.981,146.126,220.893,289.175,350.867,405.859,454.041,495.304,529.535];
setrow(e,25,'PHVS-416 / IR reported revenue',ir);setrow(e,26,'PHVS-719 / XR reported revenue',xr);
setrow(e,28,'R&D, disclosed through 2030',[138.343,139.726,141.124,142.535,135.408,null,null,null,null,null,null]);
setrow(e,29,'G&A, disclosed through 2030',[63.394,69.734,73.220,73.952,74.692,null,null,null,null,null,null]);
setrow(e,31,'Printed GAAP EBIT',[-205.742,-159.200,-44.276,145.071,310.094,503.665,662.536,814.359,959.153,1105.473,1243.994]);
setrow(e,32,'Printed taxes',[0,0,0,0,0,24.583,99.380,246.117,143.873,165.821,186.599]);
setrow(e,33,'Printed tax-effected EBIT',[-205.742,-159.200,-44.276,145.071,310.094,528.248,761.917,1060.476,1103.026,939.652,1057.395]);
linkrow(e,34,'Printed after-tax less (EBIT - taxes)',years.map((_,i)=>`=${col(5+i)}33-(${col(5+i)}31-${col(5+i)}32)`));
e.getRange('E34:O34').setNumberFormat('0.000;(0.000);0.000');cfError(e,'E34:O34','ABS(E34)>0.002');
note(e,'C36','Source 124306319 pp2–3. Printed EBIT includes figures from the pre-tax-income row; normalization uses revenue less COGS/R&D/G&A.');
note(e,'C37','The source does not clearly identify whether product revenue is already probability weighted. Inputs exposes this assumption.');
band(e,39,'Reported valuation controls and market reference');
const refs=[['Guggenheim published PT, USD',44,'124306319 p1,3'],['Guggenheim diluted shares, millions',76.125,'124306319 p3'],['Cash at 30-Jun-2026, EUR millions',318.282453,'Pharvaris 6-K, 12-Aug-2026, note12'],['Basic shares at 30-Jun-2026, millions',70.204506,'Pharvaris 6-K, note13'],['H1 cash operating + investing use, EUR m',86.645936,'Pharvaris 6-K cash-flow statement'],['PHVS quote snapshot, USD',31.83,'7-Oct-2026 16:13:58 UTC; web finance quote'],['Published Guggenheim equity value, EUR m',2932.074,'124306319 p3; not forced into normalized DCF']];
refs.forEach((r,i)=>{val(e,`C${40+i}`,r[0]);val(e,`E${40+i}`,r[1]);note(e,`G${40+i}`,r[2]);});
note(e,'C48','Company filing: https://ir.pharvaris.com/node/12571/html');
band(e,50,'Observed launch revenue (USD millions)');
e.getRange('C51:G55').values=[['Product','Period','Global sales','Basis','Source'],['Orladeyo','2021',122.6,'First full calendar year','BioCryst FY2021 final'],['Orladeyo','2022',251.6,'Second full calendar year','BioCryst FY2022'],['Orladeyo','2023',326,'Third full calendar year','BioCryst FY2023'],['Orladeyo','2024',437.7,'Fourth full calendar year','BioCryst FY2024']];
val(e,'C57','Ekterly Jul7–Dec31 2025 launch period');val(e,'E57',49.1);note(e,'G57','Global drug sales; calendar stub, not 8 commercial months');
note(e,'C58','https://ir.kalvista.com/node/12706/html; Orladeyo URLs retained in output/Orladeyo_launch_analogue_2026-10-07.txt');
band(e,60,'Wedbush product forecasts, USD millions');
setrow(e,61,'Calendar year',[2027,2028,2029,2030,2031,2032,2033,2034]);
setrow(e,62,'IR worldwide',[8.450,117.201,305.571,532.750,688.890,792.526,842.511,873.913]);
setrow(e,63,'XR worldwide',[0,6.485,209.799,538.825,816.578,971.335,1043.810,1079.803]);
setrow(e,64,'IR US',[8.450,92.102,213.307,360.960,468.327,534.767,575.839,607.363]);
setrow(e,65,'XR US',[0,6.485,202.873,410.887,549.716,644.537,720.412,759.851]);
note(e,'C67','124292125 p5 image. Top heading says EUR000s, but dollar signs and USD-to-EUR conversion rows identify product rows as USD000s.');
note(e,'C68','124292125 p4: modeled XR net annual price ~$418k. This is a broker price assumption, not realized PHVS net pricing.');
band(e,70,'BioCryst competitive cross-checks');
note(e,'C72','RBC 124314632 p1: modeled ~30% out-year Orladeyo erosion; ~$400–450m annual tail; prior survey ~30 US specialists suggested ~35%.');
note(e,'C73','Evercore 124342693 pp1–2: illustrative reverse BCRX DCF at $8.40; Orladeyo $685m peak2027 to ~$550m2037; paid patients ~1,200 to ~700.');
note(e,'C74','These are broker scenarios, not observed patient losses or a one-for-one source of PHVS revenue. Evercore printed margin summaries differ.');
e.freezePanes.freezeRows(6);

const a=sheets.Inputs;
note(a,'C4','Blue/amber cells are editable. This is a normalized, conditional DCF, not a replication of Guggenheim or a consensus model.');
const entries=[
 [5,'Target basis: 1 spot / 2 broker / 3 custom',1,'0','One active target; all downstream outputs recalculate'],
 [6,'Active target price, USD',null,price,'Selected below; stock price and sell-side PT are distinct'],
 [7,'Active target horizon, years',null,num,'Spot0; broker1; custom uses E12. Future equity target discounted to valuation date'],
 [8,'Valuation date',new Date('2026-10-07T00:00:00Z'),'d-mmm-yyyy','Snapshot date; not a live quote'],
 [9,'Observed spot price, USD',31.83,price,'Street evidence E45; quote timestamp 7-Oct-2026 16:13:58 UTC'],
 [10,'Broker target, USD',44,price,'Guggenheim 9-Sep-2026; source124306319 p1'],
 [11,'Custom target, USD',50,price,'Scenario input'],
 [12,'Custom horizon, years',0,num,'Scenario input; leave0 for present value'],
 [14,'Solve: 1 both / 2 XR only / 3 IR only',1,'0','Unsolved product stays at 1.0x reference ramp'],
 [15,'Revenue basis: 1 success / 2 risk-adjusted',1,'0','UNCONFIRMED source basis. Default treats printed revenues as success-case'],
 [16,'IR probability of success',.90,pct,'Guggenheim p1. Held fixed while solving revenue'],
 [17,'XR probability of success',.90,pct,'Guggenheim p1. Held fixed while solving revenue'],
 [18,'Source IR PoS (for basis2 de-risking)',.90,pct,'Used only to unweight already risk-adjusted source revenue'],
 [19,'Source XR PoS (for basis2 de-risking)',.90,pct,'Used only to unweight already risk-adjusted source revenue'],
 [20,'Discount rate',.125,pct,'Guggenheim p1; common discount rate for skeleton'],
 [21,'Terminal growth rate',-.05,pct,'Guggenheim p1; negative-growth perpetuity beyond2036'],
 [22,'Terminal value: 1 include / 0 exclude',1,'0','Zero switches off terminal value; do not confuse g0% with no terminal value'],
 [23,'Cash tax rate',.15,pct,'Guggenheim p1; normalized tax subtraction'],
 [24,'First cash-tax year',2031,'0','Broker stated tax start; explicit NOL schedule not available'],
 [25,'COGS / revenue',.10,pct,'Derived from Guggenheim p2 forecast COGS and sales'],
 [26,'Working capital / revenue',.05,pct,'Assumption; source DCF leaves working-capital lines blank'],
 [27,'Annual capex, EUR millions',.6,num,'Assumption, not Guggenheim disclosure'],
 [28,'Annual depreciation, EUR millions',.6,num,'Assumption; matches capex for a light-asset skeleton'],
 [29,'USD per EUR',1.16,'0.000','Explicit FX assumption; not inferred from rounded PT'],
 [30,'Diluted shares, millions',76.125,'0.000','Guggenheim p3 projection; not current basic shares'],
 [31,'Cash at last report, EUR millions',318.282453,num,'30-Jun-2026 company filing'],
 [32,'Cash report date',new Date('2026-06-30T00:00:00Z'),'d-mmm-yyyy','Roll to valuation date once; exclude elapsed spend from future DCF'],
 [33,'Estimated daily cash burn, EUR millions',86.645936/181,'0.000','H1 operating + investing cash use /181 days; proxy, not actual Q3 cash'],
 [34,'Net financing since cash report, EUR m',0,num,'Assumed zero pending dated financing reconciliation'],
 [35,'Debt excluding operating leases, EUR m',0,num,'Skeleton assumes no funded debt; operating occupancy embedded in costs'],
 [36,'Other assets rNPV, EUR millions',0,num,'HAE-only scope. Explicit zero for AAE/other pipeline, not proven no value'],
 [37,'Estimated cash at valuation, EUR m',null,num,'Derived from dated cash less estimated elapsed burn plus financing'],
 [38,'Source basis confirmation: 0 no / 1 yes',0,'0','Keep0 until source forecast probability basis is verified']
];
for(const [r,label,v,fmt,n] of entries){val(a,`C${r}`,label);if(v!==null)input(a,`E${r}`,v,fmt);note(a,`G${r}`,n);}
formula(a,'E6','=IF(ISNUMBER(CHOOSE(E5,E9,E10,E11)),CHOOSE(E5,E9,E10,E11),NA())');a.getRange('E6').setNumberFormat(price);
formula(a,'E7','=CHOOSE(E5,0,1,E12)');
formula(a,'E37','=E31-MAX(0,E8-E32)*E33+E34');
for(const r of [5,14])a.getRange(`E${r}`).dataValidation={rule:{type:'whole',operator:'between',formula1:1,formula2:3}};
a.getRange('E15').dataValidation={rule:{type:'whole',operator:'between',formula1:1,formula2:2}};
for(const r of [22,38])a.getRange(`E${r}`).dataValidation={rule:{type:'whole',operator:'between',formula1:0,formula2:1}};
for(const r of [16,17,18,19,23,25,26])a.getRange(`E${r}`).dataValidation={rule:{type:'decimal',operator:'between',formula1:0,formula2:1}};
band(a,41,'Annual costs and reported revenue path (EUR millions)');setrow(a,42,'Calendar year',years);
linkrow(a,43,'Reported IR revenue',years.map((_,i)=>`='Street evidence'!${col(5+i)}25`));
linkrow(a,44,'Reported XR revenue',years.map((_,i)=>`='Street evidence'!${col(5+i)}26`));
const rd=[138.343,139.726,141.124,142.535,135.408,135.408,135.408,135.408,135.408,135.408,135.408];
const sg=[63.394,69.734,73.220,73.952,74.692,74.692,74.692,74.692,74.692,74.692,74.692];
setrow(a,46,'R&D',rd);setrow(a,47,'G&A / commercial overhead',sg);
a.getRange('E46:O47').format.font.color=blue;a.getRange('E46:O47').format.fill=amber;
note(a,'C49','2026–2030 costs: Guggenheim p2. 2031–2036: assumed flat at2030 levels; replace when fuller model is available.');
note(a,'C50','Shared costs deducted once and not PoS-weighted. No failure-state savings, financing fees or future share issuance modeled.');
note(a,'C51','PoS weights expected revenue. Success-case revenue is the appropriate comparator to approved-drug actual sales.');
note(a,'C52','A single scale preserves the chosen ramp shape. It does not independently identify launch speed, peak sales or product mix.');
a.freezePanes.freezeRows(4);

const d=sheets.DCF;
note(d,'C4','EUR millions. Year-end discounting. 2026 costs prorated from valuation date; elapsed cash burn is reflected in cash once.');
setrow(d,5,'Calendar year',years);d.freezePanes.freezeRows(5);d.freezePanes.freezeColumns(4);
const lab={6:'Years to cash flow',7:'Year fraction included',10:'Reported IR reference revenue',11:'Reported XR reference revenue',12:'Success-case IR reference revenue',13:'Success-case XR reference revenue',15:'COGS / revenue',16:'Cash tax rate',18:'R&D cash expense',19:'G&A cash expense',20:'Capex',21:'Depreciation',22:'Shared operating cost',30:'IR expected revenue per1x ramp',31:'XR expected revenue per1x ramp',32:'IR contribution before shared costs',33:'XR contribution before shared costs',34:'IR change in working capital per1x',35:'XR change in working capital per1x',36:'IR after-tax cash-flow coefficient',37:'XR after-tax cash-flow coefficient',38:'Fixed-cost cash-flow coefficient',39:'Discount factor',40:'PV IR coefficient',41:'PV XR coefficient',42:'PV fixed coefficient'};
for(const [r,l] of Object.entries(lab))val(d,`C${r}`,l);
years.forEach((y,i)=>{const c=col(i+5),prev=col(i+4);const f={
6:`=MAX(0,(DATE(${c}5,12,31)-Inputs!$E$8)/365.25)`,
7:`=IF(${c}5=YEAR(Inputs!$E$8),MAX(0,(DATE(${c}5,12,31)-Inputs!$E$8)/365.25),IF(${c}5<YEAR(Inputs!$E$8),0,1))`,
10:`=Inputs!${c}43`,11:`=Inputs!${c}44`,
12:`=IF(Inputs!$E$15=1,${c}10,${c}10/Inputs!$E$18)`,13:`=IF(Inputs!$E$15=1,${c}11,${c}11/Inputs!$E$19)`,
15:'=Inputs!$E$25',16:`=IF(${c}5>=Inputs!$E$24,Inputs!$E$23,0)`,
18:`=Inputs!${c}46*${c}7`,19:`=Inputs!${c}47*${c}7`,20:`=Inputs!$E$27*${c}7`,21:`=Inputs!$E$28*${c}7`,22:`=SUM(${c}18:${c}19)`,
30:`=${c}12*Inputs!$E$16`,31:`=${c}13*Inputs!$E$17`,32:`=${c}30*(1-${c}15)`,33:`=${c}31*(1-${c}15)`,
34:`=Inputs!$E$26*(${c}30-${i?prev+'30':'0'})`,35:`=Inputs!$E$26*(${c}31-${i?prev+'31':'0'})`,
36:`=${c}32*(1-${c}16)-${c}34`,37:`=${c}33*(1-${c}16)-${c}35`,38:`=-${c}22*(1-${c}16)+${c}21*${c}16-${c}20`,
39:`=1/(1+Inputs!$E$20)^${c}6`,40:`=${c}36*${c}39`,41:`=${c}37*${c}39`,42:`=${c}38*${c}39`};
for(const[r,x]of Object.entries(f))formula(d,`${c}${r}`,x);
});
d.getRange('E15:O16').setNumberFormat(pct);d.getRange('E39:O39').setNumberFormat('0.000x');
band(d,45,'Reverse valuation solve');
const solver={
48:['IR rNPV coefficient','=SUM(E40:O40)+IF(Inputs!$E$22=1,O36*(1+Inputs!$E$21)/(Inputs!$E$20-Inputs!$E$21)*O39,0)'],
49:['XR rNPV coefficient','=SUM(E41:O41)+IF(Inputs!$E$22=1,O37*(1+Inputs!$E$21)/(Inputs!$E$20-Inputs!$E$21)*O39,0)'],
50:['PV of shared costs','=SUM(E42:O42)+IF(Inputs!$E$22=1,O38*(1+Inputs!$E$21)/(Inputs!$E$20-Inputs!$E$21)*O39,0)'],
51:['Target equity PV, EUR','=Inputs!E6*Inputs!E30/Inputs!E29/(1+Inputs!E20)^Inputs!E7'],
52:['Estimated cash at valuation','=Inputs!E37'],53:['Other assets rNPV','=Inputs!E36'],54:['Debt','=Inputs!E35'],
55:['Required product contribution PV','=E51-E52-E53+E54-E50'],
56:['Common revenue scale','=IF(E48+E49>0,E55/(E48+E49),NA())'],
57:['Solved IR revenue scale','=CHOOSE(Inputs!E14,E56,1,IF(E48>0,(E55-E49)/E48,NA()))'],
58:['Solved XR revenue scale','=CHOOSE(Inputs!E14,E56,IF(E49>0,(E55-E48)/E49,NA()),1)'],
59:['Normalized 1x reference value, USD/share','=(SUM(E48:E50)+E52+E53-E54)*Inputs!E29/Inputs!E30*(1+Inputs!E20)^Inputs!E7']};
for(const[r,[l,f]]of Object.entries(solver)){val(d,`C${r}`,l);formula(d,`E${r}`,f);}
d.getRange('E56:E58').setNumberFormat('0.000x');d.getRange('E59').setNumberFormat(price);
note(d,'G56','Analytical solve holds PoS, shape and cost assumptions fixed. Final cash-flow build below tests the tax domain.');
note(d,'G57','Mode2 holds IR at1x; mode3 holds XR at1x. A negative scale is not an admissible sales forecast.');
band(d,62,'Solved cash flow and valuation');
const finalLabels={64:'IR success-case revenue',65:'XR success-case revenue',66:'IR probability-weighted revenue',67:'XR probability-weighted revenue',68:'Total probability-weighted revenue',69:'COGS',70:'R&D',71:'G&A / commercial overhead',72:'Depreciation',73:'EBIT',74:'Cash taxes (no refunds)',75:'NOPAT',76:'Add depreciation',77:'Less capex',78:'Working capital balance',79:'Change in working capital',80:'Unlevered free cash flow',81:'Discount factor',82:'PV free cash flow'};
for(const[r,l]of Object.entries(finalLabels))val(d,`C${r}`,l);
years.forEach((_,i)=>{const c=col(i+5),p=col(i+4);const f={64:`=${c}12*$E$57`,65:`=${c}13*$E$58`,66:`=${c}64*Inputs!$E$16`,67:`=${c}65*Inputs!$E$17`,68:`=SUM(${c}66:${c}67)`,69:`=${c}68*${c}15`,70:`=${c}18`,71:`=${c}19`,72:`=${c}21`,73:`=${c}68-SUM(${c}69:${c}72)`,74:`=MAX(0,${c}73)*${c}16`,75:`=${c}73-${c}74`,76:`=${c}72`,77:`=${c}20`,78:`=${c}68*Inputs!$E$26`,79:`=${c}78-${i?p+'78':'0'}`,80:`=${c}75+${c}76-${c}77-${c}79`,81:`=${c}39`,82:`=${c}80*${c}81`};for(const[r,x]of Object.entries(f))formula(d,`${c}${r}`,x);});
const result={86:['PV explicit cash flows','=SUM(E82:O82)'],87:['Terminal value at2036','=IF(Inputs!E22=1,O80*(1+Inputs!E21)/(Inputs!E20-Inputs!E21),0)'],88:['PV terminal value','=E87*O81'],89:['Enterprise value incl other assets','=SUM(E86,E88,E53)'],90:['Equity value, EUR','=E89+E52-E54'],91:['Present value per share, USD','=E90*Inputs!E29/Inputs!E30'],92:['Value at selected target horizon, USD','=E91*(1+Inputs!E20)^Inputs!E7'],93:['Reconciliation to selected price','=E92-Inputs!E6']};
for(const[r,[l,f]]of Object.entries(result)){val(d,`C${r}`,l);formula(d,`E${r}`,f);}d.getRange('E91:E93').setNumberFormat(price);
band(d,96,'Model checks');
val(d,'C97','Nonnegative revenue scales');formula(d,'E97','=IF(MIN(E57:E58)>=0,"PASS","FAIL: negative ramp")');
val(d,'C98','Reprices selected target');formula(d,'E98','=IF(ABS(E93)<0.000001,"PASS","FAIL: tax / solve domain")');
val(d,'C99','Terminal discount condition');formula(d,'E99','=IF(OR(Inputs!E22=0,Inputs!E20>Inputs!E21),"PASS","FAIL: discount <= growth")');
val(d,'C100','Probability basis confirmed');formula(d,'E100','=IF(Inputs!E38=1,"Confirmed by user","UNCONFIRMED")');
note(d,'C102','Solver assumes taxable-year profits. If losses occur after tax start, no-refund taxes break the affine solve; failed reconciliation is visible.');
note(d,'C103','No NOL roll-forward. No future financing or failure-state operating savings. Correlated approval risks are not assumed independent.');
for(const r of [97,98,99])d.getRange(`E${r}`).conditionalFormats.add('containsText',{text:'FAIL',format:{fill:'#FEE4E2',font:{color:'#B42318'}}});

const s=sheets['Reverse SOTP'];
note(s,'C4','Conditional sales requirements at fixed PoS. Guggenheim-shaped reference with normalized cash flows; not recovered buy-side consensus.');
val(s,'C6','Selected target');formula(s,'E6','=Inputs!E6');s.getRange('E6').setNumberFormat(price);
val(s,'G6','Target basis');formula(s,'I6','=CHOOSE(Inputs!E5,"Spot snapshot","Broker PT","Custom")');s.getRange('I6').format.font.color=green;
val(s,'C7','Target horizon, years');formula(s,'E7','=Inputs!E7');
val(s,'C8','IR / XR PoS');formula(s,'E8','=Inputs!E16');formula(s,'F8','=Inputs!E17');s.getRange('E8:F8').setNumberFormat(pct);
val(s,'C9','IR / XR revenue scale');formula(s,'E9','=DCF!E57');formula(s,'F9','=DCF!E58');s.getRange('E9:F9').setNumberFormat('0.000x');
val(s,'C10','Source probability basis');formula(s,'E10','=IF(Inputs!E15=1,"Assume success-case","Assume risk-adjusted")');
val(s,'C11','Model result status');formula(s,'E11','=IF(OR(DCF!E97<>"PASS",DCF!E98<>"PASS",DCF!E99<>"PASS"),"INVALID: inspect DCF checks",IF(Inputs!E38=0,"Conditional: basis unconfirmed","Conditional on chosen assumptions"))');
note(s,'C13','Change Inputs E5 for spot versus PT, E14 for which product scales, and E15 for source revenue probability basis.');
band(s,15,'SOTP at the selected target horizon','J');
s.getRange('C16:F16').values=[['Component',null,'EUR m PV','USD / share']];
const components=[['IR risk-adjusted contribution','=DCF!E48*DCF!E57'],['XR risk-adjusted contribution','=DCF!E49*DCF!E58'],['Shared costs and investment','=DCF!E50'],['Estimated cash','=DCF!E52'],['Other assets less debt','=DCF!E53-DCF!E54']];
components.forEach(([l,f],i)=>{const r=17+i;val(s,`C${r}`,l);formula(s,`E${r}`,f);formula(s,`F${r}`,`=E${r}*Inputs!$E$29/Inputs!$E$30*(1+Inputs!$E$20)^Inputs!$E$7`);});
val(s,'C22','Equity value');formula(s,'E22','=SUM(E17:E21)');formula(s,'F22','=SUM(F17:F21)');s.getRange('F17:F22').setNumberFormat(price);
val(s,'C24','Normalized 1x reference value');formula(s,'E24','=DCF!E59');s.getRange('E24').setNumberFormat(price);
val(s,'C25','Published Guggenheim PT');formula(s,'E25','=Inputs!E10');s.getRange('E25').setNumberFormat(price);
note(s,'C26','The source PT is not forced to match: costs, timing, taxes and probability basis differ.');
band(s,28,'Required success-case sales (USD millions)','J');
s.getRange('C29:H29').values=[['Calendar year',null,'IR','XR','Combined','Risk-weighted combined']];
for(let i=0;i<10;i++){let r=30+i,c=col(6+i);val(s,`C${r}`,2027+i);formula(s,`E${r}`,`=DCF!${c}64*Inputs!$E$29`);formula(s,`F${r}`,`=DCF!${c}65*Inputs!$E$29`);formula(s,`G${r}`,`=SUM(E${r}:F${r})`);formula(s,`H${r}`,`=DCF!${c}68*Inputs!$E$29`);}
note(s,'C41','2036 is the last explicit year, not an independently established peak year. All sales are worldwide net revenue.');
note(s,'C42','Approved-drug launch comparisons use success-case sales. Discounted / PoS-weighted revenue is not actual commercial revenue.');
const chart=s.charts.add('line',s.getRange('C29:C39').offset(0,0).resize(11,1));
chart.setData([s.getRange('C29:C39'),s.getRange('E29:E39'),s.getRange('F29:F39')]);
chart.title='Required worldwide sales ($m)';chart.setPosition('K6','S24');chart.titleTextStyle.typeface='Arial';chart.titleTextStyle.fontSize=12;
chart.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};chart.yAxis={numberFormatCode:'$#,##0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};
chart.legend={position:'bottom',textStyle:{typeface:'Arial',fontSize:10}};
chart.series.items.forEach((x,i)=>{x.line={fill:[navy,'#A36C28'][i],style:'solid',width:2};});

const l=sheets['Launch comparison'];
note(l,'C4','Reference chronology and currencies are explicit. These comparisons do not measure a causal efficacy premium or a universal second-oral discount.');
band(l,6,'First full calendar years after launch (worldwide USD millions)','K');
l.getRange('C7:H7').values=[['Full calendar year',null,'Orladeyo actual','Solved XR','Source-shaped XR','XR / Orladeyo']];
for(let i=0;i<4;i++){const r=8+i,c=col(8+i);val(l,`C${r}`,i+1);formula(l,`E${r}`,`='Street evidence'!E${52+i}`);formula(l,`F${r}`,`=DCF!${c}65*Inputs!$E$29`);formula(l,`G${r}`,`=DCF!${c}13*Inputs!$E$29`);formula(l,`H${r}`,`=F${r}/E${r}`);}
l.getRange('H8:H11').setNumberFormat('0.00x');
note(l,'C13','Orladeyo launched Dec2020: actuals2021–2024. XR reference first sales2028: comparison2029–2032; within-year launch timing unknown.');
note(l,'C14','Different launch months, prices, geographic rollout and competitive markets limit like-for-like inference.');
band(l,17,'Early acute launch comparison','K');
val(l,'C19','Ekterly Jul7–Dec31 2025, global');formula(l,'E19',"='Street evidence'!E57");
val(l,'C20','Solved IR full calendar2027');formula(l,'E20','=DCF!F64*Inputs!E29');
val(l,'C21','Guggenheim IR2027 success-case basis');formula(l,'E21','=DCF!F12*Inputs!E29');
val(l,'C22','Wedbush IR2027 forecast, USD');formula(l,'E22',"='Street evidence'!E62");
note(l,'C24','Ekterly observation is ~6 commercial months. PHVS2027 is a calendar forecast with uncertain launch month; no precise ramp ratio is asserted.');
note(l,'C25','Wedbush starts IR sales in Q3 2027; its supplied note does not disclose whether this revenue is probability adjusted.');
band(l,28,'Additional interpretation','K');
note(l,'C30','Revenue = paid treatment volume x realized net price. A stronger sales curve can reflect pricing or persistence rather than faster switching.');
note(l,'C31','Wedbush raises XR modeled net annual price to ~$418k; Morgan Stanley raises IR net price/dose to ~$16k. These are analyst assumptions.');
note(l,'C32','IR and XR capture are economically linked. This valuation skeleton preserves a selected path; a later adoption model must reconcile rescue use.');
const lc=l.charts.add('line',[l.getRange('C7:C11'),l.getRange('E7:E11'),l.getRange('F7:F11')]);
lc.title='XR requirement versus Orladeyo ($m)';lc.setPosition('J35','R52');lc.titleTextStyle.typeface='Arial';lc.titleTextStyle.fontSize=12;
lc.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:10}};lc.yAxis={numberFormatCode:'$#,##0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:10}};lc.legend={position:'bottom',textStyle:{typeface:'Arial',fontSize:10}};
lc.series.items.forEach((x,i)=>{x.line={fill:['#8B9099','#A36C28'][i],style:i?'solid':'dashed',width:2};});

// Output views keep formula results dark. Finance colors apply to the working tabs.
for(const n of ['Reverse SOTP','Launch comparison'])sheets[n].getRange('E6:O43').format.font.color='#172B4D';
s.getRange('I6').format.font.color=green;
s.getRange('E11').conditionalFormats.add('containsText',{text:'INVALID',format:{fill:'#FEE4E2',font:{color:'#B42318',bold:true}}});
for(const sh of [s,a,d,l,e]){
 for(const r of [2,4])sh.getRange(`C${r}`).format.wrapText=false;
}
wb.recalculate();
const snapshot=()=>({target:a.getRange('E6').values[0][0],irScale:d.getRange('E57').values[0][0],xrScale:d.getRange('E58').values[0][0],price:d.getRange('E92').values[0][0],gap:d.getRange('E93').values[0][0],baseline:d.getRange('E59').values[0][0],sales:s.getRange('E30:H39').values});
const base=snapshot();
const tests=[];
for(const [cell,value,desc] of [['E5',2,'Broker target'],['E14',2,'XR-only solve'],['E15',2,'Already weighted source basis'],['E22',0,'No terminal value']]){
 const old=a.getRange(cell).values[0][0];val(a,cell,value);wb.recalculate();const x=snapshot();tests.push({desc,...x});if(Math.abs(x.gap)>1e-6)throw new Error('Repricing failed '+desc);val(a,cell,old);wb.recalculate();
}
const old=a.getRange('E11').values[0][0];val(a,'E5',3);val(a,'E11',0);wb.recalculate();const invalid=snapshot();if(d.getRange('E97').values[0][0]==='PASS'&&d.getRange('E98').values[0][0]==='PASS')throw new Error('Zero-target boundary was not flagged');val(a,'E11',old);val(a,'E5',1);wb.recalculate();
if(Math.abs(base.gap)>1e-6)throw new Error('Base repricing failed');
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},summary:'Formula error scan'});
await fs.writeFile(qa+'/formula_scan.json',errors.ndjson);
await fs.writeFile(root+'/models/reverse_dcf/verification.json',JSON.stringify({base,tests,invalid,engine:'Artifact Tool; native Excel not exercised'},null,2));
const regions={'Reverse SOTP':'C2:S42','Inputs':'C2:O38','DCF':'C45:O100','Launch comparison':'C2:R52','Street evidence':'C6:K18'};
for(const[n,range]of Object.entries(regions)){const blob=await wb.render({sheetName:n,range,scale:1.5,format:'png'});await fs.writeFile(qa+'/'+n.replaceAll(' ','_')+'.png',new Uint8Array(await blob.arrayBuffer()));}
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(out+'/PHVS_reverse_SOTP_DCF.xlsx');
console.log(JSON.stringify({base,tests:tests.map(t=>({desc:t.desc,price:t.price,irScale:t.irScale,xrScale:t.xrScale})),formulaScan:errors.ndjson,output:out+'/PHVS_reverse_SOTP_DCF.xlsx'},null,2));
