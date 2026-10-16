'use strict';

const MONTHS=['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const SHORT=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const COLORS=['#004B13','#62BAB6','#1D6DB6','#EB4A75','#D9A441','#233744','#7FA6C9','#A0522D'];
const FOCUS='#004B13',COMPARE='#62BAB6';
const charts={};
let state=null;

const normalize=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
const asNumber=value=>typeof value==='number'&&Number.isFinite(value)?value:null;
const currency=value=>value==null?'—':new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:0}).format(value);
const percent=value=>value==null?'—':new Intl.NumberFormat('pt-BR',{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1}).format(value);
const decimal=value=>value==null?'—':new Intl.NumberFormat('pt-BR',{maximumFractionDigits:1}).format(value);

function findSheet(book,fragment){const name=book.SheetNames.find(item=>normalize(item).includes(fragment));return name?book.Sheets[name]:null;}
function sheetRows(sheet){return sheet?XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,blankrows:false}):[];}
function headerIndex(rows,key){return rows.findIndex(row=>row.some(value=>normalize(value)===key));}

function parseAnnual(book){
  const rows=sheetRows(findSheet(book,'indices anuais'));const index=headerIndex(rows,'ano');
  if(index<0)throw new Error('A aba “Índices anuais” não possui uma linha de cabeçalho com Ano.');
  return rows.slice(index+1).filter(row=>Number.isInteger(row[0])).map(row=>({year:row[0],daily:asNumber(row[1]),advance:asNumber(row[2]),stay:asNumber(row[3]),cancel:asNumber(row[4])}));
}

function parseMonthly(book,fragment){
  const rows=sheetRows(findSheet(book,fragment));const index=headerIndex(rows,'mes');if(index<0)return{};
  const headers=rows[index];const series={};
  headers.forEach((header,column)=>{const match=String(header??'').match(/\b(20\d{2})\b/);if(!match)return;const year=Number(match[1]);series[year]={year,estimated:normalize(header).includes('estimado'),values:Array(12).fill(null)};for(const row of rows.slice(index+1)){const month=MONTHS.findIndex(item=>normalize(item)===normalize(row[0]));if(month>=0)series[year].values[month]=asNumber(row[column]);}});
  return series;
}

function parseGroups(book){
  const rows=sheetRows(findSheet(book,'origem e tipo'));const index=headerIndex(rows,'ano');if(index<0)return[];
  return rows.slice(index+1).filter(row=>Number.isInteger(row[0])&&asNumber(row[3])!=null).map(row=>({year:row[0],group:normalize(row[1]),category:String(row[2]??''),value:row[3]}));
}

function parseWorkbook(book){const annual=parseAnnual(book);const bookings=parseMonthly(book,'media de reservas');Object.values(bookings).forEach(series=>{series.estimated=true;});return{annual,daily:parseMonthly(book,'media de diarias'),bookings,occupancy:parseMonthly(book,'media de taxa de ocupacao'),groups:parseGroups(book),years:annual.map(item=>item.year).sort((a,b)=>a-b)};}

function selectedYears(){return{focus:Number(document.querySelector('#focus-year').value),compare:Number(document.querySelector('#compare-year').value),estimated:document.querySelector('#show-estimates').checked,period:document.querySelector('#period').value};}
function monthIndexes(series,focus,compare,period){let indexes=MONTHS.map((_,i)=>i);if(period==='first')indexes=indexes.filter(i=>i<6);if(period==='second')indexes=indexes.filter(i=>i>=6);if(period==='available'&&series[focus])indexes=indexes.filter(i=>series[focus].values[i]!=null&&(!series[compare]||series[compare].values[i]!=null));return indexes;}
function average(values){const valid=values.filter(value=>value!=null);return valid.length?valid.reduce((sum,value)=>sum+value,0)/valid.length:null;}
function averageSeries(series,year,indexes){return series[year]?average(indexes.map(i=>series[year].values[i])):null;}
function variation(current,previous){return current!=null&&previous!=null&&previous!==0?current/previous-1:null;}

function initializeFilters(){const focus=document.querySelector('#focus-year'),compare=document.querySelector('#compare-year');for(const year of [...state.years].reverse()){for(const select of [focus,compare]){const option=document.createElement('option');option.value=year;option.textContent=year;select.append(option);}}focus.value=state.years.at(-1);compare.value=state.years.at(-2)??state.years.at(-1);for(const id of ['focus-year','compare-year','period','show-estimates'])document.getElementById(id).addEventListener('change',render);}

function createChart(id,config){if(charts[id])charts[id].destroy();charts[id]=new Chart(document.getElementById(id),config);}
function lineOptions(formatter){return{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:7,padding:16,font:{size:10}}},tooltip:{backgroundColor:'#233744',padding:10,cornerRadius:8,titleFont:{weight:'700'},callbacks:{label:context=>`${context.dataset.label}: ${formatter(context.parsed.y)}`}}},scales:{x:{grid:{display:false},ticks:{font:{size:10},color:'#53778B'}},y:{grid:{color:'#EEF2F3'},border:{display:false},ticks:{font:{size:10},color:'#53778B',callback:formatter}}}};}

function lineDatasets(series,focus,compare,estimated,indexes){
  const years=[compare,focus].filter((year,i,arr)=>arr.indexOf(year)===i&&series[year]&&(!series[year].estimated||estimated));
  return years.map((year,index)=>({label:`${year}${series[year].estimated?' · estimado':''}`,data:indexes.map(i=>series[year].values[i]),borderColor:index===years.length-1?FOCUS:COMPARE,backgroundColor:index===years.length-1?FOCUS:COMPARE,pointBackgroundColor:index===years.length-1?FOCUS:COMPARE,borderWidth:index===years.length-1?3:2,tension:.35,pointRadius:3,pointHoverRadius:5,borderDash:series[year].estimated?[6,5]:[],spanGaps:false}));
}

function kpiCard(label,value,delta,comparison){const deltaText=delta==null?'Sem comparação disponível':`${delta>0?'+':''}${percent(delta)} vs. ${comparison}`;return`<article class="panel kpi"><span class="label">${label}</span><strong class="value">${value}</strong><span class="delta ${delta>0?'up':delta<0?'down':''}">${deltaText}</span></article>`;}

function renderKpis(focus,compare,period){const a=state.annual.find(item=>item.year===focus)??{},b=state.annual.find(item=>item.year===compare)??{};const indexes=monthIndexes(state.occupancy,focus,compare,period);const occ=averageSeries(state.occupancy,focus,indexes),previousOcc=averageSeries(state.occupancy,compare,indexes);document.querySelector('#kpis').innerHTML=[kpiCard('Diária média anual · Gramado',currency(a.daily),variation(a.daily,b.daily),compare),kpiCard('Ocupação · Gramado e Canela',percent(occ),variation(occ,previousOcc),compare),kpiCard('Antecedência · Gramado',`${decimal(a.advance)} dias`,variation(a.advance,b.advance),compare),kpiCard('Estadia média · Gramado',`${decimal(a.stay)} dias`,variation(a.stay,b.stay),compare),kpiCard('Cancelamento · Gramado',percent(a.cancel),variation(a.cancel,b.cancel),compare)].join('');}

function renderMonthlyCharts(focus,compare,estimated,period){for(const [id,series,type] of [['daily-chart',state.daily,'currency'],['occupancy-chart',state.occupancy,'percent'],['booking-chart',state.bookings,'number']]){const indexes=monthIndexes(series,focus,compare,period);const format=type==='currency'?value=>currency(value):type==='percent'?value=>percent(value):value=>decimal(value);const options=lineOptions(format);if(id==='booking-chart'){options.plugins.tooltip.enabled=false;options.scales.y.ticks.display=false;options.scales.y.grid.display=false;}createChart(id,{type:'line',data:{labels:indexes.map(i=>SHORT[i]),datasets:lineDatasets(series,focus,compare,estimated,indexes)},options});}}

function renderAnnualCharts(focus){const years=state.annual.map(item=>String(item.year));for(const [id,key,type] of [['annual-chart','daily','currency'],['advance-chart','advance','number'],['stay-chart','stay','number'],['cancel-chart','cancel','percent']]){const formatter=type==='currency'?currency:type==='percent'?percent:decimal;const data=state.annual.map(item=>item[key]);createChart(id,{type:id==='annual-chart'?'bar':'line',data:{labels:years,datasets:[{label:type==='percent'?'Cancelamento':key==='stay'?'Dias':key==='advance'?'Dias':'Diária média',data,backgroundColor:state.annual.map(item=>item.year===focus?FOCUS:'rgba(98,186,182,.6)'),pointBackgroundColor:state.annual.map(item=>item.year===focus?FOCUS:COMPARE),borderColor:id==='annual-chart'?state.annual.map(item=>item.year===focus?FOCUS:COMPARE):FOCUS,borderWidth:2,borderRadius:5,tension:.32,pointRadius:3,fill:false}]},options:{...lineOptions(formatter),plugins:{...lineOptions(formatter).plugins,legend:{display:false}}}});}}

function latestGroup(group,focus){const all=state.groups.filter(item=>item.group===group);if(!all.length)return[];const years=[...new Set(all.map(item=>item.year))].sort((a,b)=>b-a);const year=years.includes(focus)?focus:years[0];return all.filter(item=>item.year===year);}

function renderProfiles(focus){for(const [id,group,caption] of [['origin-chart','origem','origin-caption'],['payment-chart','pagamento','payment-caption']]){const values=latestGroup(group,focus);document.getElementById(caption).textContent=values.length?`Participação em ${values[0].year}${values[0].year!==focus?' · último ano disponível':''}.`:'Sem dados disponíveis.';createChart(id,{type:'doughnut',data:{labels:values.map(item=>item.category),datasets:[{data:values.map(item=>item.value),backgroundColor:COLORS,borderWidth:0,hoverOffset:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'68%',plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:7,padding:12,font:{size:9}}},tooltip:{callbacks:{label:context=>`${context.label}: ${percent(context.parsed)}`}}}}});}}

function renderInsights(focus,compare,period){const current=state.annual.find(item=>item.year===focus)??{},prior=state.annual.find(item=>item.year===compare)??{};const indexes=monthIndexes(state.occupancy,focus,compare,period),occ=averageSeries(state.occupancy,focus,indexes),oldOcc=averageSeries(state.occupancy,compare,indexes);const messages=[];if(current.daily!=null&&prior.daily!=null){const v=variation(current.daily,prior.daily);messages.push(`A diária média anual ${v>=0?'subiu':'caiu'} <strong>${percent(Math.abs(v))}</strong> em relação a ${compare}.`);}if(occ!=null&&oldOcc!=null){const difference=occ-oldOcc;messages.push(`A ocupação média dos meses comparáveis ${difference>=0?'avançou':'recuou'} <strong>${decimal(Math.abs(difference)*100)} p.p.</strong>`);}if(current.stay!=null&&prior.stay!=null)messages.push(`A permanência média passou de <strong>${decimal(prior.stay)}</strong> para <strong>${decimal(current.stay)} dias</strong>.`);if(current.cancel!=null)messages.push(`O cancelamento registrado para ${focus} é de <strong>${percent(current.cancel)}</strong>.`);const missing=MONTHS.length-(state.occupancy[focus]?.values.filter(v=>v!=null).length??0);if(missing>0)messages.push(`<strong>${focus} é parcial:</strong> faltam ${missing} meses na série de ocupação.`);document.querySelector('#insights').innerHTML=messages.map(message=>`<p class="insight">${message}</p>`).join('');}

function render(){const{focus,compare,estimated,period}=selectedYears();renderKpis(focus,compare,period);renderMonthlyCharts(focus,compare,estimated,period);renderAnnualCharts(focus);renderProfiles(focus);renderInsights(focus,compare,period);}

async function boot(){try{if(typeof XLSX==='undefined'||typeof Chart==='undefined')throw new Error('Não foi possível carregar as bibliotecas do dashboard. Verifique a conexão com a internet.');const response=await fetch(`dados.xlsx?v=${Date.now()}`,{cache:'no-store'});if(!response.ok)throw new Error(`Não foi possível abrir dados.xlsx (${response.status}). Coloque a planilha na mesma pasta de index.html.`);const book=XLSX.read(await response.arrayBuffer(),{type:'array',cellDates:true});state=parseWorkbook(book);if(!state.years.length)throw new Error('A planilha não contém anos válidos na aba de índices anuais.');document.querySelector('#updated-at').textContent=`${state.years[0]}–${state.years.at(-1)} · ${book.SheetNames.length} abas`;Chart.defaults.font.family="'Stolzl', system-ui, -apple-system, 'Segoe UI', sans-serif";Chart.defaults.color='#53778B';initializeFilters();render();}catch(error){console.error(error);const box=document.getElementById('error-message');box.hidden=false;box.textContent=error.message;document.querySelector('#updated-at').textContent='Erro ao carregar a base';}}

if(typeof document!=='undefined')document.addEventListener('DOMContentLoaded',boot);
if(typeof module!=='undefined')module.exports={parseWorkbook,normalize,MONTHS};
