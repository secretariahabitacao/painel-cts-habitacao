/* ============================================================
   PAINEL CTS — JS COMPARTILHADO v3.0
   Utilitários, parsers, helpers de gráficos
   ============================================================ */

const num = v => {
  if(v===null||v===undefined||v===''||v==='-') return 0;
  if(typeof v==='number') return isNaN(v)?0:v;
  const n=parseFloat(String(v).replace(/[^\d.,-]/g,'').replace(',','.'));
  return isNaN(n)?0:n;
};
const fmt = n => typeof n==='number'?n.toLocaleString('pt-BR'):String(n||'—');
const pct = (v,d=0) => {
  if(!v&&v!==0) return '—';
  const n=typeof v==='number'?v:parseFloat(v);
  if(isNaN(n)) return '—';
  return (n*(n<=1.5?100:1)).toFixed(d)+'%';
};

const normMod = v => {if(!v)return '';const s=String(v).trim();if(/^Aquisi/i.test(s))return'Aquisições';if(/^Empresa/i.test(s))return'Empresas';return s;};
const modToPrograma = m => {if(!m)return '';const s=m.toUpperCase();if(s==='PPP')return'PPP';if(s.includes('PRÓPRIA')||s.includes('PROPRIA'))return'Produção Própria';if(s==='OUCAE')return'OUCAE';return'PPE';};

function parseDate(v){
  if(!v)return null;if(v instanceof Date)return v;
  if(typeof v==='number'&&v>30000)return new Date((v-25569)*86400*1000);
  const s=String(v).trim();const m=s.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if(m)return new Date(+m[3],+m[2]-1,+m[1]);return null;
}

function fmtDateShort(v){
  if(!v)return'—';const s=String(v).trim();
  if(s==='#REF!'||s==='#N/A'||s==='-'||s==='0')return'—';
  if(/Realizada/i.test(s))return'✓ '+s.replace(/Realizada\s*em\s*/i,'');
  if(/Programada/i.test(s))return'◷ '+s.replace(/Programada\s*p\/\s*/i,'');
  if(typeof v==='number'&&v>30000&&v<60000){const d=new Date((v-25569)*86400*1000);if(!isNaN(d.getTime())){const m=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];return m[d.getMonth()]+'/'+d.getFullYear();}}
  if(v instanceof Date&&!isNaN(v.getTime())){const m=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];return v.getDate().toString().padStart(2,'0')+'/'+m[v.getMonth()]+'/'+v.getFullYear();}
  return s;
}

const COLORS={
  teal:'#009999',green:'#1B7A3D',gold:'#C9A227',red:'#CC3333',
  blue:'#2B6CB0',purple:'#6B46C1',orange:'#C05621',grey:'#888',
  palette:['#009999','#1B7A3D','#C9A227','#2B6CB0','#6B46C1','#C05621','#CC3333','#5BADA7','#888']
};
const evColors={'ENCHENTE':COLORS.teal,'INCÊNDIO':COLORS.red,'DESABAMENTO':COLORS.orange,'QUEDA DE ÁRVORE':COLORS.green,'DESLIZAMENTO':COLORS.gold,'SOLAPAMENTO':COLORS.purple,'DESTELHAMENTO':COLORS.blue,'RISCO':COLORS.grey};

function destroyCharts(arr){if(arr)arr.forEach(c=>c.destroy());return[];}

async function fetchXlsx(url){
  const r=await fetch(url);if(!r.ok)throw new Error('HTTP '+r.status);
  const buf=await r.arrayBuffer();if(buf.byteLength<500)throw new Error('Arquivo muito pequeno');
  return XLSX.read(new Uint8Array(buf),{type:'array',cellDates:true});
}

const V=Date.now(); // anti-cache global

/* --- PARSE PPE --- */
function parsePPE(wb){
  const emps=[];let quantTotals={};
  const qName=wb.SheetNames.find(n=>n.toLowerCase().includes('quantitativo'));
  if(!qName)return{emps,quantTotals,aquiMap:{}};
  const ws=wb.Sheets[qName];
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null});
  let grpIdx=-1,detIdx=-1;
  for(let j=0;j<Math.min(10,raw.length);j++){const r=raw[j];if(!r)continue;const s0=String(r[0]||'').trim();if(s0==='Empreendimento'){if(r.some(c=>String(c||'')==='Endereço'||String(c||'')==='Construtora'||String(c||'')==='COHAB'))detIdx=j;else if(grpIdx<0)grpIdx=j;}}
  if(detIdx<0)detIdx=grpIdx>=0?grpIdx+2:3;
  const hRow=raw[detIdx];
  const findCol=name=>{for(let c=0;c<hRow.length;c++)if(String(hRow[c]||'').trim()===name)return c;return -1;};
  const cEmp=findCol('Empreendimento'),cMod=findCol('Modalidade'),cEnd=findCol('Endereço'),cConst=findCol('Construtora'),cContato=findCol('Contato'),cDts=findCol('DTS'),cTotal=findCol('Total');
  let cObra=-1;for(let c=0;c<hRow.length;c++){const s=String(hRow[c]||'').toLowerCase();if(s.includes('obra')||s.includes('evolu')){cObra=c;break;}}
  let cUhCohab=-1,cUhSehab=-1,cUhTotal=-1;
  for(let c=0;c<(cDts>=0?cDts:hRow.length);c++){if(String(hRow[c]||'')==='COHAB'&&cUhCohab<0)cUhCohab=c;if(String(hRow[c]||'')==='SEHAB'&&cUhSehab<0)cUhSehab=c;if(String(hRow[c]||'')==='TOTAL'&&cUhTotal<0)cUhTotal=c;}
  const statusPairs=[];
  if(cDts>=0){for(let c=cDts+1;c<hRow.length-1;c++){if(String(hRow[c]||'')==='COHAB'&&String(hRow[c+1]||'')==='SEHAB'){statusPairs.push(c);c++;}}}
  let cEscolha=-1,cAssemb=-1,cCnpj=-1,cEntrega=-1;
  if(cTotal>=0){for(let c=cTotal+1;c<hRow.length;c++){const s=String(hRow[c]||'').toLowerCase();if(s.includes('escolha'))cEscolha=c;else if(s.includes('assemb'))cAssemb=c;else if(s.includes('cnpj'))cCnpj=c;else if(s.includes('entrega'))cEntrega=c;}}
  const gc=(r,c)=>c>=0&&r.length>c?r[c]:null;
  for(let i=detIdx+1;i<raw.length;i++){const r=raw[i];if(!r)continue;const nome=String(gc(r,cEmp)||'').trim();if(!nome||nome==='TOTAL'||nome==='Total')continue;const mod=normMod(gc(r,cMod));
    const e={nome,modalidade:mod,programa:modToPrograma(mod),endereco:gc(r,cEnd)||'',construtora:gc(r,cConst)||'',contatos:gc(r,cContato)||'',pctObra:gc(r,cObra),uhCohab:num(gc(r,cUhCohab)),uhSehab:num(gc(r,cUhSehab)),uhTotal:num(gc(r,cUhTotal)),dts:gc(r,cDts)||'',
      termoCohab:num(gc(r,statusPairs[0])),termoSehab:num(gc(r,statusPairs[0]+1)),selecCohab:num(gc(r,statusPairs[1])),selecSehab:num(gc(r,statusPairs[1]+1)),dossieCohab:num(gc(r,statusPairs[2])),dossieSehab:num(gc(r,statusPairs[2]+1)),aprovCohab:num(gc(r,statusPairs[3])),aprovSehab:num(gc(r,statusPairs[3]+1)),contratoCohab:num(gc(r,statusPairs[4])),contratoSehab:num(gc(r,statusPairs[4]+1)),reprovCohab:num(gc(r,statusPairs[5])),reprovSehab:num(gc(r,statusPairs[5]+1)),ausentCohab:num(gc(r,statusPairs[6])),ausentSehab:num(gc(r,statusPairs[6]+1)),pendCohab:num(gc(r,statusPairs[7])),pendSehab:num(gc(r,statusPairs[7]+1)),desistCohab:num(gc(r,statusPairs[8])),desistSehab:num(gc(r,statusPairs[8]+1)),desclassCohab:num(gc(r,statusPairs[9])),desclassSehab:num(gc(r,statusPairs[9]+1)),demExcCohab:num(gc(r,statusPairs[10])),demExcSehab:num(gc(r,statusPairs[10]+1)),
      total:num(gc(r,cTotal)),dataEscolha:gc(r,cEscolha),dataAssembleia:gc(r,cAssemb),cnpjEmitido:gc(r,cCnpj),entrega:gc(r,cEntrega)};
    ['termo','selec','dossie','aprov','contrato','reprov','ausent','pend','desist','desclass','demExc'].forEach(k=>{e[k]=(e[k+'Cohab']||0)+(e[k+'Sehab']||0);});
    e.aprovTotal=e.aprov;e.pendTotal=e.pend;e.q=e;emps.push(e);}
  const totRow=raw.find(r=>r&&String(gc(r,cEmp)||'').trim()==='TOTAL');
  if(totRow){quantTotals={uhTotal:num(gc(totRow,cUhTotal))};['termo','selec','dossie','aprov','contrato','reprov','ausent','pend','desist','desclass','demExc'].forEach((k,i)=>{quantTotals[k]=statusPairs[i]!==undefined?num(gc(totRow,statusPairs[i]))+num(gc(totRow,statusPairs[i]+1)):0;});}
  if(!quantTotals.termo){quantTotals={termo:0,selec:0,dossie:0,aprov:0,contrato:0,reprov:0,ausent:0,pend:0,desist:0,desclass:0,demExc:0,uhTotal:0};emps.forEach(e=>{Object.keys(quantTotals).forEach(k=>{quantTotals[k]+=(e[k]||0);});});}
  const aquiMap={};const aName=wb.SheetNames.find(n=>n.toLowerCase().includes('aquiemp'));
  if(aName){const aRaw=XLSX.utils.sheet_to_json(wb.Sheets[aName],{header:1,defval:null});for(let i=1;i<aRaw.length;i++){const r=aRaw[i];if(!r||!r[2])continue;const emp=String(r[2]).trim().toUpperCase();if(!aquiMap[emp])aquiMap[emp]=[];aquiMap[emp].push({demanda:r[0]||'',nome:r[4]||'',empreendimento:r[2]||'',inscricao:r[3]||'',devolutiva:r[7]||'',usuario:r[8]||'',dataDevolutiva:r[9]||'',telefone:r[10]||'',celular:r[11]||'',email:r[12]||'',composicao:r[19]||'',contrato:r[22]||'',motivoDevolutiva:r[28]||'',statusDossie:r[29]||'',tipologia:r[30]||'',bloco:r[31]||'',andar:r[32]||'',apto:r[33]||''});}}
  return{emps,quantTotals,aquiMap};
}

/* --- PARSE EMERGENCIAL (apenas Planilha1, expand !ref, normalize uppercase) --- */
function parseEmergencial(wb){
  const records=[];const sheetName=wb.SheetNames.find(n=>n.trim()==='Planilha1');if(!sheetName)return records;
  const ws=wb.Sheets[sheetName];
  if(ws['!ref']){const keys=Object.keys(ws).filter(k=>k[0]!=='!');let maxRow=0;keys.forEach(k=>{const m=k.match(/(\d+)/);if(m)maxRow=Math.max(maxRow,parseInt(m[1]));});if(maxRow>0){const ref=XLSX.utils.decode_range(ws['!ref']);if(maxRow-1>ref.e.r){ref.e.r=maxRow-1;ws['!ref']=XLSX.utils.encode_range(ref);}}}
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,blankrows:false});
  let hIdx=raw.findIndex(r=>r&&r.some(c=>String(c||'').toUpperCase()==='ORDEM'));if(hIdx<0)return records;
  const hRow=raw[hIdx];const col={};hRow.forEach((v,i)=>{if(v)col[String(v).toUpperCase().trim()]=i;});
  const C=k=>col[k]!==undefined?col[k]:-1;const g=(r,k)=>{const i=C(k);return i>=0&&r[i]!=null?r[i]:'';};
  for(let i=hIdx+1;i<raw.length;i++){const r=raw[i];if(!r)continue;const ordem=g(r,'ORDEM');if(!ordem)continue;
    records.push({ordem,dataOcorrencia:parseDate(g(r,'DATA OCORRÊNCIA')),evento:String(g(r,'EVENTO')||'').toUpperCase().trim(),complemento:g(r,'COMPLEMENTO EVENTO'),nomeArea:String(g(r,'NOME ÁREA')||'').trim(),regional:String(g(r,'REGIONAL')||'').toUpperCase().trim(),subprefeitura:String(g(r,'SUBPREFEITURA')||'').trim(),sei:g(r,'Nº SEI'),dataAtendimento:parseDate(g(r,'DATA ATENDIMENTO')),nome:g(r,'NOME BENEFICIÁRIO'),cpf:g(r,'CPF'),cartao:g(r,'NÚMERO CARTÃO/REFERÊNCIA'),obs:g(r,'OBSERVAÇÃO'),situacaoCpf:g(r,'SITUAÇÃO CPF'),status:g(r,'STATUS')});}
  return records;
}

/* --- PARSE INDENIZAÇÃO --- */
function parseIndenizacao(wb,regional,areaNome){
  const records=[];const sn=wb.SheetNames.find(n=>n.toUpperCase().includes('PRIORIDAD'))||wb.SheetNames[0];if(!sn)return records;
  const raw=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null});if(raw.length<2)return records;
  const hRow=raw[0];const col={};hRow.forEach((v,i)=>{if(v)col[String(v).toUpperCase().trim()]=i;});
  const fc=(...ns)=>{for(const n of ns){for(const[k,v] of Object.entries(col)){if(k.includes(n.toUpperCase()))return v;}}return -1;};
  const cArea=fc('ÁREA','AREA'),cSelo=fc('SELO'),cTitular=fc('TITULAR','NOME'),cOcup=fc('OCUPAÇÃO','OCUPACAO'),cUso=fc('USO'),cSit=fc('SITUAÇÃO','SITUACAO'),cCad=fc('CADASTRO'),cTipoAtend=fc('TIPO ATENDIMENTO','ATENDIMENTO'),cValor=fc('VALOR','BONIFICAÇÃO','BONIFICACAO'),cStatus=fc('STATUS'),cStatusPgto=fc('PAGAMENTO');
  for(let i=1;i<raw.length;i++){const r=raw[i];if(!r)continue;const g=c=>c>=0&&r.length>c?r[c]:null;const area=String(g(cArea)||'').trim();if(!area)continue;
    records.push({area,regional:regional||'',fonte:areaNome||'',selo:String(g(cSelo)||'').trim(),titular:String(g(cTitular)||'').trim(),ocupacao:String(g(cOcup)||'').trim(),usoImovel:String(g(cUso)||'').trim(),situacao:String(g(cSit)||'').trim(),cadastro:String(g(cCad)||'').trim(),tipoAtendimento:String(g(cTipoAtend)||'').trim(),valor:num(g(cValor)),status:String(g(cStatus)||'').trim(),statusPgto:String(g(cStatusPgto)||'').trim()});}
  return records;
}

/* --- PARSE DEMANDA FECHADA --- */
function parseDemanda(wb){
  const sn=wb.SheetNames.find(n=>n.toLowerCase().includes('consolida'))||wb.SheetNames[0];if(!sn)return[];
  const raw=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:null});if(raw.length<2)return[];
  const hRow=raw[0];const col={};hRow.forEach((v,i)=>{if(v)col[String(v).replace(/\n/g,' ').trim()]=i;});
  const fc=(...ns)=>{for(const n of ns){for(const[k,v] of Object.entries(col)){if(k.toLowerCase().includes(n.toLowerCase()))return v;}}return -1;};
  const cReg=fc('Regional'),cProg=fc('Prog. Habitacional','Prog.'),cMul=fc('Mulher Resp'),cPCD=fc('PCD'),cInt=fc('5 ou +'),cInf=fc('Infância'),cIdo=fc('Idoso'),cEsp=fc('Espaço habitado'),cSub=fc('Subprefeitura'),cDis=fc('Distrito'),cEmp=fc('Empreendimento'),cSta=fc('STATUS FEV','STATUS'),cInd=fc('Quem indicou');
  const records=[];
  for(let i=1;i<raw.length;i++){const r=raw[i];if(!r||!r[0])continue;const g=c=>c>=0&&r.length>c?r[c]:null;
    records.push({regional:String(g(cReg)||'').toUpperCase().trim(),programa:String(g(cProg)||'').trim(),mulher:String(g(cMul)||'')==='Sim',pcd:String(g(cPCD)||'')==='Sim',integ5:String(g(cInt)||'')==='Sim',infancia:String(g(cInf)||'')==='Sim',idoso:String(g(cIdo)||'')==='Sim',espaco:String(g(cEsp)||'').trim(),subpref:String(g(cSub)||'').trim(),distrito:String(g(cDis)||'').trim(),empreendimento:String(g(cEmp)||'').trim(),status:String(g(cSta)||'').trim(),indicou:String(g(cInd)||'').trim()});}
  return records;
}

/* --- NAV MENU BUILDER --- */
function buildNav(activePage){
  const pages=[
    {id:'index',href:'index.html',icon:'📊',label:'Dashboard',dot:'on'},
    {id:'ppe',href:'ppe.html',icon:'🏠',label:'Unidade Habitacional',dot:'on'},
    {id:'emergencial',href:'emergencial.html',icon:'🚨',label:'Emergencial',dot:'on'},
    {id:'indenizacao',href:'indenizacao.html',icon:'📋',label:'Indenização',dot:'on'},
    {id:'auxilio',href:'#',icon:'💰',label:'Auxílio Aluguel',dot:'off'},
    {id:'historico',href:'#',icon:'📊',label:'Histórico',dot:'off'},
  ];
  return '<nav class="prog-menu">'+pages.map(p=>
    `<a class="prog-menu-link${p.id===activePage?' active':''}" href="${p.href}"><span class="menu-dot ${p.dot}"></span>${p.icon} ${p.label}</a>`
  ).join('')+'</nav>';
}
