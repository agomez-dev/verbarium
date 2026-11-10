import {readTableFile} from './lib/xlsxReader.js';

const $=s=>document.querySelector(s), $$=s=>Array.from(document.querySelectorAll(s));
const STORAGE='verbarium.custom.v1', THEME='verbarium.theme', SPEECH_LOCALE='verbarium.speechLocale';
const norm=(s='')=>String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
const esc=(s='')=>String(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const firstMeaning=(meaning='')=>meaning.split(/\s+-\s+/)[0].replace(/\([^)]*\)/g,'').trim().toLowerCase();
const clean=(v='')=>String(v).split('/')[0].trim().toLowerCase();
const cap=s=>s.charAt(0).toUpperCase()+s.slice(1);
const state={verbs:[],grammar:[],config:null,query:'',type:'ALL',page:1,selected:null,dynamicHeaders:[],dataOrigin:'PDF',theme:localStorage.getItem(THEME)||'dark',speechLocale:localStorage.getItem(SPEECH_LOCALE)||'en-US'};
document.documentElement.dataset.theme=state.theme;

const icons={search:'M21 21l-4.35-4.35m1.35-5.65A7 7 0 1 1 4 11a7 7 0 0 1 14 0Z',upload:'M12 16V4m0 0 4 4m-4-4L8 8M4 16v4h16v-4',download:'M12 4v12m0 0 4-4m-4 4-4-4M4 20h16',close:'M6 6l12 12M18 6 6 18',filter:'M4 6h16M7 12h10M10 18h4',moon:'M20 15.5A8 8 0 0 1 8.5 4 8 8 0 1 0 20 15.5Z',sun:'M12 4V2m0 20v-2M4 12H2m20 0h-2M5.64 5.64 4.22 4.22m15.56 15.56-1.42-1.42M18.36 5.64l1.42-1.42M4.22 19.78l1.42-1.42M16 12a4 4 0 1 1-8 0 4 4 0 0 1 8 0Z',chevron:'m9 18 6-6-6-6',reset:'M20 11a8 8 0 1 0-2.34 5.66M20 4v7h-7',volume:'M11 5 6.5 9H3v6h3.5l4.5 4V5Zm4.5 3.5a5 5 0 0 1 0 7M18 6a8 8 0 0 1 0 12'};
const icon=(name,size=20)=>`<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${icons[name]}"></path></svg>`;
const badge=t=>`<span class="badge ${t==='I'?'irregular':'regular'}">${t==='I'?'Irregular':'Regular'}</span>`;


function getSpeechVoice(locale=state.speechLocale){
  if(!('speechSynthesis' in window))return null;
  const voices=window.speechSynthesis.getVoices();
  return voices.find(v=>v.lang===locale)||voices.find(v=>v.lang?.toLowerCase().startsWith(locale.slice(0,2).toLowerCase()))||null;
}
function speak(text,rate=0.95){
  const value=String(text||'').trim();
  if(!value||!('speechSynthesis' in window))return;
  window.speechSynthesis.cancel();
  const utterance=new SpeechSynthesisUtterance(value);
  utterance.lang=state.speechLocale;
  utterance.rate=Number(rate)||0.95;
  utterance.pitch=1;
  const voice=getSpeechVoice(state.speechLocale);if(voice)utterance.voice=voice;
  window.speechSynthesis.speak(utterance);
}
function speechControls(text){
  const value=clean(text);if(!value)return '';
  const disabled=!('speechSynthesis' in window)?'disabled title="Tu navegador no ofrece síntesis de voz"':'';
  return `<span class="audio-actions"><button class="audio-btn" type="button" data-speak="${esc(value)}" data-rate="0.95" ${disabled}>${icon('volume',14)} Normal</button><button class="audio-btn slow" type="button" data-speak="${esc(value)}" data-rate="0.60" ${disabled}>🐢 Lenta</button></span>`;
}
function ipaValue(v,key){
  if(v.ipa&&typeof v.ipa==='object')return v.ipa[key]||'';
  const map={base:'ipaBase',thirdPerson:'ipaThirdPerson',simplePast:'ipaSimplePast',pastParticiple:'ipaPastParticiple',gerund:'ipaGerund'};
  return v[map[key]]||(key==='base'&&typeof v.ipa==='string'?v.ipa:'')||'';
}

function buildTenses(v){
  const base=clean(v.simpleForm),past=clean(v.simplePast),part=clean(v.pastParticiple),ger=clean(v.gerund),gloss=firstMeaning(v.meaning)||'realizar la acción',isBe=base==='be';
  const forms={presentSimple:isBe?'I am':`I ${base}`,presentContinuous:isBe?'I am being':`I am ${ger}`,pastSimple:isBe?'I was':`I ${past}`,pastContinuous:isBe?'I was being':`I was ${ger}`,presentPerfect:`I have ${part}`,presentPerfectContinuous:`I have been ${ger}`,pastPerfect:`I had ${part}`,pastPerfectContinuous:`I had been ${ger}`,futureSimple:`I will ${base}`,futureContinuous:`I will be ${ger}`,futurePerfect:`I will have ${part}`,futurePerfectContinuous:`I will have been ${ger}`};
  const ending={presentSimple:'regularly.',presentContinuous:'right now.',pastSimple:'yesterday.',pastContinuous:'when you called.',presentPerfect:'several times.',presentPerfectContinuous:'for two hours.',pastPerfect:'before the meeting started.',pastPerfectContinuous:'for two hours before lunch.',futureSimple:'tomorrow.',futureContinuous:'this time tomorrow.',futurePerfect:'by Friday.',futurePerfectContinuous:'for two hours by noon.'};
  const es={presentSimple:`Presente habitual: «${gloss}».`,presentContinuous:`Acción en progreso: «estar ${gloss}».`,pastSimple:`Acción terminada en pasado: «${gloss}» en pasado.`,pastContinuous:`Acción en progreso en pasado: «estar ${gloss}» en pasado.`,presentPerfect:`Resultado o experiencia: «haber ${gloss}».`,presentPerfectContinuous:`Duración hasta ahora: «haber estado ${gloss}».`,pastPerfect:`Anterioridad en pasado: «haber ${gloss}» antes de otra acción.`,pastPerfectContinuous:`Duración anterior: «haber estado ${gloss}».`,futureSimple:`Acción futura: «${gloss}» en futuro.`,futureContinuous:`Acción futura en progreso: «estar ${gloss}» en futuro.`,futurePerfect:`Acción completada antes de un punto futuro: «haber ${gloss}».`,futurePerfectContinuous:`Duración acumulada hasta un punto futuro: «haber estado ${gloss}».`};
  return Object.entries(forms).map(([id,form])=>({id,form,example:`${cap(form)} ${ending[id]}`,translation:es[id]}));
}
function quickExamples(v){const b=clean(v.simpleForm),p=clean(v.simplePast),g=clean(v.gerund),m=firstMeaning(v.meaning);return [{en:`I ${b} every week.`,es:`Uso habitual relacionado con «${m}».`},{en:`I ${p} yesterday.`,es:`Uso de pasado relacionado con «${m}».`},{en:`I am ${g} now.`,es:`Acción en progreso relacionada con «${m}».`}];}
function filtered(){return state.verbs.filter(v=>{if(state.type!=='ALL'&&v.type!==state.type)return false;if(!state.query.trim())return true;const q=norm(state.query),hay=[v.simpleForm,v.meaning,v.simplePast,v.pastParticiple,v.gerund,...Object.values(v._extra||{})].map(norm).join(' ');return hay.includes(q)})}
function counts(){return {all:state.verbs.length,r:state.verbs.filter(v=>v.type==='R').length,i:state.verbs.filter(v=>v.type==='I').length}}

function render(){
  const c=counts(), f=filtered(), ps=state.config?.defaultPageSize||24,tp=Math.max(1,Math.ceil(f.length/ps));if(state.page>tp)state.page=tp;const visible=f.slice((state.page-1)*ps,state.page*ps);
  $('#app').innerHTML=`<div class="app-shell">
  <header class="topbar"><a class="brand" href="./"><span class="brand-mark">V</span><span><b>${esc(state.config.appName)}</b><small>${esc(state.config.subtitle)}</small></span></a><div class="top-actions"><button class="secondary desktop-only" data-action="export">${icon('download',17)} JSON</button><button class="primary" data-action="import">${icon('upload',17)} <span>Importar Excel</span></button><button class="icon-btn" data-action="theme" aria-label="Cambiar tema">${icon(state.theme==='dark'?'sun':'moon')}</button></div></header>
  <main><section class="hero"><div class="hero-copy"><span class="eyebrow">BIBLIOTECA DE ESTUDIO · ${esc(state.dataOrigin)}</span><h1>Verbos, sin perder el <em>contexto.</em></h1><p>Busca una forma verbal, compara sus conjugaciones y abre una ficha para revisar reglas, tiempos y ejemplos.</p></div><div class="stats"><div><b>${c.all}</b><span>verbos</span></div><div><b>${c.i}</b><span>irregulares</span></div><div><b>12</b><span>tiempos</span></div></div></section>
  <section class="workspace"><div class="controls"><label class="search">${icon('search')}<input id="searchInput" value="${esc(state.query)}" placeholder="Buscar: write, escribir, wrote…" aria-label="Buscar verbos"><button data-action="clearSearch" aria-label="Limpiar búsqueda" ${state.query?'':'hidden'}>${icon('close',16)}</button></label><div class="segmented">${[['ALL','Todos',c.all],['I','Irregulares',c.i],['R','Regulares',c.r]].map(([k,l,n])=>`<button data-type="${k}" class="${state.type===k?'active':''}">${l} <span>${n}</span></button>`).join('')}</div>${state.dataOrigin!=='PDF'?`<button class="secondary reset-data" data-action="resetData">${icon('reset',16)} Restaurar PDF</button>`:''}</div>
  ${state.dynamicHeaders.length?`<div class="dynamic-note">${icon('filter',16)}<span>Columnas dinámicas activas:</span>${state.dynamicHeaders.map(h=>`<code>${esc(h.original)}</code>`).join('')}</div>`:''}
  <div class="result-line"><span><b>${f.length}</b> resultados</span><span>Pasado · Participio · Gerundio</span></div>
  ${visible.length?`<div class="verb-grid">${visible.map(card).join('')}</div>`:`<div class="empty"><span>∅</span><h3>No hay coincidencias</h3><p>Prueba otro término o limpia los filtros activos.</p><button class="secondary" data-action="resetFilters">Limpiar filtros</button></div>`}
  ${tp>1?`<nav class="pagination"><button data-action="prev" ${state.page===1?'disabled':''}>Anterior</button><span>Página <b>${state.page}</b> de ${tp}</span><button data-action="next" ${state.page===tp?'disabled':''}>Siguiente</button></nav>`:''}</section></main>
  <footer><span>Datos base: listado proporcionado en VERBS.pdf.</span><button data-action="export">Exportar datos actuales como JSON</button></footer></div>`;
  bind();
}
function card(v){return `<button class="verb-card" data-verb="${v.id}"><span class="verb-card-top">${badge(v.type)}<span class="verb-id">#${String(v.id).padStart(3,'0')}</span></span><strong>${esc(v.simpleForm)}</strong><span class="meaning">${esc(v.meaning||'Sin traducción')}</span><span class="forms"><span>${esc(v.simplePast||'—')}</span><span>${esc(v.pastParticiple||'—')}</span><span>${esc(v.gerund||'—')}</span></span><span class="open-label">Ver ficha ${icon('chevron',15)}</span></button>`}
function bind(){
  $('#searchInput')?.addEventListener('input',e=>{const pos=e.target.selectionStart??e.target.value.length;state.query=e.target.value;state.page=1;render();const input=$('#searchInput');if(input){input.focus({preventScroll:true});input.setSelectionRange(pos,pos);}});
  $$('[data-type]').forEach(b=>b.addEventListener('click',()=>{state.type=b.dataset.type;state.page=1;render()}));
  $$('[data-verb]').forEach(b=>b.addEventListener('click',()=>openDetail(Number(b.dataset.verb))));
  $$('[data-action]').forEach(b=>b.addEventListener('click',()=>action(b.dataset.action)));
}
async function action(a){
  if(a==='export') return exportJson(); if(a==='import') return openImport(); if(a==='theme'){state.theme=state.theme==='dark'?'light':'dark';document.documentElement.dataset.theme=state.theme;localStorage.setItem(THEME,state.theme);return render()}
  if(a==='clearSearch'){state.query='';state.page=1;return render()} if(a==='resetFilters'){state.query='';state.type='ALL';state.page=1;return render()} if(a==='prev'){state.page--;return render()} if(a==='next'){state.page++;return render()}
  if(a==='resetData'){localStorage.removeItem(STORAGE);state.verbs=await fetch('./public/data/verbs.json').then(r=>r.json());state.dynamicHeaders=[];state.dataOrigin='PDF';state.query='';state.type='ALL';state.page=1;return render()}
}
function openDetail(id){
  const v=state.verbs.find(x=>Number(x.id)===id);if(!v)return;
  const ts=buildTenses(v),examples=quickExamples(v);
  const forms=[['Forma base','base',v.simpleForm],['Tercera persona','thirdPerson',v.thirdPerson],['Pasado simple','simplePast',v.simplePast],['Participio','pastParticiple',v.pastParticiple],['Gerundio','gerund',v.gerund]];
  document.body.insertAdjacentHTML('beforeend',`<div class="drawer-layer" id="drawerLayer"><aside class="drawer" role="dialog" aria-modal="true" aria-labelledby="detail-title"><div class="drawer-head"><div>${badge(v.type)}<h2 id="detail-title">${esc(v.simpleForm)}</h2><p>${esc(v.meaning)}</p></div><button class="icon-btn" id="closeDrawer" aria-label="Cerrar detalle">${icon('close')}</button></div>
  <div class="pronunciation-toolbar"><div><b>Pronunciación</b><span>La voz depende de las voces instaladas en tu navegador/sistema.</span></div><label>Acento<select id="speechLocale"><option value="en-US" ${state.speechLocale==='en-US'?'selected':''}>🇺🇸 English US</option><option value="en-GB" ${state.speechLocale==='en-GB'?'selected':''}>🇬🇧 English UK</option></select></label></div>
  <div class="core-forms">${forms.map(([l,k,x])=>`<div><span>${l}</span><b>${esc(x||'—')}</b>${ipaValue(v,k)?`<small class="ipa">${esc(ipaValue(v,k))}</small>`:''}${speechControls(x)}</div>`).join('')}</div>
  ${v._extra&&Object.keys(v._extra).length?`<section><h3>Campos del archivo</h3><div class="extra-grid">${Object.entries(v._extra).map(([k,x])=>`<div><span>${esc(k)}</span><b>${esc(x||'—')}</b></div>`).join('')}</div></section>`:''}
  <section><div class="section-heading"><h3>12 tiempos verbales</h3><span>Equivalencia orientativa</span></div><div class="tense-list">${state.grammar.map((g,i)=>{const t=ts.find(x=>x.id===g.id);return `<details ${i===0?'open':''}><summary><span><b>${esc(g.name)}</b><small>${esc(g.nameEs)}</small></span><code>${esc(t?.form)}</code></summary><div class="tense-body"><p><b>Uso.</b> ${esc(g.use)}</p><p><b>Regla.</b> ${esc(g.rule)}</p><p><b>Marcadores.</b> ${esc(g.signal)}</p><div class="example-box"><span>${esc(t?.example)}</span><small>${esc(t?.translation)}</small>${speechControls(t?.example)}</div></div></details>`}).join('')}</div></section>
  <section><h3>Ejemplos rápidos</h3><div class="example-list">${examples.map(x=>`<div><b>${esc(x.en)}</b><span>${esc(x.es)}</span>${speechControls(x.en)}</div>`).join('')}</div></section>
  <section class="source-note"><h3>Nota de datos</h3><p>Las formas base provienen del PDF proporcionado. La app conserva esos datos y no corrige automáticamente posibles erratas del documento fuente. Las equivalencias de los 12 tiempos son guías estructurales de estudio.</p></section></aside></div>`);
  const layer=$('#drawerLayer');
  $('#closeDrawer').addEventListener('click',()=>{window.speechSynthesis?.cancel();layer.remove()});
  layer.addEventListener('mousedown',e=>{if(e.target===layer){window.speechSynthesis?.cancel();layer.remove()}});
  $('#speechLocale')?.addEventListener('change',e=>{state.speechLocale=e.target.value;localStorage.setItem(SPEECH_LOCALE,state.speechLocale);window.speechSynthesis?.cancel()});
  $$('[data-speak]').forEach(b=>b.addEventListener('click',()=>speak(b.dataset.speak,b.dataset.rate)));
}
function openImport(){document.body.insertAdjacentHTML('beforeend',`<div class="modal-layer" id="importLayer"><div class="modal" role="dialog" aria-modal="true"><div class="modal-head"><div><span class="eyebrow">IMPORTAR DATOS</span><h2>Excel → aplicación</h2></div><button class="icon-btn" id="closeImport" aria-label="Cerrar">${icon('close')}</button></div><p>La primera fila se interpreta como encabezados. Los campos conocidos alimentan la ficha; cualquier columna adicional se conserva como metadato dinámico.</p><div class="header-map"><span>Ejemplos reconocidos</span><div>VERBO / SIMPLE FORM</div><div>TYPE / TIPO</div><div>SIMPLE PAST</div><div>PAST PARTICIPLE</div><div>GERUND</div><div>MEANING / SIGNIFICADO</div></div><label class="dropzone">${icon('upload',28)}<b>Selecciona .xlsx o .csv</b><span>El archivo se procesa localmente en el navegador.</span><input id="fileInput" type="file" accept=".xlsx,.csv"></label><div id="importError"></div><button class="secondary full" id="templateBtn">${icon('download',17)} Descargar plantilla CSV compatible con Excel</button></div></div>`);const layer=$('#importLayer');$('#closeImport').addEventListener('click',()=>layer.remove());layer.addEventListener('mousedown',e=>{if(e.target===layer)layer.remove()});$('#fileInput').addEventListener('change',handleImport);$('#templateBtn').addEventListener('click',downloadTemplate);}
async function handleImport(e){const file=e.target.files?.[0];if(!file)return;const err=$('#importError');err.innerHTML='';try{const rows=await readTableFile(file);const result=tableToVerbs(rows);state.verbs=result.data;state.dynamicHeaders=result.headers;state.dataOrigin='Excel local';state.query='';state.type='ALL';state.page=1;localStorage.setItem(STORAGE,JSON.stringify({data:result.data,headers:result.headers}));$('#importLayer').remove();render()}catch(ex){err.innerHTML=`<div class="error-box">${esc(ex.message||'No se pudo leer el archivo.')}</div>`}}
function tableToVerbs(rows){if(rows.length<2)throw new Error('El archivo debe tener encabezados y al menos una fila de datos.');const aliases=state.config.excel.headerAliases||{};const normHead=s=>String(s).trim().toLowerCase().replace(/[_-]+/g,' ').replace(/\s+/g,' ');const camel=s=>s.replace(/\s+(.)/g,(_,c)=>c.toUpperCase());const mapped=rows[0].map(h=>({original:String(h||'').trim(),key:aliases[normHead(h)]||camel(normHead(h))}));if(!mapped.some(h=>h.key==='simpleForm'))throw new Error('Debe existir una columna VERBO, VERB o SIMPLE FORM.');const core=['id','type','simpleForm','thirdPerson','simplePast','pastParticiple','gerund','meaning','ipa','ipaBase','ipaThirdPerson','ipaSimplePast','ipaPastParticiple','ipaGerund'];const data=rows.slice(1).filter(r=>r.some(x=>String(x||'').trim())).map((r,i)=>{const out={id:i+1,_extra:{}};mapped.forEach((h,idx)=>{if(!h.original)return;const val=r[idx]??'';if(core.includes(h.key))out[h.key]=val;else out._extra[h.key]=val});out.id=Number(out.id)||i+1;out.type=String(out.type||'R').toUpperCase().startsWith('I')?'I':'R';for(const k of ['simpleForm','thirdPerson','simplePast','pastParticiple','gerund'])out[k]=String(out[k]||'').trim().toUpperCase();out.meaning=String(out.meaning||'').trim();return out}).filter(x=>x.simpleForm);return {data,headers:mapped.filter(h=>h.original&&!core.includes(h.key))}}
function exportJson(){const blob=new Blob([JSON.stringify(state.verbs,null,2)],{type:'application/json'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='verbos.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),500)}
function downloadTemplate(){const csv='NUM,TYPE,SIMPLE FORM,THIRD PERSON,SIMPLE PAST,PAST PARTICIPLE,GERUND,MEANING,IPA BASE,IPA THIRD PERSON,IPA SIMPLE PAST,IPA PAST PARTICIPLE,IPA GERUND,NIVEL,CATEGORIA\n1,I,WRITE,WRITES,WROTE,WRITTEN,WRITING,Escribir - Redactar,/raɪt/,/raɪts/,/roʊt/,/ˈrɪtən/,/ˈraɪtɪŋ/,B1,Comunicación\n';const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='plantilla-verbos.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),500)}

async function init(){
  const [verbs,grammar,config]=await Promise.all([fetch('./public/data/verbs.json').then(r=>r.json()),fetch('./public/data/grammar.json').then(r=>r.json()),fetch('./public/data/app-config.json').then(r=>r.json())]);state.verbs=verbs;state.grammar=grammar;state.config=config;try{const saved=JSON.parse(localStorage.getItem(STORAGE)||'null');if(saved?.data?.length){state.verbs=saved.data;state.dynamicHeaders=saved.headers||[];state.dataOrigin='Excel local'}}catch{}render();
}
init().catch(err=>{$('#app').innerHTML=`<main class="empty"><h3>No se pudo iniciar la aplicación</h3><p>${esc(err.message)}</p><p>Abre el proyecto mediante un servidor local, no directamente con file://.</p></main>`});
