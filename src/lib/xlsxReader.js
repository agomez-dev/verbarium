const td = new TextDecoder();

async function inflateRaw(bytes){
  if(typeof DecompressionStream==='undefined') throw new Error('Tu navegador no soporta la descompresión necesaria para archivos .xlsx. Usa Chrome/Edge actualizado o un CSV.');
  const ds=new DecompressionStream('deflate-raw');
  const stream=new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function unzip(buffer){
  const view=new DataView(buffer); const bytes=new Uint8Array(buffer); const files={}; let p=0;
  while(p+30<=view.byteLength){
    const sig=view.getUint32(p,true); if(sig!==0x04034b50) break;
    const flags=view.getUint16(p+6,true), method=view.getUint16(p+8,true);
    const compSize=view.getUint32(p+18,true), nameLen=view.getUint16(p+26,true), extraLen=view.getUint16(p+28,true);
    if(flags&0x08) throw new Error('Este XLSX usa un formato ZIP con descriptor no compatible. Guárdalo nuevamente desde Excel y vuelve a intentarlo.');
    const name=td.decode(bytes.slice(p+30,p+30+nameLen));
    const start=p+30+nameLen+extraLen, compressed=bytes.slice(start,start+compSize);
    let out;
    if(method===0) out=compressed; else if(method===8) out=await inflateRaw(compressed); else throw new Error('Método de compresión XLSX no compatible.');
    files[name]=td.decode(out);
    p=start+compSize;
  }
  return files;
}

const colIndex=(ref='A1')=>{let n=0;for(const c of ref.match(/[A-Z]+/i)?.[0]||'A')n=n*26+(c.toUpperCase().charCodeAt(0)-64);return n-1};
const textFromNode=node=>Array.from(node.getElementsByTagName('t')).map(x=>x.textContent||'').join('');

export async function readXlsx(file){
  const files=await unzip(await file.arrayBuffer());
  const shared=[];
  if(files['xl/sharedStrings.xml']){
    const doc=new DOMParser().parseFromString(files['xl/sharedStrings.xml'],'application/xml');
    for(const si of doc.getElementsByTagName('si')) shared.push(textFromNode(si));
  }
  const sheetKey=Object.keys(files).filter(k=>/^xl\/worksheets\/sheet\d+\.xml$/.test(k)).sort()[0];
  if(!sheetKey) throw new Error('No se encontró una hoja de cálculo válida.');
  const doc=new DOMParser().parseFromString(files[sheetKey],'application/xml');
  const rows=[];
  for(const rowNode of doc.getElementsByTagName('row')){
    const row=[];
    for(const c of rowNode.getElementsByTagName('c')){
      const idx=colIndex(c.getAttribute('r')||'A1'); const type=c.getAttribute('t'); let value='';
      if(type==='inlineStr') value=textFromNode(c);
      else {const v=c.getElementsByTagName('v')[0]?.textContent??''; value=type==='s' ? (shared[Number(v)]??'') : v;}
      row[idx]=value;
    }
    rows.push(row.map(v=>v??''));
  }
  return rows;
}

export async function readTableFile(file){
  if(file.name.toLowerCase().endsWith('.xlsx')) return readXlsx(file);
  if(file.name.toLowerCase().endsWith('.csv')){
    const text=await file.text(); return parseCsv(text);
  }
  throw new Error('Formato no compatible. Usa .xlsx o .csv.');
}

function parseCsv(text){
  const out=[]; let row=[], cell='', q=false;
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(q){if(ch==='"'&&text[i+1]==='"'){cell+='"';i++;}else if(ch==='"')q=false;else cell+=ch;}
    else if(ch==='"')q=true; else if(ch===','){row.push(cell);cell='';} else if(ch==='\n'){row.push(cell.replace(/\r$/,''));out.push(row);row=[];cell='';} else cell+=ch;
  }
  if(cell.length||row.length){row.push(cell);out.push(row)}
  return out;
}
