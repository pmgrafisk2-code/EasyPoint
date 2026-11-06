javascript:(()=>{try{
/* ============================================================
   EasyPoint 2.1 / Adpoint — kontrollert kø og massetillegging
   - Norsk/engelsk Adpoint, linje- og placeholder-ID ved hvert steg
   - Én jobb om gangen; stopp/lukk venter på pågående lagring
   - Legg til antall materiellplasser uten CSV; auto-opprett ved import
   - Bekreft lagring, bevar eksisterende materiell, ingen blinde gjentakelser
   ============================================================ */

const d=document,w=window;
const previous=w.__easyPoint;
if(previous?.busy){ previous.focus?.(); alert('EasyPoint arbeider fortsatt. Stopp jobben og vent før du åpner på nytt.'); return; }
previous?.destroy?.();
const life=new AbortController();
let refreshSimpleUI=()=>{};
const app={busy:false,closing:false,focus:()=>d.getElementById('ap3p_bar')?.scrollIntoView({block:'nearest'})};
class UncertainMutationError extends Error {constructor(message){super(message);this.name='UncertainMutationError';this.uncertain=true;}}
w.__easyPoint=app;
const listen=(target,type,fn,options={})=>target.addEventListener(type,fn,{...options,signal:life.signal});

/* ========= storage ========= */
const S=(k,v)=>{try{localStorage.setItem(k,JSON.stringify(v));return true}catch{return false}};
const G=(k,f)=>{try{return JSON.parse(localStorage.getItem(k))??f}catch{return f}};

/* ========= small utils ========= */
const vis=el=>!!(el?.isConnected && !el.closest('[hidden]') && el.getClientRects().length && getComputedStyle(el).display!=='none' && getComputedStyle(el).visibility!=='hidden');
const cs=s=>String(s||'').toLowerCase().replace(/\s+/g,'').replace(/×/g,'x');
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function waitFor(fn,ms=6000,step=100){
  const t=Date.now(); let v;
  while(Date.now()-t<ms){ v=fn(); if(v) return v; await sleep(step); }
  return null;
}
function clamp(n,min,max){return Math.max(min,Math.min(max,n))}

/* ========= parsing ========= */
function parseSizeVariant(s){
  const str=String(s||'');
  const m=str.match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);
  if(!m) return{size:'',variant:null};
  const size=`${m[1]}x${m[2]}`;
  const v=(str.match(/desktop|mobil|mobile/i)||[''])[0].toLowerCase();
  const variant=v.includes('desk')?'desktop':(v?'mobil':null);
  return{size,variant};
}
function prettyCounts(entries){
  const counts=new Map();
  for(const e of entries){
    const key=e.size+(e.variant?`/${e.variant[0]}`:'');
    counts.set(key,(counts.get(key)||0)+1);
  }
  return[...counts.entries()].map(([k,n])=>n>1?`${k}×${n}`:k).join(', ');
}

/* ========= version key & line-id ========= */
function versionKey(str){
  const s=String(str||'').toLowerCase();
  const mIx=s.match(/\bix[-_][0-9a-z-]+\b/i); if(mIx) return mIx[0];
  const mV=s.match(/\b(?:ver|v)[\s._-]?(\d{1,3})\b/i); if(mV) return `v${mV[1]}`;
  return null;
}
function extractLineIds(str){
  const ids=new Set();
  const re=/L[\s\u00A0_-]*?(\d{6,10})(?!\d)/gi;
  const s=String(str||''); let m;
  while((m=re.exec(s))!==null) ids.add(m[1]);
  return [...ids];
}

// Line readiness is checked before each operation.

/* ========= detect side ========= */
function detectSide(str){
  const s=(str||'').toLowerCase().replace('høyre','hoyre');
  if(/\b(hoyre|right|r)\b/.test(s))return'right';
  if(/\b(venstre|left|l)\b/.test(s))return'left';
  return null;
}

/* ========= tile helpers ========= */
function tileClickTargets(tile){
  return [
    tile.querySelector('.set-cardLabel__main'),
    tile.querySelector('.set-cardLabel'),
    tile.querySelector('.MuiTypography-root'),
    tile.querySelector('[aria-label*="select" i]'),
    tile
  ].filter(Boolean);
}
function isTileSelected(tile){
  if(tile.classList.contains('set-creativeTile__selected'))return true;
  if(tile.getAttribute('aria-selected')==='true')return true;
  if(tile.closest('[aria-selected="true"]'))return true;
  if(tile.closest('.set-creativeTile__selected'))return true;
  return false;
}
async function selectTile(tile){
  if(!tile?.isConnected) return false;
  if(isTileSelected(tile)) return true;
  for(const target of tileClickTargets(tile)){
    if(!target.isConnected) return false;
    target.scrollIntoView({block:'center'});
    target.click();
    if(await waitFor(()=>tile.isConnected&&isTileSelected(tile),1000,70)) return true;
  }
  return false;
}

function getAllTiles(details){
  let tiles=[...details.querySelectorAll('.set-creativeTile, .set-creativeTile__selected')];
  // filter nested duplicates
  return tiles.filter(t=>!tiles.some(o=>o!==t && o.contains(t)));
}
function tileGroup(tile){
  return tile.closest('.set-creativeContainer, .set-creativeSectionContainer')||tile.parentElement;
}
function groupLabel(group){
  const n=group&&(group.querySelector('.set-cardLabel__main [title]')||
                 group.querySelector('.set-cardLabel__main div[title]')||
                 group.querySelector('.set-cardLabel__main'));
  return (n?.getAttribute?.('title')||n?.textContent||'')||'';
}
function tileOwnLabel(tile){
  const n=tile?.querySelector('.set-cardLabel__main [title],.set-cardLabel__main');
  return n?.getAttribute('title')||n?.textContent||'';
}
function dataRowForTile(tile){
  let row=tile?.closest('tr');
  while(row&&!row.matches('tr.set-matSpecDataRow')) row=row.previousElementSibling;
  return row||null;
}
function rowVariant(row){
  const headers=[...(row?.closest('table')?.tHead?.querySelectorAll('th')||[])];
  const index=headers.findIndex(th=>/module size|modul|format/i.test(th.textContent||''));
  const text=row?.children[index]?.textContent||'';
  if(/desktop/i.test(text)) return 'desktop';
  if(/mobil|mobile/i.test(text)) return 'mobil';
  return null;
}
function tileVariant(tile){
  const labels=tileOwnLabel(tile);
  if(/desktop/i.test(labels)) return 'desktop';
  if(/mobil|mobile/i.test(labels)) return 'mobil';
  return rowVariant(dataRowForTile(tile));
}
function tileSide(tile){return detectSide(tileOwnLabel(tile))}

function tileHasSize(tile,size){
  const rx=new RegExp(`\\b${String(size).replace('x','[x×]')}\\b`,'i');
  if(rx.test((tile.innerText||'')))return true;
  const lbl=tileOwnLabel(tile);
  const mappedRe=new RegExp(`mapped_${String(size).replace('x','[x×]')}(?:\\b|[_-])`,'i');
  if(mappedRe.test(lbl))return true;
  if(rx.test(lbl))return true;
  return false;
}

function setNativeValue(el, value) {
  const proto = Object.getPrototypeOf(el);
  const desc = Object.getOwnPropertyDescriptor(proto, "value");
  const setter = desc && desc.set;
  if (setter) setter.call(el, value);
  else el.value = value;
}


/* ========= sizes from expanded ========= */
function sizesFromExpanded(details){
  if(!details) return [];
  const entries=[];
  const tiles=getAllTiles(details);

  const findSizeInText=s=>{
    const m=String(s||'').match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);
    return m?`${m[1]}x${m[2]}`:'';
  };
  const findSizeInMapped=lbl=>{
    const m=String(lbl||'').match(/mapped_(\d{2,4})\s*[x×]\s*(\d{2,4})(?:\b|[_\s])/i);
    return m?`${m[1]}x${m[2]}`:'';
  };

  for(const tile of tiles){
    const text=(tile.innerText||'');
    const lbl=tileOwnLabel(tile);
    let size=findSizeInText(text);
    if(!size) size=findSizeInMapped(lbl);
    if(!size) size=findSizeInText(lbl);
    if(!size) continue;
    const variant=tileVariant(tile);
    const side=tileSide(tile);
    entries.push({size:cs(size),variant,side,tile});
  }
  return entries;
}
function pickTilesForEntry(details,entry){
  let tiles=getAllTiles(details).filter(t=>tileHasSize(t,entry.size));
  if(entry.variant) tiles=tiles.filter(t=>(tileVariant(t)||null)===entry.variant);
  if(entry.side)    tiles=tiles.filter(t=>(tileSide(t)||null)===entry.side);
  return tiles;
}

/* ========= virtualized list: mount all tiles ========= */
async function ensureAllTilesMounted(root=d){
  // Only scroll real overflowing containers. Re-read cards after scrolling.
  const scrollers=[root,...root.querySelectorAll('.set-creativeSectionContainer,.set-creativeContainer')]
    .filter(el=>el.scrollHeight>el.clientHeight+4 && /auto|scroll/.test(getComputedStyle(el).overflowY));
  for(const el of scrollers){
    const start=el.scrollTop; let last=-1,stable=0;
    for(let i=0;i<24&&stable<3;i++){
      const count=getAllTiles(root).length;
      el.scrollTop=el.scrollHeight; await sleep(120);
      stable=count===last?stable+1:0; last=count;
    }
    el.scrollTop=start;
  }
}

/* ========= info panel switching ========= */
// Panel matching uses the exact placeholder ID, including repeated sizes.

function labelledId(root,label){
  if(!root) return null;
  for(const el of root.querySelectorAll('span,div,label,dt')){
    if(el.children.length || !label.test((el.textContent||'').trim())) continue;
    let parent=el.parentElement;
    for(let depth=0;parent&&depth<3;depth++,parent=parent.parentElement){
      const remainder=(parent.textContent||'').replace(el.textContent,'').trim();
      const match=remainder.match(/^[:\s-]*(\d{5,12})(?!\d)/);
      if(match) return match[1];
    }
  }
  return null;
}
function placeholderId(tile){
  const input=tile?.querySelector('input[type="file"][id^="dropzone-"]');
  const id=input?.id.match(/-(\d+)$/)?.[1];
  return id||labelledId(tile,/^(?:placeholder\s*id|plassholder[- ]?id|materiellplass[- ]?id)$/i);
}
function panelIdentity(){
  const panel=d.querySelector('#InfoPanelContainer');
  return {
    placeholder:labelledId(panel,/^(?:placeholder\s*id|plassholder[- ]?id|materiellplass[- ]?id)$/i),
    line:labelledId(panel,/^(?:line\s*item|linje[- ]?(?:id|nummer))$/i)
  };
}
async function ensurePanelSwitchedToTile(tile,timeout=6000){
  const expected=placeholderId(tile);
  if(!expected||!tile?.isConnected) return false;
  const ok=await waitFor(()=>{
    const identity=panelIdentity();
    return tile.isConnected&&isTileSelected(tile)&&identity.placeholder===expected;
  },timeout,90);
  if(!ok) return false;
  await waitForIdle(12000);
  return panelIdentity().placeholder===expected;
}


/* ========= 3rd-party tab (FAST + stable) ========= */
function getThirdPartyPanel(){
  // Tabpanel for THIRD_PARTY_TAG (du har id/aria-labelledby i HTML)
  return d.querySelector('[role="tabpanel"][id*="THIRD_PARTY_TAG"],[role="tabpanel"][aria-labelledby*="THIRD_PARTY_TAG"]');
}

function clickThirdPartyTab(){
  const tab = [...d.querySelectorAll('#InfoPanelContainer [role="tab"]')]
    .find(el => /THIRD_PARTY_TAG/i.test(el.id||'') || /\b3(?:rd)?\s*party\b|\b3\.\s*party\b/i.test(el.textContent||''));
  tab?.click();
}

async function get3PEditorStrict(){
  clickThirdPartyTab();
  return await waitFor(()=>{
    const panel=getThirdPartyPanel();
    if(!panel||!vis(panel)) return null;
    const cm=[...panel.querySelectorAll('.third-party .CodeMirror')].find(vis);
    if(cm?.CodeMirror) return {kind:'cm',el:cm,panel,cm:cm.CodeMirror};
    const ta=[...panel.querySelectorAll('.third-party textarea')].find(vis);
    return ta?{kind:'ta',el:ta,panel}:null;
  },6000,90);
}

function pasteInto(target,value){
  if(!target) return;
  if(target.kind==='cm' && target.cm){
    try{ target.cm.setValue(String(value||'')); target.cm.refresh?.(); return; }catch{}
  }
  const ta=target.el||target;
  ta.scrollIntoView({block:'center'});
  ta.focus();
  setNativeValue(ta, value);
ta.dispatchEvent(new Event('input',  { bubbles:true }));
ta.dispatchEvent(new Event('change', { bubbles:true }));
ta.dispatchEvent(new Event('blur',   { bubbles:true })); // hjelper ofte MUI aktivere lagreknapp

}
// Saving is implemented below with acknowledgement and read-back.

/* ========= verify/retry ========= */
function normalizeTagForCompare(value){
  return String(value||'').replace(/\r\n/g,'\n').trim();
}

function getEditorValue(editor){
  if(!editor) return '';
  if(editor.kind==='cm' && editor.cm){
    try{return editor.cm.getValue()||''}catch{return''}
  }
  const ta=editor.el||editor;
  return ta?.value||'';
}
async function pasteWithVerify(editor,value,tries=2){
  const want=normalizeTagForCompare(value);
  for(let attempt=0; attempt<=tries; attempt++){
    pasteInto(editor,value);
    await sleep(120);
    if(editor.kind==='cm'){ try{editor.cm?.refresh?.()}catch{} await sleep(80); }
    const got=normalizeTagForCompare(getEditorValue(editor));
    if(got===want) return true;
    await sleep(180);
  }
  return false;
}
function saveButtons(root){
  return [...root.querySelectorAll('button')].filter(vis).filter(btn=>{
    const label=(btn.textContent||btn.getAttribute('aria-label')||btn.title||'').replace(/\s+/g,' ').trim();
    return /^(?:lagre(?: endringer)?|save(?: changes)?|oppdater|update|send inn på nytt|submit again|resubmit|reprocess)$/i.test(label);
  });
}
function visibleAdpointError(){
  const nodes=[...d.querySelectorAll('[role="alert"],.MuiAlert-standardError,.MuiAlert-filledError')]
    .filter(el=>vis(el)&&!el.closest('#ap3p_bar'));
  return nodes.find(el=>/error|failed|failure|invalid|feil|mislykk|ugyldig/i.test(el.textContent||''))||null;
}
function successNotice(){
  return [...d.querySelectorAll('[role="alert"],.MuiSnackbar-root')].filter(vis)
    .find(el=>/saved|success|lagret|vellykket|oppdatert|updated/i.test(el.textContent||''))||null;
}
async function saveWithRetry(editor,record,expected){
  // No blind write retries: a timed-out save may already have reached the server.
  const ready=await waitFor(()=>{
    const panel=getThirdPartyPanel();
    const buttons=panel&&saveButtons(panel).filter(btn=>!isDisabled(btn));
    return buttons?.length===1?buttons[0]:null;
  },6000,100);
  if(!ready) throw new Error('Fant ikke én aktiv Lagre/Update-knapp. Scriptet er ikke bekreftet lagret.');
  assertTarget(record);
  if(normalizeTagForCompare(getEditorValue(editor))!==normalizeTagForCompare(expected)) throw new Error('Editorinnholdet ble endret før lagring.');
  if(visibleAdpointError()) throw new Error('Adpoint viser en feil før lagring. Kontroller siden.');
  const oldNotice=successNotice(); const oldText=oldNotice?.textContent;
  ready.click();
  const acknowledged=await waitFor(()=>{
    const error=visibleAdpointError();
    if(error) throw new Error('Adpoint avviste lagringen: '+error.textContent.trim().slice(0,200));
    const notice=successNotice();
    if(notice&&(notice!==oldNotice||notice.textContent!==oldText)) return true;
    // Adpoint typically disables or removes the action after accepting the tag.
    return !ready.isConnected||isDisabled(ready);
  },15000,100);
  if(!acknowledged) throw new UncertainMutationError('Lagringen er usikker: ingen kvittering fra Adpoint. Kontroller materiellplassen før ny kjøring.');
  try{await waitForIdle(15000);}catch(error){throw new UncertainMutationError(error.message+' Lagringen må kontrolleres i Adpoint.');}
  assertTarget(record);
  // Remount the editor by leaving and returning to its tab, then read back.
  const summary=[...d.querySelectorAll('#InfoPanelContainer [role="tab"]')]
    .find(tab=>/SUMMARY$/i.test(tab.id)||/^(summary|sammendrag|oversikt)$/i.test(tab.textContent.trim()));
  if(!summary) throw new UncertainMutationError('Lagre ble utført, men kunne ikke åpne sammendrag for etterkontroll. Kontroller Adpoint.');
  summary.click();
  if(!await waitFor(()=>!getThirdPartyPanel()||!vis(getThirdPartyPanel()),4000,90)) throw new UncertainMutationError('Kunne ikke forlate editor for etterkontroll.');
  const readBack=await get3PEditorStrict();
  if(!readBack) throw new UncertainMutationError('Kunne ikke åpne editor igjen etter lagring.');
  assertTarget(record);
  const value=await getEditorValueStable(readBack);
  if(normalizeTagForCompare(value)!==normalizeTagForCompare(expected)) throw new Error('Etterkontrollen fant et annet script enn det som ble lagret.');
  const stillDirty=saveButtons(readBack.panel).filter(btn=>!isDisabled(btn)&&/^(lagre(?: endringer)?|save(?: changes)?|oppdater|update)$/i.test(btn.textContent.trim()));
  if(stillDirty.length&&!successNotice()) throw new UncertainMutationError('Adpoint viser fortsatt ulagrede endringer etter lagring. Kontroller før ny kjøring.');
  return true;
}


/* ========= empty editor detection ========= */
function normalizeEditorText(s){
  return String(s||'')
    .replace(/\u200B/g,'')
    .replace(/\u00A0/g,' ')
    .replace(/\r\n/g,'\n')
    .trim();
}
function isEmpty3PValue(v){ return !normalizeEditorText(v).length; }
async function getEditorValueStable(editor,tries=12,step=120){
  let last=null,stable=0;
  for(let i=0;i<tries;i++){
    if(!editor.el.isConnected) throw new Error('Editor ble byttet mens innholdet ble lest.');
    const value=getEditorValue(editor);
    stable=value===last?stable+1:0; last=value;
    if(stable>=3) return value;
    await sleep(step);
  }
  throw new Error('Editorinnholdet ble ikke stabilt. Ingen endring utført.');
}

/* ========= tile already has material? ========= */
function tileHasMaterial(tile){
  if(!tile) return false;
  const media=tile.querySelector('.MuiCardMedia-root, [class*="CardMedia"]');
  const bg=media && getComputedStyle(media).backgroundImage;
  if(bg && bg!=='none' && /url\(/i.test(bg)) return true;
  const img=tile.querySelector('img');
  if(img && (img.currentSrc || img.src)) return true;
  return false;
}

/* ========= mapping import ========= */
function parseCSV(text){
  let rows=[],row=[],f='',q=false;
  const first=(text.split(/\r?\n/).find(l=>l.trim().length>0)||'');
  const c1=(first.match(/,/g)||[]).length,c2=(first.match(/;/g)||[]).length,c3=(first.match(/\t/g)||[]).length;
  const del=c3>=c2&&c3>=c1?'\t':(c2>c1?';':',');
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(ch=='"'){ if(q&&text[i+1]=='"'){f+='"';i++} else q=!q; }
    else if(!q && ch===del){ row.push(f); f=''; }
    else if(!q && ch=='\n'){ row.push(f); rows.push(row); row=[]; f=''; }
    else if(!q && ch=='\r'){}
    else f+=ch;
  }
  row.push(f); rows.push(row);
  return rows;
}
function pickHeader(h,names){
  const L=h.map(x=>String(x||'').toLowerCase());
  for(const n of names){
    let i=L.indexOf(n); if(i>-1) return i;
    for(let k=0;k<L.length;k++) if(L[k].includes(n)) return k;
  }
  return -1;
}
function rowsToMap(rows){
  let head=[],headRow=0;
  for(let i=0;i<Math.min(rows.length,5);i++){
    const cells=(rows[i]||[]).map(c=>String(c||'').trim());
    if(cells.join('').length){ head=cells; headRow=i; break; }
  }
  const sizeIdx=pickHeader(head,['størrelse','stoerrelse','size']);
  const nameIdx=pickHeader(head,['creative name','creative','name','navn']);
  const tagIdx =pickHeader(head,['secure content','script','tag','kode','code']);
  if(tagIdx<0 || (sizeIdx<0 && nameIdx<0)) throw new Error('Mangler kolonner: Secure Content + (Size eller Creative Name)');

  let warned=false; const out=[];
  for(let r=headRow+1;r<rows.length;r++){
    const cells=rows[r]||[];
    const rawSize=sizeIdx>-1 ? (cells[sizeIdx]||'') : '';
    const rawName=nameIdx>-1 ? (cells[nameIdx]||'') : '';
    const svSize=parseSizeVariant(rawSize);
    const svName=parseSizeVariant(rawName);
    const size=cs(svSize.size || svName.size);
    const variant=svName.variant ?? svSize.variant ?? null;
    const tag=String(cells[tagIdx]||'').trim();
    const name=rawName || '';
    const vkey=versionKey(rawName);

    let lineIds=extractLineIds(rawName);
    if(!lineIds.length){
      const rowText=cells.map(c=>String(c||'')).join(' _ ');
      lineIds=extractLineIds(rowText);
    }

    if(svSize.size && svName.size && cs(svSize.size)!==cs(svName.size) && !warned){
      LOG('! CSV-advarsel: "Size" er ulik "Creative Name" — jeg stoler på Size-kolonnen.');
      warned=true;
    }
    if(size && tag) out.push({size,variant,tag,name,vkey,lineIds});
  }
  return out;
}
function normalizeJSON(arr){
  let warned=false; const out=[];
  (arr||[]).forEach(o=>{
    const rawName=o.name||o.Name||o['Creative Name']||o['creative name']||'';
    const rawSize=o.size||o.Size||o['Størrelse']||o['Stoerrelse']||'';
    const svSize=parseSizeVariant(rawSize), svName=parseSizeVariant(rawName);
    const size=cs(svSize.size || svName.size);
    const variant=(o.variant||o.Variant||svName.variant||svSize.variant)||null;
    const tag=String(o.tag||o.Tag||o['Secure Content']||o['secure content']||o.Script||o.script||'').trim();
    const vkey=versionKey(rawName);

    let lineIds=extractLineIds(rawName);
    if(!lineIds.length){
      const flat=Object.values(o||{}).map(v=>String(v||'')).join(' _ ');
      lineIds=extractLineIds(flat);
    }

    if(svSize.size && svName.size && cs(svSize.size)!==cs(svName.size) && !warned){
      LOG('! JSON-advarsel: "Size" er ulik "name" — jeg stoler på Size-feltet.');
      warned=true;
    }
    if(size && tag) out.push({size,variant,tag,name:rawName,vkey,lineIds});
  });
  return out;
}

/* ========= line items ========= */
function normLabel(s){return String(s||'').toLowerCase().replace(/\s+/g,' ').trim()}
function findRows(){
  const rows=[...d.querySelectorAll('tr.set-matSpecDataRow')].filter(vis);
  return rows.filter(el=>!rows.some(o=>o!==el && o.contains(el)));
}
function getHeaderIdxForRow(row){
  const table=row.closest('table'); if(!table) return {lineItemIdx:-1,rosenrIdx:-1};
  if(table._ap3p_headerIdx) return table._ap3p_headerIdx;
  const ths=[...(table.tHead?.querySelectorAll('th')||[])];
  let lineItemIdx=-1, rosenrIdx=-1;
  ths.forEach((th,i)=>{
    const t=normLabel(th.textContent);
    if(lineItemIdx===-1 && /(^|\s)line\s*item(\s|$)|linje[- ]?(?:id|nummer)/.test(t)) lineItemIdx=i;
    if(rosenrIdx===-1 && (/rosenr|material number|materiellnummer/.test(t))) rosenrIdx=i;
  });
  table._ap3p_headerIdx={lineItemIdx,rosenrIdx};
  return table._ap3p_headerIdx;
}
function rowId(row){
  const {lineItemIdx}=getHeaderIdxForRow(row);
  const cell=row.children[lineItemIdx];
  const value=cell?.textContent.trim()||'';
  return /^\d{6,10}$/.test(value)?value:'?';
}

function findRowsByIdAll(targetId){return findRows().filter(r=>rowId(r)===targetId)}

/* ========= expand/collapse (standardisert) ========= */
function getExpandBtnForDataRow(dataRow){
  if(!dataRow) return null;
  const btns=Array.from(dataRow.querySelectorAll('button')).filter(vis);
  for(const b of btns){
    const p=b.querySelector('svg path');
    const dd=p?.getAttribute?.('d')||'';
    if(dd.startsWith('M7.41 8.59L12 13.17')) return b; // chevron
  }
  return null;
}
function getCollapseRowsForDataRow(dataRow){
  const out=[];
  if(!dataRow) return out;

  let sib=dataRow.nextElementSibling;
  for(let i=0;i<30 && sib;i++){
    // stopp når vi møter neste data-row
    if(sib.matches?.('tr.set-matSpecDataRow')) break;

    // typisk collapse-row inneholder creative containers / creative section
    if(sib.querySelector?.('.set-matSpecCreativeSection, .set-creativeSectionContainer, .set-creativeContainer')){
      out.push(sib);
    }

    sib=sib.nextElementSibling;
  }
  return out;
}

async function ensureLineExpanded(dataRow){
  // hvis allerede expanded: returnér alle collapse rows
  let rows=getCollapseRowsForDataRow(dataRow);
  if(rows.some(row=>vis(row.querySelector('.set-matSpecCreativeSection,.set-creativeSectionContainer')))) return rows;

  const expBtn=getExpandBtnForDataRow(dataRow);
  if(expBtn){
    expBtn.scrollIntoView({block:'center'});
    expBtn.click();
  }

  rows=await waitFor(()=>{
    const r=getCollapseRowsForDataRow(dataRow);
    return r.some(row=>vis(row.querySelector('.set-matSpecCreativeSection,.set-creativeSectionContainer')))?r:null;
  }, 4000, 80);

  return rows || [];
}


/* ========= line-scoped optional placeholders, NO/EN ========= */
const ADD_LABEL=/^(?:legg til valgfri(?:tt)? materiell|add optional creative placeholder)$/i;
function getAddOptionalBtn(scope){
  return [...(scope||d).querySelectorAll('button,[role="button"]')]
    .filter(vis).find(el=>ADD_LABEL.test((el.textContent||'').replace(/\s+/g,' ').trim()))||null;
}
function optionalToggle(btn){
  return btn?.closest('[aria-haspopup]')||btn;
}
function isDisabled(el){
  return !el||el.disabled||el.getAttribute('aria-disabled')==='true'||el.classList.contains('Mui-disabled')||el.hasAttribute('disabled');
}
function dropdownRoot(btn){
  return btn.closest('[class*="dropdown---dropdown---"]')||btn.parentElement;
}
function menuForToggle(btn){
  const controlled=btn.getAttribute('aria-controls');
  const direct=controlled&&d.getElementById(controlled);
  if(direct&&vis(direct)) return direct;
  const local=[...dropdownRoot(btn).querySelectorAll('[role="menu"],[role="listbox"]')].filter(vis);
  return local.length===1?local[0]:null;
}
async function openAddSizeMenuForRow(row){
  const id=typeof row==='string'?row:rowId(row);
  const rows=findRowsByIdAll(id);
  if(rows.length!==1) throw new Error(`Linje ${id}: forventet én entydig rad, fant ${rows.length}.`);
  const details=await ensureLineExpanded(rows[0]);
  const buttons=details.map(getAddOptionalBtn).filter(Boolean);
  if(buttons.length!==1) throw new Error(`Linje ${id}: fant ikke én entydig knapp for valgfritt materiell (norsk/engelsk).`);
  const btn=optionalToggle(buttons[0]);
  if(isDisabled(btn)) throw new Error(`Linje ${id}: knappen for materiell er deaktivert.`);
  if(btn.getAttribute('aria-expanded')==='true'&&menuForToggle(btn)) return {btn,menu:menuForToggle(btn)};
  const before=new Set([...d.querySelectorAll('[role="menu"],[role="listbox"]')].filter(vis));
  btn.scrollIntoView({block:'center'}); btn.click();
  const menu=await waitFor(()=>{
    const owned=menuForToggle(btn); if(owned) return owned;
    // Only a newly opened portal, never the geometrically nearest unrelated menu.
    const fresh=[...d.querySelectorAll('[role="menu"],[role="listbox"]')].filter(el=>vis(el)&&!before.has(el));
    return fresh.length===1?fresh[0]:null;
  },5000,90);
  if(!menu) throw new Error(`Linje ${id}: størrelsesmenyen åpnet seg ikke.`);
  return {btn,menu};
}
function readMenuOptions(menu){
  return [...menu.querySelectorAll('[role="menuitem"],[role="option"]')].filter(vis).map(el=>{
    const label=(el.textContent||'').replace(/\s+/g,' ').trim();
    const meta=parseSizeVariant(label);
    return {el,label,size:cs(meta.size),variant:meta.variant,disabled:isDisabled(el)};
  }).filter(x=>x.size);
}
async function closeAddMenu(opened){
  if(!opened) return;
  const {btn,menu}=opened;
  if(!vis(menu)) return;
  if(btn.isConnected&&btn.getAttribute('aria-expanded')==='true') btn.click();
  else menu.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',code:'Escape',bubbles:true}));
  await sleep(80);
  if(vis(menu)){
    d.body.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));
    d.body.click(); await sleep(80);
  }
}
async function getAllowedAddSizesForLine(row){
  let opened;
  try{
    opened=await openAddSizeMenuForRow(row);
    const id=typeof row==='string'?row:rowId(row);
    const lineVariant=rowVariant(findRowsByIdAll(id)[0]);
    return readMenuOptions(opened.menu).filter(x=>!x.disabled).map(({size,variant,label})=>({size,variant:variant||lineVariant,label}));
  }
  finally{await closeAddMenu(opened);}
}
async function addSizeScopedToLine(line,sizeText,variant=null,exactLabel=null){
  const id=typeof line==='string'?line:rowId(line);
  let opened;
  const before=await recordsForLine(id);
  const previousIds=new Set(before.map(x=>x.placeholder));
  if(before.some(x=>!x.placeholder)) throw new Error('Kan ikke bekrefte eksisterende placeholder-ID-er. Ingen tillegging utført.');
  try{
    opened=await openAddSizeMenuForRow(id);
    let choices=readMenuOptions(opened.menu).filter(x=>x.size===cs(sizeText)&&!x.disabled);
    if(exactLabel) choices=choices.filter(x=>x.label===exactLabel);
    else if(variant){
      const exact=choices.filter(x=>x.variant===variant);
      if(exact.length) choices=exact;
      else choices=choices.filter(x=>!x.variant);
    }
    if(choices.length!==1) throw new Error(`Linje ${id}: størrelsen ${sizeText}${variant?'/'+variant:''} er utilgjengelig eller tvetydig.`);
    choices[0].el.click();
    // Mutation has happened. Never repeat this click after an ambiguous timeout.
    try{await waitForIdle(15000);}catch(error){throw new UncertainMutationError(error.message+' Tilleggingen må kontrolleres i Adpoint.');}
    const start=Date.now();
    while(Date.now()-start<12000){
      const now=await recordsForLine(id);
      const fresh=now.filter(x=>!previousIds.has(x.placeholder));
      if(fresh.length===1&&fresh[0].size===cs(sizeText)&&(!variant||fresh[0].variant===variant)) return fresh[0];
      if(fresh.length>0) throw new Error(`Linje ${id}: endringen ga uventet størrelse eller antall. Kontroller Adpoint før ny kjøring.`);
      await sleep(150);
    }
    throw new UncertainMutationError(`Linje ${id}: tillegging er usikker. Ingen automatisk gjentakelse; kontroller Adpoint.`);
  }finally{await closeAddMenu(opened);}
}


// Script sources are grouped by line and variant in workGroups below.

/* ========= image pool ========= */
let imagePool=new Map(); // key `${size}|${side||''}` -> [File,...]
function poolKey(size,side){return `${size}|${side||''}`}
function parseSizeFromName(name){
  const m=String(name||'').match(/(\d{2,4})[x×](\d{2,4})/i);
  return m?`${m[1]}x${m[2]}`:'';
}
function parseImageMeta(file){
  return {size:cs(parseSizeFromName(file.name)), side:detectSide(file.name), file};
}
function addImagesToPool(files){
  let added=0;
  for(const f of files){
    const meta=parseImageMeta(f);
    if(!meta.size) continue;
    const key=poolKey(meta.size,meta.side);
    if(!imagePool.has(key)) imagePool.set(key,[]);
    imagePool.get(key).push(f);
    added++;
  }
  renderMapChip();
  flash(mapChip); flash(drop);
  showToast('Bilder lagt til',`Antall: ${added}`);
}
function imageChoicesForEntry(entry){
  const exact=imagePool.get(`${entry.size}|${entry.side||''}`)||[];
  if(exact.length) return exact;
  if(entry.side){
    const any=imagePool.get(`${entry.size}|`)||[];
    if(any.length) return any;
  }
  return [];
}
function fileName(f){return (f&&f.name)?f.name:'(fil)'}

/* ========= upload image to tile ========= */
function getTileDropzoneInput(tile){return tile.querySelector('input[type="file"]')||null}
async function confirmReplaceIfPrompted(timeout=1200){
  const dlg=await waitFor(()=>[...d.querySelectorAll('[role="dialog"]')].filter(vis)
    .find(el=>/erstatte|erstatt|overwrite|replace/i.test(el.textContent||'')),timeout,100);
  if(!dlg) return false;
  const buttons=[...dlg.querySelectorAll('button')].filter(vis)
    .filter(btn=>!isDisabled(btn)&&/^(ok|ja|yes|fortsett|continue|erstatte|erstatt|replace|confirm|bekreft)$/i.test(btn.textContent.trim()));
  if(buttons.length!==1) throw new Error('Ukjent bekreftelsesdialog. Ingen knapp ble valgt automatisk.');
  buttons[0].click(); return true;
}
function materialSignature(tile){
  const media=[...tile.querySelectorAll('img,.MuiCardMedia-root,[class*="CardMedia"]')]
    .map(el=>el.currentSrc||el.src||getComputedStyle(el).backgroundImage||'').join('|');
  return media;
}
async function uploadImageToTile(record,file){
  const current=await getFreshRecord(record.line,record.placeholder);
  if(!current) throw new Error('Materiellplassen for bildet finnes ikke lenger.');
  if(!await selectTile(current.tile)||!await ensurePanelSwitchedToTile(current.tile)) throw new Error('Kunne ikke velge riktig materiellplass for bildet.');
  assertTarget(current);
  const inp=getTileDropzoneInput(current.tile);
  if(!inp) throw new Error('Fant ikke filopplasting på valgt materiellplass.');
  const before=materialSignature(current.tile);
  const dt=new DataTransfer(); dt.items.add(file); inp.files=dt.files;
  // Dropzone handles change; dispatching both input and change can start two uploads.
  inp.dispatchEvent(new Event('change',{bubbles:true}));
  await confirmReplaceIfPrompted(1800);
  try{await waitForIdle(20000);}catch(error){throw new UncertainMutationError(error.message+' Bildeopplastingen må kontrolleres i Adpoint.');}
  const start=Date.now();
  while(Date.now()-start<20000){
    const error=visibleAdpointError();
    if(error) throw new Error('Bildeopplasting feilet: '+error.textContent.trim().slice(0,200));
    const now=await getFreshRecord(record.line,record.placeholder);
    if(now&&tileHasMaterial(now.tile)&&materialSignature(now.tile)!==before) return true;
    await sleep(180);
  }
  throw new UncertainMutationError('Bildeopplastingen er usikker: ny forhåndsvisning ble ikke bekreftet. Kontroller Adpoint.');
}


/* ========= UI ========= */
(function injectCSS(){
  d.getElementById('ap3p_style')?.remove();
  const st=d.createElement('style'); st.id='ap3p_style';
  st.textContent=`
#ap3p_bar{box-sizing:border-box}
#ap3p_list .item{display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:6px}
#ap3p_list .item:hover{background:#121521}
#ap3p_list .id{min-width:72px;opacity:.85}
#ap3p_list .sizes{opacity:.9;font-family:ui-monospace,Consolas,monospace}
#ap3p_list .muted{opacity:.6}
#ap3p_bar .chip{padding:3px 8px;border-radius:999px;border:1px solid;font-size:12px;opacity:.95}
#ap3p_bar .chip-ok{background:#064e3b;border-color:#10b981;color:#d1fae5}
#ap3p_bar .chip-none{background:#1a1d27;border-color:#2a2d37;color:#e6e6e6}
@keyframes ap3pPulse {0%{transform:scale(1);filter:brightness(1)}30%{transform:scale(1.02);filter:brightness(1.25)}100%{transform:scale(1);filter:brightness(1)}}
.ap3p-flash{animation: ap3pPulse 550ms ease-out 1;}
#ap3p_toast{
  position:fixed; right:18px; bottom:18px; z-index:2147483648;
  background:#0b0c10; border:1px solid #2a2d37; color:#e6e6e6;
  padding:10px 12px; border-radius:10px; box-shadow:0 10px 30px rgba(0,0,0,.45);
  max-width:420px; font:12px/1.35 ui-sans-serif,system-ui; opacity:0; transform:translateY(6px);
  transition: opacity .18s ease, transform .18s ease;
}
#ap3p_toast.show{opacity:1; transform:translateY(0px);}
#ap3p_toast .t{font-weight:700;margin-bottom:4px}
#ap3p_toast .s{opacity:.85;white-space:pre-wrap}
#ap3p_bar .aprs{position:absolute;z-index:2147483648;background:transparent}
#ap3p_bar .aprs-n{top:-4px;left:10px;right:10px;height:8px;cursor:ns-resize}
#ap3p_bar .aprs-s{bottom:-4px;left:10px;right:10px;height:8px;cursor:ns-resize}
#ap3p_bar .aprs-e{top:10px;right:-4px;bottom:10px;width:8px;cursor:ew-resize}
#ap3p_bar .aprs-w{top:10px;left:-4px;bottom:10px;width:8px;cursor:ew-resize}
#ap3p_bar .aprs-ne,#ap3p_bar .aprs-nw,#ap3p_bar .aprs-se,#ap3p_bar .aprs-sw{width:12px;height:12px}
#ap3p_bar .aprs-ne{top:-6px;right:-6px;cursor:nesw-resize}
#ap3p_bar .aprs-nw{top:-6px;left:-6px;cursor:nwse-resize}
#ap3p_bar .aprs-se{bottom:-6px;right:-6px;cursor:nwse-resize}
#ap3p_bar .aprs-sw{bottom:-6px;left:-6px;cursor:nesw-resize}
`;
  d.head.appendChild(st);
})();

const UI_STATE_KEY='ap3p_ui_state_v2';
function loadUIState(){const st=G(UI_STATE_KEY,null); return (st&&typeof st==='object')?st:null;}
function saveUIState(){
  try{
    const r=ui.getBoundingClientRect();
    const st={
      left:Math.round(r.left),
      top: Math.round(r.top),
      w:   Math.round(r.width),
      h:   Math.round(r.height),
      min: ui.getAttribute('data-min')==='1'
    };
    S(UI_STATE_KEY,st);
  }catch{}
}

let ui=d.getElementById('ap3p_bar'); if(ui) ui.remove();
ui=d.createElement('div'); ui.id='ap3p_bar';
ui.style.cssText='position:fixed;right:16px;bottom:16px;z-index:2147483647;background:#0e0f13;color:#e6e6e6;border:1px solid #2a2d37;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.45);max-width:96vw;min-width:420px;max-height:90vh;';

const hdr=d.createElement('div');
hdr.style.cssText='cursor:move;user-select:none;display:flex;align-items:center;gap:10px;padding:8px 10px;background:#161922;border-radius:12px 12px 0 0;border-bottom:1px solid #2a2d37';
const title=d.createElement('div'); title.textContent='EasyPoint 2.0';
const badge=d.createElement('span'); badge.style.opacity='.8'; badge.style.marginLeft='6px'; badge.textContent='';
const mapChip=d.createElement('span'); mapChip.className='chip chip-none';
const mapClear=d.createElement('button'); mapClear.textContent='×'; mapClear.title='Tøm mapping'; mapClear.style.cssText='margin-left:4px;border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:22px;height:22px;border-radius:6px;cursor:pointer';
const flex=d.createElement('div'); flex.style.flex='1';

const btnScan=d.createElement('button'); btnScan.textContent='Skann'; btnScan.style.cssText='cursor:pointer;border:1px solid #0284c7;border-radius:8px;padding:6px 10px;background:#0ea5e9;color:#041014;font-weight:600';
const btnRun=d.createElement('button'); btnRun.textContent='Autofyll'; btnRun.style.cssText='cursor:pointer;border:1px solid #15803d;border-radius:8px;padding:6px 10px;background:#22c55e;color:#051b0a;font-weight:700';
const btnStop=d.createElement('button'); btnStop.textContent='Stopp'; btnStop.disabled=true; btnStop.style.cssText='cursor:pointer;border:1px solid #b91c1c;border-radius:8px;padding:6px 10px;background:#ef4444;color:#ffffff;font-weight:700';
const btnMin=d.createElement('button'); btnMin.textContent='–'; btnMin.title='Minimer'; btnMin.style.cssText='border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:26px;height:26px;border-radius:7px;margin-left:6px;cursor:pointer';
const btnX=d.createElement('button'); btnX.textContent='×'; btnX.title='Lukk'; btnX.style.cssText='border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:26px;height:26px;border-radius:7px;cursor:pointer';
hdr.append(title,badge,mapChip,mapClear,flex,btnScan,btnRun,btnStop,btnMin,btnX);

const body=d.createElement('div');
body.style.cssText='padding:8px 10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap';
const info=d.createElement('span'); info.textContent='Linjer: —'; info.style.opacity='.9';
const btnCSV=d.createElement('button'); btnCSV.textContent='Importer CSV/JSON'; btnCSV.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:8px;padding:6px 10px;background:#1a1d27;color:#e6e6e6';
const btnPaste=d.createElement('button'); btnPaste.textContent='Lim inn fra Excel'; btnPaste.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:8px;padding:6px 10px;background:#1a1d27;color:#e6e6e6';
const file=d.createElement('input'); file.type='file'; file.accept='.csv,.json,.txt,image/*,.psd,.tif,.tiff,.webp'; file.multiple=true; file.style.display='none';
const drop=d.createElement('div'); drop.textContent='Slipp CSV/JSON/bilder her'; drop.style.cssText='flex:1 1 100%;border:1px dashed #3a3f52;border-radius:8px;height:56px;display:flex;align-items:center;justify-content:center;opacity:.9';
const listWrap=d.createElement('div'); listWrap.style.cssText='flex:1 1 100%';
const listCtrls=d.createElement('div'); listCtrls.id='ap3p_list_ctrls'; listCtrls.style.cssText='display:flex;align-items:center;gap:8px;margin:2px 0 4px 0;opacity:.9';
const btnAll=d.createElement('button'); btnAll.textContent='Velg alle';
const btnNone=d.createElement('button'); btnNone.textContent='Velg ingen';
const btnInv=d.createElement('button'); btnInv.textContent='Inverter';
[btnAll,btnNone,btnInv].forEach(b=>b.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:6px;padding:4px 8px;background:#1a1d27;color:#e6e6e6');
listCtrls.append(btnAll,btnNone,btnInv);
const list=d.createElement('div'); list.id='ap3p_list'; list.style.cssText='max-height:160px;overflow:auto;background:#0b0c10;border:1px solid #1b1e28;border-radius:8px;padding:6px';
listWrap.append(listCtrls,list);
const log=d.createElement('div'); log.style.cssText='height:200px;overflow:auto;background:#0b0c10;border-top:1px solid #1b1e28;border-radius:0 0 12px 12px;padding:8px 10px;font-family:ui-monospace,Consolas,monospace;white-space:pre-wrap';
ui.append(hdr,body,listWrap,log); d.body.appendChild(ui);

const LOG=m=>{log.textContent+=m+'\n'; log.scrollTop=log.scrollHeight;};
const setInfo=o=>{app.lastStats={...o};info.textContent=`Linjer:${o.items}  Ferdig:${o.done}  Lagret:${o.saved||0}  Nye:${o.added||0}  Hoppet:${o.skip||0}  Feil:${o.err||0}  Usikkert:${o.uncertain||0}`;refreshSimpleUI();};
body.append(info,btnCSV,btnPaste,file,drop);

/* ===== Apply saved UI state ===== */
(function applyInitialUIState(){
  const st=loadUIState();
  const defW=560, defH=660;
  const vw=Math.max(420, Math.min(defW, Math.floor(window.innerWidth*0.92)));
  const vh=Math.max(260, Math.min(defH, Math.floor(window.innerHeight*0.85)));

  if(st?.w && st?.h){
    const ww=clamp(st.w,420,Math.floor(window.innerWidth*0.96));
    const hh=clamp(st.h,160,Math.floor(window.innerHeight*0.90));
    ui.style.width=ww+'px'; ui.style.height=hh+'px';
  }else{
    ui.style.width=vw+'px'; ui.style.height=vh+'px';
  }

  if(typeof st?.left==='number' && typeof st?.top==='number'){
    const ww=parseInt(ui.style.width,10)||ui.getBoundingClientRect().width;
    const hh=parseInt(ui.style.height,10)||ui.getBoundingClientRect().height;
    ui.style.left=clamp(st.left,6,Math.floor(window.innerWidth-ww-6))+'px';
    ui.style.top =clamp(st.top, 6,Math.floor(window.innerHeight-hh-6))+'px';
    ui.style.right='auto'; ui.style.bottom='auto';
  }
})();

/* ===== toast + flash ===== */
let toast=d.getElementById('ap3p_toast');
if(!toast){
  toast=d.createElement('div');
  toast.id='ap3p_toast';
  toast.innerHTML=`<div class="t"></div><div class="s"></div>`;
  d.body.appendChild(toast);
}
let toastTimer=null;
function showToast(t,s=''){
  try{
    toast.querySelector('.t').textContent=t||'';
    toast.querySelector('.s').textContent=s||'';
    toast.classList.add('show');
    if(toastTimer) clearTimeout(toastTimer);
    toastTimer=setTimeout(()=>toast.classList.remove('show'),2000);
  }catch{}
}
function flash(el){
  if(!el) return;
  el.classList.remove('ap3p-flash');
  void el.offsetWidth;
  el.classList.add('ap3p-flash');
}

/* ===== resize/drag/minimize ===== */
function adjustHeights(){
  if(ui.getAttribute('data-min')==='1')return;
  ui.style.overflowY='auto';ui.style.overflowX='hidden';
}
(function makeDraggable(handle,box){
  let sx=0,sy=0,ox=0,oy=0,drag=false;
  handle.addEventListener('pointerdown',e=>{
    if(e.button!==0 || e.target.closest('button,input,select')) return;
    drag=true;
    const r=box.getBoundingClientRect();
    ox=r.left; oy=r.top; sx=e.clientX; sy=e.clientY;
    box.style.left=ox+'px'; box.style.top=oy+'px';
    box.style.right='auto'; box.style.bottom='auto';
  });
  listen(w,'pointermove',e=>{
    if(!drag)return;
    box.style.left=(ox+e.clientX-sx)+'px';
    box.style.top =(oy+e.clientY-sy)+'px';
  });
  listen(w,'pointerup',()=>{
    if(drag){ drag=false; saveUIState(); }
  });
})(hdr,ui);

(function makeResizable(box){
  const dirs=['n','s','e','w','ne','nw','se','sw']; const els={};
  dirs.forEach(dn=>{const h=d.createElement('div');h.className='aprs aprs-'+dn;box.appendChild(h);els[dn]=h;});
  let rs=null;
  function onDown(e,dir){
    if(e.button!==0)return;
    e.preventDefault();
    const r=box.getBoundingClientRect();
    rs={dir,sx:e.clientX,sy:e.clientY,x:r.left,y:r.top,w:r.width,h:r.height};
    listen(w,'pointermove',onMove);
    listen(w,'pointerup',onUp);
  }
  function onMove(e){
    if(!rs)return;
    const dx=e.clientX-rs.sx, dy=e.clientY-rs.sy; let x=rs.x,y=rs.y,wv=rs.w,hv=rs.h;
    if(/e/.test(rs.dir)) wv=Math.max(420,rs.w+dx);
    if(/s/.test(rs.dir)) hv=Math.max(180,rs.h+dy);
    if(/w/.test(rs.dir)){ wv=Math.max(420,rs.w-dx); x=rs.x+dx; }
    if(/n/.test(rs.dir)){ hv=Math.max(180,rs.h-dy); y=rs.y+dy; }
    Object.assign(box.style,{width:wv+'px',height:hv+'px',left:x+'px',top:y+'px',right:'auto',bottom:'auto'});
    adjustHeights(hv);
  }
  function onUp(){
    if(!rs) return;
    rs=null;
    w.removeEventListener('pointermove',onMove);
    w.removeEventListener('pointerup',onUp);
    saveUIState();
  }
  dirs.forEach(dn=>els[dn].addEventListener('pointerdown',e=>onDown(e,dn)));
  box._resizerEls=Object.values(els);
})(ui);

function setMinimized(min){
  const isMin=ui.getAttribute('data-min')==='1';
  if(min){
    if(!isMin){
      try{
        const r=ui.getBoundingClientRect();
        ui.setAttribute('data-last-h',String(Math.round(r.height)));
        ui.setAttribute('data-last-w',String(Math.round(r.width)));
      }catch{}
    }
    ui.setAttribute('data-min','1');
    body.style.display='none';
    listWrap.style.display='none';
    log.style.display='none';
    ui.style.height='38px';
    ui.style.overflow='hidden';
    (ui._resizerEls||[]).forEach(el=>el.style.display='none');
    btnMin.textContent='+';
    saveUIState();
    return;
  }

  ui.setAttribute('data-min','0');
  body.style.display='flex';
  listWrap.style.display='block';
  log.style.display='block';
  (ui._resizerEls||[]).forEach(el=>el.style.display='');
  btnMin.textContent='–';

  const st=loadUIState()||{};
  const lastH=parseInt(ui.getAttribute('data-last-h')||'',10);
  const lastW=parseInt(ui.getAttribute('data-last-w')||'',10);

  const targetH=clamp((Number.isFinite(lastH)&&lastH>80)?lastH:(st.h||660),180,Math.floor(window.innerHeight*0.90));
  const targetW=clamp((Number.isFinite(lastW)&&lastW>200)?lastW:(st.w||560),420,Math.floor(window.innerWidth*0.96));

  ui.style.height=targetH+'px';
  ui.style.width =targetW+'px';
  adjustHeights(targetH);
  saveUIState();
}
btnMin.onclick=()=>setMinimized(ui.getAttribute('data-min')!=='1');

let _bestiltInt=null;
app.destroy=()=>{
  if(app.busy){ app.closing=true; stopping=true; showToast('Lukker når pågående steg er ferdig'); return; }
  life.abort(); clearInterval(_bestiltInt); clearTimeout(toastTimer);
  ui.remove(); toast?.remove(); d.getElementById('ap3p_style')?.remove();
  if(w.__easyPoint===app) delete w.__easyPoint;
};
btnX.onclick=()=>app.destroy();


adjustHeights(ui.getBoundingClientRect().height);

/* ========= mapping state ========= */
const MAP_KEY='ap3p_map_v2';
let mapping=G(MAP_KEY,[]); if(!Array.isArray(mapping)) mapping=[];
function renderMapChip(){
  const csvCount=mapping.length;
  let imgCount=0; imagePool.forEach(a=>imgCount+=a.length);
  mapChip.className=(csvCount||imgCount)?'chip chip-ok':'chip chip-none';
  mapChip.textContent=(csvCount||imgCount)?`CSV: ${csvCount}  Bilder: ${imgCount} ✓`:'CSV/Bilder: ingen';
  refreshSimpleUI();
}
function saveMap(arr){
  mapping=arr;
  S(MAP_KEY,mapping);
  renderMapChip();
  flash(mapChip); flash(drop);
  showToast('Importert',`CSV-rader: ${mapping.length}`);
}
function clearMap(){
  mapping=[];
  S(MAP_KEY,mapping);
  renderMapChip();
  flash(mapChip);
  showToast('Importen er tømt');
}
mapClear.onclick=()=>{clearMap();imagePool.clear();renderMapChip()};
renderMapChip();

/* ========= toggles ========= */
const REUSE_KEY_NEW='ap3p_reuse_assets';
const REUSE_KEY_OLD='ap3p_reuse_images';
let reuseAssets=G(REUSE_KEY_NEW, G(REUSE_KEY_OLD, true)); // default ON

const reuseWrap=d.createElement('label');
reuseWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const reuseCb=d.createElement('input'); reuseCb.type='checkbox'; reuseCb.checked=reuseAssets;
const reuseTxt=d.createElement('span');
reuseWrap.append(reuseCb,reuseTxt);
body.insertBefore(reuseWrap,drop);

const ADD_KEY='ap3p_add_extra_sizes';
let addExtraSizes=!!G(ADD_KEY,false);

const addWrap=d.createElement('label');
addWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const addCb=d.createElement('input'); addCb.type='checkbox'; addCb.checked=addExtraSizes;
const addTxt=d.createElement('span'); addTxt.textContent='Opprett manglende materiellplasser ved autofyll';
addWrap.append(addCb,addTxt);
body.insertBefore(addWrap,drop);

const OVERWRITE_KEY='ap3p_overwrite_mode'; // 'all' | 'empty'
let overwriteMode=G(OVERWRITE_KEY,'empty');

const owWrap=d.createElement('label');
owWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const owCb=d.createElement('input'); owCb.type='checkbox'; owCb.checked=(overwriteMode==='all');
const owTxt=d.createElement('span'); owTxt.textContent='Erstatt eksisterende materiell';
owWrap.append(owCb,owTxt);
body.insertBefore(owWrap,drop);

let _lastPreview=null;
function previewPlannedPlacements(){if(_lastPreview) renderPreview(_lastPreview);}

function updateReuseLabel(){
  if(addExtraSizes){
    reuseTxt.textContent='Gjenbruk';
  }else{
    reuseTxt.textContent='Gjenbruk';
  }
}
const REUSE_LAST_KEY='ap3p_reuse_assets_last';

function syncReuseWithAddExtra(){
  if(addExtraSizes){
    if(!reuseCb.disabled) S(REUSE_LAST_KEY, reuseAssets); // lagre kun første gang
    reuseAssets=false;
    reuseCb.checked=false;
    reuseCb.disabled=true;
    S(REUSE_KEY_NEW,reuseAssets);
  } else {
    reuseCb.disabled=false;
    reuseAssets=!!G(REUSE_LAST_KEY, true);
    reuseCb.checked=reuseAssets;
    S(REUSE_KEY_NEW,reuseAssets);
  }
  updateReuseLabel();
}


reuseCb.onchange=()=>{
  reuseAssets=reuseCb.checked;
  S(REUSE_KEY_NEW,reuseAssets);
  showToast('Gjenbruk',reuseAssets?'På (fyll alle størrelser)':'Av (bruk hver rad én gang)');
  if(_lastPreview){ log.textContent=''; previewPlannedPlacements(_lastPreview); }
};
addCb.onchange=()=>{
  addExtraSizes=addCb.checked;
  S(ADD_KEY,addExtraSizes);
  syncReuseWithAddExtra();
  showToast('Legg til ekstra',addExtraSizes?'På (kan legge til størrelser)':'Av');
  if(_lastPreview){ log.textContent=''; previewPlannedPlacements(_lastPreview); }
};
owCb.onchange=()=>{
  overwriteMode=owCb.checked?'all':'empty';
  S(OVERWRITE_KEY,overwriteMode);
  showToast('Erstatt-modus',overwriteMode==='all'?'Erstatt alle':'Kun tomme størrelser');
  if(_lastPreview){ log.textContent=''; previewPlannedPlacements(_lastPreview); }
};
syncReuseWithAddExtra();

/* ========= import UI ========= */
btnCSV.onclick=()=>file.click();
file.onchange=async e=>{
  const files=[...(e.target.files||[])];file.value='';await importFiles(files);
};
btnPaste.onclick=()=>{
  const go=t=>{try{saveMap(rowsToMap(parseCSV(t)))}catch(e){alert(e.message)}};
  if(navigator.clipboard?.readText){
    navigator.clipboard.readText().then(go).catch(()=>{
      const t=prompt('Lim inn rader (Excel: Creative Name/Size + Secure Content)');
      if(t) go(t);
    });
  }else{
    const t=prompt('Lim inn rader (Excel: Creative Name/Size + Secure Content)');
    if(t) go(t);
  }
};
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='#171a24';flash(drop)}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='transparent'}));
function isCSVFile(f){return /\.csv$/i.test(f.name)||f.type==='text/csv'}
function isJSONFile(f){return /\.json$/i.test(f.name)||f.type==='application/json'}
function readFileText(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result||''));r.onerror=rej;r.readAsText(file,'utf-8')})}
async function importFiles(files){
  if(app.busy||!files.length)return;
  const images=files.filter(f=>/^image\//.test(f.type)||/\.(psd|tiff?|webp|png|jpe?g|gif|bmp|svg)$/i.test(f.name));
  const data=files.filter(f=>isCSVFile(f)||isJSONFile(f)||/\.txt$/i.test(f.name));
  try{
    const entries=[];
    for(const f of data){const text=await readFileText(f);entries.push(...(isJSONFile(f)?normalizeJSON(JSON.parse(text)):rowsToMap(parseCSV(text))));}
    if(app.busy)return;
    if(entries.length)saveMap(entries);
    if(images.length)addImagesToPool(images);
    if(!data.length&&!images.length)showToast('Ukjent filtype','Velg CSV, JSON eller bildefiler.');
    refreshSimpleUI();
  }catch(error){showToast('Kunne ikke importere',error.message);LOG('! Import feilet: '+error.message);}
}
drop.addEventListener('drop',async e=>{e.preventDefault();drop.style.background='transparent';await importFiles([...(e.dataTransfer?.files||[])]);});

/* ========= shared queue, scan, bulk placeholders and autofill ========= */
let scanned=[];
let selected=new Set(Array.isArray(G('ap3p_sel_ids',[]))?G('ap3p_sel_ids',[]):[]);
let stopping=false;
function tagLabel(entry){return typeof entry==='string'?entry.slice(0,80):(entry?.name||'(uten navn)');}

function saveSelection(){S('ap3p_sel_ids',[...selected]);}
function refreshList(){
  refreshSimpleUI();
  list.replaceChildren();
  if(!scanned.length){const empty=d.createElement('div');empty.className='muted';empty.textContent='Trykk Skann for å hente linjer og tilgjengelige størrelser.';list.append(empty);return;}
  for(const item of scanned){
    const row=d.createElement('label');row.className='item';
    const cb=d.createElement('input');cb.type='checkbox';cb.checked=selected.has(item.id);cb.disabled=app.busy;
    const id=d.createElement('span');id.className='id';id.textContent=item.id;
    const sizes=d.createElement('span');sizes.className='sizes';sizes.textContent=item.label;
    row.append(cb,id,sizes);list.append(row);
    cb.onchange=()=>{if(cb.checked)selected.add(item.id);else selected.delete(item.id);saveSelection();refreshSimpleUI();};
  }
}
btnAll.onclick=()=>{if(app.busy)return;selected=new Set(scanned.map(x=>x.id));saveSelection();refreshList();};
btnNone.onclick=()=>{if(app.busy)return;selected.clear();saveSelection();refreshList();};
btnInv.onclick=()=>{if(app.busy)return;selected=new Set(scanned.filter(x=>!selected.has(x.id)).map(x=>x.id));saveSelection();refreshList();};

const bulk=d.createElement('div');bulk.style.cssText='flex:1 1 100%;border:1px solid #3a3f52;border-radius:8px;padding:8px;display:flex;gap:7px;align-items:center;flex-wrap:wrap';
const bulkLabel=d.createElement('strong');bulkLabel.textContent='Nye materiellplasser';bulkLabel.style.fontSize='12px';
const bulkRows=d.createElement('div');bulkRows.style.cssText='flex:1 1 100%;display:grid;gap:5px';
const bulkPlus=d.createElement('button');bulkPlus.textContent='+ Størrelse';
const btnBulk=d.createElement('button');btnBulk.textContent='Legg til på valgte linjer';
const bulkNote=d.createElement('span');bulkNote.textContent='Antall er nye plasser per valgt linje. Kun størrelser linjen tilbyr blir brukt.';bulkNote.style.cssText='font-size:11px;opacity:.8';
[bulkPlus,btnBulk].forEach(btn=>btn.style.cssText='cursor:pointer;border:1px solid #0284c7;border-radius:6px;padding:5px 8px;background:#1a1d27;color:#e6e6e6');
bulk.append(bulkLabel,bulkPlus,btnBulk,bulkNote,bulkRows);body.insertBefore(bulk,drop);
const bulkControls=[];
function availableBulkOptions(){
  const map=new Map();
  for(const line of scanned)for(const option of line.allowed||[]){
    const key=option.size+'|'+(option.variant||'')+'|'+option.label;
    if(!map.has(key))map.set(key,{...option,key});
  }
  return [...map.values()];
}
function refreshBulkOptions(){
  const options=availableBulkOptions();
  for(const control of bulkControls){
    const chosen=control.select.value;control.select.replaceChildren();
    for(const option of options){const el=d.createElement('option');el.value=option.key;el.textContent=option.label+(option.variant&&!parseSizeVariant(option.label).variant?' ('+option.variant+')':'');control.select.append(el);}
    if(options.some(x=>x.key===chosen))control.select.value=chosen;
    if(!options.length){const el=d.createElement('option');el.textContent='Skann for å hente størrelser';el.value='';control.select.append(el);}
  }
}
function addBulkRow(){
  if(app.busy)return;
  const row=d.createElement('div');row.style.cssText='display:flex;gap:6px;align-items:center';
  const select=d.createElement('select');select.setAttribute('aria-label','Størrelse for nye materiellplasser');select.style.cssText='flex:1;min-width:140px;background:#161922;color:#eee;border:1px solid #3a3f52;padding:4px;border-radius:5px';
  const count=d.createElement('input');count.type='number';count.min='1';count.step='1';count.value='1';count.setAttribute('aria-label','Antall nye materiellplasser per linje');count.style.cssText='width:68px;background:#161922;color:#eee;border:1px solid #3a3f52;padding:4px;border-radius:5px';
  const remove=d.createElement('button');remove.textContent='×';remove.title='Fjern størrelse';
  const control={row,select,count,remove};bulkControls.push(control);
  remove.onclick=()=>{if(app.busy)return;row.remove();bulkControls.splice(bulkControls.indexOf(control),1);adjustHeights(ui.getBoundingClientRect().height);};
  const quantity=d.createElement('label');quantity.style.cssText='display:flex;align-items:center;gap:5px;font-size:12px;color:#a9b3c4';const quantityText=d.createElement('span');quantityText.textContent='Antall';quantity.append(quantityText,count);remove.setAttribute('aria-label','Fjern størrelse');row.append(select,quantity,remove);bulkRows.append(row);refreshBulkOptions();adjustHeights(ui.getBoundingClientRect().height);
}
bulkPlus.onclick=addBulkRow;addBulkRow();

function setBusy(value){
  app.busy=value;
  const controls=[btnScan,btnRun,btnBulk,bulkPlus,btnCSV,btnPaste,mapClear,reuseCb,addCb,owCb,btnAll,btnNone,btnInv,
    ...bulkControls.flatMap(x=>[x.select,x.count,x.remove])];
  for(const control of controls)control.disabled=value;
  if(!value){reuseCb.disabled=addExtraSizes;btnBulk.disabled=!scanned.length;}
  btnStop.disabled=!value;
  drop.style.pointerEvents=value?'none':'';
  refreshList();
}
async function withJob(label,fn){
  if(app.busy)return;
  stopping=false;app.jobLabel=label;app.lastError='';app.lastStats=null;setBusy(true);showToast(label);
  try{await fn();}
  catch(error){app.lastError=error?.message||String(error);LOG('! '+app.lastError);showToast('Jobben er stoppet',app.lastError);}
  finally{setBusy(false);if(app.closing)app.destroy();}
}
function requireRunning(){if(stopping)throw new Error('Stoppet av bruker. Pågående endring er avsluttet.');}
btnStop.onclick=()=>{
  if(!app.busy)return;
  stopping=true;btnStop.disabled=true;refreshSimpleUI();
  showToast('Stopper…','Venter på at pågående tillegging eller lagring blir avsluttet.');
};

async function recordsForLine(id){
  if(!await ensureCampaignView())throw new Error('Fant ikke kampanjetabellen.');
  const rows=findRowsByIdAll(String(id));
  if(rows.length!==1)throw new Error(`Linje ${id}: fant ${rows.length} rader; kan ikke velge entydig.`);
  const details=await ensureLineExpanded(rows[0]);
  if(!details.length)throw new Error(`Linje ${id}: kunne ikke åpne materiellområdet.`);
  const records=[],seen=new Set();
  for(const detail of details){
    await ensureAllTilesMounted(detail);
    for(const entry of sizesFromExpanded(detail)){
      const placeholder=placeholderId(entry.tile);
      if(!placeholder)throw new Error(`Linje ${id}: fant en materiellplass uten lesbar placeholder-ID.`);
      if(seen.has(placeholder))continue;seen.add(placeholder);
      records.push({...entry,line:String(id),placeholder});
    }
  }
  return records;
}
async function getFreshRecord(line,placeholder){return (await recordsForLine(line)).find(x=>x.placeholder===placeholder)||null;}
function assertTarget(record){
  const identity=panelIdentity();
  if(identity.placeholder!==record.placeholder || (identity.line&&identity.line!==record.line))throw new Error(`Valgt materiell samsvarer ikke med linje ${record.line}, placeholder ${record.placeholder}. Ingen videre skriving.`);
  const selectedTiles=[...d.querySelectorAll('.set-creativeTile__selected,.set-creativeTile[aria-selected="true"]')].filter(vis);
  if(selectedTiles.length!==1||placeholderId(selectedTiles[0])!==record.placeholder)throw new Error('Riktig materiellplass er ikke entydig valgt.');
}
async function editorForRecord(record){
  const fresh=await getFreshRecord(record.line,record.placeholder);
  if(!fresh||!await selectTile(fresh.tile)||!await ensurePanelSwitchedToTile(fresh.tile))throw new Error(`Kunne ikke bekrefte valgt materiellplass ${record.placeholder}.`);
  const editor=await get3PEditorStrict();
  if(!editor)throw new Error(`Fant ikke 3rd party-editor på materiellplass ${record.placeholder}.`);
  assertTarget(record);return editor;
}
async function inspectRecord(record){
  const editor=await editorForRecord(record);
  const value=await getEditorValueStable(editor);
  assertTarget(record);
  const fresh=await getFreshRecord(record.line,record.placeholder);
  return {record:fresh,value,occupied:!isEmpty3PValue(value)||tileHasMaterial(fresh.tile)};
}
function sameTag(a,b){return normalizeTagForCompare(a)===normalizeTagForCompare(b);}
function takeExistingMatches(assets,states,kind){
  const remaining=assets.slice();let matched=0;
  if(kind==='script')for(const state of states){
    if(isEmpty3PValue(state.value))continue;
    const index=remaining.findIndex(item=>sameTag(item.tag,state.value));
    if(index!==-1){remaining.splice(index,1);matched++;}
  }
  return {remaining,matched};
}
function lineScriptSpecs(id,map){
  const eligible=map.filter(item=>item.size&&item.tag&&(!(item.lineIds||[]).length||(item.lineIds||[]).map(String).includes(id)));
  const keys=[...new Set(eligible.map(item=>`${cs(item.size)}|${item.variant||''}`))];
  return keys.map(key=>{
    const [size,variant]=key.split('|');
    const exact=eligible.filter(item=>cs(item.size)===size&&(item.variant||'')===variant);
    const scoped=exact.filter(item=>(item.lineIds||[]).map(String).includes(id));
    return {kind:'script',size,variant:variant||null,side:null,assets:scoped.length?scoped:exact.filter(item=>!(item.lineIds||[]).length)};
  }).sort((a,b)=>Number(!!b.variant)-Number(!!a.variant));
}
function workGroups(id,records,map,images){
  const scripts=lineScriptSpecs(id,map);
  const imageSpecs=[...images].map(([key,assets])=>{const [size,side]=key.split('|');return {kind:'image',size,side:side||null,variant:null,assets};})
    .sort((a,b)=>Number(!!b.side)-Number(!!a.side));
  const used=new Set(),groups=[];
  for(const spec of [...imageSpecs,...scripts]){
    // Retain the original image-first rule for a size supplied as images.
    if(spec.kind==='script'&&imageSpecs.some(image=>image.size===spec.size))continue;
    const matching=records.filter(record=>record.size===spec.size&&(!spec.variant||record.variant===spec.variant)&&(!spec.side||record.side===spec.side)&&!used.has(record.placeholder));
    matching.forEach(record=>used.add(record.placeholder));
    groups.push({...spec,records:matching});
  }
  return groups;
}
function optionForGroup(allowed,group){
  let options=allowed.filter(option=>option.size===group.size);
  if(group.variant){const exact=options.filter(x=>x.variant===group.variant);options=exact.length?exact:options.filter(x=>!x.variant);}
  if(group.side)options=options.filter(x=>detectSide(x.label)===group.side);
  return options.length===1?options[0]:null;
}
function renderPreview(byId){
  LOG(`Plan: ${overwriteMode==='all'?'erstatt eksisterende':'bevar eksisterende'} · ${addExtraSizes?'opprett manglende plasser':'bruk eksisterende plasser'} · gjenbruk ${reuseAssets?'på':'av'}`);
  for(const [id,records]of byId){
    LOG(`→ Linje ${id}: ${prettyCounts(records)||'ingen plasser'}`);
    const groups=workGroups(id,records,mapping,imagePool);
    for(const group of groups){
      LOG(`   ${group.size}${group.variant?'/'+group.variant:''}${group.side?'/'+group.side:''}: ${group.assets.length} ${group.kind==='image'?'bilder':'scripts'}, ${group.records.length} plasser.`);
    }
  }
  LOG('Tom/utfylt-status, eksisterende scripts og nøyaktig tilleggsbehov kontrolleres ved autofyll.');
}

btnScan.onclick=()=>withJob('Skanner Adpoint…',async()=>{
  log.textContent='';
  await waitForIdle();
  const rows=findRows();if(!rows.length)throw new Error('Fant ingen kampanjelinjer på denne siden.');
  const ids=rows.map(rowId);
  if(ids.includes('?')||new Set(ids).size!==ids.length)throw new Error('Linje-ID-ene kunne ikke leses entydig fra kolonnen Line item.');
  const report=[],byId=new Map();
  for(const id of ids){
    if(stopping)break;
    const records=await recordsForLine(id);
    let allowed=[];
    try{allowed=await getAllowedAddSizesForLine(id);}
    catch(error){LOG(`! Linje ${id}: autofyll kan brukes, men størrelsesmenyen kunne ikke leses: ${error.message}`);}
    const row=findRowsByIdAll(id)[0];const {rosenrIdx}=getHeaderIdxForRow(row);
    const material=row.children[rosenrIdx]?.textContent.trim()||'';
    const label=`${prettyCounts(records)||'ingen plasser'}${material?' · materiell '+material:''}`;
    report.push({id,label,allowed});byId.set(id,records);LOG(`• ${id}: ${label}`);
  }
  scanned=report;selected=new Set(report.map(x=>x.id));saveSelection();refreshList();refreshBulkOptions();
  _lastPreview=byId;renderPreview(byId);
  setInfo({items:ids.length,done:report.length});
  showToast(stopping?'Skann stoppet':'Skann ferdig',`${report.length} linjer lest`);
  adjustHeights(ui.getBoundingClientRect().height);
});

function selectedIds(){return scanned.filter(x=>selected.has(x.id)).map(x=>x.id);}
function bulkRequests(){
  const options=availableBulkOptions(),requests=new Map();
  for(const control of bulkControls){
    const option=options.find(x=>x.key===control.select.value),count=Number(control.count.value);
    if(!option||!Number.isSafeInteger(count)||count<1)throw new Error('Velg en størrelse og et positivt heltall for hver rad.');
    const existing=requests.get(option.key);
    if(existing){existing.count+=count;if(!Number.isSafeInteger(existing.count))throw new Error('Antallet er for stort.');}
    else requests.set(option.key,{...option,count});
  }
  if(!requests.size)throw new Error('Legg til minst én størrelse.');
  return [...requests.values()];
}
btnBulk.onclick=()=>withJob('Legger til materiellplasser…',async()=>{
  const ids=selectedIds();if(!ids.length)throw new Error('Skann og velg minst én linje først.');
  const requests=bulkRequests(),plan=[];
  // Refresh allowed menus before mutating; the scan can have become stale.
  for(const id of ids){
    requireRunning();const allowed=await getAllowedAddSizesForLine(id);
    for(const request of requests){
      const option=allowed.find(x=>x.size===request.size&&x.label===request.label&&(x.variant||null)===(request.variant||null));
      if(option)plan.push({id,request,option});else LOG(`• Linje ${id}: hopper over ${request.label}; linjen tilbyr ikke dette valget.`);
    }
  }
  const total=plan.reduce((n,x)=>n+x.request.count,0);
  if(!total)throw new Error('Ingen av valgte linjer tilbyr størrelsene du valgte.');
  LOG(`Legger til ${total} nye materiellplasser på ${new Set(plan.map(x=>x.id)).size} linjer.`);
  const stats={items:ids.length,done:0,added:0,err:0};
  try{
    for(const {id,request,option} of plan){
      for(let n=0;n<request.count;n++){
        requireRunning();
        const fresh=await addSizeScopedToLine(id,request.size,request.variant,option.label);
        stats.added++;setInfo(stats);LOG(`+ ${stats.added}/${total}: linje ${id}, ${request.label}, placeholder ${fresh.placeholder}`);
      }
    }
    stats.done=ids.length;LOG(`Ferdig: ${stats.added} nye materiellplasser.`);showToast('Materiellplasser lagt til',`${stats.added} nye plasser`);
  }catch(error){if(error.uncertain)stats.uncertain=(stats.uncertain||0)+1;else if(!stopping)stats.err++;throw error;}
  finally{setInfo(stats);await refreshAfterMutation(ids);}
});
async function refreshAfterMutation(ids){
  const byId=new Map();
  for(const id of ids){
    try{
      const records=await recordsForLine(id);byId.set(id,records);
      const item=scanned.find(x=>x.id===id);if(item)item.label=prettyCounts(records)||'ingen plasser';
    }catch(error){LOG(`! Kunne ikke oppdatere visning for ${id}: ${error.message}`);}
  }
  _lastPreview=byId;refreshList();
}

btnRun.onclick=()=>withJob('Legger inn materiell…',async()=>{
  const ids=selectedIds();if(!ids.length)throw new Error('Skann og velg minst én linje først.');
  if(!mapping.length&&!imagePool.size)throw new Error('Importer CSV/JSON eller bilder først.');
  const map=mapping.map(x=>({...x,lineIds:[...(x.lineIds||[])]}));
  const images=new Map([...imagePool].map(([key,files])=>[key,files.slice()]));
  const settings={overwrite:overwriteMode,add:addExtraSizes,reuse:reuseAssets};
  const stats={items:ids.length,done:0,saved:0,added:0,skip:0,err:0,matched:0};setInfo(stats);
  try{
    for(const id of ids){
      requireRunning();await waitForIdle();
      const records=await recordsForLine(id);
      const groups=workGroups(id,records,map,images);
      const allowed=settings.add?await getAllowedAddSizesForLine(id):[];
      if(!groups.length){stats.skip++;stats.done++;setInfo(stats);continue;}
      for(const group of groups){
        requireRunning();
        if(!group.records.length&&(!settings.add||!optionForGroup(allowed,group))){
          stats.skip+=group.assets.length;
          LOG(`• Linje ${id}: hopper over ${group.size}${group.variant?'/'+group.variant:''}; ingen passende plass eller tilgjengelig størrelsesvalg.`);
          continue;
        }
        const states=[];
        for(const record of group.records){requireRunning();states.push(await inspectRecord(record));}
        const matches=settings.overwrite==='empty'?takeExistingMatches(group.assets,states,group.kind):{remaining:group.assets.slice(),matched:0};
        let remaining=matches.remaining;stats.matched+=matches.matched;
        if(matches.matched)LOG(`✓ Linje ${id}, ${group.size}: ${matches.matched} scripts finnes allerede og bevares.`);
        const slots=settings.overwrite==='all'?states:states.filter(x=>!x.occupied);
        stats.skip+=states.length-slots.length;
        const need=settings.add?Math.max(0,remaining.length-slots.length):0;
        if(need){
          const option=optionForGroup(allowed,group);
          if(!option)throw new Error(`Linje ${id}: kan ikke entydig opprette ${group.size}${group.variant?'/'+group.variant:''}. Ingen scripts i denne gruppen ble skrevet.`);
          LOG(`Linje ${id}, ${group.size}: oppretter ${need} manglende plasser.`);
          for(let i=0;i<need;i++){
            requireRunning();
            const record=await addSizeScopedToLine(id,group.size,group.variant,option.label);
            if(group.side&&record.side!==group.side)throw new Error('Ny materiellplass har ikke forventet side. Kontroller Adpoint.');
            stats.added++;setInfo(stats);
            const state=await inspectRecord(record);
            if(state.occupied)throw new Error('Ny materiellplass inneholder allerede materiell. Kontroller Adpoint.');
            slots.push(state);
          }
        }
        let placed=0;
        for(const state of slots){
          requireRunning();
          const asset=settings.reuse&&remaining.length?remaining[placed%remaining.length]:remaining[placed];
          if(!asset)break;
          // Re-read immediately before writing: UI and server may have changed since planning.
          const check=await inspectRecord(state.record);
          if(settings.overwrite==='empty'&&check.occupied)throw new Error('En planlagt tom plass er blitt fylt. Jobben stoppet før overskriving.');
          if(!sameTag(check.value,state.value))throw new Error('Materiellinnholdet endret seg under kjøringen. Kontroller Adpoint.');
          if(group.kind==='image'){
            await uploadImageToTile(check.record,asset);
          }else{
            const editor=await editorForRecord(check.record);
            assertTarget(check.record);
            if(!await pasteWithVerify(editor,asset.tag,1))throw new Error('Scriptet festet seg ikke i editoren.');
            assertTarget(check.record);
            await saveWithRetry(editor,check.record,asset.tag);
          }
          placed++;stats.saved++;setInfo(stats);
          LOG(`✓ Linje ${id}, ${group.size}, placeholder ${state.record.placeholder}: ${group.kind==='image'?fileName(asset):tagLabel(asset)} — bekreftet.`);
        }
        if(!settings.reuse&&placed<remaining.length){
          stats.skip+=remaining.length-placed;
          LOG(`! ${remaining.length-placed} ${group.kind==='image'?'bilder':'scripts'} mangler plass på linje ${id}. Slå på «Opprett manglende materiellplasser» for å legge til.`);
        }
      }
      stats.done++;setInfo(stats);
    }
    LOG(`Ferdig. Bekreftet lagret/opplastet: ${stats.saved}. Allerede på plass: ${stats.matched}. Nye plasser: ${stats.added}. Hoppet over: ${stats.skip}.`);
    showToast('Autofyll ferdig',`Lagret: ${stats.saved} · Allerede på plass: ${stats.matched} · Nye: ${stats.added}`);
  }catch(error){if(error.uncertain)stats.uncertain=(stats.uncertain||0)+1;else if(!stopping)stats.err++;throw error;}
  finally{setInfo(stats);await refreshAfterMutation(ids);}
});
refreshList();btnBulk.disabled=true;


/* ========= simple task-based interface ========= */
(function simplifyInterface(){
  ui.classList.add('ap-simple');
  const css=d.createElement('style');css.textContent=`
#ap3p_bar.ap-simple{font:13px/1.45 system-ui,-apple-system,Segoe UI,sans-serif;min-width:420px;isolation:isolate}
#ap3p_bar.ap-simple *{box-sizing:border-box}
#ap3p_bar.ap-simple [hidden]{display:none!important}
#ap3p_bar.ap-simple button,#ap3p_bar.ap-simple input,#ap3p_bar.ap-simple select{font:inherit}
#ap3p_bar.ap-simple button:disabled{opacity:.45;cursor:default!important}
#ap3p_bar.ap-simple button:focus-visible,#ap3p_bar.ap-simple summary:focus-visible,#ap3p_bar.ap-simple input:focus-visible,#ap3p_bar.ap-simple select:focus-visible{outline:2px solid #67caff;outline-offset:3px}
#ap3p_bar .ep-card{border:1px solid #303644;border-radius:10px;padding:13px;background:#141820;display:flex;flex-direction:column;gap:10px}
#ap3p_bar .ep-heading{font-weight:650;color:#f4f6fa;font-size:14px;margin:0}
#ap3p_bar .ep-help{color:#a9b3c4;font-size:12px;margin:0}
#ap3p_bar .ep-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
#ap3p_bar .ep-mode{display:flex;gap:5px;padding:4px;border:1px solid #303644;border-radius:10px;background:#0b0e14}
#ap3p_bar .ep-mode button{flex:1;border:0;border-radius:7px;padding:10px 8px;background:transparent;color:#a9b3c4;cursor:pointer;font-weight:600}
#ap3p_bar .ep-mode button[aria-pressed="true"]{background:#253348;color:#f4f8ff;box-shadow:0 1px 4px #0005}
#ap3p_bar summary{cursor:pointer;color:#c4cedc;font-weight:500;list-style-position:inside;padding:8px 0}
#ap3p_bar summary::marker{color:#7d8ba0}
#ap3p_bar .ep-options>div{display:flex;flex-direction:column;gap:12px;padding:8px 0 5px}
#ap3p_bar .ep-options label{margin:0!important;align-items:flex-start!important;gap:8px!important;opacity:1!important}
#ap3p_bar .ep-options label input{margin-top:3px;flex:none}
#ap3p_bar .ep-options small{display:block;color:#9ba8ba;margin-top:3px;font-size:11px;line-height:1.4}
#ap3p_bar .ep-action{display:flex;flex-direction:column;gap:9px;padding-top:12px;border-top:1px solid #303644}
#ap3p_bar .ep-action>button{width:100%;padding:11px 14px!important;font-size:14px!important;border-radius:8px!important}
#ap3p_bar .ep-status{color:#a9b3c4;font-size:12px;white-space:pre-wrap;overflow-wrap:anywhere}
#ap3p_bar .ep-policy{color:#a9b3c4;font-size:12px}
#ap3p_bar .ep-policy[data-replace="true"]{color:#ffca8b}
#ap3p_bar #ap3p_list .item{padding:7px 4px;gap:8px;align-items:flex-start}
#ap3p_bar #ap3p_list .sizes{font:12px/1.4 system-ui,sans-serif;color:#a9b3c4}
#ap3p_bar #ap3p_list .id{font-size:12px;min-width:66px;color:#e1e7ef}
#ap3p_bar #ap3p_list input{margin-top:2px}
`;
  // Store the extra style inside the panel so the existing close/reopen cleanup owns it.
  ui.append(css);
  title.textContent='EasyPoint';title.style.cssText='font-weight:700;letter-spacing:-.2px';
  badge.style.display='none';
  hdr.style.cssText='cursor:move;user-select:none;display:flex;align-items:center;gap:8px;padding:10px 14px;background:#161b24;border-radius:12px 12px 0 0;border-bottom:1px solid #303644;position:sticky;top:0;z-index:2';
  hdr.replaceChildren(title,flex,btnMin,btnX);
  body.style.cssText='padding:14px;display:flex;flex-direction:column;gap:12px;align-items:stretch';
  const hidden=(el,value)=>{if(value)el.setAttribute('hidden','');else el.removeAttribute('hidden');};
  const make=(tag,text,className)=>{const el=d.createElement(tag);if(text)el.textContent=text;if(className)el.className=className;return el;};
  const taskSwitch=make('div','','ep-mode');taskSwitch.setAttribute('role','group');taskSwitch.setAttribute('aria-label','Hva vil du gjøre?');
  const fillMode=make('button','Scripts og bilder');fillMode.type='button';fillMode.setAttribute('aria-controls','ep-fill-card');
  const emptyMode=make('button','Tomme materiellplasser');emptyMode.type='button';emptyMode.setAttribute('aria-controls','ep-empty-card');
  taskSwitch.append(fillMode,emptyMode);

  const scanCard=make('section','','ep-card');
  const scanRow=make('div','','ep-row');
  const scanHeading=make('h2','1. Hent annonser','ep-heading');scanHeading.style.flex='1';
  btnScan.textContent='Hent fra Adpoint';btnScan.style.cssText='cursor:pointer;border:1px solid #44516a;border-radius:7px;padding:7px 11px;background:#253348;color:#e8f1ff;font-weight:600';
  scanRow.append(scanHeading,btnScan);
  const scanHelp=make('p','Henter annonselinjene fra kampanjen du har åpen.','ep-help');
  const selectionDetails=make('details');
  const selectionSummary=make('summary','Velg annonselinjer');selectionSummary.setAttribute('aria-live','polite');
  selectionDetails.append(selectionSummary,listWrap);
  listWrap.style.cssText='width:100%';list.style.maxHeight='150px';
  btnAll.textContent='Alle';btnNone.textContent='Ingen';hidden(btnInv,true);
  scanCard.append(scanRow,scanHelp,selectionDetails);

  const fillCard=make('section','','ep-card');fillCard.id='ep-fill-card';
  const importHeading=make('h2','2. Velg scripts eller bilder','ep-heading');
  const importHelp=make('p','Importer filen med scripts, lim inn fra Excel eller velg bildefiler.','ep-help');
  const importButtons=make('div','','ep-row');btnCSV.textContent='Velg fil';btnPaste.textContent='Lim inn fra Excel';importButtons.append(btnCSV,btnPaste,file);
  drop.textContent='Du kan også slippe CSV, JSON eller bilder her';drop.style.cssText='border:1px dashed #455169;border-radius:7px;padding:13px;text-align:center;color:#a9b3c4;font-size:12px';
  const importedRow=make('div','','ep-row');mapChip.className='';mapChip.style.cssText='flex:1;font-size:12px;color:#b4dec7';
  mapClear.textContent='Tøm import';mapClear.title='Fjern importerte scripts og bilder';mapClear.style.cssText='cursor:pointer;border:0;background:transparent;color:#a9b3c4;padding:3px;font-size:12px';
  importedRow.append(mapChip,mapClear);
  fillCard.append(importHeading,importHelp,importButtons,drop,importedRow);

  bulk.id='ep-empty-card';bulk.className='ep-card';bulk.style.cssText='';
  bulkLabel.textContent='2. Velg størrelse og antall';bulkLabel.className='ep-heading';bulkLabel.style.fontSize='14px';
  bulkNote.textContent='Antallet er per valgt linje. Størrelsen må være tilgjengelig i Adpoint.';bulkNote.className='ep-help';bulkNote.style.cssText='';
  bulkPlus.textContent='+ Legg til en størrelse';bulkPlus.style.alignSelf='flex-start';
  bulk.replaceChildren(bulkLabel,bulkNote,bulkRows,bulkPlus);

  const options=make('details','','ep-options');
  const optionsSummary=make('summary','Flere valg');
  const optionsBody=make('div');
  const explain=(wrap,text,hint)=>{
    const input=wrap.querySelector('input');const content=make('span',text);content.append(make('small',hint));wrap.replaceChildren(input,content);return content;
  };
  const reuseContent=explain(reuseWrap,'Gjenta scripts og bilder ved behov','Bruker importen flere ganger hvis det er flere plasser enn filer.');
  explain(addWrap,'Opprett manglende materiellplasser','Legger til plasser hvis importen har flere versjoner enn det er plass til.');
  explain(owWrap,'Erstatt eksisterende materiell','Tillater at scripts og bilder som allerede ligger inne blir byttet ut.');
  optionsBody.append(addWrap,owWrap,reuseWrap);options.append(optionsSummary,optionsBody);

  const action=make('div','','ep-action');
  const policy=make('div','','ep-policy');
  const status=make('div','Start med å hente annonsene fra Adpoint.','ep-status');status.setAttribute('role','status');status.setAttribute('aria-live','polite');
  action.append(policy,btnRun,btnBulk,btnStop,status);
  btnBulk.style.cssText=btnRun.style.cssText;
  btnStop.style.cssText='cursor:pointer;border:1px solid #a84d52;border-radius:8px;padding:8px 12px;background:#4c222a;color:#ffe4e8;font-weight:600';
  const diagnostics=make('details');const diagnosticsSummary=make('summary','Vis detaljer');
  info.style.cssText='display:block;font-size:11px;color:#9ba8ba;padding:0 0 7px';
  log.style.cssText='max-height:180px;overflow:auto;background:#0b0e14;border:1px solid #303644;border-radius:7px;padding:9px;font:11px/1.5 ui-monospace,Consolas,monospace;white-space:pre-wrap;overflow-wrap:anywhere';
  diagnostics.append(diagnosticsSummary,info,log);
  body.replaceChildren(taskSwitch,scanCard,fillCard,bulk,options,action,diagnostics);
  let mode='fill';
  refreshSimpleUI=()=>{
    const ids=selectedIds(),count=ids.length;
    let images=0;imagePool.forEach(items=>images+=items.length);
    const imports=mapping.length+images;
    fillMode.setAttribute('aria-pressed',String(mode==='fill'));emptyMode.setAttribute('aria-pressed',String(mode==='empty'));
    fillMode.disabled=app.busy;emptyMode.disabled=app.busy;
    hidden(fillCard,mode!=='fill');hidden(bulk,mode!=='empty');hidden(options,mode!=='fill');
    hidden(btnRun,mode!=='fill');hidden(btnBulk,mode!=='empty');hidden(btnStop,!app.busy);
    hidden(selectionDetails,!scanned.length);hidden(importedRow,!imports);
    hidden(policy,mode!=='fill'||!imports);
    btnScan.textContent=scanned.length?'Hent på nytt':'Hent fra Adpoint';
    scanHelp.textContent=scanned.length?`${scanned.length} annonselinjer hentet fra Adpoint.`:'Henter annonselinjene fra kampanjen du har åpen.';
    selectionSummary.textContent=count===scanned.length?`Alle ${count} linjer er valgt · endre utvalg`:`${count} av ${scanned.length} linjer er valgt · endre utvalg`;
    mapChip.className='';mapChip.textContent=[mapping.length?`${mapping.length} scripts`:null,images?`${images} bilder`:null].filter(Boolean).join(' og ')+' klare';
    btnRun.textContent=count?`Legg inn på ${count} ${count===1?'linje':'linjer'}`:'Legg inn materiell';
    btnBulk.textContent=count?`Opprett plasser på ${count} ${count===1?'linje':'linjer'}`:'Opprett materiellplasser';
    btnRun.disabled=app.busy||!count||!imports;btnBulk.disabled=app.busy||!count||!availableBulkOptions().length;
    bulkPlus.disabled=app.busy||!scanned.length;
    policy.textContent=overwriteMode==='all'?'Erstatter eksisterende materiell.':'Bevarer eksisterende materiell.';
    if(addExtraSizes)policy.textContent+=' Oppretter manglende plasser.';
    policy.setAttribute('data-replace',String(overwriteMode==='all'));
    reuseContent.style.opacity=addExtraSizes?'.55':'1';
    if(app.lastError){status.textContent=app.lastError;status.style.color='#ffb7b7';diagnostics.open=true;}
    else if(app.busy){
      const progress=app.lastStats;
      status.textContent=stopping?'Stopper etter pågående steg…':(app.jobLabel||'Arbeider…')+(progress?.saved?' · '+progress.saved+' lagret':progress?.added?' · '+progress.added+' plasser opprettet':'');status.style.color='#a9cff3';
    }
    else if(app.lastStats&&((app.lastStats.saved||0)+(app.lastStats.added||0)+(app.lastStats.err||0)+(app.lastStats.uncertain||0)>0)){
      const result=app.lastStats;status.textContent=[result.saved?`${result.saved} materiell lagret`:null,result.added?`${result.added} nye plasser`:null,result.err?`${result.err} feil`:null,result.uncertain?`${result.uncertain} usikre endringer`:null].filter(Boolean).join(' · ');status.style.color='#b4dec7';
    }else{
      status.style.color='#a9b3c4';status.textContent=!scanned.length?'Start med å hente annonsene fra Adpoint.':!count?'Velg minst én annonselinje.':mode==='empty'?'Velg størrelse og antall nye plasser.':!imports?'Velg fil eller lim inn scripts for å fortsette.':'Klar til å legge inn materiell.';
    }
  };
  fillMode.onclick=()=>{if(app.busy)return;mode='fill';refreshSimpleUI();};
  emptyMode.onclick=()=>{if(app.busy)return;mode='empty';refreshSimpleUI();};
  app.runSelected=()=>{if(mode==='empty')btnBulk.click();else btnRun.click();};
  for(const control of [reuseCb,addCb,owCb]){
    const previousChange=control.onchange;
    control.onchange=event=>{previousChange?.(event);refreshSimpleUI();};
  }
  listen(options,'toggle',()=>adjustHeights(ui.getBoundingClientRect().height));
  listen(diagnostics,'toggle',()=>adjustHeights(ui.getBoundingClientRect().height));
  refreshSimpleUI();
})();


/* ========= campaign view helpers ========= */
function inCampaignTableView(){ return !!d.querySelector('tr.set-matSpecDataRow'); }
function findBackToCampaignButton(){
  const btns=[...d.querySelectorAll('button,[role="button"]')];
  for(const b of btns){
    const p=b.querySelector('svg path');
    const dd=p?.getAttribute?.('d')||'';
    if(dd.startsWith('M9.4 16.6L4.8 12')) return b;
  }
  return null;
}
async function ensureCampaignView(){
  if(inCampaignTableView()) return true;
  const back=findBackToCampaignButton();
  if(back){
    back.scrollIntoView({block:'center'});
    back.click();
    await waitFor(()=>inCampaignTableView(),8000,120);
  }
  return inCampaignTableView();
}
async function waitForIdle(timeout=12000){
  let quietSince=0; const started=Date.now();
  while(Date.now()-started<timeout){
    const busy=[...d.querySelectorAll('.MuiBackdrop-root,[role="progressbar"],.MuiCircularProgress-root')]
      .some(el=>vis(el)&&getComputedStyle(el).opacity!=='0'&&!el.closest('#ap3p_bar'));
    if(!busy){ if(!quietSince) quietSince=Date.now(); if(Date.now()-quietSince>=200) return true; }
    else quietSince=0;
    await sleep(90);
  }
  throw new Error('Adpoint ble ikke ferdig innen tidsfristen. Jobben er stoppet.');
}

/* ========= selection shortcuts + badge ========= */
_bestiltInt=setInterval(()=>{
  const panel=d.querySelector('#InfoPanelContainer');
  const t=panel?.innerText||'';
  const m=t.match(/Bestilt(?:\s*\(BxH\))?\s*[:\-]?\s*(\d{2,4})\s*[x×]\s*(\d{2,4})/i);
  badge.textContent=m?`— ${m[1]}x${m[2]}`:'';
},900);
listen(w,'keydown',e=>{ if(e.altKey && e.key.toLowerCase()==='a'){ if(e.target?.closest('input,textarea,[contenteditable="true"],.CodeMirror')) return; e.preventDefault(); if(app.runSelected)app.runSelected();else btnRun.click(); }});

/* restore minimized state */
(function restoreMinState(){
  const st=loadUIState();
  if(st?.min) setMinimized(true);
})();

}catch(e){console.error(e);alert('Autofill-feil: '+(e&&e.message?e.message:e));}})();






