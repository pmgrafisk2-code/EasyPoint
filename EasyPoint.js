javascript:(()=>{try{
/* ========= helpers ========= */
const d=document,w=window;
const S=(k,v)=>localStorage.setItem(k,JSON.stringify(v));
const G=(k,f)=>{try{return JSON.parse(localStorage.getItem(k))??f}catch{return f}};
const vis=el=>!!(el&&(el.offsetParent||(el.getClientRects&&el.getClientRects().length)));
const cs=s=>String(s||'').toLowerCase().replace(/\s+/g,'').replace(/×/g,'x');
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function waitFor(fn,ms=6000,step=100){const t=Date.now();let v;while(Date.now()-t<ms){v=fn();if(v) return v;await sleep(step)}return null}
function parseSizeVariant(s){const str=String(s||'');const m=str.match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);if(!m) return{size:'',variant:null};const size=`${m[1]}x${m[2]}`;const tail=str.slice(m.index+m[0].length);const v=(tail.match(/desktop|mobil|mobile/i)||[''])[0].toLowerCase();const variant=v.includes('desk')?'desktop':(v?'mobil':null);return{size,variant}}
function prettyCounts(entries){const counts=new Map();for(const e of entries){const key=e.size+(e.variant?`/${e.variant[0]}`:'');counts.set(key,(counts.get(key)||0)+1)}return[...counts.entries()].map(([k,n])=>n>1?`${k}×${n}`:k).join(', ')}

/* ========= version keys & line-id ========= */
function versionKey(str){const s=String(str||'').toLowerCase();const mIx=s.match(/\bix[-_][0-9a-z-]+\b/i);if(mIx) return mIx[0];const mV=s.match(/\b(?:ver|v)[\s._-]?(\d{1,3})\b/i);if(mV) return `v${mV[1]}`;return null}
function extractLineIds(str){
  const ids = new Set();
  const re = /L[\s\u00A0_-]*?(\d{6,10})(?!\d)/gi;
  const s = String(str || '');
  let m;
  while ((m = re.exec(s)) !== null) ids.add(m[1]);
  return [...ids];
}

// Ensures all creative tiles are actually mounted (virtualized list)
async function ensureAllTilesMounted(root = d){
  const scroller =
    root.querySelector('.set-creativeSectionContainer') ||
    root.querySelector('.set-creativeContainer') ||
    d.querySelector('.set-creativeSectionContainer') ||
    d.scrollingElement;

  if (!scroller) return;

  let prev = -1, still = 0;
  for (let i = 0; i < 32 && still < 3; i++) {
    scroller.scrollBy(0, 1e6);
    await sleep(300);
    const count = d.querySelectorAll('.set-creativeTile').length;
    if (count === prev) still++; else { prev = count; still = 0; }
  }
  scroller.scrollTo(0, 0);
}

/* ========= image side helpers ========= */
function detectSide(str){const s=(str||'').toLowerCase().replace('høyre','hoyre');if(/\b(hoyre|right|r)\b/.test(s))return'right';if(/\b(venstre|left|l)\b/.test(s))return'left';return null}

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
  if (!tile) return false;
  if (isTileSelected(tile)) return true;

  const targets = tileClickTargets(tile);

  for (let attempt = 0; attempt < 3; attempt++) {
    for (const t of targets) {
      if (isTileSelected(tile)) return true;

      try { t.scrollIntoView({ block: 'center' }); } catch {}
      t.dispatchEvent(new MouseEvent('pointerdown', { bubbles:true }));
      t.dispatchEvent(new MouseEvent('mousedown',  { bubbles:true }));
      t.dispatchEvent(new MouseEvent('mouseup',    { bubbles:true }));
      t.click();

      for (let i = 0; i < 10; i++) {
        await sleep(80);
        if (isTileSelected(tile)) return true;
      }
    }

    try {
      tile.scrollIntoView({ block:'center' });
      tile.focus?.();
      tile.dispatchEvent(new KeyboardEvent('keydown', { bubbles:true, key:' ' }));
      tile.dispatchEvent(new KeyboardEvent('keyup',   { bubbles:true, key:' ' }));
    } catch {}

    for (let i = 0; i < 8; i++) {
      await sleep(90);
      if (isTileSelected(tile)) return true;
    }
  }
  return isTileSelected(tile);
}
function getAllTiles(details){
  let tiles = [...details.querySelectorAll('.set-creativeTile, .set-creativeTile__selected')];
  return tiles.filter(t => !tiles.some(o => o !== t && o.contains(t)));
}
function tileGroup(tile){return tile.closest('.set-creativeContainer, .set-creativeSectionContainer')||tile.parentElement}
function groupLabel(group){
  const n=group&&(group.querySelector('.set-cardLabel__main [title]')||group.querySelector('.set-cardLabel__main div[title]')||group.querySelector('.set-cardLabel__main'));
  return(n?.getAttribute?.('title')||n?.textContent||'')||''
}
function tileVariant(tile){const lbl=groupLabel(tileGroup(tile)).toLowerCase();if(/desktop/.test(lbl))return'desktop';if(/mobil|mobile/.test(lbl))return'mobil';return null}
function tileSide(tile){return detectSide(groupLabel(tileGroup(tile)))}
function tileHasSize(tile,size){
  const rx=new RegExp(`\\b${size.replace('x','[x×]')}\\b`,'i');
  if(rx.test((tile.innerText||'')))return true;
  const lbl=groupLabel(tileGroup(tile));
  const mappedRe=new RegExp(`mapped_${size.replace('x','[x×]')}(?:\\b|[_-])`,'i');
  if(mappedRe.test(lbl))return true;
  if(rx.test(lbl))return true;
  return false;
}

/* ========= mapping import (CSV/JSON + Excel paste + DnD) ========= */
function parseCSV(text){
  let rows=[],row=[],f='',q=false;
  const first=(text.split(/\r?\n/).find(l=>l.trim().length>0)||'');
  const c1=(first.match(/,/g)||[]).length,c2=(first.match(/;/g)||[]).length,c3=(first.match(/\t/g)||[]).length;
  const del=c3>=c2&&c3>=c1?'\t':(c2>c1?';':',');
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(ch=='"'){if(q&&text[i+1]=='"'){f+='"';i++}else q=!q}
    else if(!q&&ch===del){row.push(f);f=''}
    else if(!q&&ch=='\n'){row.push(f);rows.push(row);row=[];f=''}
    else if(!q&&ch=='\r'){}
    else f+=ch
  }
  row.push(f);rows.push(row);return rows
}
function pickHeader(h,names){
  const L=h.map(x=>String(x||'').toLowerCase());
  for(const n of names){
    let i=L.indexOf(n);if(i>-1)return i;
    for(let k=0;k<L.length;k++) if(L[k].includes(n)) return k
  }
  return -1
}
function rowsToMap(rows){
  let head=[],headRow=0;
  for(let i=0;i<Math.min(rows.length,5);i++){
    const cells=(rows[i]||[]).map(c=>String(c||'').trim());
    if(cells.join('').length){head=cells;headRow=i;break}
  }
  const sizeIdx=pickHeader(head,['størrelse','stoerrelse','size']);
  const nameIdx=pickHeader(head,['creative name','creative','name','navn']);
  const tagIdx=pickHeader(head,['secure content','script','tag','kode','code']);
  if(tagIdx<0||(sizeIdx<0&&nameIdx<0))throw new Error('Mangler kolonner: Secure Content + (Size eller Creative Name)');

  let warned=false; const out=[];
  for(let r=headRow+1;r<rows.length;r++){
    const cells = rows[r] || [];
    const rawSize = sizeIdx>-1 ? (cells[sizeIdx]||'') : '';
    const rawName = nameIdx>-1 ? (cells[nameIdx]||'') : '';
    const svSize = parseSizeVariant(rawSize);
    const svName = parseSizeVariant(rawName);
    const size   = cs(svSize.size || svName.size);
    const variant= svName.variant ?? svSize.variant ?? null;
    const tag    = String(cells[tagIdx]||'').trim();
    const name   = rawName || '';
    const vkey   = versionKey(rawName);

    let lineIds = extractLineIds(rawName);
    if (!lineIds.length){
      const rowText = cells.map(c => String(c||'')).join(' _ ');
      lineIds = extractLineIds(rowText);
    }

    if (svSize.size && svName.size && cs(svSize.size)!==cs(svName.size) && !warned){
      LOG('! CSV-advarsel: "Size" er ulik "Creative Name" — jeg stoler på Size-kolonnen.');
      warned = true;
    }
    if (size && tag) out.push({ size, variant, tag, name, vkey, lineIds });
  }
  return out;
}
function normalizeJSON(arr){
  let warned=false; const out=[];
  (arr||[]).forEach(o=>{
    const rawName = o.name||o.Name||o['Creative Name']||o['creative name']||'';
    const rawSize = o.size||o.Size||o['Størrelse']||o['Stoerrelse']||'';
    const svSize  = parseSizeVariant(rawSize), svName=parseSizeVariant(rawName);
    const size    = cs(svSize.size || svName.size);
    const variant = (o.variant||o.Variant||svName.variant||svSize.variant)||null;
    const tag     = String(o.tag||o.Tag||o['Secure Content']||o['secure content']||o.Script||o.script||'').trim();
    const vkey    = versionKey(rawName);

    let lineIds = extractLineIds(rawName);
    if (!lineIds.length){
      const flat = Object.values(o||{}).map(v=>String(v||'')).join(' _ ');
      lineIds = extractLineIds(flat);
    }

    if (svSize.size && svName.size && cs(svSize.size)!==cs(svName.size) && !warned){
      LOG('! JSON-advarsel: "Size" er ulik "name" — jeg stoler på Size-feltet.');
      warned=true;
    }
    if (size && tag) out.push({ size, variant, tag, name: rawName, vkey, lineIds });
  });
  return out;
}

/* ========= line items ========= */
function normLabel(s){return String(s||'').toLowerCase().replace(/\s+/g,' ').trim()}
function findRows(){const rows=[...d.querySelectorAll('tr.set-matSpecDataRow')].filter(vis);return rows.filter(el=>!rows.some(o=>o!==el&&o.contains(el)))}
function getHeaderIdxForRow(row){
  const table=row.closest('table');if(!table) return{lineItemIdx:-1,rosenrIdx:-1};
  if(table._ap3p_headerIdx) return table._ap3p_headerIdx;
  const ths=[...(table.tHead?.querySelectorAll('th')||[])];
  let lineItemIdx=-1,rosenrIdx=-1;
  ths.forEach((th,i)=>{
    const t=normLabel(th.textContent);
    if(lineItemIdx===-1&&/(^|\s)line\s*item(\s|$)/.test(t)) lineItemIdx=i;
    if(rosenrIdx===-1&&(/rosenr/.test(t)||(/materiell/.test(t)&&/rosenr/.test(t)))) rosenrIdx=i
  });
  table._ap3p_headerIdx={lineItemIdx,rosenrIdx};return table._ap3p_headerIdx
}
function rowId(row){
  const {lineItemIdx,rosenrIdx}=getHeaderIdxForRow(row);
  const tds=[...row.querySelectorAll('td')];
  if(lineItemIdx>-1&&tds[lineItemIdx]){
    const txt=tds[lineItemIdx].textContent.trim();
    const m=txt.match(/\d{6,10}/);
    return(m&&m[0])||(txt||'?')
  }
  const rosenVal=(rosenrIdx>-1&&tds[rosenrIdx])?tds[rosenrIdx].textContent.trim():'';
  const nums=tds.map(td=>td.textContent.trim()).filter(t=>/^\d{6,10}$/.test(t));
  const hit=nums.find(n=>n!==rosenVal);
  return hit||nums[0]||'?'
}
function findRowsByIdAll(targetId){return findRows().filter(r=>rowId(r)===targetId)}
async function expandRow(row){
  if(!row) return null;
  row.scrollIntoView({block:'center'});
  row.dispatchEvent(new MouseEvent('pointerdown',{bubbles:true}));
  row.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));
  row.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));
  row.click();
  await sleep(120);
  row.click();
  await sleep(160);
  for(let i=0;i<14;i++){
    const sib=row.nextElementSibling;
    if(sib&&(sib.querySelector('.set-creativeSectionContainer, .set-creativeContainer'))) return sib;
    await sleep(90)
  }
  return row.nextElementSibling||null
}

/* ========= detect placeholders ========= */
function sizesFromExpanded(details){
  if(!details) return[];
  const entries=[];
  const tiles=getAllTiles(details);
  function findSizeInText(s){const m=String(s||'').match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);return m?`${m[1]}x${m[2]}`:''}
  function findSizeInMapped(lbl){const m=String(lbl||'').match(/mapped_(\d{2,4})\s*[x×]\s*(\d{2,4})(?:\b|[_\s])/i);return m?`${m[1]}x${m[2]}`:''}
  for(const tile of tiles){
    const text=(tile.innerText||'');
    const lbl=groupLabel(tileGroup(tile));
    let size=findSizeInText(text);
    if(!size) size=findSizeInMapped(lbl);
    if(!size) size=findSizeInText(lbl);
    if(!size) continue;
    const variant=tileVariant(tile);
    const side=tileSide(tile);
    entries.push({size:cs(size),variant,side,tile})
  }
  return entries
}
function pickTilesForEntry(details,entry){
  let tiles=getAllTiles(details).filter(t=>tileHasSize(t,entry.size));
  if(entry.variant) tiles=tiles.filter(t=>(tileVariant(t)||null)===entry.variant);
  if(entry.side) tiles=tiles.filter(t=>(tileSide(t)||null)===entry.side);
  return tiles
}

/* ========= 3rd-party tab + save/reprocess ========= */
async function open3PTab(){
  const t=[...d.querySelectorAll('button[role="tab"],a[role="tab"],button,a')].filter(vis).find(b=>/\b3(?:rd)?\s*party\s*tag\b/i.test(b.textContent||''));
  if(t){t.click()}
  const editor=await waitFor(()=>{
    const panel=d.querySelector('#InfoPanelContainer')||d;
    const cm=[...panel.querySelectorAll('.CodeMirror')].find(vis);
    if(cm) return{kind:'cm',el:cm,panel,cm:cm.CodeMirror};
    const ta=[...panel.querySelectorAll('textarea')].find(vis);
    if(ta) return{kind:'ta',el:ta,panel};
    return null
  },4000,120);
  return editor||null
}
function pasteInto(target,value){
  if(!target) return;
  if(target.kind==='cm'&&target.cm){
    try{const cm=target.cm;cm.setValue((value||'')+'');cm.refresh?.();return}catch{}
  }
  const ta=target.el||target;
  ta.scrollIntoView({block:'center'});
  ta.focus();
  ta.value=value;
  ta.dispatchEvent(new Event('input',{bubbles:true}));
  ta.dispatchEvent(new Event('change',{bubbles:true}))
}
async function clickSaveOrReprocess(scope){
  const root=scope?.panel||d.querySelector('#InfoPanelContainer')||d;
  const btn=[...root.querySelectorAll('button')].filter(vis).find(b=>{
    const t=(b.textContent||'').toLowerCase();
    return /lagre|save|oppdater|update|send inn på nytt|reprocess/i.test(t)&&!b.disabled&&!b.classList.contains('Mui-disabled')
  });
  if(btn){btn.scrollIntoView({block:'center'});btn.click();await sleep(700);return true}
  return false
}

/* ========= Verify + retry helpers ========= */
function normalizeTagForCompare(s){
  return String(s||'')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
function getEditorValue(editor){
  if (!editor) return '';
  if (editor.kind === 'cm' && editor.cm) {
    try { return editor.cm.getValue() || ''; } catch { return ''; }
  }
  const ta = editor.el || editor;
  return ta?.value || '';
}
async function pasteWithVerify(editor, value, tries = 2){
  const want = normalizeTagForCompare(value);
  for (let attempt = 0; attempt <= tries; attempt++){
    pasteInto(editor, value);
    await sleep(120);
    if (editor.kind === 'cm') {
      try { editor.cm?.refresh?.(); } catch {}
      await sleep(80);
    }
    const got = normalizeTagForCompare(getEditorValue(editor));
    if (got === want) return true;
    if (got.replace(/\s+/g,'') === want.replace(/\s+/g,'')) return true;
    await sleep(180);
  }
  return false;
}
async function saveWithRetry(editor, tries = 2){
  for (let i = 0; i <= tries; i++){
    const ok = await clickSaveOrReprocess(editor);
    if (!ok) {
      await sleep(250);
      continue;
    }

    await waitForIdle(15000);
    await ensureCampaignView();
    await waitForIdle(8000);
    return true;
  }
  return false;
}

async function checkpoint(label=''){
  await ensureCampaignView();
  await waitForIdle(15000);
  if (label) LOG(`   · sjekkpunkt ${label}`);
}
function isEmpty3PValue(v){ return !String(v||'').trim(); }

/* ========= UI ========= */
(function injectCSS(){
  const st=d.createElement('style');
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
@keyframes ap3pPulse { 0%{transform:scale(1);filter:brightness(1)} 30%{transform:scale(1.02);filter:brightness(1.25)} 100%{transform:scale(1);filter:brightness(1)} }
.ap3p-flash{animation: ap3pPulse 550ms ease-out 1;}
#ap3p_toast{
  position:fixed; right:18px; bottom:18px; z-index:2147483648;
  background:#0b0c10; border:1px solid #2a2d37; color:#e6e6e6;
  padding:10px 12px; border-radius:10px; box-shadow:0 10px 30px rgba(0,0,0,.45);
  max-width: 420px; font: 12px/1.35 ui-sans-serif,system-ui; opacity:0; transform:translateY(6px);
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
  d.head.appendChild(st)
})();

/* ===== Window state (size/pos) ===== */
const UI_STATE_KEY='ap3p_ui_state_v1';
function clamp(n,min,max){return Math.max(min,Math.min(max,n))}
function loadUIState(){
  const st=G(UI_STATE_KEY,null);
  if(!st||typeof st!=='object') return null;
  return st;
}
function saveUIState(){
  try{
    const r=ui.getBoundingClientRect();
    const st={
      left: Math.round(r.left),
      top:  Math.round(r.top),
      w:    Math.round(r.width),
      h:    Math.round(r.height),
      min:  ui.getAttribute('data-min')==='1'
    };
    S(UI_STATE_KEY, st);
  }catch{}
}

let ui=d.getElementById('ap3p_bar'); if(ui) ui.remove();
ui=d.createElement('div'); ui.id='ap3p_bar';
ui.style.cssText='position:fixed;right:16px;bottom:16px;z-index:2147483647;background:#0e0f13;color:#e6e6e6;border:1px solid #2a2d37;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.45);max-width:96vw;min-width:520px;max-height:90vh;';

const hdr=d.createElement('div');
hdr.style.cssText='cursor:move;user-select:none;display:flex;align-items:center;gap:10px;padding:8px 10px;background:#161922;border-radius:12px 12px 0 0;border-bottom:1px solid #2a2d37';
const title=d.createElement('div'); title.textContent='EasyPoint';
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
const file=d.createElement('input'); file.type='file'; file.accept='.csv,.json,.txt'; file.style.display='none';
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
const setInfo=o=>{info.textContent=`Linjer:${o.items}  Ferdig:${o.done}  Treff:${o.hit}  Hoppet:${o.skip}  Feil:${o.err}`};
body.append(info,btnCSV,btnPaste,file,drop);

/* ===== Apply saved UI state (don’t auto-bloat) ===== */
(function applyInitialUIState(){
  const st=loadUIState();
  const defW=660, defH=520;
  const vw=Math.max(520, Math.min(defW, Math.floor(window.innerWidth*0.92)));
  const vh=Math.max(260, Math.min(defH, Math.floor(window.innerHeight*0.85)));

  if(st?.w && st?.h){
    const ww=clamp(st.w, 520, Math.floor(window.innerWidth*0.96));
    const hh=clamp(st.h, 160, Math.floor(window.innerHeight*0.90));
    ui.style.width=ww+'px';
    ui.style.height=hh+'px';
  }else{
    ui.style.width=vw+'px';
    ui.style.height=vh+'px';
  }

  if(typeof st?.left==='number' && typeof st?.top==='number'){
    const ww=parseInt(ui.style.width,10)||ui.getBoundingClientRect().width;
    const hh=parseInt(ui.style.height,10)||ui.getBoundingClientRect().height;
    ui.style.left=clamp(st.left, 6, Math.floor(window.innerWidth-ww-6))+'px';
    ui.style.top =clamp(st.top,  6, Math.floor(window.innerHeight-hh-6))+'px';
    ui.style.right='auto'; ui.style.bottom='auto';
  }
})();

/* ===== UX: toast + flash ===== */
let toast = d.getElementById('ap3p_toast');
if(!toast){
  toast = d.createElement('div');
  toast.id='ap3p_toast';
  toast.innerHTML = `<div class="t"></div><div class="s"></div>`;
  d.body.appendChild(toast);
}
let toastTimer=null;
function showToast(t,s=''){
  try{
    toast.querySelector('.t').textContent = t||'';
    toast.querySelector('.s').textContent = s||'';
    toast.classList.add('show');
    if(toastTimer) clearTimeout(toastTimer);
    toastTimer=setTimeout(()=>toast.classList.remove('show'), 2000);
  }catch{}
}
function flash(el){
  if(!el) return;
  el.classList.remove('ap3p-flash');
  void el.offsetWidth;
  el.classList.add('ap3p-flash');
}

/* ===== Drag, Resize, Minimize, Close ===== */
function adjustHeights(totalH){
  if(ui.getAttribute('data-min')==='1') return;
  const hdrH=hdr.getBoundingClientRect().height;
  const bodyH=body.offsetParent?body.getBoundingClientRect().height:0;
  const lstH=(listWrap.style.display==='none')?0:listWrap.getBoundingClientRect().height;
  const pad=24;
  const logH=Math.max(110,totalH-(hdrH+bodyH+lstH+pad));
  log.style.height=logH+'px';
  ui.style.overflow='hidden';
}

(function makeDraggable(handle,box){
  let sx=0,sy=0,ox=0,oy=0,drag=false;
  handle.addEventListener('pointerdown',e=>{
    if(e.button!==0) return;
    drag=true;
    const r=box.getBoundingClientRect();
    ox=r.left; oy=r.top; sx=e.clientX; sy=e.clientY;
    box.style.left=ox+'px'; box.style.top=oy+'px';
    box.style.right='auto'; box.style.bottom='auto';
  });
  w.addEventListener('pointermove',e=>{
    if(!drag)return;
    box.style.left=(ox+e.clientX-sx)+'px';
    box.style.top =(oy+e.clientY-sy)+'px';
  });
  w.addEventListener('pointerup',()=>{
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
    w.addEventListener('pointermove',onMove);
    w.addEventListener('pointerup',onUp);
  }
  function onMove(e){
    if(!rs)return;
    const dx=e.clientX-rs.sx, dy=e.clientY-rs.sy; let x=rs.x,y=rs.y,wv=rs.w,hv=rs.h;
    if(/e/.test(rs.dir)) wv=Math.max(520,rs.w+dx);
    if(/s/.test(rs.dir)) hv=Math.max(180,rs.h+dy);
    if(/w/.test(rs.dir)){ wv=Math.max(520,rs.w-dx); x=rs.x+dx; }
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
  const isMin = ui.getAttribute('data-min') === '1';

  // --- going MINIMIZED ---
  if(min){
    if(!isMin){
      // store last expanded size before shrinking
      try{
        const r = ui.getBoundingClientRect();
        ui.setAttribute('data-last-h', String(Math.round(r.height)));
        ui.setAttribute('data-last-w', String(Math.round(r.width)));
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

  // --- going EXPANDED ---
  ui.setAttribute('data-min','0');
  body.style.display='flex';
  listWrap.style.display='block';
  log.style.display='block';

  (ui._resizerEls||[]).forEach(el=>el.style.display='');
  btnMin.textContent='–';

  // restore last expanded height (fallback to saved state, else a sane default)
  const st = loadUIState() || {};
  const lastH = parseInt(ui.getAttribute('data-last-h') || '', 10);
  const lastW = parseInt(ui.getAttribute('data-last-w') || '', 10);

  const targetH = clamp(
    (Number.isFinite(lastH) && lastH > 80) ? lastH : (st.h || 520),
    180,
    Math.floor(window.innerHeight * 0.90)
  );

  const targetW = clamp(
    (Number.isFinite(lastW) && lastW > 200) ? lastW : (st.w || 660),
    520,
    Math.floor(window.innerWidth * 0.96)
  );

  ui.style.height = targetH + 'px';
  ui.style.width  = targetW + 'px';

  adjustHeights(targetH);
  saveUIState();
}

btnMin.onclick=()=>setMinimized(ui.getAttribute('data-min')!=='1');

let _bestiltInt=null;
btnX.onclick=()=>{
  if(_bestiltInt) clearInterval(_bestiltInt);
  ui.remove(); toast?.remove?.();
};

adjustHeights(ui.getBoundingClientRect().height);

/* ========= mapping state ========= */
const MAP_KEY='ap3p_map_v2';
let mapping=G(MAP_KEY,[]);
let imagePool=new Map(); // key `${size}|${side||''}` -> [File,...]

function renderMapChip(){
  const csvCount=mapping.length;
  let imgCount=0; imagePool.forEach(a=>imgCount+=a.length);
  mapChip.className=(csvCount||imgCount)?'chip chip-ok':'chip chip-none';
  mapChip.textContent=(csvCount||imgCount)?`CSV: ${csvCount}  Bilder: ${imgCount} ✓`:'CSV/Bilder: ingen';
}
function saveMap(arr){
  mapping=arr;
  S(MAP_KEY,mapping);
  renderMapChip();
  flash(mapChip); flash(drop);
  showToast('Mapping oppdatert', `CSV-rader: ${mapping.length}`);
}
function clearMap(){
  mapping=[];
  S(MAP_KEY,mapping);
  renderMapChip();
  flash(mapChip);
  showToast('Mapping tømt');
}
mapClear.onclick=()=>{clearMap();imagePool.clear();renderMapChip()};
renderMapChip();

/* ========= toggles: gjenbruk / legg til / erstatt ========= */
const REUSE_KEY_NEW = 'ap3p_reuse_assets';
const REUSE_KEY_OLD = 'ap3p_reuse_images';
let reuseAssets = G(REUSE_KEY_NEW, G(REUSE_KEY_OLD, true)); // default ON

const reuseWrap = document.createElement('label');
reuseWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const reuseCb = document.createElement('input');
reuseCb.type='checkbox';
reuseCb.checked = reuseAssets;
const reuseTxt = document.createElement('span');
reuseTxt.textContent = 'Gjenbruk samme scripts for å fylle alt';
reuseWrap.append(reuseCb, reuseTxt);
body.insertBefore(reuseWrap, drop);

let _lastPreview = null;

reuseCb.onchange = () => {
  reuseAssets = reuseCb.checked;
  S(REUSE_KEY_NEW, reuseAssets);
  showToast('Gjenbruk', reuseAssets ? 'På (fyll alle størrelser)' : 'Av (bruk hver rad én gang)');
  if (_lastPreview) { log.textContent=''; previewPlannedPlacements(_lastPreview); }
};

// Legg til ekstra størrelser
const ADD_KEY = 'ap3p_add_extra_sizes';
let addExtraSizes = !!G(ADD_KEY, false);

const addWrap = document.createElement('label');
addWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const addCb = document.createElement('input');
addCb.type='checkbox';
addCb.checked = addExtraSizes;
const addTxt = document.createElement('span');
addTxt.textContent = 'Tillat å legge til størrelser (hvis flere scripts enn plasser)';
addWrap.append(addCb, addTxt);
body.insertBefore(addWrap, drop);

// Erstatt-modus: alle vs kun tomme
const OVERWRITE_KEY = 'ap3p_overwrite_mode'; // 'all' | 'empty'
let overwriteMode = G(OVERWRITE_KEY, 'empty');

const owWrap = document.createElement('label');
owWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const owCb = document.createElement('input');
owCb.type='checkbox';
owCb.checked = (overwriteMode === 'all');
const owTxt = document.createElement('span');
owTxt.textContent = 'Erstatt materiell hvis det finnes allerede';
owWrap.append(owCb, owTxt);
body.insertBefore(owWrap, drop);

owCb.onchange = () => {
  overwriteMode = owCb.checked ? 'all' : 'empty';
  S(OVERWRITE_KEY, overwriteMode);
  showToast('Erstatt-modus', overwriteMode === 'all' ? 'Erstatt alle' : 'Kun tomme størrelser');
  if (_lastPreview) { log.textContent=''; previewPlannedPlacements(_lastPreview); }
};

function syncReuseWithAddExtra(){
  if(addExtraSizes){
    reuseAssets = false;
    reuseCb.checked = false;
    reuseCb.disabled = true;
    reuseTxt.textContent = 'Gjenbruk';
    S(REUSE_KEY_NEW, reuseAssets);
  }else{
    reuseCb.disabled = false;
    reuseTxt.textContent = 'Gjenbruk';
    reuseAssets = !!G(REUSE_KEY_NEW, true);
    reuseCb.checked = reuseAssets;
  }
}

addCb.onchange = () => {
  addExtraSizes = addCb.checked;
  S(ADD_KEY, addExtraSizes);
  syncReuseWithAddExtra();
  showToast('Legg til ekstra', addExtraSizes ? 'På (kan legge til størrelser)' : 'Av');
  if (_lastPreview) { log.textContent=''; previewPlannedPlacements(_lastPreview); }
};
syncReuseWithAddExtra();

const reuseImages = () => reuseAssets;

/* ========= image helpers ========= */
function getTileDropzoneInput(tile){return tile.querySelector('input[type="file"]')||null}
async function confirmReplaceIfPrompted(timeout=5000){
  const dlg = await waitFor(() => {
    const el = document.querySelector('.MuiDialog-container, .MuiDialog-root, [role="dialog"]');
    if (!el) return null;
    const txt = (el.innerText || '').toLowerCase();
    if (/(bekreftelse|erstatte|erstatt|overwrite|replace|confirm)/.test(txt)) return el;
    return null;
  }, timeout, 120);
  if (!dlg) return false;

  const btns = [...dlg.querySelectorAll('button')];
  const isOK = b => /^(ok|ja|fortsett|erstatte|erstatt|update|replace|confirm)$/i.test((b.innerText || '').trim());
  const okBtn = btns.find(isOK) || btns.at(-1);
  if (okBtn) {
    okBtn.scrollIntoView({block:'center'});
    okBtn.click();
    await sleep(200);
    return true;
  }
  return false;
}
async function uploadImageToTile(tile,file){
  const inp=getTileDropzoneInput(tile);
  if(!inp) return false;
  const dt=new DataTransfer();
  dt.items.add(file);
  inp.files=dt.files;
  inp.dispatchEvent(new Event('input',{bubbles:true}));
  inp.dispatchEvent(new Event('change',{bubbles:true}));
  await sleep(200);
  await confirmReplaceIfPrompted(5000);
  await sleep(250);
  return true;
}
function parseSizeFromName(name){const m=String(name||'').match(/(\d{2,4})[x×](\d{2,4})/i);return m?`${m[1]}x${m[2]}`:''}
function parseImageMeta(file){return{size:cs(parseSizeFromName(file.name)),side:detectSide(file.name),file}}
function poolKey(size,side){return `${size}|${side||''}`}
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
  showToast('Bilder lagt til', `Antall: ${added}`);
}

/* ========= import UI ========= */
btnCSV.onclick=()=>file.click();
file.onchange=e=>{
  const f=e.target.files?.[0];
  if(!f) return;
  const r=new FileReader();
  r.onload=ev=>{
    const text=String(ev.target.result||'');
    if(f.name.toLowerCase().endsWith('.json')){
      try{saveMap(normalizeJSON(JSON.parse(text)))}catch{alert('Ugyldig JSON')}
    }else{
      try{saveMap(rowsToMap(parseCSV(text)))}catch(err){alert(err.message)}
    }
  };
  r.readAsText(f,'utf-8');
  file.value='';
};
btnPaste.onclick=()=>{
  const go=t=>{try{saveMap(rowsToMap(parseCSV(t)))}catch(e){alert(e.message)}};
  if(navigator.clipboard?.readText){
    navigator.clipboard.readText().then(go).catch(()=>{
      const t=prompt('Lim inn rader (Excel: Creative Name/Size + Secure Content)');
      if(t) go(t);
    })
  }else{
    const t=prompt('Lim inn rader (Excel: Creative Name/Size + Secure Content)');
    if(t) go(t);
  }
};
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='#171a24';flash(drop)}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='transparent'}));
function isCSV(f){return /\.csv$/i.test(f.name)||f.type==='text/csv'}
function isJSON(f){return /\.json$/i.test(f.name)||f.type==='application/json'}
function readFileText(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result||''));r.onerror=rej;r.readAsText(file,'utf-8')})}
drop.addEventListener('drop',async e=>{
  e.preventDefault();drop.style.background='transparent';
  const files=[...(e.dataTransfer?.files||[])]; if(!files.length) return;
  const csvs=files.filter(isCSV), jsons=files.filter(isJSON),
        imgs=files.filter(f=>/^image\//.test(f.type)||/\.(psd|tiff?|webp)$/i.test(f.name));

  try{for(const f of csvs){const text=await readFileText(f);saveMap(rowsToMap(parseCSV(text)))}}catch(err){alert('CSV-import feilet: '+(err?.message||err))}
  try{for(const f of jsons){const text=await readFileText(f);saveMap(normalizeJSON(JSON.parse(text)))}}catch(err){alert('JSON-import feilet: '+(err?.message||err))}
  if(imgs.length) addImagesToPool(imgs);
});

/* ========= pools (scripts) ========= */
function buildTagPools(mapping){
  const tagsByExactGlobal = new Map();
  const tagsBySizeGlobal  = new Map();
  const tagsByExactLine   = new Map();
  const tagsBySizeLine    = new Map();

  const keyExact = (lineId, size, variant) => (lineId ? `${lineId}|` : '') + `${size}|${variant || ''}`;
  const keySize  = (lineId, size) => (lineId ? `${lineId}|` : '') + size;

  let lineScopedCount = 0, globalCount = 0;

  for (const m of (mapping || [])){
    const ids = Array.isArray(m.lineIds) ? m.lineIds.filter(Boolean) : [];
    const payload = { tag: m.tag, name: m.name || '' };

    if (ids.length){
      for (const lid of ids){
        const k1 = keyExact(lid, m.size, m.variant || null);
        const k2 = keySize(lid, m.size);
        if (!tagsByExactLine.has(k1)) tagsByExactLine.set(k1, []);
        if (!tagsBySizeLine.has(k2))  tagsBySizeLine.set(k2, []);
        tagsByExactLine.get(k1).push(payload);
        tagsBySizeLine.get(k2).push(payload);
        lineScopedCount++;
      }
    }else{
      const k1 = `${m.size}|${m.variant || ''}`;
      const k2 = m.size;
      if (!tagsByExactGlobal.has(k1)) tagsByExactGlobal.set(k1, []);
      if (!tagsBySizeGlobal.has(k2))  tagsBySizeGlobal.set(k2, []);
      tagsByExactGlobal.get(k1).push(payload);
      tagsBySizeGlobal.get(k2).push(payload);
      globalCount++;
    }
  }

  const hasLineIds = lineScopedCount > 0;
  LOG(`CSV map: linje=${lineScopedCount}, global=${globalCount}`);

  return {
    hasLineIds,
    kExact: keyExact,
    kSize : keySize,
    getExact(lineId, size, variant){
      const kLine = keyExact(lineId, size, variant || null);
      const kGlob = `${size}|${variant || ''}`;
      return (tagsByExactLine.get(kLine) || tagsByExactGlobal.get(kGlob) || []);
    },
    getSize(lineId, size){
      const kLine = keySize(lineId, size);
      return (tagsBySizeLine.get(kLine) || tagsBySizeGlobal.get(size) || []);
    }
  };
}

/* ========= preview ========= */
function imageChoicesForEntry(entry){
  const exact=imagePool.get(`${entry.size}|${entry.side||''}`)||[];
  if(exact.length) return exact;
  if(entry.side){
    const any=imagePool.get(`${entry.size}|`)||[];
    if(any.length) return any
  }
  return[]
}
function fileName(f){return(f&&f.name)?f.name:'(fil)'}
function tagLabel(entry){
  if (!entry) return '(ukjent)';
  if (typeof entry === 'string') return entry.slice(0,60) + (entry.length>60?'…':'');
  return entry.name ? entry.name : '(uten navn)';
}

function countPlannedAdds(byId){
  if(!addExtraSizes) return new Map();
  const pools = buildTagPools(mapping);
  const planned = new Map();

  for(const [id, arr] of byId.entries()){
    const tileCounts = new Map();
    for(const e of arr){
      const key = `${e.size}|${e.variant||''}|${e.side||''}`;
      tileCounts.set(key, (tileCounts.get(key)||0) + 1);
    }

    const uniqSizes = new Map();
    for(const e of arr){
      const k = `${e.size}|${e.variant||''}`;
      if(!uniqSizes.has(k)) uniqSizes.set(k, {size:e.size,variant:e.variant||null});
    }

    const adds=[];
    for(const sv of uniqSizes.values()){
      const exactLine = pools.getExact(id, sv.size, sv.variant||null);
      const sizeLine  = pools.getSize(id, sv.size);
      const exactAny  = pools.getExact(null, sv.size, sv.variant||null);
      const sizeAny   = pools.getSize(null, sv.size);

      const list = (pools.hasLineIds
        ? (exactLine.length ? exactLine : (sizeLine.length ? sizeLine : (exactAny.length ? exactAny : sizeAny)))
        : (exactAny.length ? exactAny : sizeAny)
      );

      const scriptsCount = list?.length || 0;

      let tilesCount = 0;
      for(const [k,n] of tileCounts.entries()){
        const [sz,vr] = k.split('|');
        if(sz===sv.size && (vr||'')===(sv.variant||'')) tilesCount += n;
      }

      const need = Math.max(0, scriptsCount - tilesCount);
      if(need>0) adds.push({size:sv.size,variant:sv.variant||null,need});
    }

    if(adds.length) planned.set(id, adds);
  }
  return planned;
}

function previewPlannedPlacements(byId) {
  const pools = buildTagPools(mapping);
  const reuse = reuseAssets;

  LOG(`Modus: ${pools.hasLineIds ? 'Linje-spesifikk (ID i CSV).' : 'Global (ingen linje-ID i CSV).'}  •  Erstatt: ${overwriteMode==='all'?'alle':'kun tomme'}  •  Gjenbruk: ${reuse?'på':'av'}  •  Legg til ekstra: ${addExtraSizes?'på':'av'}`);

  const plannedAdds = countPlannedAdds(byId);
  let addTotal = 0;
  plannedAdds.forEach(v=>v.forEach(x=>addTotal+=x.need));
  if(addTotal){
    LOG(`\nPLAN: Legg til ekstra størrelser totalt: ${addTotal}`);
    for(const [id, adds] of plannedAdds.entries()){
      const txt = adds.map(a=>`${a.size}${a.variant?'/'+a.variant:''}×${a.need}`).join(', ');
      LOG(`  • ${id}: ${txt}`);
    }
    LOG('');
  }

  for (const [id, arr] of byId.entries()) {
    const groups = new Map();
    for (const e of arr) {
      const key = `${e.size}|${e.variant||''}|${e.side||''}`;
      if (!groups.has(key)) groups.set(key, { entry: { size:e.size, variant:e.variant||null, side:e.side||null }, tiles: [] });
      if (e.tile) groups.get(key).tiles.push(e.tile);
    }

    LOG(`→ Forhåndsvis ${id}`);
    if (!groups.size) { LOG('   (ingen størrelser funnet)'); continue; }

    for (const { entry, tiles } of groups.values()) {
      const label = [entry.size, entry.variant, entry.side].filter(Boolean).join('/');

      const imgsExact = imagePool.get(`${entry.size}|${entry.side||''}`) || [];
      const imgsAny   = entry.side ? (imagePool.get(`${entry.size}|`) || []) : [];
      const imgs = imgsExact.length ? imgsExact : imgsAny;

      if (imgs.length) {
        const lines = tiles.map((_, i) => {
          const f = reuse ? imgs[i % imgs.length] : (i < imgs.length ? imgs[i] : null);
          return f ? `størrelse#${i+1} ← ${fileName(f)}` : `størrelse#${i+1} ← (ingen bilde)`;
        });
        LOG(`   • ${label}  (bilder) ${tiles.length} størrelse(er)\n      ${lines.join('\n      ')}`);
        continue;
      }

      let poolItems = pools.getExact(id, entry.size, entry.variant || null);
      if (!poolItems.length) poolItems = pools.getSize(id, entry.size);
      if (!poolItems.length) poolItems = pools.getExact(null, entry.size, entry.variant || null);
      if (!poolItems.length) poolItems = pools.getSize(null, entry.size);

      if (!poolItems.length) { LOG(`   • ${label}  (ingen match på bilder eller scripts)`); continue; }

      const lines = tiles.map((_, i) => {
        const payload = reuse ? poolItems[i % poolItems.length] : (i < poolItems.length ? poolItems[i] : null);
        const nice = payload ? tagLabel(payload) : '(ingen script)';
        return `størrelse#${i+1} ← ${nice}`;
      });

      LOG(`   • ${label}  (scripts) ${tiles.length} størrelse(er)\n      ${lines.join('\n      ')}`);
    }
  }
}

/* ========= scan ========= */
let scanned=[]; let selected=new Set(G('ap3p_sel_ids',[]));
function saveSelection(){S('ap3p_sel_ids',Array.from(selected))}
function refreshList(){
  list.innerHTML='';
  if(!scanned.length){list.innerHTML='<div class="muted">Ingen funnet — trykk Skann.</div>';return}
  scanned.forEach(it=>{
    const row=d.createElement('label');row.className='item';
    const cb=d.createElement('input');cb.type='checkbox';cb.checked=selected.has(it.id);
    const id=d.createElement('span');id.className='id';id.textContent=it.id;
    const sz=d.createElement('span');sz.className='sizes';sz.textContent=it.label;
    row.append(cb,id,sz);list.appendChild(row);
    cb.onchange=()=>{if(cb.checked)selected.add(it.id);else selected.delete(it.id);saveSelection()}
  })
}
btnAll.onclick=()=>{selected=new Set(scanned.map(s=>s.id));saveSelection();refreshList()};
btnNone.onclick=()=>{selected=new Set();saveSelection();refreshList()};
btnInv.onclick=()=>{const next=new Set();scanned.forEach(s=>{if(!selected.has(s.id)) next.add(s.id)});selected=next;saveSelection();refreshList()};

let stopping=false;

btnScan.onclick=async()=>{
  log.textContent='';
  scanned.length=0;

  stopping=false;
  btnRun.disabled=true; btnStop.disabled=false; btnScan.disabled=true;
  showToast('Skanner…', 'Finner linjer og størrelser');

  const rows=findRows();
  const byId=new Map();

  for (const r of rows) {
    if (stopping) break;

    await ensureCampaignView();
    await waitForIdle();

    const id  = rowId(r);
    const det = await expandRow(r);

    await waitForIdle();
    await ensureAllTilesMounted(det || d);
    await sleep(80);

    let entries = sizesFromExpanded(det);
    const prev = byId.get(id) || [];
    byId.set(id, prev.concat(entries));

    await sleep(40);
  }

  const report=[];
  for(const [id,arr] of byId.entries()){
    const pretty=prettyCounts(arr.map(({size,variant})=>({size,variant})));
    LOG(`• ${id}  størrelser:[${pretty||'-'}]`);
    const uniq=new Map();
    for(const e of arr){
      const k=`${e.size}|${e.variant||''}|${e.side||''}`;
      if(!uniq.has(k)) uniq.set(k,{size:e.size,variant:e.variant,side:e.side})
    }
    report.push({id,entries:[...uniq.values()],label:pretty||'-'})
  }

  LOG(`\nFant ${report.length} linje(r).`);
  try{previewPlannedPlacements(byId)}catch(e){LOG('! Forhåndsvis feilet: '+(e?.message||e))}

  _lastPreview = byId;
  setInfo({items:report.length,done:0,hit:0,skip:0,err:0});
  scanned=report.slice();
  selected=new Set(report.map(r=>r.id));
  saveSelection(); refreshList();

  btnRun.disabled=false; btnStop.disabled=true; btnScan.disabled=false;
  showToast('Skann ferdig', `Linjer: ${report.length}`);
};

/* ========= run ========= */
btnStop.onclick=()=>{
  stopping=true;
  btnStop.disabled=true;
  btnRun.disabled=false;
  showToast('Stopper…', 'Avbryter etter nåværende steg');
};

btnRun.onclick=()=>{
  const ids=[...new Set(scanned.filter(s=>selected.has(s.id)).map(s=>s.id))];
  if(!ids.length){LOG('Ingen valgt — trykk Skann og velg linjer.');return}
  if(!mapping.length&&imagePool.size===0){LOG('Ingen mapping — importer CSV/JSON eller bilder først.');return}
  runOnIds(ids)
};

/* ========= (BEST EFFORT) add extra tiles (dropdown-menu) ========= */
/**
 * For din UI:
 * - Knapp: "Legg til valgfri materiell"
 * - Klikk åpner <div role="menu"> med <div role="menuitem">580x500</div>
 * - Vi klikker direkte på menuitem (ingen combobox/select i dette UI-et)
 */

// ---- add size (LINJE-SCOPED) ----
// Klikker "Legg til valgfri materiell" i riktig linje,
// finner riktig "portal"-meny som dukker opp etter klikk,
// klikker ønsket størrelse INNI den menyen, og verifiserer at linja fikk +1 tile.

function isVisibleMenu(el){
  if(!el) return false;
  const cs = getComputedStyle(el);
  if (cs.display === "none" || cs.visibility === "hidden") return false;
  const r = el.getBoundingClientRect();
  return r.width > 0 && r.height > 0;
}

function countTilesInCollapse(collapseRow){
  return collapseRow
    ? collapseRow.querySelectorAll(".set-creativeTile, .set-creativeTile__selected").length
    : 0;
}

function getAddOptionalBtn(collapseRow){
  if(!collapseRow) return null;
  const btns = Array.from(collapseRow.querySelectorAll("button")).filter(vis);
  return btns.find(b => (b.textContent || "").includes("Legg til valgfri materiell")) || null;
}

async function addSizeScopedToLine(dataRow, sizeText){
  // 1) Sørg for at linja er åpen og vi har riktig collapseRow
  const collapseRow = await expandRow(dataRow);
  if (!collapseRow || !collapseRow.querySelector?.(".set-matSpecCreativeSection")) {
    return { ok:false, reason:"Linja ble ikke åpnet / fant ikke creative-seksjon" };
  }

  await ensureAllTilesMounted(collapseRow);
  await sleep(120);

  const btn = getAddOptionalBtn(collapseRow);
  if (!btn) return { ok:false, reason:"Fant ikke 'Legg til valgfri materiell' i den åpne linja" };

  const tilesBefore = countTilesInCollapse(collapseRow);

  // 2) Snapshot av synlige menyer før klikk
  const visibleMenusBefore = Array.from(document.querySelectorAll('[role="menu"]')).filter(isVisibleMenu);

  btn.scrollIntoView({ block:'center' });
  btn.click();

  // 3) Finn menyen robust:
  // - ny synlig meny som dukket opp
  // - eller en synlig meny som inneholder ønsket sizeText
  const menu = await waitFor(() => {
    const menusNow = Array.from(document.querySelectorAll('[role="menu"]')).filter(isVisibleMenu);

    // prøv: ny meny
    const newOnes = menusNow.filter(m => !visibleMenusBefore.includes(m));
    const candidates = newOnes.length ? newOnes : menusNow;

    // velg den som faktisk inneholder sizeText (best signal)
    const want = String(sizeText).trim();
    const hit = candidates.find(m => (m.innerText || "").split("\n").some(line => line.trim() === want));
    if (hit) return hit;

    // fallback: siste synlige meny
    return candidates.length ? candidates[candidates.length - 1] : null;
  }, 3000, 80);

  if (!menu) return { ok:false, reason:"Meny dukket ikke opp" };

  // 4) Klikk riktig item
  const want = String(sizeText).trim();
  const wanted = Array.from(menu.querySelectorAll('[role="menuitem"], [role="option"], button, div'))
    .find(el => (el.textContent || "").trim() === want);

  if (!wanted) {
    document.body.click();
    return { ok:false, reason:`Fant ikke størrelse '${sizeText}' i menyen` };
  }

  wanted.scrollIntoView({ block:'center' });
  wanted.click();

  // 5) Verifiser at DENNE linja fikk ny tile
  const ok = await waitFor(() => countTilesInCollapse(collapseRow) >= tilesBefore + 1, 4000, 100);
  if (!ok) return { ok:false, reason:"Klikk utført, men ingen ny tile dukket opp i denne linja" };

  return { ok:true };
}




function countTilesForSizeVariant(detailsList, size, variant){
  const want = cs(size);
  const vWant = variant || null;
  const seen = new Set();
  let n = 0;

  for (const det of (detailsList || [])) {
    for (const t of getAllTiles(det)) {
      if (seen.has(t)) continue;
      if (!tileHasSize(t, want)) continue;

      const tv = tileVariant(t) || null;
      if (vWant && tv !== vWant) continue;

      seen.add(t);
      n++;
    }
  }
  return n;
}


async function rebuildDetailsAndEntries(rowsForId){
  // Expand all rows for this ID and collect all "collapse rows" (details)
  const detailsList = [];

  for (const r of (rowsForId || [])) {
    if (!r) continue;

    await ensureCampaignView();
    await waitForIdle();

    const det = await expandRow(r);
    if (det) {
      detailsList.push(det);
      await waitForIdle();
      await ensureAllTilesMounted(det);
      await sleep(80);
    }
  }

  // Collect entries (size/variant/side) from all details
  const all = [];
  for (const det of detailsList) {
    try {
      all.push(...sizesFromExpanded(det));
    } catch {}
  }

  // De-dupe: keep only unique size+variant+side combos (but still preserve a representative tile)
  const uniq = new Map();
  for (const e of all) {
    const k = `${e.size}|${e.variant || ''}|${e.side || ''}`;
    if (!uniq.has(k)) uniq.set(k, { size: e.size, variant: e.variant || null, side: e.side || null, tile: e.tile || null });
  }

  return {
    detailsList,
    entries: [...uniq.values()]
  };
}


async function runOnIds(ids){
  stopping=false;
  btnRun.disabled=true; btnStop.disabled=false; btnScan.disabled=true;
  const stats={items:ids.length,done:0,hit:0,skip:0,err:0};setInfo(stats);
  LOG(`Starter… (${ids.length} linje(r))`);
  showToast('Autofyll startet', `Linjer: ${ids.length}`);
  const pools=buildTagPools(mapping);

  for(const id of ids){
    if(stopping) break;

    try{
      const rowsForId=findRowsByIdAll(id);

      // initial build
      let {detailsList, entries} = await rebuildDetailsAndEntries(rowsForId);

      LOG(`→ ${id} størrelser:[${prettyCounts(entries)||'—'}]`);
      if(!entries.length){stats.skip++;stats.done++;setInfo(stats);continue}

      let wrote=false;

      // --- Optional: add extra sizes if needed ---
      if(addExtraSizes && mapping.length){
        const tileCountBySV = new Map();
        for(const det of detailsList){
          const tiles = getAllTiles(det);
          for(const t of tiles){
            const txt = (t.innerText||'') + ' ' + groupLabel(tileGroup(t));
            const m = txt.match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);
            if(!m) continue;
            const sz = cs(`${m[1]}x${m[2]}`);
            const vr = tileVariant(t)||'';
            const key = `${sz}|${vr}`;
            tileCountBySV.set(key, (tileCountBySV.get(key)||0)+1);
          }
        }

        const scriptsBySV = new Map();
        for(const m of (mapping||[])){
          const idsArr=(m.lineIds||[]).map(String).filter(Boolean);
          if(idsArr.length && !idsArr.includes(String(id))) continue;
          if(!m.size || !m.tag) continue;
          const k = `${m.size}|${m.variant||''}`;
          scriptsBySV.set(k, (scriptsBySV.get(k)||0)+1);
        }

        const needList=[];
        for(const [k, scriptsCount] of scriptsBySV.entries()){
          const [sz, vr] = k.split('|');
const tilesCount = countTilesForSizeVariant(detailsList, sz, vr || null);

          const need = Math.max(0, scriptsCount - tilesCount);
          if(need>0){
            const [sz]=k.split('|');
            needList.push({key:k, size:sz, need});
          }
        }

        if(needList.length){
          LOG(`   · Legg-til plan (${id}): ` + needList.map(x=>`${x.size}×${x.need}`).join(', '));
          for(const x of needList){
            if(stopping) break;
            for (let n = 0; n < x.need; n++){
  if (stopping) break;

  await ensureCampaignView();
  await waitForIdle();

  // bruk EN av radene for denne id-en (vanligvis bare 1)
  const dataRow = rowsForId[0];
  if (!dataRow) { LOG(`   ! Fant ikke dataRow for ${id}`); break; }

  LOG(`   · prøver å legge til ${x.size} (linje ${id})`);
  const res = await addSizeScopedToLine(dataRow, x.size);

  if (!res.ok) {
    LOG(`   ! Klarte ikke å legge til ${x.size}: ${res.reason}`);
    break;
  }

  LOG(`   + La til størrelse: ${x.size}`);

  await sleep(350);
  // Rebuild etter add slik at scripts kan matches på de nye tilesene
  ({detailsList, entries} = await rebuildDetailsAndEntries(rowsForId));
}

          }

          ({detailsList, entries} = await rebuildDetailsAndEntries(rowsForId));
          LOG(`   · Etter legg-til: størrelser=[${prettyCounts(entries)||'—'}]`);
        }
      }

      for(const entry of entries){
        if(stopping) break;

        await ensureCampaignView();
        await waitForIdle();

        let tiles = [];
        for (const det of detailsList) tiles = tiles.concat(pickTilesForEntry(det, entry));
        tiles = Array.from(new Set(tiles));

        if (!tiles.length) {
          await ensureAllTilesMounted();
          tiles = [];
          for (const det of detailsList) tiles = tiles.concat(pickTilesForEntry(det, entry));
          tiles = Array.from(new Set(tiles));
        }

        if (!tiles.length) {
          LOG(`   • ${entry.size}${entry.side?('/'+entry.side):''} (ingen matchende størrelser)`);
          continue;
        }

        const imgs = imageChoicesForEntry(entry);
        if (imgs.length) {
          let placed = 0;
          for (let i = 0; i < tiles.length; i++) {
            if(stopping) break;
            const t = tiles[i];
            const f = reuseImages() ? imgs[i % imgs.length] : (i < imgs.length ? imgs[i] : null);
            if (!f) continue;

            t.scrollIntoView({ block: 'center' });
            let ok = await uploadImageToTile(t, f);
            await sleep(160);
            if(ok) placed++;
          }
          LOG(`   • ${entry.size}${entry.side?('/'+entry.side):''}  lagt inn bilder=${placed}${placed<tiles.length?` (hoppet over ${tiles.length-placed})`:''}`);
          wrote = placed > 0 || wrote;
          continue;
        }

        const exactLine = pools.getExact(id, entry.size, entry.variant || null);
        const sizeLine  = pools.getSize(id, entry.size);
        const exactAny  = pools.getExact(null, entry.size, entry.variant || null);
        const sizeAny   = pools.getSize(null, entry.size);

        let list;
        if (pools.hasLineIds) {
          if (exactLine.length) list = exactLine;
          else if (sizeLine.length) list = sizeLine;
          else if (exactAny.length) list = exactAny;
          else list = sizeAny;
        } else {
          list = exactAny.length ? exactAny : sizeAny;
        }

        if (!list || !list.length) {
          LOG(`   • ${entry.size}${entry.variant?('/'+entry.variant):''} (ingen script for denne størrelsen/linjen)`);
          continue;
        }

        LOG(`   • ${entry.size}${entry.variant?('/'+entry.variant):''}  størrelser=${tiles.length}, scripts=${list.length}`);

        let used = 0, skipped = 0;

        for (let i = 0; i < tiles.length; i++) {
          if(stopping) break;

          const payload = reuseAssets ? list[i % list.length] : (i < list.length ? list[i] : null);
          if (!payload) { skipped++; continue; }

          if (i > 0 && (i % 8 === 0)) await checkpoint(`størrelse ${i}/${tiles.length}`);

          await ensureCampaignView();
          await waitForIdle();
          await selectTile(tiles[i]);

          const editor = await open3PTab();
          if (!editor) { LOG('   ! Fant ikke 3rd-party editor'); stats.err++; continue; }

          if (overwriteMode === 'empty') {
            const current = getEditorValue(editor);
            if (!isEmpty3PValue(current)) { skipped++; continue; }
          }

          const tagStr = (typeof payload === 'string') ? payload : (payload?.tag || '');
          const pasted = await pasteWithVerify(editor, tagStr, 2);
          if (!pasted) { LOG('   ! Lim inn festet ikke (etter retries) — hopper over størrelse'); stats.err++; continue; }

          used++;

          if (G('ap3p_auto', true)) {
            const saved = await saveWithRetry(editor, 2);
            if (!saved) LOG('   ! Lagre/Oppdater/Send inn på nytt feilet (fortsetter)');
          }

          await sleep(160);
        }

        LOG(`     → brukt scripts=${used}${skipped?` (hoppet over ${skipped})`:''}`);
        wrote = used > 0 || wrote;
      }

      if(wrote) stats.hit++; else stats.skip++;
      stats.done++; setInfo(stats);
      await sleep(160);

    }catch(e){
      LOG('   ! feil: '+(e&&e.message?e.message:e));
      stats.err++;stats.done++;setInfo(stats)
    }
  }

  LOG(`Ferdig. Linjer=${stats.items}  Treff=${stats.hit}  Hoppet=${stats.skip}  Feil=${stats.err}`);
  showToast('Ferdig', `Treff: ${stats.hit} • Feil: ${stats.err}`);
  btnStop.disabled=true; btnRun.disabled=false; btnScan.disabled=false;
}

/* ========= niceties ========= */
_bestiltInt=setInterval(()=>{
  const panel=d.querySelector('#InfoPanelContainer');
  const t=panel?.innerText||'';
  const m=t.match(/Bestilt(?:\s*\(BxH\))?\s*[:\-]?\s*(\d{2,4})\s*[x×]\s*(\d{2,4})/i);
  badge.textContent=m?`— ${m[1]}x${m[2]}`:''
},900);

w.addEventListener('keydown',e=>{if(e.altKey&&e.key.toLowerCase()==='a'){e.preventDefault();btnRun.click()}});

function inCampaignTableView(){ return !!document.querySelector('tr.set-matSpecDataRow'); }

function findBackToCampaignButton(){
  const btns = [...document.querySelectorAll('button, [role="button"]')];
  for (const b of btns) {
    const p = b.querySelector('svg path');
    const dd = p?.getAttribute?.('d') || '';
    if (dd.startsWith('M9.4 16.6L4.8 12')) return b;
  }
  return null;
}
async function ensureCampaignView(){
  if (inCampaignTableView()) return true;
  const back = findBackToCampaignButton();
  if (back) {
    back.scrollIntoView({ block:'center' });
    back.click();
    await waitFor(() => inCampaignTableView(), 8000, 120);
  }
  return inCampaignTableView();
}
async function waitForIdle(timeout=12000){
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const backdrop = document.querySelector('.MuiBackdrop-root');
    const progress = document.querySelector('[role="progressbar"], .MuiCircularProgress-root');

    const backdropVisible = backdrop &&
      getComputedStyle(backdrop).visibility !== 'hidden' &&
      getComputedStyle(backdrop).display !== 'none' &&
      getComputedStyle(backdrop).opacity !== '0';

    const progressVisible = progress &&
      getComputedStyle(progress).visibility !== 'hidden' &&
      getComputedStyle(progress).display !== 'none' &&
      getComputedStyle(progress).opacity !== '0';

    if (!backdropVisible && !progressVisible) return true;
    await sleep(120);
  }
  return true;
}

/* restore minimized state if previously minimized */
(function restoreMinState(){
  const st=loadUIState();
  if(st?.min) setMinimized(true);
})();

}catch(e){console.error(e);alert('Autofill-feil: '+(e&&e.message?e.message:e));}})();
