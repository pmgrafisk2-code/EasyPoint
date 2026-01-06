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
// Find ALL line IDs like "... L1173675 ..." anywhere in the name
// Extract ALL line IDs anywhere in the string, tolerant to spacing/underscore
function extractLineIds(str){
  const ids = new Set();
  // L + optional spaces/nbsp/_/- then 6–10 digits; stop before another digit
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
    scroller.scrollBy(0, 1e6);           // jump down
    await sleep(300);
    const count = d.querySelectorAll('.set-creativeTile').length;
    if (count === prev) still++; else { prev = count; still = 0; }
  }
  scroller.scrollTo(0, 0);               // return to top
}




/* ========= image side helpers ========= */
function detectSide(str){const s=(str||'').toLowerCase().replace('høyre','hoyre');if(/\b(hoyre|right|r)\b/.test(s))return'right';if(/\b(venstre|left|l)\b/.test(s))return'left';return null}

/* ========= tile helpers ========= */
function tileClickTargets(tile){
  // Prefer label/text areas over the media/thumbnail area (thumbnail often opens preview)
  return [
    tile.querySelector('.set-cardLabel__main'),          // best: label main
    tile.querySelector('.set-cardLabel'),               // label container
    tile.querySelector('.MuiTypography-root'),          // any visible text area
    tile.querySelector('[aria-label*="select" i]'),      // if they have a dedicated select control
    tile                                                     // fallback
  ].filter(Boolean);
}

async function selectTile(tile){
  if (!tile) return false;

  // If already selected, we’re done
  if (isTileSelected(tile)) return true;

  const targets = tileClickTargets(tile);

  // Try a few times (MUI sometimes ignores the first click when busy)
  for (let attempt = 0; attempt < 3; attempt++) {
    for (const t of targets) {
      if (isTileSelected(tile)) return true;

      try { t.scrollIntoView({ block: 'center' }); } catch {}
      // Click sequence that tends to trigger selection without "opening"
      t.dispatchEvent(new MouseEvent('pointerdown', { bubbles:true }));
      t.dispatchEvent(new MouseEvent('mousedown',  { bubbles:true }));
      t.dispatchEvent(new MouseEvent('mouseup',    { bubbles:true }));
      t.click();

      for (let i = 0; i < 10; i++) {
        await sleep(80);
        if (isTileSelected(tile)) return true;
      }
    }

    // Keyboard fallback: focus tile and press Space (often "selects" without navigation)
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

function isTileSelected(tile){if(tile.classList.contains('set-creativeTile__selected'))return true;if(tile.getAttribute('aria-selected')==='true')return true;if(tile.closest('[aria-selected="true"]'))return true;if(tile.closest('.set-creativeTile__selected'))return true;return false}
function getAllTiles(details){
  let tiles = [...details.querySelectorAll('.set-creativeTile, .set-creativeTile__selected')];
  // only keep top-level tiles (no nested dups)
  return tiles.filter(t => !tiles.some(o => o !== t && o.contains(t)));
}
function tileGroup(tile){return tile.closest('.set-creativeContainer, .set-creativeSectionContainer')||tile.parentElement}
function groupLabel(group){const n=group&&(group.querySelector('.set-cardLabel__main [title]')||group.querySelector('.set-cardLabel__main div[title]')||group.querySelector('.set-cardLabel__main'));return(n?.getAttribute?.('title')||n?.textContent||'')||''}
function tileVariant(tile){const lbl=groupLabel(tileGroup(tile)).toLowerCase();if(/desktop/.test(lbl))return'desktop';if(/mobil|mobile/.test(lbl))return'mobil';return null}
function tileSide(tile){return detectSide(groupLabel(tileGroup(tile)))}
function tileHasSize(tile,size){const rx=new RegExp(`\\b${size.replace('x','[x×]')}\\b`,'i');if(rx.test((tile.innerText||'')))return true;const lbl=groupLabel(tileGroup(tile));const mappedRe=new RegExp(`mapped_${size.replace('x','[x×]')}(?:\\b|[_-])`,'i');if(mappedRe.test(lbl))return true;if(rx.test(lbl))return true;return false}

/* ========= mapping import (CSV/JSON + Excel paste + DnD) ========= */
function parseCSV(text){let rows=[],row=[],f='',q=false;const first=(text.split(/\r?\n/).find(l=>l.trim().length>0)||'');const c1=(first.match(/,/g)||[]).length,c2=(first.match(/;/g)||[]).length,c3=(first.match(/\t/g)||[]).length;const del=c3>=c2&&c3>=c1?'\t':(c2>c1?';':',');for(let i=0;i<text.length;i++){const ch=text[i];if(ch=='"'){if(q&&text[i+1]=='"'){f+='"';i++}else q=!q}else if(!q&&ch===del){row.push(f);f=''}else if(!q&&ch=='\n'){row.push(f);rows.push(row);row=[];f=''}else if(!q&&ch=='\r'){}else f+=ch}row.push(f);rows.push(row);return rows}
function pickHeader(h,names){const L=h.map(x=>String(x||'').toLowerCase());for(const n of names){let i=L.indexOf(n);if(i>-1)return i;for(let k=0;k<L.length;k++) if(L[k].includes(n)) return k}return -1}
function rowsToMap(rows){let head=[],headRow=0;for(let i=0;i<Math.min(rows.length,5);i++){const cells=(rows[i]||[]).map(c=>String(c||'').trim());if(cells.join('').length){head=cells;headRow=i;break}}const sizeIdx=pickHeader(head,['størrelse','stoerrelse','size']);const nameIdx=pickHeader(head,['creative name','creative','name','navn']);const tagIdx=pickHeader(head,['secure content','script','tag','kode','code']);if(tagIdx<0||(sizeIdx<0&&nameIdx<0))throw new Error('Missing columns: Secure Content + (Size or Creative Name)'); let warned=false; const out=[];
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

    // primary: from Creative Name
    let lineIds = extractLineIds(rawName);
    // fallback: scan the entire row text if none found
    if (!lineIds.length){
      const rowText = cells.map(c => String(c||'')).join(' _ ');
      lineIds = extractLineIds(rowText);
    }

    if (svSize.size && svName.size && cs(svSize.size)!==cs(svName.size) && !warned){
      LOG('! CSV warning: Size column disagrees with Creative Name — trusting the Size column.');
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
      LOG('! JSON warning: Size field disagrees with name — trusting the Size field.');
      warned=true;
    }
    if (size && tag) out.push({ size, variant, tag, name: rawName, vkey, lineIds });
  });
  return out;
}

/* ========= line items ========= */
function normLabel(s){return String(s||'').toLowerCase().replace(/\s+/g,' ').trim()}
function findRows(){const rows=[...d.querySelectorAll('tr.set-matSpecDataRow')].filter(vis);return rows.filter(el=>!rows.some(o=>o!==el&&o.contains(el)))}
function getHeaderIdxForRow(row){const table=row.closest('table');if(!table) return{lineItemIdx:-1,rosenrIdx:-1};if(table._ap3p_headerIdx) return table._ap3p_headerIdx;const ths=[...(table.tHead?.querySelectorAll('th')||[])];let lineItemIdx=-1,rosenrIdx=-1;ths.forEach((th,i)=>{const t=normLabel(th.textContent);if(lineItemIdx===-1&&/(^|\s)line\s*item(\s|$)/.test(t)) lineItemIdx=i;if(rosenrIdx===-1&&(/rosenr/.test(t)||(/materiell/.test(t)&&/rosenr/.test(t)))) rosenrIdx=i});table._ap3p_headerIdx={lineItemIdx,rosenrIdx};return table._ap3p_headerIdx}
function rowId(row){const {lineItemIdx,rosenrIdx}=getHeaderIdxForRow(row);const tds=[...row.querySelectorAll('td')];if(lineItemIdx>-1&&tds[lineItemIdx]){const txt=tds[lineItemIdx].textContent.trim();const m=txt.match(/\d{6,10}/);return(m&&m[0])||(txt||'?')}const rosenVal=(rosenrIdx>-1&&tds[rosenrIdx])?tds[rosenrIdx].textContent.trim():'';const nums=tds.map(td=>td.textContent.trim()).filter(t=>/^\d{6,10}$/.test(t));const hit=nums.find(n=>n!==rosenVal);return hit||nums[0]||'?'}
function findRowsByIdAll(targetId){return findRows().filter(r=>rowId(r)===targetId)}
async function expandRow(row){if(!row) return null;row.scrollIntoView({block:'center'});row.dispatchEvent(new MouseEvent('pointerdown',{bubbles:true}));row.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));row.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));row.click();await sleep(120);row.click();await sleep(160);for(let i=0;i<14;i++){const sib=row.nextElementSibling;if(sib&&(sib.querySelector('.set-creativeSectionContainer, .set-creativeContainer'))) return sib;await sleep(90)}return row.nextElementSibling||null}

/* ========= detect placeholders ========= */
function sizesFromExpanded(details){if(!details) return[];const entries=[];const tiles=getAllTiles(details);function findSizeInText(s){const m=String(s||'').match(/(\d{2,4})\s*[x×]\s*(\d{2,4})/i);return m?`${m[1]}x${m[2]}`:''}function findSizeInMapped(lbl){const m=String(lbl||'').match(/mapped_(\d{2,4})\s*[x×]\s*(\d{2,4})(?:\b|[_\s])/i);return m?`${m[1]}x${m[2]}`:''}for(const tile of tiles){const text=(tile.innerText||'');const lbl=groupLabel(tileGroup(tile));let size=findSizeInText(text);if(!size) size=findSizeInMapped(lbl);if(!size) size=findSizeInText(lbl);if(!size) continue;const variant=tileVariant(tile);const side=tileSide(tile);entries.push({size:cs(size),variant,side,tile})}return entries}
function pickTilesForEntry(details,entry){let tiles=getAllTiles(details).filter(t=>tileHasSize(t,entry.size));if(entry.variant) tiles=tiles.filter(t=>(tileVariant(t)||null)===entry.variant);if(entry.side) tiles=tiles.filter(t=>(tileSide(t)||null)===entry.side);return tiles}

/* ========= 3rd-party tab + save/reprocess ========= */
async function open3PTab(){const t=[...d.querySelectorAll('button[role="tab"],a[role="tab"],button,a')].filter(vis).find(b=>/\b3(?:rd)?\s*party\s*tag\b/i.test(b.textContent||''));if(t){t.click()}const editor=await waitFor(()=>{const panel=d.querySelector('#InfoPanelContainer')||d;const cm=[...panel.querySelectorAll('.CodeMirror')].find(vis);if(cm) return{kind:'cm',el:cm,panel,cm:cm.CodeMirror};const ta=[...panel.querySelectorAll('textarea')].find(vis);if(ta) return{kind:'ta',el:ta,panel};return null},4000,120);return editor||null}
function pasteInto(target,value){if(!target) return;if(target.kind==='cm'&&target.cm){try{const cm=target.cm;cm.setValue((value||'')+'');cm.refresh?.();return}catch{}}const ta=target.el||target;ta.scrollIntoView({block:'center'});ta.focus();ta.value=value;ta.dispatchEvent(new Event('input',{bubbles:true}));ta.dispatchEvent(new Event('change',{bubbles:true}))}
async function clickSaveOrReprocess(scope){const root=scope?.panel||d.querySelector('#InfoPanelContainer')||d;const btn=[...root.querySelectorAll('button')].filter(vis).find(b=>{const t=(b.textContent||'').toLowerCase();return /lagre|save|oppdater|update|send inn på nytt|reprocess/i.test(t)&&!b.disabled&&!b.classList.contains('Mui-disabled')});if(btn){btn.scrollIntoView({block:'center'});btn.click();await sleep(700);return true}return false}


/* ========= UI ========= */
(function injectCSS(){const st=d.createElement('style');st.textContent=`
#ap3p_bar{box-sizing:border-box}
#ap3p_list .item{display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:6px}
#ap3p_list .item:hover{background:#121521}
#ap3p_list .id{min-width:72px;opacity:.85}
#ap3p_list .sizes{opacity:.9;font-family:ui-monospace,Consolas,monospace}
#ap3p_list .muted{opacity:.6}
#ap3p_bar .chip{padding:3px 8px;border-radius:999px;border:1px solid;font-size:12px;opacity:.95}
#ap3p_bar .chip-ok{background:#064e3b;border-color:#10b981;color:#d1fae5}
#ap3p_bar .chip-none{background:#1a1d27;border-color:#2a2d37;color:#e6e6e6}
/* Resize handles */
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
`;d.head.appendChild(st)})();

let ui=d.getElementById('ap3p_bar'); if(ui) ui.remove();
ui=d.createElement('div'); ui.id='ap3p_bar';
ui.style.cssText='position:fixed;right:16px;bottom:16px;z-index:2147483647;background:#0e0f13;color:#e6e6e6;border:1px solid #2a2d37;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.45);min-width:780px;max-width:96vw;';

const hdr=d.createElement('div');
hdr.style.cssText='cursor:move;user-select:none;display:flex;align-items:center;gap:10px;padding:8px 10px;background:#161922;border-radius:12px 12px 0 0;border-bottom:1px solid #2a2d37';
const title=d.createElement('div'); title.textContent='EasyPoint';
const badge=d.createElement('span'); badge.style.opacity='.8'; badge.style.marginLeft='6px'; badge.textContent='';
const mapChip=d.createElement('span'); mapChip.className='chip chip-none';
const mapClear=d.createElement('button'); mapClear.textContent='×'; mapClear.title='Clear mapping'; mapClear.style.cssText='margin-left:4px;border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:22px;height:22px;border-radius:6px;cursor:pointer';
const flex=d.createElement('div'); flex.style.flex='1';
const btnScan=d.createElement('button'); btnScan.textContent='Scan'; btnScan.style.cssText='cursor:pointer;border:1px solid #0284c7;border-radius:8px;padding:6px 10px;background:#0ea5e9;color:#041014;font-weight:600';
const btnRun=d.createElement('button'); btnRun.textContent='Autofill'; btnRun.style.cssText='cursor:pointer;border:1px solid #15803d;border-radius:8px;padding:6px 10px;background:#22c55e;color:#051b0a;font-weight:700';
const btnStop=d.createElement('button'); btnStop.textContent='Stop'; btnStop.disabled=true; btnStop.style.cssText='cursor:pointer;border:1px solid #b91c1c;border-radius:8px;padding:6px 10px;background:#ef4444;color:#ffffff;font-weight:700';
const btnMin=d.createElement('button'); btnMin.textContent='–'; btnMin.title='Minimize'; btnMin.style.cssText='border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:26px;height:26px;border-radius:7px;margin-left:6px;cursor:pointer';
const btnX=d.createElement('button'); btnX.textContent='×'; btnX.title='Close'; btnX.style.cssText='border:1px solid #2a2d37;background:#1a1d27;color:#e6e6e6;width:26px;height:26px;border-radius:7px;cursor:pointer';
hdr.append(title,badge,mapChip,mapClear,flex,btnScan,btnRun,btnStop,btnMin,btnX);

const body=d.createElement('div');
body.style.cssText='padding:8px 10px;display:flex;gap:10px;align-items:center;flex-wrap:wrap';
const info=d.createElement('span'); info.textContent='Items: —'; info.style.opacity='.9';
const btnCSV=d.createElement('button'); btnCSV.textContent='Import CSV/JSON'; btnCSV.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:8px;padding:6px 10px;background:#1a1d27;color:#e6e6e6';
const btnPaste=d.createElement('button'); btnPaste.textContent='Paste from Excel'; btnPaste.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:8px;padding:6px 10px;background:#1a1d27;color:#e6e6e6';
const file=d.createElement('input'); file.type='file'; file.accept='.csv,.json,.txt'; file.style.display='none';
const drop=d.createElement('div'); drop.textContent='Drop CSV/JSON/Images here'; drop.style.cssText='flex:1 1 100%;border:1px dashed #3a3f52;border-radius:8px;height:64px;display:flex;align-items:center;justify-content:center;opacity:.9';
const listWrap=d.createElement('div'); listWrap.style.cssText='flex:1 1 100%';
const listCtrls=d.createElement('div'); listCtrls.id='ap3p_list_ctrls'; listCtrls.style.cssText='display:flex;align-items:center;gap:8px;margin:2px 0 4px 0;opacity:.9';
const btnAll=d.createElement('button'); btnAll.textContent='Select all';
const btnNone=d.createElement('button'); btnNone.textContent='Select none';
const btnInv=d.createElement('button'); btnInv.textContent='Invert';
[btnAll,btnNone,btnInv].forEach(b=>b.style.cssText='cursor:pointer;border:1px solid #2a2d37;border-radius:6px;padding:4px 8px;background:#1a1d27;color:#e6e6e6');
listCtrls.append(btnAll,btnNone,btnInv);
const list=d.createElement('div'); list.id='ap3p_list'; list.style.cssText='max-height:180px;overflow:auto;background:#0b0c10;border:1px solid #1b1e28;border-radius:8px;padding:6px';
listWrap.append(listCtrls,list);
const log=d.createElement('div'); log.style.cssText='height:220px;overflow:auto;background:#0b0c10;border-top:1px solid #1b1e28;border-radius:0 0 12px 12px;padding:8px 10px;font-family:ui-monospace,Consolas,monospace;white-space:pre-wrap';
ui.append(hdr,body,listWrap,log); d.body.appendChild(ui);

const LOG=m=>{log.textContent+=m+'\n'; log.scrollTop=log.scrollHeight;};
const setInfo=o=>{info.textContent=`Items:${o.items}  Done:${o.done}  Hit:${o.hit}  Skip:${o.skip}  Err:${o.err}`};
body.append(info,btnCSV,btnPaste,file,drop);

/* ===== Drag, Resize, Minimize, Close ===== */
function adjustHeights(totalH){
  if(ui.getAttribute('data-min')==='1') return;
  const hdrH=hdr.getBoundingClientRect().height;
  const bodyH=body.offsetParent?body.getBoundingClientRect().height:0;
  const lstH=(listWrap.style.display==='none')?0:listWrap.getBoundingClientRect().height;
  const pad=24;
  const logH=Math.max(100,totalH-(hdrH+bodyH+lstH+pad));
  log.style.height=logH+'px';
  ui.style.overflow='hidden';
}
(function makeDraggable(handle,box){
  let sx=0,sy=0,ox=0,oy=0,drag=false;
  handle.addEventListener('pointerdown',e=>{ if(e.button!==0) return; drag=true; const r=box.getBoundingClientRect(); ox=r.left; oy=r.top; sx=e.clientX; sy=e.clientY; box.style.left=ox+'px'; box.style.top=oy+'px'; box.style.right='auto'; box.style.bottom='auto' });
  w.addEventListener('pointermove',e=>{if(!drag)return;box.style.left=(ox+e.clientX-sx)+'px';box.style.top=(oy+e.clientY-sy)+'px';});
  w.addEventListener('pointerup',()=>drag=false);
})(hdr,ui);
(function makeResizable(box){
  const dirs=['n','s','e','w','ne','nw','se','sw']; const els={};
  dirs.forEach(dn=>{const h=d.createElement('div');h.className='aprs aprs-'+dn;box.appendChild(h);els[dn]=h;});
  let rs=null;
  function onDown(e,dir){ if(e.button!==0)return; e.preventDefault(); const r=box.getBoundingClientRect(); rs={dir,sx:e.clientX,sy:e.clientY,x:r.left,y:r.top,w:r.width,h:r.height}; w.addEventListener('pointermove',onMove); w.addEventListener('pointerup',onUp); }
  function onMove(e){ if(!rs)return; const dx=e.clientX-rs.sx, dy=e.clientY-rs.sy; let x=rs.x,y=rs.y,wv=rs.w,hv=rs.h;
    if(/e/.test(rs.dir)) wv=Math.max(420,rs.w+dx);
    if(/s/.test(rs.dir)) hv=Math.max(120,rs.h+dy);
    if(/w/.test(rs.dir)){ wv=Math.max(420,rs.w-dx); x=rs.x+dx; }
    if(/n/.test(rs.dir)){ hv=Math.max(120,rs.h-dy); y=rs.y+dy; }
    Object.assign(box.style,{width:wv+'px',height:hv+'px',left:x+'px',top:y+'px',right:'auto',bottom:'auto'}); adjustHeights(hv);
  }
  function onUp(){ rs=null; w.removeEventListener('pointermove',onMove); w.removeEventListener('pointerup',onUp); }
  dirs.forEach(dn=>els[dn].addEventListener('pointerdown',e=>onDown(e,dn)));
  box._resizerEls=Object.values(els);
})(ui);
function setMinimized(min){
  ui.setAttribute('data-min',min?'1':'0');
  body.style.display=min?'none':'flex';
  listWrap.style.display=min?'none':'block';
  log.style.display=min?'none':'block';
  if(min){
    ui.style.height='38px'; ui.style.minWidth='0'; ui.style.width='max-content'; ui.style.maxWidth='96vw'; ui.style.overflow='hidden';
    (ui._resizerEls||[]).forEach(el=>el.style.display='none'); btnMin.textContent='+';
  }else{
    ui.style.height=''; ui.style.minWidth='780px'; ui.style.width=''; ui.style.overflow='';
    (ui._resizerEls||[]).forEach(el=>el.style.display=''); btnMin.textContent='–'; adjustHeights(ui.getBoundingClientRect().height);
  }
}
btnMin.onclick=()=>setMinimized(ui.getAttribute('data-min')!=='1');
let _bestiltInt=null;
btnX.onclick=()=>{ if(_bestiltInt) clearInterval(_bestiltInt); ui.remove(); };
adjustHeights(ui.getBoundingClientRect().height);

/* ========= mapping state ========= */
const MAP_KEY='ap3p_map_v2';
let mapping=G(MAP_KEY,[]);
let imagePool=new Map(); // key `${size}|${side||''}` -> [File,...]
function renderMapChip(){const csvCount=mapping.length;let imgCount=0;imagePool.forEach(a=>imgCount+=a.length);mapChip.className=(csvCount||imgCount)?'chip chip-ok':'chip chip-none';mapChip.textContent=(csvCount||imgCount)?`CSV: ${csvCount}  Images: ${imgCount} ✓`:'CSV/Images: none'}
function saveMap(arr){mapping=arr;S(MAP_KEY,mapping);renderMapChip()}
function clearMap(){mapping=[];S(MAP_KEY,mapping);renderMapChip()}
mapClear.onclick=()=>{clearMap();imagePool.clear();renderMapChip()};
renderMapChip();


/* ========= reuse assets setting (images + scripts) ========= */
const REUSE_KEY_NEW = 'ap3p_reuse_assets';
const REUSE_KEY_OLD = 'ap3p_reuse_images';
let reuseAssets = G(REUSE_KEY_NEW, G(REUSE_KEY_OLD, true)); // default ON

const reuseWrap = document.createElement('label');
reuseWrap.style.cssText='display:flex;align-items:center;gap:6px;margin-left:8px;font-size:12px;opacity:.95';
const reuseCb = document.createElement('input');
reuseCb.type='checkbox';
reuseCb.checked = reuseAssets;
const reuseTxt = document.createElement('span');
reuseTxt.textContent = 'Fill all tiles (reuse images/scripts)';
reuseWrap.append(reuseCb, reuseTxt);
body.insertBefore(reuseWrap, drop);

let _lastPreview = null; // if you already have this, keep just one definition

reuseCb.onchange = () => {
  reuseAssets = reuseCb.checked;
  S(REUSE_KEY_NEW, reuseAssets);
  if (_lastPreview) {
    log.textContent = '';
    LOG(`Mode: ${reuseAssets ? 'Reuse to fill all tiles' : "Don't reuse (use each once)"}`);
    previewPlannedPlacements(_lastPreview);
  }
};

const reuseImages = reuseAssets;



/* ========= image helpers ========= */
function getTileDropzoneInput(tile){return tile.querySelector('input[type="file"]')||null}
async function uploadImageToInput(inp, file){
  if (!inp) return false;
  const dt = new DataTransfer();
  dt.items.add(file);
  inp.files = dt.files;
  inp.dispatchEvent(new Event('input',{bubbles:true}));
  inp.dispatchEvent(new Event('change',{bubbles:true}));
  await sleep(300);
  return true;
}


async function confirmReplaceIfPrompted(timeout=5000){
  // Wait until a confirm/overwrite dialog shows up (MUI or similar)
  const dlg = await waitFor(() => {
    const el = document.querySelector('.MuiDialog-container, .MuiDialog-root, [role="dialog"]');
    if (!el) return null;
    const txt = (el.innerText || '').toLowerCase();
    // Norwegian + English keywords to be safe
    if (/(bekreftelse|erstatte|erstatt|overwrite|replace|confirm)/.test(txt)) return el;
    return null;
  }, timeout, 120);

  if (!dlg) return false;

  // Prefer an explicit OK/Yes/Replace button; otherwise click the last button (usually primary)
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

  // give the UI a beat to render the dialog
  await sleep(200);

  // auto-confirm "replace" if prompted
  await confirmReplaceIfPrompted(5000);

  // short settle delay so the upload attaches before moving on
  await sleep(250);
  return true;
}

function parseSizeFromName(name){const m=String(name||'').match(/(\d{2,4})[x×](\d{2,4})/i);return m?`${m[1]}x${m[2]}`:''}
function parseImageMeta(file){return{size:cs(parseSizeFromName(file.name)),side:detectSide(file.name),file}}
function poolKey(size,side){return `${size}|${side||''}`}
function addImagesToPool(files){for(const f of files){const meta=parseImageMeta(f);if(!meta.size) continue;const key=poolKey(meta.size,meta.side);if(!imagePool.has(key)) imagePool.set(key,[]);imagePool.get(key).push(f)}}

/* ========= import UI ========= */
btnCSV.onclick=()=>file.click();
file.onchange=e=>{const f=e.target.files?.[0];if(!f) return;const r=new FileReader();r.onload=ev=>{const text=String(ev.target.result||'');if(f.name.toLowerCase().endsWith('.json')){try{saveMap(normalizeJSON(JSON.parse(text)))}catch{alert('Invalid JSON')}}else{try{saveMap(rowsToMap(parseCSV(text)))}catch(err){alert(err.message)}}};r.readAsText(f,'utf-8');file.value=''};
btnPaste.onclick=()=>{const go=t=>{try{saveMap(rowsToMap(parseCSV(t)))}catch(e){alert(e.message)}};if(navigator.clipboard?.readText){navigator.clipboard.readText().then(go).catch(()=>{const t=prompt('Paste rows (Excel: Creative Name/Size, Secure Content)');if(t) go(t)})}else{const t=prompt('Paste rows (Excel: Creative Name/Size, Secure Content)');if(t) go(t)}};
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='#171a24'}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.style.background='transparent'}));
function isCSV(f){return /\.csv$/i.test(f.name)||f.type==='text/csv'}
function isJSON(f){return /\.json$/i.test(f.name)||f.type==='application/json'}
function readFileText(file){return new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(String(r.result||''));r.onerror=rej;r.readAsText(file,'utf-8')})}
drop.addEventListener('drop',async e=>{e.preventDefault();drop.style.background='transparent';const files=[...(e.dataTransfer?.files||[])];if(!files.length) return;const csvs=files.filter(isCSV),jsons=files.filter(isJSON),imgs=files.filter(f=>/^image\//.test(f.type)||/\.(psd|tiff?|webp)$/i.test(f.name));try{for(const f of csvs){const text=await readFileText(f);saveMap(rowsToMap(parseCSV(text)))}}catch(err){alert('CSV import failed: '+(err?.message||err))}try{for(const f of jsons){const text=await readFileText(f);saveMap(normalizeJSON(JSON.parse(text)))}}catch(err){alert('JSON import failed: '+(err?.message||err))}if(imgs.length) addImagesToPool(imgs);renderMapChip()});

/* ========= pools (scripts) ========= */
function buildTagPools(mapping){
  const tagsByExactGlobal = new Map(); // "size|variant" -> [{tag,name},...]
  const tagsBySizeGlobal  = new Map(); // "size"        -> [{tag,name},...]
  const tagsByExactLine   = new Map(); // "line|size|variant" -> [{tag,name},...]
  const tagsBySizeLine    = new Map(); // "line|size"         -> [{tag,name},...]

  const keyExact = (lineId, size, variant) =>
    (lineId ? `${lineId}|` : '') + `${size}|${variant || ''}`;
  const keySize  = (lineId, size) =>
    (lineId ? `${lineId}|` : '') + size;

  let lineScopedCount = 0, globalCount = 0;

  for (const m of (mapping || [])){
    const ids = Array.isArray(m.lineIds) ? m.lineIds.filter(Boolean) : [];
    const payload = { tag: m.tag, name: m.name || '' };

    if (ids.length){
      // strictly line-scoped: DO NOT mirror into global
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
      // global-only row
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

  // Quick debug summary in the log to verify classification
  LOG(`CSV map: line-scoped=${lineScopedCount}, global=${globalCount}`);

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
    },
    maps: { tagsByExactGlobal, tagsBySizeGlobal, tagsByExactLine, tagsBySizeLine }
  };
}





/* ========= preview ========= */
function imageChoicesForEntry(entry){const exact=imagePool.get(`${entry.size}|${entry.side||''}`)||[];if(exact.length) return exact;if(entry.side){const any=imagePool.get(`${entry.size}|`)||[];if(any.length) return any}return[]}
function fileName(f){return(f&&f.name)?f.name:'(file)'}
function tagLabel(entry){ 
  if (!entry) return '(unknown)';
  if (typeof entry === 'string') return entry.slice(0,60) + (entry.length>60?'…':'');
  return entry.name ? entry.name : '(unnamed creative)';
}

function previewPlannedPlacements(byId) {
  const pools = buildTagPools(mapping);
  LOG(`Mode: ${pools.hasLineIds ? 'Line-targeted (IDs in CSV).' : 'Global (no line IDs in CSV).'}`);

  for (const [id, arr] of byId.entries()) {
    // group by size|variant|side
    const groups = new Map();
    for (const e of arr) {
      const key = `${e.size}|${e.variant||''}|${e.side||''}`;
      if (!groups.has(key)) groups.set(key, { entry: { size:e.size, variant:e.variant||null, side:e.side||null }, tiles: [] });
      if (e.tile) groups.get(key).tiles.push(e.tile);
    }

    LOG(`→ Preview ${id}`);
    if (!groups.size) { LOG('   (no tiles detected)'); continue; }

    for (const { entry, tiles } of groups.values()) {
      const label = [entry.size, entry.variant, entry.side].filter(Boolean).join('/');

      // Prefer images
      const imgsExact = imagePool.get(`${entry.size}|${entry.side||''}`) || [];
      const imgsAny   = entry.side ? (imagePool.get(`${entry.size}|`) || []) : [];
      const imgs = imgsExact.length ? imgsExact : imgsAny;

if (imgs.length) {
  const lines = tiles.map((_, i) => {
    const f = reuseAssets ? imgs[i % imgs.length] : (i < imgs.length ? imgs[i] : null);
    return f ? `tile#${i+1} ← ${fileName(f)}` : `tile#${i+1} ← (no image)`;
  });
  LOG(`   • ${label}  (images) ${tiles.length} tile(s)\n      ${lines.join('\n      ')}`);
  continue;
}



      // Scripts
      let poolItems = pools.getExact(id, entry.size, entry.variant || null);
      let poolLabel = poolItems.length
        ? (pools.maps.tagsByExactLine.has(pools.kExact(id, entry.size, entry.variant || null)) ? 'script line:exact' : 'script global:exact')
        : '';

      if (!poolItems.length) {
        poolItems = pools.getSize(id, entry.size);
        poolLabel = poolItems.length
          ? (pools.maps.tagsBySizeLine.has(pools.kSize(id, entry.size)) ? 'script line:size' : 'script global:size')
          : '';
      }

      if (!poolItems.length) {
  LOG(`   • ${label}  (no matching images or scripts)`);
  continue;
}

const lines = tiles.map((_, i) => {
  const payload = reuseAssets ? poolItems[i % poolItems.length]
                              : (i < poolItems.length ? poolItems[i] : null);
  const nice = payload ? tagLabel(payload) : '(no script)';
  return `tile#${i+1} ← ${nice}`;
});

LOG(`   • ${label}  (scripts) ${tiles.length} tile(s)  pool=${poolLabel}\n      ${lines.join('\n      ')}`);

    }
  }
}




/* ========= scan ========= */
let scanned=[]; let selected=new Set(G('ap3p_sel_ids',[]));
function saveSelection(){S('ap3p_sel_ids',Array.from(selected))}
function refreshList(){list.innerHTML='';if(!scanned.length){list.innerHTML='<div class="muted">No items — run Scan.</div>';return}
  scanned.forEach(it=>{const row=d.createElement('label');row.className='item';const cb=d.createElement('input');cb.type='checkbox';cb.checked=selected.has(it.id);const id=d.createElement('span');id.className='id';id.textContent=it.id;const sz=d.createElement('span');sz.className='sizes';sz.textContent=it.label;row.append(cb,id,sz);list.appendChild(row);cb.onchange=()=>{if(cb.checked)selected.add(it.id);else selected.delete(it.id);saveSelection()}})}
btnAll.onclick=()=>{selected=new Set(scanned.map(s=>s.id));saveSelection();refreshList()};
btnNone.onclick=()=>{selected=new Set();saveSelection();refreshList()};
btnInv.onclick=()=>{const next=new Set();scanned.forEach(s=>{if(!selected.has(s.id)) next.add(s.id)});selected=next;saveSelection();refreshList()};

btnScan.onclick=async()=>{log.textContent='';scanned.length=0;const rows=findRows();const byId=new Map();
for (const r of rows) {
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
  const report=[];for(const [id,arr] of byId.entries()){const pretty=prettyCounts(arr.map(({size,variant})=>({size,variant})));LOG(`• ${id}  sizes:[${pretty||'-'}]`);const uniq=new Map();for(const e of arr){const k=`${e.size}|${e.variant||''}|${e.side||''}`;if(!uniq.has(k)) uniq.set(k,{size:e.size,variant:e.variant,side:e.side})}report.push({id,entries:[...uniq.values()],label:pretty||'-'})}
  LOG(`\nFound ${report.length} line items.`);try{previewPlannedPlacements(byId)}catch(e){LOG('! Preview failed: '+(e?.message||e))}
  _lastPreview = byId;
  setInfo({items:report.length,done:0,hit:0,skip:0,err:0});scanned=report.slice();selected=new Set(report.map(r=>r.id));saveSelection();refreshList();
};

/* ========= run ========= */
let stopping=false; btnStop.onclick=()=>{stopping=true;btnStop.disabled=true;btnRun.disabled=false};
btnRun.onclick=()=>{const ids=[...new Set(scanned.filter(s=>selected.has(s.id)).map(s=>s.id))];if(!ids.length){LOG('No items selected — run Scan and tick rows to include.');return}if(!mapping.length&&imagePool.size===0){LOG('No mapping loaded — import CSV/JSON or images first.');return}runOnIds(ids)};

async function runOnIds(ids){
  stopping=false;btnRun.disabled=true;btnStop.disabled=false;
  const stats={items:ids.length,done:0,hit:0,skip:0,err:0};setInfo(stats);
  LOG(`Starting… (${ids.length} items)`);
  const pools=buildTagPools(mapping);

  for(const id of ids){
    if(stopping) break;
    try{
      const rowsForId=findRowsByIdAll(id); const detailsList=[]; let liveEntries=[];
      for(const r of rowsForId){await ensureCampaignView();
const det = await expandRow(r);
await waitForIdle();
await ensureAllTilesMounted(det || d);
await sleep(80);

if (det) detailsList.push(det);
liveEntries = liveEntries.concat(sizesFromExpanded(det))}
      const uniq=new Map(); for(const e of liveEntries){const k=`${e.size}|${e.variant||''}|${e.side||''}`;if(!uniq.has(k)) uniq.set(k,{size:e.size,variant:e.variant,side:e.side})}
      const entries=[...uniq.values()];
      LOG(`→ ${id} sizes:[${prettyCounts(entries)||'—'}]`);
      if(!entries.length){stats.skip++;stats.done++;setInfo(stats);continue}

      let wrote=false;
      for(const entry of entries){
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
  LOG(`   • ${entry.size}${entry.side?('/'+entry.side):''} (no tile match)`);
  continue;
}
const imgs = imageChoicesForEntry(entry);
if (imgs.length) {
  let placed = 0;
  for (let i = 0; i < tiles.length; i++) {
    const t = tiles[i];
    const f = reuseImages ? imgs[i % imgs.length] : (i < imgs.length ? imgs[i] : null);
    if (!f) continue; // no reuse and no image left → skip tile

    t.scrollIntoView({ block: 'center' });

    // Do NOT click the tile — avoids navigating to the preview
    let ok = await uploadImageToTile(t, f);
    if (!ok) {
      const alt = (t.closest('.set-creativeContainer') || document)
                    .querySelector('input[type="file"]');
      if (alt) ok = await uploadImageToInput(alt, f);
    }

    await sleep(160);
    placed++;
  }
  LOG(`   • ${entry.size}${entry.side?('/'+entry.side):''}  placed images=${placed}${placed<tiles.length?` (skipped ${tiles.length-placed})`:''}`);
  wrote = placed > 0 || wrote;
  continue;
}


        // Line-aware script pool: prefer exact(size+variant) for this line, then size for this line,
// then global exact, then global size.
const exactLine = pools.getExact(id, entry.size, entry.variant || null);
const sizeLine  = pools.getSize(id, entry.size);
const exactAny  = pools.getExact(null, entry.size, entry.variant || null);
const sizeAny   = pools.getSize(null, entry.size);

let list, pool;
if (pools.hasLineIds) {
  if (exactLine.length) { list = exactLine; pool = 'line:exact'; }
  else if (sizeLine.length) { list = sizeLine; pool = 'line:size'; }
  else if (exactAny.length) { list = exactAny; pool = 'global:exact'; }
  else { list = sizeAny; pool = 'global:size'; }
} else {
  list = exactAny.length ? exactAny : sizeAny;
  pool = exactAny.length ? 'global:exact' : 'global:size';
}

if (!list || !list.length) {
  LOG(`   • ${entry.size}${entry.variant?('/'+entry.variant):''} (no tags for this size/line)`);
  continue;
}

LOG(`   • ${entry.size}${entry.variant?('/'+entry.variant):''}  tiles=${tiles.length}, pool=${pool}, tags=${list.length}`);

let used = 0, skipped = 0;
for (let i = 0; i < tiles.length; i++) {
  const payload = reuseAssets ? list[i % list.length]
                              : (i < list.length ? list[i] : null);
  if (!payload) { skipped++; continue; }

  await selectTile(tiles[i]);
  // If selection accidentally navigated somewhere, go back and retry once
if (!(await ensureCampaignView())) {
  LOG('   ! Lost campaign view; attempted to go back.');
  await sleep(250);
}

await waitForIdle();

// Optional: retry selection once if it still isn't selected
if (!isTileSelected(tiles[i])) {
  await selectTile(tiles[i]);
  await waitForIdle();
}

  const editor = await open3PTab();
  if (!editor) { LOG('   ! 3rd-party editor not found'); stats.err++; continue; }

  const tagStr = (typeof payload === 'string') ? payload : (payload?.tag || '');
  pasteInto(editor, tagStr);

  used++;
  if (G('ap3p_auto', true)) {
    const ok = await clickSaveOrReprocess(editor);
    await waitForIdle(15000);
await ensureCampaignView(); // if it navigated, come back before next tile

    if (!ok) LOG('   ! Save/Update/Reprocess not found (continuing)');
  }
  await sleep(160);
}

LOG(`     → used scripts=${used}${skipped?` (skipped ${skipped})`:''}`);
wrote = used > 0 || wrote;

}
      if(wrote) stats.hit++; else stats.skip++; stats.done++; setInfo(stats); await sleep(160);
    }catch(e){LOG('   ! error: '+(e&&e.message?e.message:e));stats.err++;stats.done++;setInfo(stats)}
  }
  LOG(`Done. Items=${stats.items}  Hit=${stats.hit}  Skip=${stats.skip}  Err=${stats.err}`);
  btnStop.disabled=true; btnRun.disabled=false;
}

/* ========= niceties ========= */
_bestiltInt=setInterval(()=>{const panel=d.querySelector('#InfoPanelContainer');const t=panel?.innerText||'';const m=t.match(/Bestilt(?:\s*\(BxH\))?\s*[:\-]?\s*(\d{2,4})\s*[x×]\s*(\d{2,4})/i);badge.textContent=m?`— ${m[1]}x${m[2]}`:''},900);
w.addEventListener('keydown',e=>{if(e.altKey&&e.key.toLowerCase()==='a'){e.preventDefault();btnRun.click()}});

function inCampaignTableView(){
  // Campaign list view has visible line rows
  return !!document.querySelector('tr.set-matSpecDataRow');
}

function findBackToCampaignButton(){
  // In your snapshot, the back arrow button has an SVG path starting with:
  // M9.4 16.6L4.8 12...
  // We'll match that safely.
  const btns = [...document.querySelectorAll('button, [role="button"]')];
  for (const b of btns) {
    const p = b.querySelector('svg path');
    const d = p?.getAttribute?.('d') || '';
    if (d.startsWith('M9.4 16.6L4.8 12')) return b;
  }
  return null;
}

async function ensureCampaignView(){
  if (inCampaignTableView()) return true;

  const back = findBackToCampaignButton();
  if (back) {
    back.scrollIntoView({ block:'center' });
    back.click();
    // wait until the table view is back
    await waitFor(() => inCampaignTableView(), 8000, 120);
  }

  return inCampaignTableView();
}

// Wait for "busy" overlays to clear (MUI Backdrop / progress)
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



}catch(e){console.error(e);alert('Autofill error: '+(e&&e.message?e.message:e));}})();

