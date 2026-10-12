// Gophers CRM — standalone (Microsoft sign-in, Excel on OneDrive is the single source of truth)
(() => {
const DRIVE_ID = "b!y07j3zs0wUORORx62WJ844NU9RJFZu9JsXYsPHd-SIEociUqeJtSTq6ZENgXtiJP";
const CRM_FILE = "01HMEMOXUZW5ZBHFP6IBEKYO5KFYJMBAFM";
const REHBER_FOLDER_ID = "01HMEMOXRD3AFZOVQW4NBLGY2OPUCEW5JB";
const WB = `/drives/${DRIVE_ID}/items/${CRM_FILE}/workbook`;
const COLS = [["d","Date"],["by","Call By"],["ct","Call Type"],["rb","Referral by"],["n","Contact name"],["c1","Column1"],["kt","Contact Type"],["oc","Occupation"],["t","Telephone"],["em","E Mail"],["ad","Adress"],["pb","Price / Budget"],["rf","Reference"],["cma","CMA"],["lu","Listing URL"],["co","Contracted"],["no","Notes"],["k","Keep (Y/N)"],["tf","Touch Frequency"],["tt","Touch Type"],["p","Potential Level"],["nd","Next Action Date"],["na","Next Action Type"],["x","Ek Alan"],["sd","Sözleşme DB"]];
const POTS = ["VIP","Very High","High","Medium","Low"];
const SHEET = {soz:"📋 Sözleşme Veritabanı", kira:"🔑 Kiralık Takip", firsat:"🎯 Fırsat Takip"};
const START = {soz:5, kira:3, firsat:3};
const LASTC = {soz:"BR", kira:"U", firsat:"K"};

const S = {rows:[], leads:[], contacts:new Map(), sel:null, ct:"", kt:"", p:"", histAll:false, account:null};
const $ = id => document.getElementById(id);

// ---------- Microsoft sign-in + Graph ----------
const SCOPES = ["User.Read", "Files.ReadWrite"];
const CONFIG = window.GT_CONFIG;
const BASE = location.origin + location.pathname.replace(/[^/]*$/, "");   // registered redirect URI (app root)
const msalApp = new msal.PublicClientApplication({
  auth: { clientId: CONFIG.clientId, authority: "https://login.microsoftonline.com/" + CONFIG.tenantId, redirectUri: BASE },
  cache: { cacheLocation: "localStorage" }
});
async function token(){
  try{ return (await msalApp.acquireTokenSilent({scopes:SCOPES, account:S.account})).accessToken; }
  catch(e){ await msalApp.acquireTokenRedirect({scopes:SCOPES, account:S.account, state:"crm"}); throw {code:"redirect"}; }
}
async function graph(path, opts = {}){
  for(let attempt = 1;; attempt++){
    const res = await fetch("https://graph.microsoft.com/v1.0" + path, {method: opts.method || "GET", headers: {"Authorization": "Bearer " + await token(), "Content-Type": "application/json"}, body: opts.body ? JSON.stringify(opts.body) : undefined});
    if(res.ok) return res.status === 204 ? {} : res.json();
    if(attempt < 4 && [409, 423, 429, 500, 502, 503, 504].includes(res.status)){ await new Promise(r => setTimeout(r, 1500 * attempt)); continue; }
    let msg = ""; try{ msg = (await res.json()).error.message; }catch(_){}
    throw {code: "http_" + res.status, message: msg};
  }
}
function errMsg(err, prefix){
  const c = err && err.code;
  return c === "http_401" ? "Oturum süresi doldu. Sayfayı yenile ve Microsoft ile tekrar giriş yap."
    : c === "http_403" ? "Bu Microsoft hesabının CRM dosyasına erişimi yok. Gophers Realty hesabıyla giriş yaptığından emin ol."
    : c === "http_404" ? "CRM dosyası OneDrive'da bulunamadı. Dosyanın adı ya da yeri değişmiş olabilir; Selin'e haber ver."
    : c === "http_423" || c === "http_409" ? prefix + ": dosya şu an başka bir yerde kilitli. Birkaç saniye sonra tekrar dene."
    : c === "not_found" ? prefix + ": kayıt Excel'de bulunamadı (başka yerden değişmiş olabilir). Yenile'ye basıp tekrar dene."
    : prefix + ". İnternet bağlantını kontrol edip tekrar dene." + (err && err.message ? " (" + String(err.message).slice(0,120) + ")" : "");
}
const wsPath = name => `${WB}/worksheets('${encodeURIComponent(name)}')`;
function serial(iso){ const [y,m,d]=iso.split("-").map(Number); return Math.round((Date.UTC(y,m-1,d)-Date.UTC(1899,11,30))/86400000); }
function fromSerial(n){ const d=new Date(Date.UTC(1899,11,30)+Math.floor(n)*86400000); return d.getUTCFullYear()+"-"+String(d.getUTCMonth()+1).padStart(2,"0")+"-"+String(d.getUTCDate()).padStart(2,"0"); }
function toISO(v){
  if(typeof v==="number" && v>20000 && v<80000) return fromSerial(v);
  const s=String(v??"").trim();
  let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return m[1]+"-"+m[2]+"-"+m[3];
  m=s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/); if(m) return m[3]+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");
  return "";
}
function colNum(c){ let n=0; for(const ch of c) n=n*26+(ch.charCodeAt(0)-64); return n; }
// Excel would turn "+90 532…", "0532…", "-not" or "=x" into numbers/formulas: keep such text as text
function cellIn(v){
  if(v===undefined || v===null) return "";
  if(typeof v==="number") return isFinite(v) ? v : "";
  const s=String(v); if(!s) return "";
  if(/^[=+\-@]/.test(s) || (/\d/.test(s) && /^[\d\s.,:\/()%]+$/.test(s))) return "'"+s;
  return s;
}

// ---------- loaders ----------
async function readCrmTable(){
  const r = await graph(`${WB}/tables('CRM')/range?$select=address,values`);
  const head = r.values[0].map(h => String(h).trim());
  const idx = COLS.map(([,h]) => head.indexOf(h));
  return {r, head, idx};
}
function crmRec(c, idx, i){
  const o = {seq:i+1};
  COLS.forEach(([k], j) => {
    const ix = idx[j]; if(ix < 0) return;
    let v = c[ix]; if(v === "" || v === null || v === undefined) return;
    if(k === "d" || k === "nd"){ v = toISO(v); if(!v) return; }
    else v = String(v).trim();
    if(v !== "") o[k] = v;
  });
  return o;
}
async function loadCrm(){
  const {r, head, idx} = await readCrmTable();
  S.crmHead = head; S.crmIdx = idx;
  S.rows = r.values.slice(1).map((c, i) => { const o = crmRec(c, idx, i); o._id = "r" + i; return o; }).filter(o => o.n || o.no);
}
async function loadMod(mod){
  const st = START[mod];
  const r = await graph(`${wsPath(SHEET[mod])}/range(address='A${st}:${LASTC[mod]}${st+599}')?$select=values`);
  const out = [];
  r.values.forEach((c, i) => {
    if(c[0] === "" || c[0] === null) return;
    const d = {_r:st+i, _id:mod+(st+i), seq:i+1};
    for(const f of sch(mod)){
      if(f.t === "comp") continue;
      const v = c[colNum(f.c)-1]; if(v === "" || v === null || v === undefined) continue;
      if(f.t === "date") d[f.c] = toISO(v) || String(v);
      else if(f.t === "money" || f.t === "num" || f.t === "pct") d[f.c] = typeof v === "number" ? v : String(v);
      else d[f.c] = String(v);
    }
    out.push(d);
  });
  S[mod] = out;
}
function saveCache(){ try{ localStorage.setItem("gc_cache", JSON.stringify({rows:S.rows, soz:S.soz, kira:S.kira, firsat:S.firsat, at:Date.now()})); }catch(_){} }
function loadCache(){ try{ const c = JSON.parse(localStorage.getItem("gc_cache") || "null"); if(!c) return false; S.rows=c.rows||[]; S.soz=c.soz||[]; S.kira=c.kira||[]; S.firsat=c.firsat||[]; return true; }catch(_){ return false; } }

// ---------- helpers ----------
const pad = n => String(n).padStart(2,"0");
function isoLocal(d){ return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
function todayISO(){ return isoLocal(new Date()); }
function addDays(iso,n){ const [y,m,d]=iso.split("-").map(Number); const x=new Date(y,m-1,d); x.setDate(x.getDate()+n); return isoLocal(x); }
function dayDiff(iso){ if(!iso) return null; const [y,m,d]=iso.split("-").map(Number); const a=new Date(y,m-1,d); const t=new Date(); const b=new Date(t.getFullYear(),t.getMonth(),t.getDate()); return Math.round((a-b)/86400000); }
function tr(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return d+"."+m+"."+y; }
function norm(s){ return String(s||"").toLocaleLowerCase("tr").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/ı/g,"i").replace(/[^a-z0-9 ]+/g," ").replace(/\s+/g," ").trim(); }
function digits(s){ return String(s||"").replace(/\D/g,""); }
function phoneKey(s){ const d=digits(s); return d.length>=10 ? d.slice(-10) : ""; }
function waNum(s){ let d=digits(s); if(d.length===10 && d[0]==="5") d="90"+d; if(d.length===11 && d.startsWith("05")) d="9"+d; return (d.length===12 && d.startsWith("905")) ? d : ""; }
function esc(s){ return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
function potCls(p){ return p ? "pot pot-"+p.replace(/\s/g,"") : ""; }
function toast(msg){ const t=$("toast"); t.textContent=msg; t.classList.add("on"); clearTimeout(toast._t); toast._t=setTimeout(()=>t.classList.remove("on"),2600); }
function banner(msg){ const b=$("banner"); if(!msg){b.classList.add("hidden");return;} b.textContent=msg; b.classList.remove("hidden"); }

// ---------- data model ----------
function rebuild(){
  const map=new Map();
  const sorted=[...S.rows].sort((a,b)=>((a.d||"").localeCompare(b.d||"")||a.seq-b.seq));
  for(const r of sorted){
    if(!r.n) continue;
    const key=norm(r.n); if(!key) continue;
    let c=map.get(key);
    if(!c){ c={key,recs:[]}; map.set(key,c); }
    c.recs.push(r);
  }
  for(const c of map.values()){
    const last=c.recs[c.recs.length-1];
    const pick=f=>{ for(let i=c.recs.length-1;i>=0;i--){ if(c.recs[i][f]) return c.recs[i][f]; } return ""; };
    c.last=last; c.name=last.n; c.t=pick("t"); c.rf=pick("rf"); c.p=pick("p"); c.kt=pick("kt"); c.em=pick("em"); c.tf=pick("tf");
    c.lastDate=c.recs.reduce((m,r)=>r.d && r.d>m ? r.d : m,"");
    c.hay=norm(c.name)+" "+c.recs.map(r=>norm(r.rf)).join(" ");
    c.dig=c.recs.map(r=>digits(r.t)).join(" ");
    c.pk=phoneKey(c.t);
  }
  S.contacts=map;
  const rfs=new Set(); S.rows.forEach(r=>r.rf&&rfs.add(r.rf));
  $("rfList").innerHTML=[...rfs].slice(-400).map(v=>`<option value="${esc(v)}">`).join("");
  $("contactNames").innerHTML=[...map.values()].map(c=>`<option value="${esc(c.name)}">`).join("");
  buildChips();
  renderResults(); renderToday();
}

function topValues(field, fallback, n){
  const cnt=new Map(); S.rows.forEach(r=>{ if(r[field]) cnt.set(r[field],(cnt.get(r[field])||0)+1); });
  const top=[...cnt.entries()].sort((a,b)=>b[1]-a[1]).map(e=>e[0]);
  const out=[]; [...fallback,...top].forEach(v=>{ if(v && !out.includes(v)) out.push(v); });
  return out.slice(0,n);
}
function buildChips(){
  chipGroup("ctChips", topValues("ct",["Call In","Call Out","WhatsApp","Face to Face","E-mail"],8), "ct");
  chipGroup("ktChips", topValues("kt",["Alıcı","Satıcı","Alıcı/Satıcı","Personal Network","Emlakçı"],10), "kt", true);
  chipGroup("pChips", POTS, "p");
  $("quick").innerHTML=[["Yarın",1],["3 gün",3],["1 hafta",7],["2 hafta",14],["1 ay",30]].map(([l,n])=>`<button type="button" class="chip" data-n="${n}">${l}</button>`).join("");
}
function chipGroup(id, vals, field, allowOther){
  const cur=S[field];
  const list=[...vals]; if(cur && !list.includes(cur)) list.push(cur);
  $(id).innerHTML=list.map(v=>`<button type="button" class="chip" aria-pressed="${v===cur}" data-v="${esc(v)}">${esc(v)}</button>`).join("")+(allowOther?`<button type="button" class="chip add" data-other="1">Diğer…</button>`:"");
  $(id).onclick=e=>{
    const b=e.target.closest("button"); if(!b) return;
    if(b.dataset.other){
      const inp=document.createElement("input"); inp.className="otherin"; inp.placeholder="Yaz, Enter";
      b.replaceWith(inp); inp.focus();
      const done=()=>{ const v=inp.value.trim(); if(v) S[field]=v; chipGroup(id, vals, field, allowOther); };
      inp.addEventListener("keydown",ev=>{ if(ev.key==="Enter"){ ev.preventDefault(); done(); } });
      inp.addEventListener("blur",done);
      return; }
    S[field] = S[field]===b.dataset.v ? "" : b.dataset.v;
    chipGroup(id, vals, field, allowOther);
  };
}

// ---------- search ----------
function search(q){
  const all=[...S.contacts.values()];
  const nq=norm(q), dq=digits(q);
  if(!nq && !dq) return all.sort((a,b)=>(b.lastDate||"").localeCompare(a.lastDate||"")||(b.last.seq-a.last.seq)).slice(0,25);
  const toks=nq.split(" ").filter(Boolean);
  const res=all.filter(c=>{
    if(dq.length>=4 && c.dig.includes(dq)) return true;
    return toks.length && toks.every(t=>c.hay.includes(t));
  });
  res.sort((a,b)=>{
    const an=norm(a.name).startsWith(nq)?0:1, bn=norm(b.name).startsWith(nq)?0:1;
    return an-bn || (b.lastDate||"").localeCompare(a.lastDate||"");
  });
  return res.slice(0,40);
}
function renderResults(){
  const q=$("q").value;
  const res=search(q);
  $("qHint").textContent = q ? (res.length ? `${res.length} kişi bulundu` : "Kayıtlı kişi yok. Yeni kişi olarak ekleyebilirsin.") : "Son görüşülen kişiler. Yazdıkça CRM'de arar.";
  $("newBtnTxt").textContent = q && !res.length ? `"${q}" yeni kişi olarak ekle` : "Yeni kişi ekle";
  $("results").innerHTML = res.map(c=>`
    <button type="button" class="person" data-k="${esc(c.key)}" aria-current="${S.sel===c.key}">
      <div class="nm">${esc(c.name)}${c.p?`<span class="${potCls(c.p)}">${esc(c.p)}</span>`:""}</div>
      <div class="sub">${esc([c.rf, c.t, c.lastDate?tr(c.lastDate):""].filter(Boolean).join("  ·  "))}</div>
    </button>`).join("") || (S.rows.length?"":`<div class="empty">Kayıtlar yükleniyor…</div>`);
}

// ---------- form ----------
function setVal(id,v){ $(id).value=v||""; if(id==="d" && window.__markD) window.__markD(); }
function clearForm(keepName){
  S.sel=null; S.histAll=false;
  const nm = keepName || "";
  ["n","t","rf","no","na","em","pb","tf","ad"].forEach(i=>setVal(i,""));
  setVal("n",nm); setVal("d",todayISO()); setVal("nd","");
  S.ct=""; S.kt=""; S.p=""; $("kN").checked=false;
  $("fTitle").textContent="Yeni kişi"; $("fMeta").textContent="Kayıtlı değilse bilgileri gir ve kaydet.";
  $("histBox").innerHTML=""; $("dupWarn").classList.add("hidden");
  buildChips(); renderResults();
}
function selectContact(key, focusNote=true){
  const c=S.contacts.get(key); if(!c) return;
  S.sel=key; S.histAll=false;
  setVal("n",c.name); setVal("t",c.t); setVal("rf",c.rf); setVal("em",c.em); setVal("tf",c.tf);
  setVal("no",""); setVal("na",""); setVal("nd",""); setVal("pb",""); setVal("ad",""); setVal("d",todayISO());
  S.ct=""; S.kt=c.kt||""; S.p=c.p||""; $("kN").checked=false;
  $("fTitle").textContent=c.name;
  $("fMeta").textContent=`${c.recs.length} kayıt · son görüşme ${c.lastDate?tr(c.lastDate):"—"}`;
  $("dupWarn").classList.add("hidden");
  renderHist(); buildChips(); renderResults();
  document.body.classList.add("editing");
  if(focusNote) setTimeout(()=>$("no").focus({preventScroll:true}),50);
}
function renderHist(){
  const c=S.contacts.get(S.sel); if(!c){ $("histBox").innerHTML=""; return; }
  const recs=[...c.recs].sort((a,b)=>(b.d||"").localeCompare(a.d||"")||b.seq-a.seq);
  const show=S.histAll?recs:recs.slice(0,3);
  $("histBox").innerHTML=`<div class="hist">${show.map(r=>`
    <div class="hist-item">
      <div class="when">${tr(r.d)||"tarihsiz"}${r.ct?" · "+esc(r.ct):""}${r.rf?" · "+esc(r.rf):""}</div>
      ${r.no?`<div class="what">${esc(r.no)}</div>`:""}
      ${(r.na||r.nd)?`<div class="next">Sonraki: ${esc(r.na||"")}${r.nd?" ("+tr(r.nd)+")":""}</div>`:""}
      <span class="delwrap" data-id="${esc(r._id)}"><button type="button" class="del">Sil</button></span>
    </div>`).join("")}</div>${recs.length>3?`<button type="button" class="more-hist" id="moreHist">${S.histAll?"Daha az göster":`Tüm geçmiş (${recs.length})`}</button>`:""}`;
}
function checkDup(){
  const w=$("dupWarn");
  const pk=phoneKey($("t").value);
  if(!pk){ w.classList.add("hidden"); return; }
  const hit=[...S.contacts.values()].find(c=>c.pk===pk && c.key!==S.sel);
  if(hit){ w.innerHTML=`Bu telefon <b>${esc(hit.name)}</b> adına kayıtlı.<button type="button" data-k="${esc(hit.key)}">O kişiyi aç</button>`; w.classList.remove("hidden"); }
  else w.classList.add("hidden");
}


async function save(e){
  e.preventDefault();
  const n=$("n").value.trim(), no=$("no").value.trim();
  if(!n){ $("n").focus(); toast("Ad soyad gerekli"); return; }
  if(!no){ $("no").focus(); toast("Görüşme notu gerekli"); return; }
  const rec={d:$("d").value||todayISO(), by:"Gültekin", ct:S.ct, n, kt:S.kt, t:$("t").value.trim(), em:$("em").value.trim(), ad:$("ad").value.trim(), pb:$("pb").value.trim(), rf:$("rf").value.trim(), no, k:$("kN").checked?"N":"Y", tf:$("tf").value.trim(), p:S.p, nd:$("nd").value, na:$("na").value.trim()};
  Object.keys(rec).forEach(k=>{ if(rec[k]===""||rec[k]==null) delete rec[k]; });
  $("saveBtn").disabled=true; $("saveBtn").textContent="Excel'e yazılıyor…";
  try{
    const head=S.crmHead, idx=S.crmIdx;
    if(!head) throw {code:"not_loaded", message:"CRM henüz yüklenmedi"};
    const vals=Array(head.length).fill("");
    COLS.forEach(([k],j)=>{ const ix=idx[j]; if(ix<0) return; let v=rec[k]; if((k==="d"||k==="nd") && isISO(v)) v=serial(v); vals[ix]=cellIn(v); });
    await graph(`${WB}/tables('CRM')/rows/add`,{method:"POST",body:{values:[vals]}});
    rec.seq=S.rows.reduce((m,r)=>Math.max(m,r.seq||0),0)+1; rec._id="new"+rec.seq;
    S.rows.push(rec);
    const key=norm(n);
    rebuild(); saveCache();
    toast("Kaydedildi");
    $("q").value=""; selectContact(key,false); setVal("no",""); setVal("na",""); setVal("nd","");
    syncLbl();
    loadCrm().then(()=>{ rebuild(); saveCache(); if(S.sel) renderHist(); }).catch(()=>{});
  }catch(err){
    console.error(err); if(err && err.code==="redirect") return;
    banner(errMsg(err,"Kaydedilemedi")); toast("Kaydedilemedi");
  }finally{ $("saveBtn").disabled=false; $("saveBtn").textContent="Kaydet"; }
}
async function delRec(id){
  const r=S.rows.find(x=>x._id===id); if(!r) return;
  const key=S.sel;
  toast("Siliniyor…");
  try{
    const {r:rng, idx}=await readCrmTable();
    const rows=rng.values.slice(1);
    const same=(o)=>o.d===r.d && o.n===r.n && (o.no||"")===(r.no||"") && (o.nd||"")===(r.nd||"");
    let i=rows.findIndex((c,j)=>same(crmRec(c,idx,j)));
    if(i<0) throw {code:"not_found"};
    try{ await graph(`${WB}/tables('CRM')/rows/itemAt(index=${i})`,{method:"DELETE"}); }
    catch(e1){
      if(e1 && e1.code==="redirect") throw e1;
      const m=String(rng.address).match(/!([A-Z]+)(\d+):([A-Z]+)\d+$/); if(!m) throw e1;
      const row=Number(m[2])+1+i, sheet=String(rng.address).split("!")[0].replace(/^'|'$/g,"").replace(/''/g,"'");
      await graph(`${wsPath(sheet)}/range(address='${m[1]}${row}:${m[3]}${row}')/delete`,{method:"POST",body:{shift:"Up"}});
    }
    await loadCrm(); rebuild(); saveCache(); syncLbl();
    if(S.contacts.has(key)) selectContact(key,false); else { clearForm(); document.body.classList.remove("editing"); }
    toast("Kayıt silindi");
  }catch(e){
    console.error(e); if(e && e.code==="redirect") return;
    renderHist(); banner(errMsg(e,"Kayıt silinemedi"));
  }
}

// ---------- today ----------
function buckets(){
  const POT={"VIP":1,"Very High":2,"High":3,"Medium":4,"Low":5};
  const a=[],b=[],c=[];
  for(const ct of S.contacts.values()){
    const r=ct.last;
    if(!r.nd || r.k==="N" || norm(r.na).includes("olumsuz")) continue;
    const dd=dayDiff(r.nd);
    const item={ct,r,dd};
    if(dd<-7) b.push(item); else if(dd<=0) a.push(item); else if(dd<=7) c.push(item);
  }
  a.sort((x,y)=>x.dd-y.dd);
  b.sort((x,y)=>(POT[x.ct.p]||6)-(POT[y.ct.p]||6) || y.dd-x.dd);
  c.sort((x,y)=>x.dd-y.dd);
  return {a,b,c};
}
function dueLbl(dd){ if(dd===0) return ["Bugün","now"]; if(dd<0) return [(-dd)+" gün gecikti","late"]; return [dd===1?"Yarın":dd+" gün sonra","soon"]; }
function callRow({ct,r,dd}){
  const [l,cls]=dueLbl(dd); const wa=waNum(ct.t);
  return `<div class="call" data-k="${esc(ct.key)}" tabindex="0" role="button">
    <span class="due ${cls}">${l}</span>
    <div><div class="nm">${esc(ct.name)}${ct.p?`<span class="${potCls(ct.p)}">${esc(ct.p)}</span>`:""}</div>
    <div class="sub">${esc([r.na, ct.rf].filter(Boolean).join("  ·  "))}</div></div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}
function renderToday(){
  const {a,b,c}=buckets();
  const rem=sozReminders();
  const leads=todaysLeads();
  const nb=a.length+rem.filter(r=>r.dd<=7).length;
  const badge=$("bugunBadge"); if(nb){ badge.textContent=nb; badge.classList.remove("hidden"); } else badge.classList.add("hidden");
  const sec=(title,items,extra)=>`<div class="today-sec"><h3>${title}<span class="n">${extra||items.length+" kişi"}</span></h3>${items.map(callRow).join("")||`<div class="empty">Bu bölümde kimse yok.</div>`}</div>`;
  const leadSec=`<div class="today-sec"><h3>Bugün görüşülecek yeni adaylar<span class="n">${leads.length} kişi</span></h3>${leads.map(leadRow).join("")||`<div class="empty">Rehberden eklenecek yeni aday yok.</div>`}</div>`;
  const ft=firsatTotals();
  const pipeSec = ft.total ? `<div class="today-sec"><h3>Fırsat Pipeline'ı<span class="n">${ft.open} açık / ${ft.total} toplam</span></h3><div class="call" style="cursor:default"><span></span><div><div class="nm">Ağırlıklı gelir tahmini: ${fmtTL(ft.I)}</div><div class="sub">Toplam potansiyel komisyon: ${fmtTL(ft.F)}</div></div><span></span></div></div>` : "";
  $("todayBox").innerHTML =
    pipeSec +
    (rem.length?`<div class="today-sec"><h3>Sözleşme & kira hatırlatmaları<span class="n">${rem.length} hatırlatma</span></h3>${rem.map(remRow).join("")}</div>`:"") +
    sec("Bugün ve son 7 gün",a) +
    sec("Öncelikli birikmiş takip",b.slice(0,10),`ilk 10 · toplam ${b.length} birikmiş`) +
    sec("Önümüzdeki 7 gün",c) +
    leadSec;
}

// ---------- new leads (Rehber) ----------
function leadPool(){
  const used=new Set([...S.contacts.values()].map(c=>c.pk).filter(Boolean));
  return (S.leads||[]).filter(l=>{
    const pk=phoneKey(l.phone1||l.phone2||l.phone3||"");
    return !pk || !used.has(pk);
  });
}
function todaysLeads(){
  const pool=leadPool();
  if(!pool.length) return [];
  const n=Math.min(10,pool.length);
  const days=Math.floor(Date.now()/86400000);
  const start=(days*10)%pool.length;
  const out=[];
  for(let i=0;i<n;i++) out.push(pool[(start+i)%pool.length]);
  return out;
}
function leadRow(l){
  const wa=waNum(l.phone1||l.phone2||l.phone3||"");
  const sub=[l.org,l.title,l.phone1].filter(Boolean).join(" · ");
  return `<div class="call" data-lead="${esc(l.id)}" tabindex="0" role="button">
    <span class="due soon">Rehber</span>
    <div><div class="nm">${esc(l.name)}${l.isForeign?` <span class="pot pot-Low">Yabancı</span>`:""}</div>
    <div class="sub">${esc(sub)}</div></div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}
function openLead(id){
  const l=(S.leads||[]).find(x=>x.id===id); if(!l) return;
  clearForm(l.name);
  setVal("t", l.phone1||l.phone2||l.phone3||"");
  setVal("em", l.email||"");
  setVal("rf", [l.org,l.title].filter(Boolean).join(" · "));
  $("fTitle").textContent=l.name;
  $("fMeta").textContent="Rehberden yeni aday · henüz CRM'de kayıtlı değil.";
  document.body.classList.add("editing");
  checkDup();
  setTimeout(()=>$("no").focus({preventScroll:true}),50);
}
function isForeignPhone(s){
  const d=digits(s); if(!d) return false;
  if(d.length===10) return false;
  if(d.length===11 && d[0]==="0") return false;
  if(d.length===12 && d.startsWith("90")) return false;
  return true;
}
function vcardLines(block){
  return block.split(/\r\n|\n|\r/).reduce((out,line)=>{
    if(/^[ \t]/.test(line) && out.length) out[out.length-1]+=line.slice(1);
    else out.push(line);
    return out;
  },[]);
}
function vField(lines,key){ const re=new RegExp("^(?:item\\d+\\.)?"+key+"[;:]","i"); for(const l of lines){ if(re.test(l)) return l.slice(l.indexOf(":")+1).trim(); } return ""; }
function vFieldAll(lines,key){ const re=new RegExp("^(?:item\\d+\\.)?"+key+"[;:]","i"); const out=[]; for(const l of lines){ if(re.test(l)) out.push(l.slice(l.indexOf(":")+1).trim()); } return out; }
function parseVCards(text){
  const out=[]; const blocks=text.split(/BEGIN:VCARD/i).slice(1);
  for(const raw of blocks){
    const lines=vcardLines(raw.split(/END:VCARD/i)[0]);
    let name=vField(lines,"FN");
    if(!name){ const n=vField(lines,"N"); if(n){ const p=n.split(";"); name=[p[1],p[0]].filter(Boolean).join(" ").trim(); } }
    if(!name) continue;
    const tels=vFieldAll(lines,"TEL").map(t=>t.trim()).filter(Boolean);
    const emails=vFieldAll(lines,"EMAIL").map(t=>t.trim()).filter(Boolean);
    const org=vField(lines,"ORG").replace(/;+$/,"").replace(/;/g," ").trim();
    const title=vField(lines,"TITLE").trim();
    out.push({name:name.trim(),phones:tels,email:emails[0]||"",org,title});
  }
  return out;
}

// ---------- Rehber (phone contacts) → new leads, read straight from OneDrive ----------
// Every .vcf / .txt contact export in the Rehber folder is parsed once and cached on this device;
// a new export dropped into the folder is picked up automatically the next time Bugün opens.
async function loadLeads(force){
  const lbl=$("leadImportLbl");
  let cache={files:{}}; try{ cache=JSON.parse(localStorage.getItem("gc_leads")||"null")||{files:{}}; }catch(_){}
  try{
    if(lbl) lbl.textContent="Rehber kontrol ediliyor…";
    const res=await graph(`/drives/${DRIVE_ID}/items/${REHBER_FOLDER_ID}/children?$select=id,name,eTag,createdDateTime,file&$top=200`);
    const files=(res.value||[]).filter(f=>f.file && /\.(vcf|txt)$/i.test(f.name)).sort((a,b)=>String(a.createdDateTime).localeCompare(String(b.createdDateTime)));
    const next={files:{}};
    for(const f of files){
      const old=cache.files[f.id];
      if(old && old.eTag===f.eTag && !force){ next.files[f.id]=old; continue; }
      if(lbl) lbl.textContent="İşleniyor: "+f.name;
      const meta=await graph(`/drives/${DRIVE_ID}/items/${f.id}?$select=id,@microsoft.graph.downloadUrl`);
      const text=await (await fetch(meta["@microsoft.graph.downloadUrl"])).text();
      next.files[f.id]={eTag:f.eTag, name:f.name, at:f.createdDateTime, cards:parseVCards(text).map(c=>[c.name,c.phones.slice(0,3),c.email,c.org,c.title])};
    }
    cache=next;
    try{ localStorage.setItem("gc_leads", JSON.stringify(cache)); }catch(_){}
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; }
  // build the lead list: one entry per phone, first file wins
  const used=new Set(); const out=[]; let seq=0;
  for(const f of Object.values(cache.files).sort((a,b)=>String(a.at).localeCompare(String(b.at)))){
    for(const [name,phones,email,org,title] of f.cards){
      const pks=phones.map(phoneKey).filter(Boolean);
      if(pks.length && pks.some(pk=>used.has(pk))) continue;
      pks.forEach(pk=>used.add(pk));
      seq++;
      out.push({id:"lead_"+seq, seq, name, phone1:phones[0]||"", phone2:phones[1]||"", phone3:phones[2]||"", email:email||"", org:org||"", title:title||"", isForeign:isForeignPhone(phones[0]||""), noPhone:!phones.length});
    }
  }
  S.leads=out;
  if(lbl) lbl.textContent=out.length ? `Rehber: ${leadPool().length} aday` : "";
  renderToday();
}

// ================= MODULES: Sözleşme & Kira Takip =================
const SCH = JSON.parse(document.getElementById("schemaJson").textContent);
const SOZ_TYPES=["SATIŞ","VİLLA SATIŞ","ARSA/ARAZİ SATIŞ","TİCARİ SATIŞ","ENDÜSTRİYEL SATIŞ","PROJE SATIŞ","KİRA","VİLLA KİRA","ARSA/ARAZİ KİRA","TİCARİ KİRA","ENDÜSTRİYEL KİRA","DEVİR"];
const KIRA_TYPES=["KİRA","VİLLA KİRA","ARSA/ARAZİ KİRA","TİCARİ KİRA","ENDÜSTRİYEL KİRA","DEVİR"];
const SOZ_DURUM=["🟡 Devam Ediyor","✅ Tamamlandı","❌ İptal"];
const KIRA_DURUM=["🔄 Kiracı Aranıyor","✅ Aktif","⏳ Çıkış Hazırlığı","⛔ Pasif"];
const FIRSAT_DURUM=["🆕 Yeni Lead","📞 İlk Görüşme","🔎 İlgileniyor / Takip","📝 Teklif / Görüşme","📋 Sözleşme Bekliyor","✅ Kazanıldı","❌ Kaybedildi","⏸️ Donduruldu"];
const FIRSAT_PROB_DEFAULT={"🆕 Yeni Lead":10,"📞 İlk Görüşme":20,"🔎 İlgileniyor / Takip":35,"📝 Teklif / Görüşme":55,"📋 Sözleşme Bekliyor":75,"✅ Kazanıldı":100,"❌ Kaybedildi":0,"⏸️ Donduruldu":5};
const PAIRS={soz:{N:["O","P"],R:["S","T"],W:["X"],Y:["Z"],AW:["AY","AZ"],BB:["BD"],BF:["BH"]}, kira:{D:["E"],F:["G"],J:["K"],L:["M"]}, firsat:{A:["B",null]}};
const PERSON={soz:["N","R","W","Y","Q","V","AW","BB","BF"], kira:["D","F","J","L"], firsat:["A"]};
const LISTS={J:["Sahibinden","Hepsiemlak","Emlakjet","KW","Instagram","Referans"], AA:["Nakit","Havale/EFT","Kredi"]};
const COLL={soz:"soz",kira:"kira",firsat:"firsat"};
S.soz=[]; S.kira=[]; S.firsat=[]; S.modSel={soz:null,kira:null,firsat:null};
const isKira = d => KIRA_TYPES.includes(d && d.D);
const isISO = v => typeof v==="string" && /^\d{4}-\d{2}-\d{2}$/.test(v);
const sch = mod => mod==="soz" ? SCH.soz : mod==="kira" ? SCH.kira : SCH.firsat;
function firsatCalc(d){ const F=nz(d.D)*nz(d.E); const I=F*nz(d.H); return {F,I}; }
function firsatTotals(){
  let F=0,I=0,open=0;
  for(const d of S.firsat){ const c=firsatCalc(d); F+=c.F; I+=c.I; if(!["✅ Kazanıldı","❌ Kaybedildi"].includes(d.G)) open++; }
  return {F,I,open,total:S.firsat.length};
}
function nz(v){ const n=Number(v); return isFinite(n)?n:0; }
function fmtNum(n,dec){ if(n===""||n==null||!isFinite(Number(n))) return n??""; return Number(n).toLocaleString("tr-TR",{maximumFractionDigits:dec??2}); }
function fmtTL(n){ return fmtNum(Math.round(nz(n)))+" ₺"; }
function parseNum(v){
  let s=String(v).replace(/[₺%\s]/g,""); if(!s) return "";
  if(s.includes(",") && s.includes(".")) s=s.replace(/\./g,"").replace(",",".");
  else if(s.includes(",")) s=s.replace(",",".");
  else if(/^\d{1,3}(\.\d{3})+$/.test(s)) s=s.replace(/\./g,"");
  const n=Number(s); return isFinite(n) ? n : String(v).trim();
}
function calc(d){
  const AF=nz(d.AC)*(nz(d.AD)+nz(d.AE)), BJ=AF, BK=BJ*nz(d.BA), BL=BJ*nz(d.BE), BM=BJ*nz(d.BI), BN=BK+BL+BM, BO=BJ-BN;
  const BP=(d.BP===""||d.BP==null)?0.3:nz(d.BP), BQ=BO*BP, BR=BO-BQ;
  return {AF,BJ,BK,BL,BM,BN,BO,BQ,BR};
}
const FORMULA={AF:r=>`IFERROR(AC${r}*(AD${r}+AE${r}),"")`,BJ:r=>`AF${r}`,BK:r=>`IFERROR(BJ${r}*BA${r},0)`,BL:r=>`IFERROR(BJ${r}*BE${r},0)`,BM:r=>`IFERROR(BJ${r}*BI${r},0)`,BN:r=>`BK${r}+BL${r}+BM${r}`,BO:r=>`BJ${r}-BN${r}`,BQ:r=>`BO${r}*BP${r}`,BR:r=>`BO${r}-BQ${r}`,F:r=>`IFERROR(D${r}*E${r},"")`,I:r=>`IFERROR(F${r}*H${r},"")`};

// ---------- form building ----------
function fieldHtml(mod,f){
  const id=`${mod}_${f.c}`, k=f.k?` data-k="1"`:"";
  const lab=`<label for="${id}">${esc(f.l)}</label>`;
  if(f.t==="comp") return `<div class="f comp"${k}><span class="lab">${esc(f.l)}</span><output id="${id}">—</output></div>`;
  if(f.t==="sel"){
    let opts=[];
    if(mod==="soz"&&f.c==="D") opts=SOZ_TYPES; else if(mod==="soz"&&f.c==="AK") opts=SOZ_DURUM; else if(mod==="kira"&&f.c==="T") opts=KIRA_DURUM;
    else if(mod==="firsat"&&f.c==="C") opts=SOZ_TYPES; else if(mod==="firsat"&&f.c==="G") opts=FIRSAT_DURUM;
    if(mod==="kira"&&f.c==="U") return `<div class="f"${k}>${lab}<select id="${id}"></select></div>`;
    return `<div class="f"${k}>${lab}<select id="${id}"><option value="">Seç…</option>${opts.map(o=>`<option>${esc(o)}</option>`).join("")}</select></div>`;
  }
  const ro = (mod==="soz" && (f.c==="A"||f.c==="B")) ? " readonly" : "";
  const type = f.t==="date" ? "date" : "text";
  const im = (f.t==="money"||f.t==="num"||f.t==="pct") ? ` inputmode="decimal"` : "";
  const list = PERSON[mod].includes(f.c) ? ` list="contactNames"` : LISTS[f.c]&&mod==="soz" ? ` list="dl_${f.c}"` : "";
  const suf = f.t==="pct" ? `<span class="suf">%</span>` : f.t==="money" ? `<span class="suf">₺</span>` : "";
  return `<div class="f"${k}>${lab}<div class="inwrap"><input id="${id}" type="${type}"${im}${list}${ro}>${suf}</div></div>`;
}
function buildModForm(mod){
  const fs=sch(mod);
  let html=`<button type="button" class="mobile-back" data-back="1">‹ Listeye dön</button><h2 id="${mod}Title"></h2><div class="meta" id="${mod}Meta"></div><div id="${mod}Extra"></div>`;
  if(mod==="soz"){
    for(const s of SCH.secs){
      const flds=fs.filter(f=>f.s===s.k); const kAll=flds.every(f=>f.k);
      html+=`<details class="sec" data-sec="${s.k}"${kAll?' data-k="1"':""}${s.k==="A"||s.k==="F"?" open":""}><summary>${esc(s.n)}<span class="cnt" id="cnt_${s.k}"></span></summary><div class="fgrid">${flds.map(f=>fieldHtml(mod,f)).join("")}</div></details>`;
    }
  } else {
    html+=`<div class="fgrid">${fs.map(f=>fieldHtml(mod,f)).join("")}</div>`;
  }
  html+=`<button class="save" type="submit">Kaydet</button><div class="delrow" id="${mod}Del"></div><button class="ghost" type="button" data-clear="1">Yeni kayıt / formu temizle</button>`;
  $(mod+"Form").innerHTML=html;
  if(mod==="soz"){ Object.entries(LISTS).forEach(([c,vals])=>{ if(!$("dl_"+c)) document.body.insertAdjacentHTML("beforeend",`<datalist id="dl_${c}">${vals.map(v=>`<option value="${esc(v)}">`).join("")}</datalist>`); }); }
  const form=$(mod+"Form");
  form.addEventListener("submit",e=>saveMod(mod,e));
  form.addEventListener("input",e=>onModInput(mod,e));
  form.addEventListener("change",e=>onModInput(mod,e));
  form.addEventListener("click",e=>{
    if(e.target.closest("[data-back]")) document.body.classList.remove("editing");
    if(e.target.closest("[data-clear]")) openMod(mod,null);
    const w=e.target.closest(".delwrap");
    if(w){ e.preventDefault();
      if(e.target.closest(".yes")){ w.innerHTML='<span class="ask">Siliniyor…</span>'; return delMod(mod,w.dataset.id); }
      if(e.target.closest(".no")){ w.innerHTML='<button type="button" class="del">Bu kaydı sil</button>'; return; }
      if(e.target.closest(".del")) w.innerHTML='<span class="ask">Silinsin mi?</span><button type="button" class="yes">Evet, sil</button><button type="button" class="no">Vazgeç</button>';
    }
    const oc=e.target.closest("[data-open-soz]"); if(oc){ showView("s"); openMod("soz",oc.dataset.openSoz); }
    const cv=e.target.closest("[data-firsat-convert]"); if(cv){ e.preventDefault(); convertFirsatToSoz(cv.dataset.firsatConvert); }
  });
}
function onModInput(mod,e){
  const id=e.target.id||""; const col=id.split("_")[1];
  if(mod==="soz"){
    if(col==="D") applyKiraVis();
    if(col==="C"){ const v=$("soz_C").value; if(isISO(v)) $("soz_B").value=v.slice(0,4); }
    if(["AC","AD","AE","BA","BE","BI","BP","D"].includes(col)) showCalc();
  }
  if(mod==="firsat"){
    if(col==="G"){ const v=$("firsat_G").value; if(v && FIRSAT_PROB_DEFAULT[v]!=null && !$("firsat_H").value) $("firsat_H").value=FIRSAT_PROB_DEFAULT[v]; }
    if(["D","E","H","G"].includes(col)) showFirsatCalc();
  }
  if(e.type==="change" && PAIRS[mod][col]){
    const c=S.contacts.get(norm(e.target.value));
    if(c){ const [tel,mail]=PAIRS[mod][col]; if(tel && !$(mod+"_"+tel).value) $(mod+"_"+tel).value=c.t||""; if(mail && $(mod+"_"+mail) && !$(mod+"_"+mail).value) $(mod+"_"+mail).value=c.em||""; }
  }
  if(mod==="kira" && col==="U") renderKiraLink();
  updCounts(mod);
}
function applyKiraVis(){
  const k=KIRA_TYPES.includes($("soz_D").value);
  document.querySelectorAll("#sozForm [data-k]").forEach(el=>el.classList.toggle("hidden",!k));
}
function readMod(mod){
  const d={};
  for(const f of sch(mod)){
    if(f.t==="comp") continue;
    const el=$(mod+"_"+f.c); if(!el) continue;
    let v=String(el.value||"").trim(); if(v==="") continue;
    if(f.t==="money"||f.t==="num") v=parseNum(v);
    else if(f.t==="pct"){ const n=parseNum(v); v=typeof n==="number" ? Math.round(n*10000)/1000000 : n; }
    if(v==="") continue;
    d[f.c]=v;
  }
  return d;
}
function showCalc(){
  const d=readMod("soz"); const c=calc(d);
  for(const [k,v] of Object.entries(c)){ const o=$("soz_"+k); if(o) o.textContent=d.AC?fmtTL(v):"—"; }
}
function showFirsatCalc(){
  const d=readMod("firsat"); const c=firsatCalc(d);
  const f=$("firsat_F"); if(f) f.textContent=d.D?fmtTL(c.F):"—";
  const i=$("firsat_I"); if(i) i.textContent=d.D?fmtTL(c.I):"—";
}
function updCounts(mod){
  if(mod!=="soz") return;
  for(const s of SCH.secs){
    const n=SCH.soz.filter(f=>f.s===s.k && f.t!=="comp" && $("soz_"+f.c) && String($("soz_"+f.c).value||"").trim()).length;
    const el=$("cnt_"+s.k); if(el) el.textContent=n?` · ${n} alan dolu`:"";
  }
}
function fillMod(mod,d){
  for(const f of sch(mod)){
    if(f.t==="comp") continue;
    const el=$(mod+"_"+f.c); if(!el) continue;
    let v=d?d[f.c]:"";
    if(f.t==="date"){ el.type = (v && !isISO(v)) ? "text" : "date"; }
    if(v==null) v="";
    if(f.t==="money") v = (typeof v==="number") ? fmtNum(v,0) : v;
    else if(f.t==="pct") v = (typeof v==="number") ? fmtNum(v*100,4) : v;
    else if(f.t==="num") v = (typeof v==="number") ? String(v) : v;
    el.value=v;
  }
}
function kiraSozOptions(){
  const opts=S.soz.filter(isKira).sort((a,b)=>String(b.A).localeCompare(String(a.A)));
  $("kira_U").innerHTML=`<option value="">Bağlı değil</option>`+opts.map(s=>`<option value="${esc(s.A)}">${esc(s.A+" · "+(s.E||"")+(s.R?" · "+s.R:""))}</option>`).join("");
}
function renderKiraLink(){
  const id=$("kira_U").value; const s=S.soz.find(x=>x.A===id);
  $("kiraExtra").innerHTML = s ? `<div class="linkcard"><b>Bağlı sözleşme ${esc(s.A)}</b><br>${esc([s.R?"Kiracı: "+s.R:"", s.AC?"Kira: "+fmtTL(s.AC):"", isISO(s.AJ)?"Bitiş: "+tr(s.AJ):""].filter(Boolean).join(" · "))}<br><button type="button" class="linkbtn" data-open-soz="${esc(s._id)}">Sözleşmeyi aç ›</button></div>` : "";
}
function openMod(mod,id){
  const arr=S[mod]; const d=id?arr.find(x=>x._id===id):null;
  S.modSel[mod]=d?d._id:null;
  if(mod==="kira") kiraSozOptions();
  fillMod(mod,d);
  if(mod==="soz"){
    if(!d){ $("soz_C").value=todayISO(); $("soz_B").value=todayISO().slice(0,4); $("soz_AK").value="🟡 Devam Ediyor"; $("soz_BP").value="30"; $("soz_A").value="(kayıtta verilecek)"; }
    $("sozTitle").textContent=d?`${d.A} · ${d.E||""}`:"Yeni sözleşme";
    $("sozMeta").textContent=d?[d.D,d.AK,d.R?("Karşı taraf: "+d.R):""].filter(Boolean).join(" · "):"İşlem türünü seç; kira seçilince kefil, bitiş ve kira artışı alanları açılır.";
    applyKiraVis(); showCalc();
  } else if(mod==="kira"){
    $("kiraTitle").textContent=d?d.A:"Yeni kiralık mülk";
    $("kiraMeta").textContent=d?[d.T,d.D?("Mal sahibi: "+d.D):""].filter(Boolean).join(" · "):"Mülk bilgilerini gir; varsa kira sözleşmesine bağla.";
    if(!d) $("kira_T").value="🔄 Kiracı Aranıyor";
    renderKiraLink();
  } else {
    $("firsatTitle").textContent=d?(d.A||"Fırsat"):"Yeni fırsat";
    $("firsatMeta").textContent=d?[d.C,d.G,d.J?("Beklenen kapanış: "+tr(d.J)):""].filter(Boolean).join(" · "):"Lead bilgilerini gir; durum seçince olasılık otomatik önerilir, istersen elle değiştir.";
    if(!d) $("firsat_G").value="🆕 Yeni Lead", $("firsat_H").value="10";
    showFirsatCalc();
    $("firsatExtra").innerHTML = d ? `<div class="linkcard">Sözleşme imzalandı mı? Bilgileri Sözleşme sekmesine aktarabilirsin.<br><button type="button" class="linkbtn" data-firsat-convert="${esc(d._id)}">📋 Sözleşmeye Aktar ›</button></div>` : "";
  }
  $(mod+"Del").innerHTML = d ? `<span class="delwrap" data-id="${esc(d._id)}"><button type="button" class="del">Bu kaydı sil</button></span>` : "";
  updCounts(mod); renderModList(mod);
  document.body.classList.add("editing");
}
function modHay(mod,d){ return norm(Object.values(d).filter(v=>typeof v==="string"||typeof v==="number").join(" ")); }
function renderModList(mod){
  const q=norm($(mod+"Q").value), dq=digits($(mod+"Q").value);
  let arr=[...S[mod]];
  if(q||dq) arr=arr.filter(d=>{ const h=modHay(mod,d); return (q && q.split(" ").every(t=>h.includes(t))) || (dq.length>=4 && digits(Object.values(d).join(" ")).includes(dq)); });
  if(mod==="soz") arr.sort((a,b)=>String(b.A||"").localeCompare(String(a.A||"")));
  else if(mod==="kira") { const ord={"⏳ Çıkış Hazırlığı":0,"🔄 Kiracı Aranıyor":1,"✅ Aktif":2,"⛔ Pasif":3}; arr.sort((a,b)=>(ord[a.T]??9)-(ord[b.T]??9)||String(a.A).localeCompare(String(b.A),"tr")); }
  else { const ord={"✅ Kazanıldı":9,"❌ Kaybedildi":10,"⏸️ Donduruldu":8}; arr.sort((a,b)=>(ord[a.G]??0)-(ord[b.G]??0)||(firsatCalc(b).I-firsatCalc(a).I)); }
  $(mod+"Hint").textContent = `${arr.length} kayıt`;
  $(mod+"List").innerHTML = arr.map(d=> mod==="soz"
    ? `<button type="button" class="person" data-id="${esc(d._id)}" aria-current="${S.modSel.soz===d._id}"><div class="nm">${esc(d.A||"")} · ${esc(d.E||"")}</div><div class="sub">${esc([d.D, d.R, d.AC?fmtTL(d.AC):"", d.AK].filter(Boolean).join("  ·  "))}</div></button>`
    : mod==="kira"
    ? `<button type="button" class="person" data-id="${esc(d._id)}" aria-current="${S.modSel.kira===d._id}"><div class="nm">${esc(d.A||"")}</div><div class="sub">${esc([d.T, d.F?("Kiracı: "+d.F):"", d.P?fmtTL(d.P):"", d.U?("Sözl. "+d.U):""].filter(Boolean).join("  ·  "))}</div></button>`
    : `<button type="button" class="person" data-id="${esc(d._id)}" aria-current="${S.modSel.firsat===d._id}"><div class="nm">${esc(d.A||"")}${d.G?`<span class="${d.G==="✅ Kazanıldı"?"pot pot-Low":d.G==="❌ Kaybedildi"?"pot pot-VIP":"pot pot-Medium"}">${esc(d.G)}</span>`:""}</div><div class="sub">${esc([d.C, firsatCalc(d).I?("Ağırlıklı: "+fmtTL(firsatCalc(d).I)):"", d.J?("Kapanış "+tr(d.J)):""].filter(Boolean).join("  ·  "))}</div></button>`
  ).join("") || `<div class="empty">Kayıt yok.</div>`;
  if(mod==="firsat") renderFirsatSummary();
}
function renderFirsatSummary(){
  const el=$("firsatSummary"); if(!el) return;
  const t=firsatTotals();
  el.innerHTML = t.total
    ? `<b>${t.open}</b> açık fırsat (toplam ${t.total}) · Potansiyel komisyon: <b>${fmtTL(t.F)}</b> · Ağırlıklı gelir tahmini: <b>${fmtTL(t.I)}</b>`
    : `Henüz fırsat kaydı yok. "＋ Yeni fırsat" ile ekleyebilirsin.`;
}

async function convertFirsatToSoz(id){
  const d=S.firsat.find(x=>x._id===id); if(!d) return;
  showView("s"); openMod("soz",null);
  if(d.C) $("soz_D").value=d.C;
  if(d.A) $("soz_R").value=d.A;
  if(d.B) $("soz_S").value=d.B;
  if(d.D) $("soz_AC").value=fmtNum(d.D,0);
  if(d.E) $("soz_AE").value=fmtNum(d.E*100,4);
  applyKiraVis(); showCalc(); updCounts("soz");
  banner(`"${d.A||"Fırsat"}" bilgileri Sözleşme formuna aktarıldı — kontrol edip Kaydet'e bas.`);
  if(d.G!=="✅ Kazanıldı"){
    try{ await writeMod("firsat",{...d,G:"✅ Kazanıldı",H:1},d); await loadMod("firsat"); renderModList("firsat"); renderToday(); saveCache(); }catch(e){ console.error(e); }
  }
}
// locate a record again in a fresh read (rows may have moved since the page loaded)
function findFresh(mod,cur){
  const arr=S[mod];
  if(mod==="soz") return arr.find(x=>x.A===cur.A);
  return arr.find(x=>x._r===cur._r && x.A===cur.A) || arr.find(x=>x.A===cur.A && (x.B||x.D||"")===(cur.B||cur.D||"")) || null;
}
function modRow(mod,d,r){
  const row=Array(colNum(LASTC[mod])).fill("");
  for(const f of sch(mod)){
    const ix=colNum(f.c)-1;
    if(f.t==="comp"){ row[ix]="="+FORMULA[f.c](r); continue; }
    let v=d[f.c];
    if(mod==="soz" && f.c==="BP" && (v===""||v==null)) v=0.3;
    if(f.t==="date" && isISO(v)) v=serial(v);
    row[ix]=cellIn(v);
  }
  return row;
}
// write d into the row of cur (update) or a new row at the end (cur = null); returns the sheet row
async function writeMod(mod,d,cur){
  await loadMod(mod);
  let r;
  if(cur){ const f=findFresh(mod,cur); if(!f) throw {code:"not_found"}; r=f._r; }
  else {
    if(mod==="soz"){ d.A=nextSozId(d.B); }
    r=S[mod].reduce((m,x)=>Math.max(m,x._r),START[mod]-1)+1;
    if(r>START[mod]){ try{ await graph(`${wsPath(SHEET[mod])}/range(address='A${r}:${LASTC[mod]}${r}')/insert`,{method:"POST",body:{shift:"Down"}}); }catch(e){ if(e&&e.code==="redirect") throw e; } }
  }
  await graph(`${wsPath(SHEET[mod])}/range(address='A${r}:${LASTC[mod]}${r}')`,{method:"PATCH",body:{formulas:[modRow(mod,d,r)]}});
  return r;
}
async function saveMod(mod,e){
  e.preventDefault();
  const d=readMod(mod);
  const cur=S.modSel[mod]?S[mod].find(x=>x._id===S.modSel[mod]):null;
  if(mod==="soz"){
    if(!d.D){ toast("İşlem türünü seç"); $("soz_D").focus(); return; }
    if(!d.E){ toast("Portföy adı gerekli"); $("soz_E").focus(); return; }
    if(!isISO(d.C)){ toast("Sözleşme tarihi gerekli"); $("soz_C").focus(); return; }
    d.B=Number(d.C.slice(0,4));
    if(cur) d.A=cur.A;
  } else if(mod==="kira"){
    if(!d.A){ toast("Portföy adı gerekli"); $("kira_A").focus(); return; }
  } else {
    if(!d.A){ toast("Ad soyad / lead gerekli"); $("firsat_A").focus(); return; }
    if(!d.G){ d.G="🆕 Yeni Lead"; }
  }
  const btn=$(mod+"Form").querySelector(".save"); btn.disabled=true; btn.textContent="Excel'e yazılıyor…";
  try{
    const r=await writeMod(mod,d,cur);
    await loadMod(mod); saveCache();
    toast(cur?"Güncellendi":"Kaydedildi");
    openMod(mod,mod+r); renderToday();
  }catch(err){ console.error(err); if(err&&err.code==="redirect") return; banner(errMsg(err,"Kaydedilemedi")); toast("Kaydedilemedi"); }
  finally{ btn.disabled=false; btn.textContent="Kaydet"; }
}
async function delMod(mod,id){
  const d=S[mod].find(x=>x._id===id); if(!d) return;
  toast("Siliniyor…");
  try{
    await loadMod(mod);
    const f=findFresh(mod,d); if(!f) throw {code:"not_found"};
    await graph(`${wsPath(SHEET[mod])}/range(address='A${f._r}:${LASTC[mod]}${f._r}')/delete`,{method:"POST",body:{shift:"Up"}});
    await loadMod(mod); saveCache();
    openMod(mod,null); renderToday();
    toast("Kayıt silindi");
  }catch(e){ console.error(e); if(e&&e.code==="redirect") return; openMod(mod,id); banner(errMsg(e,"Kayıt silinemedi")); }
}

function nextSozId(year){
  let mx=0; S.soz.forEach(d=>{ const m=String(d.A||"").match(/^(\d{4})-(\d+)$/); if(m && m[1]===String(year)) mx=Math.max(mx,Number(m[2])); });
  return `${year}-${String(mx+1).padStart(3,"0")}`;
}
// ---------- reminders ----------
function sozReminders(){
  const out=[]; const t=todayISO();
  for(const d of S.soz){
    if(!isKira(d) || String(d.AK||"").includes("İptal")) continue;
    // aynı mülk için daha yeni bir kira sözleşmesi varsa eskisini atla
    const newer=S.soz.some(x=>x!==d && isKira(x) && x.E && x.E===d.E && String(x.A)>String(d.A));
    if(newer) continue;
    if(isISO(d.AJ)){ const dd=dayDiff(d.AJ); if(dd>=-30 && dd<=60) out.push({dd,kind:"Kira bitişi",title:d.E,sub:[d.R?"Kiracı: "+d.R:"","Bitiş "+tr(d.AJ)].filter(Boolean).join(" · "),mod:"soz",id:d._id}); }
    const active=!isISO(d.AJ) || dayDiff(d.AJ)>=-30;
    const base=isISO(d.AH)?d.AH:(isISO(d.C)?d.C:"");
    const endingSoon=isISO(d.AJ) && dayDiff(d.AJ)<=30;
    if(active && base && !endingSoon){
      let y=Number(t.slice(0,4)); let ann=y+base.slice(4); if(dayDiff(ann)<-7){ y++; ann=y+base.slice(4); }
      const yrs=y-Number(base.slice(0,4)); const dd=dayDiff(ann);
      if(yrs>0 && dd>=-7 && dd<=30) out.push({dd,kind:`Kira artışı (${yrs}. yıl)`,title:d.E,sub:[d.R?"Kiracı: "+d.R:"", d.AC?"Mevcut kira: "+fmtTL(d.AC):"", "Yıldönümü "+tr(ann)].filter(Boolean).join(" · "),mod:"soz",id:d._id});
    }
  }
  for(const k of S.kira){
    if(String(k.T||"").includes("Çıkış")) out.push({dd:isISO(k.R)?dayDiff(k.R):0,kind:"Çıkış hazırlığı",title:k.A,sub:[k.F?"Kiracı: "+k.F:"", k.R?(isISO(k.R)?"Çıkış "+tr(k.R):k.R):""].filter(Boolean).join(" · "),mod:"kira",id:k._id});
  }
  return out.sort((a,b)=>a.dd-b.dd);
}
function remRow(r){
  const [l,cls]=dueLbl(r.dd);
  return `<div class="call" data-mod="${r.mod}" data-id="${esc(r.id)}" tabindex="0" role="button"><span class="due ${cls}">${l}</span><div><div class="nm">${esc(r.kind)} · ${esc(r.title||"")}</div><div class="sub">${esc(r.sub)}</div></div><span></span></div>`;
}
function showView(v){
  const map={k:"viewKayit",b:"viewBugun",s:"viewSoz",r:"viewKira",p:"viewFirsat"}, tabs={k:"tabKayit",b:"tabBugun",s:"tabSoz",r:"tabKira",p:"tabFirsat"};
  for(const [key,id] of Object.entries(map)){ $(id).classList.toggle("hidden",key!==v); $(tabs[key]).setAttribute("aria-selected",key===v); }
  document.body.classList.remove("editing");
  if(v==="b"){ renderToday(); if(!S.leadsAt || Date.now()-S.leadsAt>600000){ S.leadsAt=Date.now(); loadLeads(false); } }
  if(v==="s"){ renderModList("soz"); if(!S.modSel.soz) openMod("soz",null), document.body.classList.remove("editing"); }
  if(v==="r"){ renderModList("kira"); if(!S.modSel.kira) openMod("kira",null), document.body.classList.remove("editing"); }
  if(v==="p"){ renderModList("firsat"); if(!S.modSel.firsat) openMod("firsat",null), document.body.classList.remove("editing"); }
  try{ localStorage.setItem("gc_tab", v); }catch(_){}
  window.scrollTo(0,0);
}
function wire(){
  $("todayLbl").textContent=new Date().toLocaleDateString("tr-TR",{weekday:"long",day:"numeric",month:"long"});
  setVal("d",todayISO());
  $("q").addEventListener("input",renderResults);
  $("q").addEventListener("keydown",e=>{ if(e.key==="Enter"){ e.preventDefault(); const first=$("results").querySelector(".person"); if(first) selectContact(first.dataset.k); } });
  $("results").addEventListener("click",e=>{ const b=e.target.closest(".person"); if(b) selectContact(b.dataset.k); });
  $("newBtn").addEventListener("click",()=>{ const q=$("q").value.trim(); clearForm(/\d{4,}/.test(q)?"":q); if(/\d{4,}/.test(q)) setVal("t",q); document.body.classList.add("editing"); $(q && !/\d{4,}/.test(q)?"t":"n").focus(); });
  $("backBtn").addEventListener("click",()=>document.body.classList.remove("editing"));
  $("clearBtn").addEventListener("click",()=>clearForm());
  $("t").addEventListener("input",checkDup);
  $("dupWarn").addEventListener("click",e=>{ const b=e.target.closest("button"); if(b) selectContact(b.dataset.k); });
  $("histBox").addEventListener("click",e=>{
    const w=e.target.closest(".delwrap");
    if(w){
      e.preventDefault(); e.stopPropagation();
      if(e.target.closest(".yes")){ w.innerHTML='<span class="ask">Siliniyor…</span>'; return delRec(w.dataset.id); }
      if(e.target.closest(".no")){ w.innerHTML='<button type="button" class="del">Sil</button>'; return; }
      if(e.target.closest(".del")){ w.innerHTML='<span class="ask">Silinsin mi?</span><button type="button" class="yes">Evet, sil</button><button type="button" class="no">Vazgeç</button>'; }
      return;
    }
    if(e.target.id==="moreHist"){ S.histAll=!S.histAll; renderHist(); }
  });
  $("dQuick").innerHTML=[["Bugün",0],["Dün",1],["2 gün önce",2],["1 hafta önce",7]].map(([l,n])=>`<button type="button" class="chip" data-n="${n}">${l}</button>`).join("");
  const markD=()=>{ const v=$("d").value; $("dPast").classList.toggle("hidden", !v || v>=todayISO()); };
  $("dQuick").addEventListener("click",e=>{ const b=e.target.closest("button"); if(b){ setVal("d",addDays(todayISO(),-Number(b.dataset.n))); markD(); } });
  $("d").addEventListener("change",markD); $("d").addEventListener("input",markD);
  window.__markD=markD;
  $("quick").addEventListener("click",e=>{ const b=e.target.closest("button"); if(b) setVal("nd",addDays(todayISO(),Number(b.dataset.n))); });
  $("form").addEventListener("submit",save);
  $("syncBtn").addEventListener("click",()=>{ banner(""); refresh(true); });
  $("signoutBtn").addEventListener("click",()=>msalApp.logoutRedirect({account:S.account, postLogoutRedirectUri:BASE+"crm.html"}));
  const show=showView;
  $("tabKayit").onclick=()=>show("k"); $("tabBugun").onclick=()=>show("b"); $("tabSoz").onclick=()=>show("s"); $("tabKira").onclick=()=>show("r"); $("tabFirsat").onclick=()=>show("p");
  const openFromToday=e=>{ if(e.target.closest(".wa")) return; const r=e.target.closest(".call"); if(!r) return;
    if(r.dataset.mod){ show(r.dataset.mod==="soz"?"s":"r"); openMod(r.dataset.mod,r.dataset.id); return; }
    if(r.dataset.lead){ show("k"); openLead(r.dataset.lead); return; }
    show("k"); selectContact(r.dataset.k); };
  for(const mod of ["soz","kira","firsat"]){
    buildModForm(mod);
    $(mod+"Q").addEventListener("input",()=>renderModList(mod));
    $(mod+"New").addEventListener("click",()=>openMod(mod,null));
    $(mod+"List").addEventListener("click",e=>{ const b=e.target.closest(".person"); if(b) openMod(mod,b.dataset.id); });
  }
  $("todayBox").addEventListener("click",openFromToday);
  $("todayBox").addEventListener("keydown",e=>{ if(e.key==="Enter") openFromToday(e); });
  buildChips(); renderResults(); renderToday();
}


function syncLbl(){ const at=new Date(); $("syncLbl").textContent=`${S.contacts.size} kişi · ${S.rows.length} kayıt · Excel ${pad(at.getHours())}:${pad(at.getMinutes())}`; }
async function refresh(manual){
  $("syncLbl").textContent="Excel okunuyor…";
  try{
    await Promise.all([loadCrm(), loadMod("soz"), loadMod("kira"), loadMod("firsat")]);
    S.modsLoaded=true;
    const busy=document.activeElement && ["INPUT","TEXTAREA","SELECT"].includes(document.activeElement.tagName);
    rebuild(); if(!busy){ renderModList("soz"); renderModList("kira"); renderModList("firsat"); }
    saveCache(); syncLbl(); S.lastRefresh=Date.now();
    if(manual){ S.leadsAt=Date.now(); loadLeads(false); toast("Excel'den yenilendi"); }
  }catch(err){
    console.error(err); if(err && err.code==="redirect") return;
    $("syncLbl").textContent="Excel okunamadı"; banner(errMsg(err,"Excel okunamadı"));
  }
}
async function boot(){
  await msalApp.initialize();
  const rr=await msalApp.handleRedirectPromise().catch(e=>{ console.error(e); return null; });
  S.account=(rr && rr.account) || msalApp.getAllAccounts()[0] || null;
  if(!S.account){ $("signin").hidden=false; $("app").hidden=true; $("signinBtn").onclick=()=>msalApp.loginRedirect({scopes:SCOPES, prompt:"select_account", state:"crm"}); return; }
  msalApp.setActiveAccount(S.account);
  $("signin").hidden=true; $("app").hidden=false;
  wire();
  const had=loadCache();
  if(had){ rebuild(); renderModList("soz"); renderModList("kira"); renderModList("firsat"); $("syncLbl").textContent="Excel okunuyor…"; }
  const qs=new URLSearchParams(location.search);
  let tab="k"; try{ tab=localStorage.getItem("gc_tab")||"k"; }catch(_){}
  if(qs.get("tab")) tab=qs.get("tab");
  showView(["k","b","s","r","p"].includes(tab)?tab:"k");
  await refresh(false);
  const c=qs.get("c"); if(c){ const key=norm(c); if(S.contacts.has(key)){ showView("k"); selectContact(key); } else { showView("k"); $("q").value=c; renderResults(); } history.replaceState(null,"",location.pathname); }
  document.addEventListener("visibilitychange",()=>{ if(document.visibilityState==="visible" && Date.now()-(S.lastRefresh||0)>60000){ const busy=document.activeElement && ["INPUT","TEXTAREA","SELECT"].includes(document.activeElement.tagName); if(!busy) refresh(false); } });
}
boot();
})();
