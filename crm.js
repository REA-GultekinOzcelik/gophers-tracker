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
function saveCache(){ try{ localStorage.setItem("gc_cache", JSON.stringify({rows:S.rows, soz:S.soz, kira:S.kira, firsat:S.firsat, og:S.og, at:Date.now()})); }catch(_){} }
function loadCache(){ try{ const c = JSON.parse(localStorage.getItem("gc_cache") || "null"); if(!c) return false; S.rows=c.rows||[]; S.soz=c.soz||[]; S.kira=c.kira||[]; S.firsat=c.firsat||[]; S.og=c.og||[]; return true; }catch(_){ return false; } }

// ---------- helpers ----------
const pad = n => String(n).padStart(2,"0");
function isoLocal(d){ return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
function todayISO(){ return isoLocal(new Date()); }
function addDays(iso,n){ const [y,m,d]=iso.split("-").map(Number); const x=new Date(y,m-1,d); x.setDate(x.getDate()+n); return isoLocal(x); }
function dayDiff(iso){ if(!iso) return null; const [y,m,d]=iso.split("-").map(Number); const a=new Date(y,m-1,d); const t=new Date(); const b=new Date(t.getFullYear(),t.getMonth(),t.getDate()); return Math.round((a-b)/86400000); }
function tr(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return d+"."+m+"."+y; }
function norm(s){ return String(s||"").toLocaleLowerCase("tr").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/ı/g,"i").replace(/[^a-z0-9 ]+/g," ").replace(/\s+/g," ").trim(); }
function digits(s){ return String(s||"").replace(/\D/g,""); }
// one Telephone cell may hold several numbers: "cep | ikinci" (primary first)
function telParts(s){ const parts=String(s||"").split(/\s*\|\s*|\s+-\s+|\s*[\/;,]\s*|\s+\+\s+/).map(x=>x.trim()).filter(x=>digits(x).length>=7); return parts.length ? parts : (String(s||"").trim() ? [String(s).trim()] : []); }
function trMobile(s){ let d=digits(s); if(d.startsWith("9005")) d="90"+d.slice(3); if(d.length===10 && d[0]==="5") d="90"+d; if(d.length===11 && d.startsWith("05")) d="9"+d; return (d.length===12 && d.startsWith("905")) ? d : ""; }
function mobileFirst(arr){ return arr.slice().sort((a,b)=>(trMobile(b)?1:0)-(trMobile(a)?1:0)); }
function phoneKey(s){ const p=mobileFirst(telParts(s))[0]||""; const d=digits(p); return d.length>=10 ? d.slice(-10) : ""; }
function allPhoneKeys(s){ return telParts(s).map(p=>{ const d=digits(p); return d.length>=10 ? d.slice(-10) : ""; }).filter(Boolean); }
function waNum(s){ for(const p of mobileFirst(telParts(s))){ const w=trMobile(p); if(w) return w; } return ""; }
function joinTel(a,b){ return [a,b].map(x=>String(x||"").trim()).filter(Boolean).join(" | "); }
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
// ---------- Dokunuş planı (stored in the CRM "Touch Frequency" column) ----------
const PLANS=[["8x8","8x8 · yeni tanışma"],["Weekly","Haftalık"],["Bi-Weekly","2 haftada bir"],["Monthly","Aylık"],["2 Months","2 ayda bir"],["3 Months","3 ayda bir"],["Special Days","Sadece özel günler"]];
const PLAN_HINT={"8x8":"İlk 8 haftada her hafta bir dokunuş (8 dokunuş), sonra 2 ayda bir aramaya geçer.","Weekly":"Her hafta aranır.","Bi-Weekly":"2 haftada bir aranır.","Monthly":"Her ay aranır (A sınıfı).","2 Months":"2 ayda bir aranır (B sınıfı).","3 Months":"3 ayda bir aranır (C sınıfı).","Special Days":"Aranmaz; sadece doğum günü, yıldönümü ve bayramlarda dokunulur."};
const planLabel=v=>(PLANS.find(p=>p[0]===v)||[v,v])[1];
// special days a contact gets a message on. Kept in the same CRM cell: "<plan> + Special Days: a, b"
// Religious holidays move every year: first days per Diyanet calendar (update each year)
const RELIGIOUS={"Ramazan Bayramı":["2026-03-20","2027-03-09"],"Kurban Bayramı":["2026-05-27","2027-05-16"]};
const SD=[
  {id:"Doğum günü", own:true}, {id:"Ev alış yıldönümü", own:true}, {id:"Evlilik yıldönümü", own:true},
  {id:"Ramazan Bayramı", msg:"Ramazan Bayramınız mübarek olsun. Sevdiklerinizle sağlık ve huzur dolu nice bayramlara."},
  {id:"Kurban Bayramı", msg:"Kurban Bayramınız mübarek olsun. Sevdiklerinizle sağlık ve huzur dolu nice bayramlara."},
  {id:"Yılbaşı", md:"01-01", msg:"Yeni yılınız kutlu olsun. Sağlık, huzur ve bereket dolu bir yıl dilerim."},
  {id:"Cumhuriyet Bayramı", md:"10-29", msg:"Cumhuriyet Bayramımız kutlu olsun."},
  {id:"Noel", md:"12-25", msg:"Merry Christmas and happy holidays to you and your family!"},
  {id:"Thanksgiving", fn:y=>{ const d=new Date(y,10,1); const off=(4-d.getDay()+7)%7; return y+"-11-"+pad(1+off+21); }, msg:"Happy Thanksgiving to you and your family!"}
];
const SD_DEFAULT=["Doğum günü","Ev alış yıldönümü","Ramazan Bayramı","Kurban Bayramı","Yılbaşı"];   // old "Special Days" rows without a list
function parseTf(tf){
  const v=String(tf||"").trim(); if(!v) return {plan:"", sd:[]};
  const m=v.match(/special days\s*:\s*(.*)$/i);
  let plan=v.split("+")[0].trim(); if(/^special days/i.test(plan)) plan="Special Days";
  let sd = m ? m[1].split(",").map(x=>x.trim()).filter(Boolean) : (/special days/i.test(v) ? SD_DEFAULT.slice() : []);
  return {plan, sd};
}
function formatTf(plan, sd){
  if(!sd.length) return plan;
  if(!plan || plan==="Special Days") return "Special Days: "+sd.join(", ");
  return plan+" + Special Days: "+sd.join(", ");
}
function sdNext(d){
  const t=todayISO(), y=Number(t.slice(0,4));
  const cands=[];
  for(const yy of [y,y+1]){
    if(d.md) cands.push(yy+"-"+d.md);
    else if(d.fn) cands.push(d.fn(yy));
    else if(RELIGIOUS[d.id]) RELIGIOUS[d.id].filter(x=>x.startsWith(String(yy))).forEach(x=>cands.push(x));
  }
  return cands.filter(x=>dayDiff(x)>=0).sort()[0]||"";
}
// holiday groups within `days`: who gets a message
function holidaysUpcoming(days){
  const out=[];
  for(const d of SD){
    if(d.own) continue;
    const iso=sdNext(d); if(!iso) continue; const dd=dayDiff(iso); if(dd>days) continue;
    const people=[...S.contacts.values()].filter(c=>c.last.k!=="N" && parseTf(c.tf).sd.includes(d.id));
    if(people.length) out.push({d, iso, dd, people});
  }
  return out.sort((a,b)=>a.dd-b.dd);
}
function holidayHtml(h){
  const [l,cls]=h.dd===0?["Bugün","now"]:[h.dd===1?"Yarın":h.dd+" gün sonra","soon"];
  const text=h.d.msg+"\n\nGültekin";
  return `<details class="hol"><summary class="call" style="cursor:pointer"><span class="due ${cls}">${l}</span><div><div class="nm">🎉 ${esc(h.d.id)}</div><div class="sub">${tr(h.iso)} · ${h.people.length} kişiye mesaj · listeyi aç</div></div><span></span></summary>
    ${h.people.map(c=>{ const wa=waNum(c.t); return `<div class="call" data-k="${esc(c.key)}" data-hol="${esc(h.d.id)}" tabindex="0" role="button"><span></span><div><div class="nm">${esc(c.name)}</div><div class="sub">${esc(c.t||"telefon yok")}</div><div class="acts"><button type="button" class="act ok" data-act="holSent">✓ Gönderildi</button></div></div>${wa?`<a class="wa" href="https://wa.me/${wa}?text=${encodeURIComponent(text)}" target="_blank" rel="noopener" aria-label="WhatsApp mesajı">WA</a>`:`<span></span>`}</div>`; }).join("")}</details>`;
}
function planDays(tf){
  const v=parseTf(tf).plan.toLowerCase();
  if(!v || v.startsWith("special")) return 0;
  if(v.includes("daily")||v.startsWith("weekly")) return 7;
  if(v.startsWith("bi-weekly")) return 14;
  if(v.startsWith("monthly")) return 30;
  if(v.startsWith("2 month")) return 60;
  if(v.startsWith("3 month")) return 90;
  return 0;
}
// contacts whose plan says it is time to touch them again (and who have no explicit next action pending)
function planDue(skip){
  const out=[]; let unplanned=0;
  for(const c of S.contacts.values()){
    const r=c.last; if(r.k==="N" || norm(r.na).includes("olumsuz")) continue;
    if(skip.has(c.key)) continue;
    if(r.nd && dayDiff(r.nd)>=-7) continue;                      // explicit next action is pending / just due: other lists show it
    const tf=c.tf||""; if(!tf){ unplanned++; continue; }
    const since=c.lastDate?-dayDiff(c.lastDate):999;
    let cad, tag;
    if(/^8x8/i.test(parseTf(tf).plan)){
      const first=c.recs.reduce((m,x)=>x.d && (!m||x.d<m)?x.d:m,"");
      const el=first?-dayDiff(first):999;
      if(el<=63 && c.recs.length<8){ cad=7; tag=`8x8 · ${c.recs.length}/8`; } else { cad=60; tag="8x8 bitti · 2 ayda"; }
    } else { cad=planDays(tf); tag=planLabel(parseTf(tf).plan); }
    if(!cad || since<cad) continue;
    out.push({c, since, cad, tag, over:since/cad});
  }
  out.sort((x,y)=>x.cad-y.cad || y.over-x.over);
  return {due:out, unplanned};
}
const PLAN_DAILY=12;
// ---------- quick plan assignment for contacts without a plan ----------
const QP=[["Monthly","Aylık"],["2 Months","2 ayda"],["3 Months","3 ayda"],["Special Days","Özel gün"],["DROP","Çıkar"]];
function suggestPlan(c){
  const p=c.p||"", kt=norm(c.kt), since=c.lastDate?-dayDiff(c.lastDate):999;
  let plan;
  if(["VIP","Very High","High"].includes(p)) plan="Monthly";
  else if(p==="Medium") plan="2 Months";
  else if(p==="Low") plan="3 Months";
  else if(kt.includes("emlakci")) plan="3 Months";
  else if(kt.includes("personal")) plan="2 Months";
  else if(since<=180) plan="2 Months";
  else plan="3 Months";
  const sd = isForeignPhone(c.t||"") ? ["Noel","Thanksgiving"] : ["Ramazan Bayramı","Kurban Bayramı","Yılbaşı"];
  return {plan, sd};
}
function unplannedList(){
  return [...S.contacts.values()].filter(c=>c.last.k!=="N" && !norm(c.last.na).includes("olumsuz") && !c.tf)
    .sort((a,b)=>({"VIP":1,"Very High":2,"High":3,"Medium":4,"Low":5}[a.p]||6)-({"VIP":1,"Very High":2,"High":3,"Medium":4,"Low":5}[b.p]||6) || (b.lastDate||"").localeCompare(a.lastDate||""));
}
function renderQuickPlan(){
  const list=unplannedList(), page=list.slice(0,20);
  S.qp=S.qp||{};
  $("todayBox").innerHTML=`<div class="today-sec"><h3>Hızlı plan atama<span class="n">${list.length} kişinin planı yok · 20'şer 20'şer</span></h3>
    <p class="hint" style="margin:0 16px 8px">Her kişi için önerilen plan seçili geldi. Gerekirse değiştir, sonra en alttan kaydet. "Çıkar" kişiyi takipten çıkarır. Bayram ve yılbaşı mesajları da otomatik eklenir (yabancı numaralara Noel ve Thanksgiving).</p>
    ${page.map(c=>{ const sg=S.qp[c.key]||suggestPlan(c).plan; return `<div class="call qprow" data-k="${esc(c.key)}" style="cursor:default"><span></span><div><div class="nm">${esc(c.name)}${c.p?`<span class="${potCls(c.p)}">${esc(c.p)}</span>`:""}</div><div class="sub">${esc([c.kt, c.rf, c.lastDate?("son temas "+tr(c.lastDate)):"temas yok"].filter(Boolean).join("  ·  "))}</div><div class="chips qpchips" style="margin-top:6px">${QP.map(([v,l])=>`<button type="button" class="chip" data-qp="${v}" aria-pressed="${v===sg}">${l}</button>`).join("")}</div></div><span></span></div>`; }).join("") || `<div class="empty">Planı olmayan kişi kalmadı. 👏</div>`}
    <div class="qbtns" style="padding:12px 16px 16px">${page.length?`<button type="button" class="act ok" id="qpSave">Bu ${page.length} kişiyi kaydet</button>`:""}<button type="button" class="act" id="qpBack">Bugün'e dön</button></div></div>`;
  S.qpPage=page.map(c=>c.key);
}
async function saveQuickPlan(){
  const btn=$("qpSave"); btn.disabled=true; btn.textContent="Excel'e yazılıyor…";
  try{
    const head=S.crmHead, idx=S.crmIdx; if(!head) throw {code:"not_loaded"};
    const recs=[];
    for(const key of S.qpPage){
      const c=S.contacts.get(key); if(!c) continue; const last=c.last||{};
      const choice=S.qp[key]||suggestPlan(c).plan, sg=suggestPlan(c);
      const base={d:todayISO(), by:"Gültekin", n:c.name, t:c.t, rf:c.rf, p:c.p, kt:c.kt, ct:"Plan", nd:last.nd||"", na:last.na||""};
      recs.push(choice==="DROP" ? {...base, tf:"", k:"N", no:"Takipten çıkarıldı (toplu plan atama)"} : {...base, k:last.k||"Y", tf:formatTf(choice, sg.sd), no:"Dokunuş planı: "+planLabel(choice)+" (toplu atama)"});
    }
    const rows=recs.map(rec=>{ const vals=Array(head.length).fill(""); COLS.forEach(([k],j)=>{ const ix=idx[j]; if(ix<0) return; let v=rec[k]; if((k==="d"||k==="nd") && isISO(v)) v=serial(v); vals[ix]=cellIn(v); }); return vals; });
    await graph(`${WB}/tables('CRM')/rows/add`,{method:"POST",body:{values:rows}});
    let seq=S.rows.reduce((m,r)=>Math.max(m,r.seq||0),0);
    recs.forEach(r=>{ Object.keys(r).forEach(k=>{ if(r[k]==="") delete r[k]; }); r.seq=++seq; r._id="new"+seq; S.rows.push(r); });
    rebuild(); saveCache(); syncLbl();
    toast(`${recs.length} kişinin planı kaydedildi`);
    S.qp={}; renderQuickPlan();
    loadCrm().then(()=>{ rebuild(); saveCache(); if(S.quickPlan) renderQuickPlan(); }).catch(()=>{});
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; banner(errMsg(e,"Planlar kaydedilemedi")); btn.disabled=false; btn.textContent="Tekrar dene"; }
}
function planRow(x){
  const c=x.c, wa=waNum(c.t);
  return `<div class="call" data-k="${esc(c.key)}" tabindex="0" role="button">
    <span class="due ${x.over>=1.5?"late":"now"}">${esc(x.tag)}</span>
    <div><div class="nm">${esc(c.name)}${c.p?`<span class="${potCls(c.p)}">${esc(c.p)}</span>`:""}</div>
    <div class="sub">${esc([c.lastDate?("son temas "+x.since+" gün önce"):"hiç temas yok", c.rf].filter(Boolean).join("  ·  "))}</div>${actsHtml(false)}</div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}
function renderPlanChips(){
  const {plan, sd}=parseTf($("tf").value);
  const known=PLANS.some(p=>p[0]===plan);
  $("tfChips").innerHTML=PLANS.map(([v,l])=>`<button type="button" class="chip" data-v="${esc(v)}" aria-pressed="${v===plan}">${esc(l)}</button>`).join("")+(plan && !known?`<button type="button" class="chip" data-v="${esc(plan)}" aria-pressed="true">${esc(plan)}</button>`:"");
  $("tfHint").textContent = plan ? (PLAN_HINT[plan]||"Mevcut değer: "+plan) : "Plan seçilmedi: kişi sadece 'sonraki aksiyon tarihi' girersen listeye gelir.";
  $("sdChips").innerHTML=SD.map(d=>`<button type="button" class="chip" data-sd="${esc(d.id)}" aria-pressed="${sd.includes(d.id)}">${esc(d.id)}</button>`).join("");
  const c=S.sel?S.contacts.get(S.sel):null;
  const missing=c?sd.filter(x=>SD.find(d=>d.id===x&&d.own) && !ogFor(c.key).some(o=>o.type===x)):[];
  $("sdHint").textContent = missing.length ? `${missing.join(", ")} seçili ama tarihi girilmemiş; aşağıdaki "Önemli günler"den tarihini ekle.` : (sd.length?"Bu günlerde Bugün listesinde mesaj hatırlatması çıkar.":"Hiçbiri seçili değil.");
}
function buildChips(){
  renderPlanChips();
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
  // not (or not only) in CRM: look in the phone contacts exported to the Rehber folder
  const hits=q.trim().length>=2 ? rehberSearch(q) : [];
  if(hits.length){
    $("results").insertAdjacentHTML("beforeend", `<div class="rehh">📇 Telefon rehberinde · ${hits.length}${hits.length>=25?"+":""}</div>` + hits.map(l=>`
    <button type="button" class="person rehp" data-lead="${esc(l.id)}">
      <div class="nm">${esc(l.name)}<span class="pot pot-Low">Rehber</span></div>
      <div class="sub">${esc([l.phone1, l.email, l.org].filter(Boolean).join("  ·  "))}</div>
    </button>`).join(""));
    if(!res.length) $("qHint").textContent="CRM'de yok · telefon rehberinde bulundu. Seçince bilgileri forma gelir.";
  } else if(q.trim().length>=2 && S.leadsState==="ready" && !res.length){
    $("qHint").textContent = `CRM'de ve telefon rehberinde yok (${rehberInfo()}). Yeni kişi olarak ekleyebilirsin.`;
  } else if(q.trim().length>=2 && S.leadsState!=="ready"){
    $("qHint").textContent += S.leadsState==="loading" ? " · Telefon rehberi yükleniyor…" : S.leadsState==="error" ? " · Telefon rehberi okunamadı ("+S.leadsErr+")" : " · Telefon rehberi boş";
  }
}
// phone-contacts search (all Rehber entries not already in CRM by phone)
function rehberSearch(q){
  const leads=S.leads||[]; if(!leads.length) return [];
  const nq=norm(q), dq=digits(q), toks=nq.split(" ").filter(Boolean);
  const out=[];
  for(const l of leads){
    if(S.contacts.has(norm(l.name))) continue;            // already in CRM under the same name
    const ok=(dq.length>=4 && digits([l.phone1,l.phone2,l.phone3].join(" ")).includes(dq)) || (toks.length && toks.every(t=>norm(l.name+" "+(l.org||"")+" "+(l.email||"")).includes(t)));
    if(ok){ out.push(l); if(out.length>=25) break; }
  }
  return out;
}
function rehberSuggest(){
  const box=$("rehSug"); if(!box) return;
  const n=$("n").value.trim();
  if(S.sel || n.length<3){ box.innerHTML=""; box.hidden=true; return; }
  if(S.contacts.has(norm(n))){ box.innerHTML=""; box.hidden=true; return; }
  const hits=rehberSearch(n).slice(0,5);
  if(!hits.length && S.leadsState==="ready"){ box.hidden=false; box.innerHTML=`<p class="hint">📇 Telefon rehberinde bulunamadı (${esc(rehberInfo())}).</p>`; return; }
  if(!hits.length){ box.hidden=S.leadsState==="ready"; box.innerHTML=S.leadsState==="loading"?`<p class="hint">Telefon rehberi yükleniyor…</p>`:S.leadsState==="error"?`<p class="hint">Telefon rehberi okunamadı (${esc(S.leadsErr)}).</p>`:""; return; }
  box.hidden=false;
  box.innerHTML=`<div class="rehh">📇 Telefon rehberinde</div>`+hits.map(l=>`<button type="button" class="person rehp" data-fill="${esc(l.id)}"><div class="nm">${esc(l.name)}</div><div class="sub">${esc([l.phone1,l.email,l.org].filter(Boolean).join("  ·  "))}</div></button>`).join("");
}
function fillFromLead(id){
  const l=(S.leads||[]).find(x=>x.id===id); if(!l) return;
  setVal("n",l.name); setVal("t",l.phone1||l.phone2||l.phone3||""); setVal("t2",l.phone1?l.phone2:"");
  if(!$("em").value) setVal("em",l.email||""); if(!$("rf").value) setVal("rf",[l.org,l.title].filter(Boolean).join(" · "));
  if(isForeignPhone(l.phone1||"")){ setVal("tf",formatTf(parseTf($("tf").value).plan||"8x8",["Noel","Thanksgiving"])); renderPlanChips(); }
  $("fTitle").textContent=l.name; $("rehSug").innerHTML=""; $("rehSug").hidden=true; checkDup(); toast("Rehberden dolduruldu");
  setTimeout(()=>$("no").focus({preventScroll:true}),50);
}
function rehberInfo(){
  const n=(S.leads||[]).length, last=(S.leadFiles||[]).slice(-1)[0];
  return `rehber kopyası: ${last?tr(String(last.at).slice(0,10)):"?"}, ${n.toLocaleString("tr-TR")} kişi`;
}
function rehberByPhone(v){
  const pk=phoneKey(v); if(!pk) return null;
  return (S.leads||[]).find(l=>[l.phone1,l.phone2,l.phone3].some(p=>phoneKey(p)===pk)) || null;
}

// ---------- form ----------
function setVal(id,v){ $(id).value=v||""; if(id==="d" && window.__markD) window.__markD(); }
function clearForm(keepName){
  vReset();
  S.sel=null; S.histAll=false;
  const nm = keepName || "";
  ["n","t","t2","rf","no","na","em","pb","tf","ad"].forEach(i=>setVal(i,""));
  setVal("tf",formatTf("8x8",["Ramazan Bayramı","Kurban Bayramı","Yılbaşı"]));   // new people start on 8x8 + holidays
  S.sdTouched=false;
  setVal("n",nm); setVal("d",todayISO()); setVal("nd","");
  S.ct=""; S.kt=""; S.p=""; $("kN").checked=false;
  $("fTitle").textContent="Yeni kişi"; $("fMeta").textContent="Bilgileri gir, dokunuş planını ve özel günleri seç (hazır önerildi), sonra kaydet. Doğum günü gibi tarihli günleri kayıttan sonra açılan \"Önemli günler\"den eklersin.";
  $("histBox").innerHTML=""; $("ogBox").innerHTML=""; $("docBox").innerHTML=""; $("rehSug").innerHTML=""; $("rehSug").hidden=true; S.pendingPhoto=null; $("dupWarn").classList.add("hidden");
  S.ih=ihEmpty(); S.ihOrig=null; S.ihSrc=""; S.ihFilled=false; S.ihOpen=false; S.ihDirty=false; renderIh(); renderIhSum();
  buildChips(); renderResults();
}
function selectContact(key, focusNote=true){
  const c=S.contacts.get(key); if(!c) return;
  if(S.sel!==key) vReset();
  S.sel=key; S.histAll=false; S.ogOpen=false; S.docOpen=false; S.pendingPhoto=null;
  setVal("n",c.name); { const tp=mobileFirst(telParts(c.t)); setVal("t",tp[0]||""); setVal("t2",tp.slice(1).join(" | ")); } setVal("rf",c.rf); setVal("em",c.em); setVal("tf",c.tf);
  setVal("no",""); setVal("na",""); setVal("nd",""); setVal("pb",""); setVal("ad",""); setVal("d",todayISO());
  S.ct=""; S.kt=c.kt||""; S.p=c.p||""; $("kN").checked=false;
  $("fTitle").textContent=c.name;
  $("fMeta").textContent=`${c.recs.length} kayıt · son görüşme ${c.lastDate?tr(c.lastDate):"—"}`;
  $("dupWarn").classList.add("hidden");
  $("rehSug").innerHTML=""; $("rehSug").hidden=true;
  renderHist(); renderOg(); renderDocs(); buildChips(); renderResults();
  S.ihFilled=false; ihLoadFor(c.name, c.t);
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
      ${r.no?`<div class="what">${esc(r.no)}${isURL(r.x)?` · <a class="doclink" href="${esc(r.x)}" target="_blank" rel="noopener">Klasörü aç ›</a>`:""}</div>`:""}
      ${(r.na||r.nd)?`<div class="next">Sonraki: ${esc(r.na||"")}${r.nd?" ("+tr(r.nd)+")":""}</div>`:""}
      <span class="delwrap" data-id="${esc(r._id)}"><button type="button" class="del">Sil</button></span>
    </div>`).join("")}</div>${recs.length>3?`<button type="button" class="more-hist" id="moreHist">${S.histAll?"Daha az göster":`Tüm geçmiş (${recs.length})`}</button>`:""}`;
}
function checkDup(){
  const w=$("dupWarn");
  const pk=phoneKey($("t").value);
  if(!pk){ w.classList.add("hidden"); return; }
  const hit=[...S.contacts.values()].find(c=>c.pk===pk && c.key!==S.sel);
  if(hit){ w.innerHTML=`Bu telefon <b>${esc(hit.name)}</b> adına kayıtlı.<button type="button" data-k="${esc(hit.key)}">O kişiyi aç</button>`; w.classList.remove("hidden"); return; }
  const l=!S.sel && rehberByPhone($("t").value);
  if(l && (norm(l.name)!==norm($("n").value) || (!$("em").value && l.email))){ w.innerHTML=`Telefon rehberinde: <b>${esc(l.name)}</b>${l.email?" · "+esc(l.email):""}.<button type="button" data-fill="${esc(l.id)}">Bilgileri getir</button>`; w.classList.remove("hidden"); }
  else w.classList.add("hidden");
}


// append one log row to the CRM table and update the local model
async function addCrmRow(rec){
  Object.keys(rec).forEach(k=>{ if(rec[k]===""||rec[k]==null) delete rec[k]; });
  const head=S.crmHead, idx=S.crmIdx;
  if(!head) throw {code:"not_loaded", message:"CRM henüz yüklenmedi"};
  const vals=Array(head.length).fill("");
  COLS.forEach(([k],j)=>{ const ix=idx[j]; if(ix<0) return; let v=rec[k]; if((k==="d"||k==="nd") && isISO(v)) v=serial(v); vals[ix]=cellIn(v); });
  await graph(`${WB}/tables('CRM')/rows/add`,{method:"POST",body:{values:[vals]}});
  rec.seq=S.rows.reduce((m,r)=>Math.max(m,r.seq||0),0)+1; rec._id="new"+rec.seq;
  S.rows.push(rec);
  rebuild(); saveCache(); syncLbl();
  loadCrm().then(()=>{ rebuild(); saveCache(); if(S.sel) renderHist(); }).catch(()=>{});
}
// quick actions from Bugün: "Arandı" (log the call, optional next date) and "Takipten çıkar" (Keep = N)
async function quickAction(row, kind, opt){
  const key=row.dataset.k, lead=row.dataset.lead;
  const ct=key?S.contacts.get(key):null, l=lead?(S.leads||[]).find(x=>x.id===lead):null;
  let rec;
  if(ct){
    const base={d:todayISO(), by:"Gültekin", n:ct.name, t:ct.t, rf:ct.rf, p:ct.p, kt:ct.kt, tf:ct.tf};
    rec = kind==="done"
      ? {...base, ct:"Call Out", no:opt.note||"Arandı", k:"Y", nd:opt.nd||"", na:opt.nd?(opt.na||"Takip"):""}
      : {...base, no:opt.note||"Takipten çıkarıldı (ilgisiz)", k:"N"};
  } else if(l){
    rec={d:todayISO(), by:"Gültekin", n:l.name, t:joinTel(l.phone1,l.phone2)||l.phone3||"", em:l.email, rf:[l.org,l.title].filter(Boolean).join(" · "), kt:"Rehber", no:"Rehber adayı: ilgisiz, listeden çıkarıldı", k:"N"};
  } else return;
  row.querySelectorAll("button").forEach(b=>b.disabled=true);
  try{ await addCrmRow(rec); toast(kind==="done"?(opt.nd?"Arandı · sonraki "+tr(opt.nd):"Arandı olarak kaydedildi"):"Listeden çıkarıldı"); }
  catch(err){ console.error(err); if(err && err.code==="redirect") return; banner(errMsg(err,"Kaydedilemedi")); renderToday(); }
}
function actsHtml(lead){
  return lead
    ? `<div class="acts"><button type="button" class="act ok" data-act="leadAdd">＋ CRM'e ekle</button><button type="button" class="act" data-act="drop">İlgisiz</button></div>`
    : `<div class="acts"><button type="button" class="act ok" data-act="done">✓ Arandı</button><button type="button" class="act" data-act="drop">Takipten çıkar</button></div>`;
}
function onAct(e){
  const chip=e.target.closest(".qa .chip");
  if(chip){ e.stopPropagation(); if(chip.dataset.ls){ chip.setAttribute("aria-pressed", chip.getAttribute("aria-pressed")!=="true"); return true; } chip.parentNode.querySelectorAll(".chip").forEach(c=>c.setAttribute("aria-pressed", c===chip)); return true; }
  const b=e.target.closest("[data-act]"); if(!b) return false;
  e.preventDefault(); e.stopPropagation();
  const row=b.closest(".call"), box=row.querySelector(".acts"), a=b.dataset.act;
  if(a==="leadAdd"){
    const l=(S.leads||[]).find(x=>x.id===row.dataset.lead); if(!l) return true;
    const sd=isForeignPhone(l.phone1||"")?["Noel","Thanksgiving"]:["Ramazan Bayramı","Kurban Bayramı","Yılbaşı"];
    box.innerHTML=`<div class="qa"><span class="lab">Dokunuş planı</span><div class="chips lp">${PLANS.map(([v,lab],i)=>`<button type="button" class="chip" data-lp="${esc(v)}" aria-pressed="${i===0}">${esc(lab)}</button>`).join("")}</div>
      <span class="lab">Özel günler</span><div class="chips ls">${SD.filter(d=>!d.own).map(d=>`<button type="button" class="chip" data-ls="${esc(d.id)}" aria-pressed="${sd.includes(d.id)}">${esc(d.id)}</button>`).join("")}</div>
      <input class="qnote" placeholder="Not (isteğe bağlı): örn. tanışma araması yapıldı" aria-label="Not"><div class="qbtns"><button type="button" class="act ok" data-act="leadSave">Kaydet</button><button type="button" class="act" data-act="cancel">Vazgeç</button></div></div>`;
    return true;
  }
  if(a==="leadSave"){
    const l=(S.leads||[]).find(x=>x.id===row.dataset.lead); if(!l) return true;
    const plan=(box.querySelector('.lp .chip[aria-pressed="true"]')||{dataset:{}}).dataset.lp||"8x8";
    const sd=[...box.querySelectorAll('.ls .chip[aria-pressed="true"]')].map(x=>x.dataset.ls);
    const note=(box.querySelector(".qnote").value||"").trim();
    row.querySelectorAll("button").forEach(x=>x.disabled=true);
    addCrmRow({d:todayISO(), by:"Gültekin", n:l.name, t:joinTel(l.phone1,l.phone2)||l.phone3||"", em:l.email, rf:[l.org,l.title].filter(Boolean).join(" · "), kt:"Rehber", ct:"Plan", tf:formatTf(plan, sd), k:"Y", no:note||("Rehberden CRM'e eklendi · dokunuş planı: "+planLabel(plan))})
      .then(()=>toast(l.name+" CRM'e eklendi · "+planLabel(plan))).catch(err=>{ console.error(err); banner(errMsg(err,"Kaydedilemedi")); renderToday(); });
    return true;
  }
  if(a==="holSent"){
    const c=S.contacts.get(row.dataset.k); if(!c) return true; const last=c.last||{};
    b.disabled=true;
    addCrmRow({d:todayISO(), by:"Gültekin", n:c.name, t:c.t, rf:c.rf, p:c.p, kt:c.kt, tf:c.tf, ct:"WhatsApp", no:"🎉 "+row.dataset.hol+" mesajı gönderildi", k:last.k||"Y", nd:last.nd||"", na:last.na||""})
      .then(()=>{ b.textContent="✓ Kaydedildi"; }).catch(err=>{ b.disabled=false; banner(errMsg(err,"Kaydedilemedi")); });
    return true;
  }
  if(a==="done"){
    box.innerHTML=`<div class="qa"><span class="lab">Sonraki aksiyon</span><div class="chips">${[["Yok",0],["3 gün",3],["1 hafta",7],["2 hafta",14],["1 ay",30]].map(([t,n],i)=>`<button type="button" class="chip" data-nd="${n}" aria-pressed="${i===0}">${t}</button>`).join("")}</div><input class="qnote" placeholder="Kısa not (isteğe bağlı)" aria-label="Kısa not"><div class="qbtns"><button type="button" class="act ok" data-act="doneSave">Kaydet</button><button type="button" class="act" data-act="cancel">Vazgeç</button></div></div>`;
    setTimeout(()=>{ const i=box.querySelector(".qnote"); if(i) i.focus({preventScroll:true}); },30);
  } else if(a==="drop"){
    box.innerHTML=`<div class="qbtns"><span class="ask">${row.dataset.lead?"Bu aday bir daha gösterilmesin mi?":"Takipten çıkarılsın mı? Bir daha gelmez."}</span><button type="button" class="act warn" data-act="dropYes">Evet, çıkar</button><button type="button" class="act" data-act="cancel">Vazgeç</button></div>`;
  } else if(a==="cancel"){ renderToday(); }
  else if(a==="doneSave"){
    const sel=box.querySelector('.chip[aria-pressed="true"]'); const n=sel?Number(sel.dataset.nd):0;
    quickAction(row,"done",{nd:n?addDays(todayISO(),n):"", note:(box.querySelector(".qnote").value||"").trim()});
  } else if(a==="dropYes"){ quickAction(row,"drop",{}); }
  return true;
}
async function save(e){
  e.preventDefault();
  if(V.on){ await vStop(); }
  const n=$("n").value.trim(); let no=$("no").value.trim();
  if(!n){ $("n").focus(); toast("Ad soyad gerekli"); return; }
  const cur=S.sel?S.contacts.get(S.sel):null, tfNow=$("tf").value.trim();
  let carry=null;
  if(!no && cur && tfNow!==(cur.tf||"")){            // only the plan changed: log it, keep the pending follow-up
    const t=parseTf(tfNow); no="Dokunuş planı: "+(tfNow?[t.plan?planLabel(t.plan):"", t.sd.length?"özel günler: "+t.sd.join(", "):""].filter(Boolean).join(" · "):"kaldırıldı"); carry=cur.last||{};
  }
  if(!no && ihChanged()) no="İhtiyaç profili güncellendi";
  if(!no){ $("no").focus(); toast("Görüşme notu gerekli"); return; }
  const rec={d:$("d").value||todayISO(), by:"Gültekin", ct:S.ct, n, kt:S.kt, t:joinTel($("t").value,$("t2").value), em:$("em").value.trim(), ad:$("ad").value.trim(), pb:$("pb").value.trim(), rf:$("rf").value.trim(), no, k:$("kN").checked?"N":"Y", tf:$("tf").value.trim(), p:S.p, nd:$("nd").value, na:$("na").value.trim()};
  if(carry){ if(!rec.nd && carry.nd) rec.nd=carry.nd; if(!rec.na && carry.na) rec.na=carry.na; if(!S.ct) rec.ct="Plan"; }
  $("saveBtn").disabled=true; $("saveBtn").textContent="Excel'e yazılıyor…";
  try{
    await addCrmRow(rec);
    const key=norm(n);
    toast("Kaydedildi");
    if(S.pendingPhoto) afterSavePhoto(n);
    const ihSrcNow=V.used?"Ses notu":S.ihFilled?"Not":"Seçenek";
    vAfterSave(n);
    try{ if(await ihSave(n, rec.t, ihSrcNow)) toast("Kaydedildi · ihtiyaç profili güncellendi"); }
    catch(e2){ console.error(e2); if(e2 && e2.code==="redirect") return; banner(errMsg(e2,"İhtiyaç profili kaydedilemedi (görüşme kaydedildi)")); }
    $("q").value=""; selectContact(key,false); setVal("no",""); setVal("na",""); setVal("nd","");
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
    <div class="sub">${esc([r.na, ct.rf].filter(Boolean).join("  ·  "))}</div>${actsHtml(false)}</div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}
function renderToday(){
  if(S.quickPlan){ renderQuickPlan(); return; }
  const {a,b,c}=buckets();
  const shown=new Set([...a,...b.slice(0,10),...c].map(x=>x.ct.key));
  const pd=planDue(shown);
  const rem=sozReminders();
  const leads=todaysLeads();
  const nb=a.length+Math.min(PLAN_DAILY,pd.due.length)+rem.filter(r=>r.dd<=7).length+ogUpcoming(0).length;
  const badge=$("bugunBadge"); if(nb){ badge.textContent=nb; badge.classList.remove("hidden"); } else badge.classList.add("hidden");
  const sec=(title,items,extra)=>`<div class="today-sec"><h3>${title}<span class="n">${extra||items.length+" kişi"}</span></h3>${items.map(callRow).join("")||`<div class="empty">Bu bölümde kimse yok.</div>`}</div>`;
  const leadSec=`<div class="today-sec"><h3>Bugün görüşülecek yeni adaylar<span class="n">${leads.length} kişi</span></h3>${leads.map(leadRow).join("")||`<div class="empty">Rehberden eklenecek yeni aday yok.</div>`}</div>`;
  const ogs=ogUpcoming(7), hols=holidaysUpcoming(3);
  const ogSec = (ogs.length||hols.length) ? `<div class="today-sec"><h3>Önemli günler<span class="n">önümüzdeki 7 gün · ${ogs.length+hols.length}</span></h3>${hols.map(holidayHtml).join("")}${ogs.map(ogRow).join("")}</div>` : "";
  const ft=firsatTotals();
  const pipeSec = ft.total ? `<div class="today-sec"><h3>Fırsat Pipeline'ı<span class="n">${ft.open} açık / ${ft.total} toplam</span></h3><div class="call" style="cursor:default"><span></span><div><div class="nm">Ağırlıklı gelir tahmini: ${fmtTL(ft.I)}</div><div class="sub">Toplam potansiyel komisyon: ${fmtTL(ft.F)}</div></div><span></span></div></div>` : "";
  $("todayBox").innerHTML =
    ogSec + pipeSec +
    (rem.length?`<div class="today-sec"><h3>Sözleşme & kira hatırlatmaları<span class="n">${rem.length} hatırlatma</span></h3>${rem.map(remRow).join("")}</div>`:"") +
    sec("Bugün ve son 7 gün",a) +
    `<div class="today-sec"><h3>Dokunuş planı · sırası gelenler<span class="n">${Math.min(PLAN_DAILY,pd.due.length)} / ${pd.due.length} kişi</span></h3>${pd.due.slice(0,PLAN_DAILY).map(planRow).join("")||`<div class="empty">Bugün planı gelen kimse yok.</div>`}${unplannedList().length?`<div class="empty">Planı seçilmemiş ${unplannedList().length} kişi var. <button type="button" class="act ok" id="qpOpen">Hızlı plan ata ›</button></div>`:""}</div>` +
    sec("Öncelikli birikmiş takip",b.slice(0,10),`ilk 10 · toplam ${b.length} birikmiş`) +
    sec("Önümüzdeki 7 gün",c) +
    leadSec;
}

// ---------- new leads (Rehber) ----------
function leadPool(){
  const used=new Set(); for(const c of S.contacts.values()) for(const r of c.recs) allPhoneKeys(r.t).forEach(k=>used.add(k));
  return (S.leads||[]).filter(l=>{
    const ks=[l.phone1,l.phone2,l.phone3].map(p=>phoneKey(p)).filter(Boolean);
    return ks.length && !ks.some(k=>used.has(k));                    // daily call candidates need a phone; email-only people stay searchable
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
    <div class="sub">${esc(sub)}</div>${actsHtml(true)}</div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}
function openLead(id){
  const l=(S.leads||[]).find(x=>x.id===id); if(!l) return;
  clearForm(l.name);
  setVal("t", l.phone1||l.phone2||l.phone3||""); setVal("t2", l.phone1?l.phone2:"");
  setVal("em", l.email||"");
  setVal("rf", [l.org,l.title].filter(Boolean).join(" · "));
  $("fTitle").textContent=l.name;
  $("fMeta").textContent="Rehberden yeni aday · henüz CRM'de kayıtlı değil. Aşağıdan dokunuş planını ve özel günleri seç, sonra kaydet.";
  if(isForeignPhone(l.phone1||"")){ setVal("tf",formatTf("8x8",["Noel","Thanksgiving"])); renderPlanChips(); }
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
    const tels=mobileFirst(vFieldAll(lines,"TEL").map(t=>t.trim()).filter(Boolean));
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
  let cache={files:{}}; try{ cache=JSON.parse(localStorage.getItem("gc_leads2")||"null")||{files:{}}; }catch(_){}
  S.leadsState="loading"; S.leadsErr="";
  try{
    if(lbl) lbl.textContent="Rehber kontrol ediliyor…";
    const res=await graph(`/drives/${DRIVE_ID}/items/${REHBER_FOLDER_ID}/children?$select=id,name,eTag,createdDateTime,file&$top=200`);
    const files=(res.value||[]).filter(f=>f.file && /\.(vcf|txt)$/i.test(f.name)).sort((a,b)=>String(a.createdDateTime).localeCompare(String(b.createdDateTime)));
    const next={files:{}};
    for(const f of files){
      const old=cache.files[f.id];
      if(old && old.eTag===f.eTag && !force){ next.files[f.id]=old; continue; }
      if(lbl) lbl.textContent="İşleniyor: "+f.name;
      const meta=await graph(`/drives/${DRIVE_ID}/items/${f.id}`);
      const dl=meta["@microsoft.graph.downloadUrl"]; if(!dl) throw {code:"no_url", message:"indirme linki alınamadı"};
      const resp=await fetch(dl); if(!resp.ok) throw {code:"http_"+resp.status, message:"rehber dosyası indirilemedi"};
      const text=await resp.text();
      next.files[f.id]={eTag:f.eTag, name:f.name, at:f.createdDateTime, cards:parseVCards(text).map(c=>[c.name,c.phones.slice(0,3),c.email,c.org,c.title])};
    }
    cache=next;
    try{ localStorage.setItem("gc_leads2", JSON.stringify(cache)); }catch(_){}
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; S.leadsErr=(e&&(e.message||e.code))||"bilinmeyen hata"; }
  S.leadFiles=Object.values(cache.files).map(f=>({name:f.name, at:f.at})).sort((a,b)=>String(a.at).localeCompare(String(b.at)));
  // build the lead list: one entry per phone, first file wins
  const used=new Set(); const out=[]; let seq=0;
  for(const f of Object.values(cache.files).sort((a,b)=>String(a.at).localeCompare(String(b.at)))){
    for(const [name,phones,email,org,title] of f.cards){
      // drop only true duplicates (same name and phone); shared office numbers must not hide people
      const sig=norm(name)+"|"+(phones.map(phoneKey).filter(Boolean)[0]||"");
      if(used.has(sig)) continue;
      used.add(sig);
      seq++;
      out.push({id:"lead_"+seq, seq, name, phone1:phones[0]||"", phone2:phones[1]||"", phone3:phones[2]||"", email:email||"", org:org||"", title:title||"", isForeign:isForeignPhone(phones[0]||""), noPhone:!phones.length});
    }
  }
  S.leads=out; S.leadsState = out.length ? "ready" : (S.leadsErr ? "error" : "empty");
  if(S.contacts.size) syncLbl();
  if($("q").value) renderResults();
  if($("n").value && !S.sel) rehberSuggest();
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
    $("sozExtra").innerHTML=sozDocHtml(d);
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

// ================= Önemli günler (birthdays, home-purchase anniversaries …) =================
// Stored in their own sheet/table of the CRM workbook ("🎂 Özel Günler" / OzelGunler), created on first use.
// Sales contracts add the buyer's title-deed anniversary automatically (read-only, from Sözleşme).
const OG_SHEET="🎂 Özel Günler", OG_TABLE="OzelGunler", OG_HEAD=["Contact name","Telephone","Occasion","Date","Note","Added"];
const OG_TYPES=["Doğum günü","Ev alış yıldönümü","Evlilik yıldönümü","Diğer"];
const OG_ICON={"Doğum günü":"🎂","Ev alış yıldönümü":"🏠","Evlilik yıldönümü":"💍"};
const MON_TR=["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran","Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"];
S.og=[]; S.ogType="Doğum günü";
function ogParse(v){
  if(typeof v==="number" && v>2000){ const iso=fromSerial(v); const [y,m,d]=iso.split("-").map(Number); return {d,m,y}; }
  const s=String(v||"").trim().replace(/^'/,""); let x;
  if((x=s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return {d:+x[3],m:+x[2],y:+x[1]};
  if((x=s.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{4}))?$/))) return {d:+x[1],m:+x[2],y:x[3]?+x[3]:0};
  return null;
}
function ogText(p){ return pad(p.d)+"."+pad(p.m)+(p.y?"."+p.y:""); }
function ogLabel(p){ return p.d+" "+MON_TR[p.m-1]+(p.y?" "+p.y:""); }
function ogNext(p){
  const t=todayISO(); let y=Number(t.slice(0,4));
  const mk=yy=>{ const leap=(yy%4===0&&yy%100!==0)||yy%400===0; const d=(p.m===2&&p.d===29&&!leap)?28:p.d; return yy+"-"+pad(p.m)+"-"+pad(d); };
  let iso=mk(y); if(dayDiff(iso)<0){ y++; iso=mk(y); }
  return {iso, dd:dayDiff(iso), years:p.y?y-p.y:0};
}
async function loadOg(){
  try{
    const r=await graph(`${WB}/tables('${OG_TABLE}')/range?$select=values`);
    const head=r.values[0].map(h=>String(h).trim()), ix=h=>head.indexOf(h);
    S.og=r.values.slice(1).map((c,i)=>({i, n:String(c[ix("Contact name")]??"").trim(), t:String(c[ix("Telephone")]??""), type:String(c[ix("Occasion")]??"").trim(), p:ogParse(c[ix("Date")]), note:String(c[ix("Note")]??"")})).filter(o=>o.n && o.p);
    S.ogMissing=false;
  }catch(e){ if(e && e.code==="http_404"){ S.og=[]; S.ogMissing=true; } else throw e; }
}
async function ensureOg(){
  if(!S.ogMissing) return;
  try{ await graph(`${WB}/worksheets/add`,{method:"POST",body:{name:OG_SHEET}}); }catch(e){ if(e && e.code==="redirect") throw e; }
  await graph(`${wsPath(OG_SHEET)}/range(address='A1:F1')`,{method:"PATCH",body:{values:[OG_HEAD]}});
  const tb=await graph(`${wsPath(OG_SHEET)}/tables/add`,{method:"POST",body:{address:"A1:F1",hasHeaders:true}});
  await graph(`${WB}/tables/${encodeURIComponent(tb.id||tb.name)}`,{method:"PATCH",body:{name:OG_TABLE}});
  S.ogMissing=false;
}
// all occasions known for the app: the table + buyers' deed anniversaries from sales contracts
function ogAll(){
  const out=S.og.map(o=>({...o, src:"og"}));
  for(const d of S.soz||[]){
    if(isKira(d) || String(d.AK||"").includes("İptal") || !d.R) continue;
    const base=isISO(d.AH)?d.AH:(isISO(d.C)?d.C:""); if(!base) continue;
    const [y,m,dd]=base.split("-").map(Number);
    out.push({n:d.R, t:d.S||"", type:"Ev alış yıldönümü", p:{d:dd,m,y}, note:(d.E||"")+" · sözleşme "+(d.A||""), src:"soz"});
  }
  return out;
}
function ogFor(key){ return ogAll().filter(o=>norm(o.n)===key); }
function ogUpcoming(days){
  return ogAll().map(o=>({...o, nx:ogNext(o.p)})).filter(o=>o.nx.dd>=0 && o.nx.dd<=days).sort((a,b)=>a.nx.dd-b.nx.dd);
}
function ogWhat(o){
  const yrs=o.nx?o.nx.years:ogNext(o.p).years;
  const extra = yrs>0 ? (o.type==="Doğum günü" ? ` · ${yrs}. yaş` : ` · ${yrs}. yıl`) : "";
  return (OG_ICON[o.type]||"📌")+" "+o.type+extra;
}
function renderOg(){
  const box=$("ogBox"); if(!box) return;
  const c=S.contacts.get(S.sel);
  if(!c){ box.innerHTML=""; return; }
  const items=ogFor(c.key).map(o=>({...o, nx:ogNext(o.p)})).sort((a,b)=>a.nx.dd-b.nx.dd);
  const types=[...OG_TYPES]; if(S.ogType && !types.includes(S.ogType)) types.splice(3,0,S.ogType);
  box.innerHTML=`<details class="sec ogsec"${items.length||S.ogOpen?" open":""}><summary>Önemli günler<span class="cnt">${items.length?" · "+items.length:""}</span></summary>
    ${items.map(o=>`<div class="ogitem"><div><div class="nm">${esc(ogWhat(o))}</div><div class="sub">${esc(ogLabel(o.p))} · ${o.nx.dd===0?"bugün":o.nx.dd+" gün sonra"}${o.note?" · "+esc(o.note):""}</div></div>${o.src==="og"?`<span class="delwrap ogdel" data-i="${o.i}"><button type="button" class="del">Sil</button></span>`:`<span class="sub">otomatik</span>`}</div>`).join("") || `<p class="hint">Henüz kayıtlı gün yok.</p>`}
    <div class="ogadd">
      <div class="chips" id="ogTypes">${types.map(t=>`<button type="button" class="chip" data-v="${esc(t)}" aria-pressed="${t===S.ogType}">${esc(t)}</button>`).join("")}</div>
      <input id="ogOther" class="otherin${S.ogType==="Diğer"?"":" hidden"}" placeholder="Ne günü? (örn. Şirket kuruluşu)" aria-label="Gün türü">
      <div class="row"><div class="f"><label for="ogDate">Tarih</label><input id="ogDate" type="date"></div>
      <div class="f"><label for="ogNote">Not (isteğe bağlı)</label><input id="ogNote" placeholder="Örn. eşi Ayşe"></div></div>
      <label class="keep" style="margin-top:8px"><input type="checkbox" id="ogNoYear"> Yılı bilinmiyor</label>
      <button type="button" class="act ok" id="ogAddBtn" style="margin-top:10px">Günü ekle</button>
    </div></details>`;
}
async function addOg(){
  const c=S.contacts.get(S.sel); if(!c) return;
  let type=S.ogType; if(type==="Diğer"){ type=($("ogOther").value||"").trim(); if(!type){ toast("Günün adını yaz"); $("ogOther").focus(); return; } }
  const v=$("ogDate").value; if(!isISO(v)){ toast("Tarihi seç"); $("ogDate").focus(); return; }
  const [y,m,d]=v.split("-").map(Number); const p={d,m,y:$("ogNoYear").checked?0:y};
  if(ogFor(c.key).some(o=>o.type===type && o.p.d===p.d && o.p.m===p.m)){ toast("Bu gün zaten kayıtlı"); return; }
  const btn=$("ogAddBtn"); btn.disabled=true; btn.textContent="Excel'e yazılıyor…";
  try{
    await loadOg(); await ensureOg();
    await graph(`${WB}/tables('${OG_TABLE}')/rows/add`,{method:"POST",body:{values:[[c.name, cellIn(c.t||""), type, "'"+ogText(p), $("ogNote").value.trim(), serial(todayISO())]]}});
    await loadOg(); saveCache();
    S.ogOpen=true; toast("Önemli gün eklendi"); renderOg(); renderToday();
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; banner(errMsg(e,"Önemli gün eklenemedi")); btn.disabled=false; btn.textContent="Günü ekle"; }
}
async function delOg(i){
  const o=S.og.find(x=>x.i===Number(i)); if(!o) return;
  try{
    await loadOg();
    const f=S.og.find(x=>x.n===o.n && x.type===o.type && x.p.d===o.p.d && x.p.m===o.p.m); if(!f) throw {code:"not_found"};
    await graph(`${WB}/tables('${OG_TABLE}')/rows/itemAt(index=${f.i})`,{method:"DELETE"});
    await loadOg(); saveCache(); toast("Silindi"); renderOg(); renderToday();
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; banner(errMsg(e,"Silinemedi")); renderOg(); }
}
function wireOg(){
  $("ogBox").addEventListener("keydown",e=>{ if(e.key==="Enter" && e.target.tagName==="INPUT"){ e.preventDefault(); if(e.target.id!=="ogOther") addOg(); } });
  $("ogBox").addEventListener("click",e=>{
    const ch=e.target.closest("#ogTypes .chip");
    if(ch){ S.ogType=ch.dataset.v; document.querySelectorAll("#ogTypes .chip").forEach(x=>x.setAttribute("aria-pressed",x===ch)); $("ogOther").classList.toggle("hidden",S.ogType!=="Diğer"); if(S.ogType==="Diğer") $("ogOther").focus(); return; }
    if(e.target.id==="ogAddBtn"){ addOg(); return; }
    const w=e.target.closest(".ogdel");
    if(w){ e.preventDefault();
      if(e.target.closest(".yes")){ w.innerHTML='<span class="ask">Siliniyor…</span>'; return delOg(w.dataset.i); }
      if(e.target.closest(".no")){ w.innerHTML='<button type="button" class="del">Sil</button>'; return; }
      if(e.target.closest(".del")) w.innerHTML='<span class="ask">Silinsin mi?</span><button type="button" class="yes">Evet</button><button type="button" class="no">Vazgeç</button>';
    }
  });
}
function ogRow(o){
  const ct=S.contacts.get(norm(o.n)); const wa=waNum(o.t||(ct&&ct.t)||"");
  const [l,cls]=o.nx.dd===0?["Bugün","now"]:[o.nx.dd===1?"Yarın":o.nx.dd+" gün sonra","soon"];
  return `<div class="call" ${ct?`data-k="${esc(ct.key)}"`:""} tabindex="0" role="button">
    <span class="due ${cls}">${l}</span>
    <div><div class="nm">${esc(o.n)}</div><div class="sub">${esc(ogWhat(o))} · ${esc(ogLabel(o.p))}${o.note?" · "+esc(o.note):""}</div></div>
    ${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:`<span></span>`}
  </div>`;
}

// ================= Dokümanlar (files → OneDrive) & Fotoğraftan kişi (OCR) =================
// Files go to OneDrive: Gophers CRM/Dokümanlar/Kişiler/<ad> or Gophers CRM/Dokümanlar/Sözleşmeler/<ID portföy>
const CRM_FOLDER_ID = "01HMEMOXXXL3AP2YVEZ5ALRH7V52RZSG3K";
const safeName = s => String(s||"").replace(/[\\/:*?"<>|#%]/g," ").replace(/\s+/g," ").trim().replace(/^\.+|\.+$/g,"").slice(0,80) || "Adsız";
const encPath = segs => segs.map(encodeURIComponent).join("/");
function docPath(kind,label){ return ["Dokümanlar", kind==="soz"?"Sözleşmeler":"Kişiler", safeName(label)]; }
async function ensureFolder(segs){
  let parent=CRM_FOLDER_ID, item=null;
  for(const seg of segs){
    try{ item=await graph(`/drives/${DRIVE_ID}/items/${parent}:/${encodeURIComponent(seg)}?$select=id,webUrl,folder`); }
    catch(e){
      if(!(e && e.code==="http_404")) throw e;
      try{ item=await graph(`/drives/${DRIVE_ID}/items/${parent}/children`,{method:"POST",body:{name:seg, folder:{}, "@microsoft.graph.conflictBehavior":"fail"}}); }
      catch(e2){ if(e2 && e2.code==="http_409") item=await graph(`/drives/${DRIVE_ID}/items/${parent}:/${encodeURIComponent(seg)}?$select=id,webUrl,folder`); else throw e2; }
    }
    parent=item.id;
  }
  return item;
}
async function listDocs(segs){
  try{ const r=await graph(`/drives/${DRIVE_ID}/items/${CRM_FOLDER_ID}:/${encPath(segs)}:/children?$select=name,webUrl,size,lastModifiedDateTime,file&$orderby=lastModifiedDateTime desc&$top=50`); return (r.value||[]).filter(x=>x.file); }
  catch(e){ if(e && e.code==="http_404") return []; throw e; }
}
async function uploadFile(folderId, file, nameOverride){
  const name=safeName(nameOverride||file.name).slice(0,120);
  const tok=await token();
  if(file.size <= 3.5*1024*1024){
    const res=await fetch(`https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=rename`,{method:"PUT",headers:{"Authorization":"Bearer "+tok,"Content-Type":file.type||"application/octet-stream"},body:file});
    if(!res.ok){ let m=""; try{ m=(await res.json()).error.message; }catch(_){} throw {code:"http_"+res.status, message:m}; }
    return res.json();
  }
  const ses=await graph(`/drives/${DRIVE_ID}/items/${folderId}:/${encodeURIComponent(name)}:/createUploadSession`,{method:"POST",body:{item:{"@microsoft.graph.conflictBehavior":"rename"}}});
  const CH=5*1024*1024; let last=null;          // multiple of 320 KiB
  for(let off=0; off<file.size; off+=CH){
    const end=Math.min(off+CH,file.size);
    const res=await fetch(ses.uploadUrl,{method:"PUT",headers:{"Content-Range":`bytes ${off}-${end-1}/${file.size}`},body:file.slice(off,end)});
    if(!res.ok && res.status!==202){ throw {code:"http_"+res.status, message:"yükleme kesildi"}; }
    last=res;
  }
  return last.json();
}
function fmtSize(n){ return n>1048576 ? (n/1048576).toFixed(1)+" MB" : Math.max(1,Math.round(n/1024))+" KB"; }
// contact documents panel (Kayıt form)
async function renderDocs(){
  const box=$("docBox"); if(!box) return;
  const c=S.contacts.get(S.sel);
  if(!c){ box.innerHTML=$("n").value.trim()?`<p class="hint">📎 Doküman eklemek için önce kişiyi kaydet.</p>`:""; return; }
  const key=c.key;
  box.innerHTML=`<details class="sec docsec" open><summary>📎 Dokümanlar<span class="cnt" id="docCnt"></span></summary>
    <div id="docList"><p class="hint">Okunuyor…</p></div>
    <label class="act ok filebtn">📎 Doküman ekle<input type="file" id="docFile" multiple hidden></label>
    <p class="hint">Dosya OneDrive'da "Gophers CRM › Dokümanlar › Kişiler › ${esc(safeName(c.name))}" klasörüne kaydedilir.</p></details>`;
  try{
    const docs=await listDocs(docPath("kisi",c.name));
    if(S.sel!==key) return;
    $("docCnt").textContent=docs.length?" · "+docs.length:"";
    $("docList").innerHTML=docs.map(d=>`<div class="ogitem"><div><a class="doclink" href="${esc(d.webUrl)}" target="_blank" rel="noopener">${esc(d.name)}</a><div class="sub">${fmtSize(d.size)} · ${tr(String(d.lastModifiedDateTime).slice(0,10))}</div></div></div>`).join("") || `<p class="hint">Henüz doküman yok.</p>`;
  }catch(e){ console.error(e); if($("docList")) $("docList").innerHTML=`<p class="hint">Dokümanlar okunamadı.</p>`; }
}
async function uploadDocs(files, kind, label, onDone){
  if(!files || !files.length) return;
  banner(""); toast(files.length>1?`${files.length} dosya yükleniyor…`:"Yükleniyor…");
  try{
    const folder=await ensureFolder(docPath(kind,label));
    let i=0; for(const f of files){ i++; if(files.length>1) toast(`Yükleniyor ${i}/${files.length}: ${f.name}`); await uploadFile(folder.id,f); }
    toast(files.length>1?`${files.length} dosya OneDrive'a eklendi`:"Doküman OneDrive'a eklendi");
    if(onDone) onDone(folder);
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; banner(errMsg(e,"Doküman yüklenemedi")); }
}
function wireDocs(){
  $("docBox").addEventListener("toggle",e=>{ if(e.target.matches(".docsec")) S.docOpen=e.target.open; },true);
  $("docBox").addEventListener("change",e=>{
    if(e.target.id!=="docFile") return;
    const c=S.contacts.get(S.sel); if(!c) return;
    const files=[...e.target.files]; e.target.value="";
    S.docOpen=true;
    uploadDocs(files,"kisi",c.name,folder=>{
      renderDocs();
      // leave a trace in the CRM history; keep the contact's current follow-up (next action, keep flag) unchanged
      const last=c.last||{};
      addCrmRow({d:todayISO(), by:"Gültekin", n:c.name, t:c.t, rf:c.rf, p:c.p, kt:c.kt, tf:c.tf, ct:"Doküman", no:"📎 Doküman eklendi: "+files.map(f=>f.name).join(", "), x:folder.webUrl||"", k:last.k||"Y", nd:last.nd||"", na:last.na||""})
        .then(()=>{ if(S.sel===c.key) renderHist(); }).catch(e=>{ console.error(e); banner(errMsg(e,"Doküman yüklendi ama CRM geçmişine yazılamadı")); });
    });
  });
  // contract documents (Sözleşme form): upload + fill the "Cloud Belge Linki" field
  $("sozForm").addEventListener("change",e=>{
    if(e.target.id!=="sozDocFile") return;
    const d=S.soz.find(x=>x._id===S.modSel.soz); if(!d) return;
    const files=[...e.target.files]; e.target.value="";
    uploadDocs(files,"soz",`${d.A} ${d.E||""}`,folder=>{
      if(!$("soz_AL").value.trim() && folder.webUrl){ $("soz_AL").value=folder.webUrl; updCounts("soz"); banner('Belge klasörünün linki "📎 Cloud Belge Linki" alanına yazıldı. Kalıcı olması için Kaydet\'e bas.'); }
    });
  });
}
function sozDocHtml(d){
  return d ? `<div class="linkcard">📎 Sözleşme belgelerini (imzalı sözleşme, tapu, kimlik…) OneDrive'a ekle.<br><label class="linkbtn filebtn">Belge ekle ›<input type="file" id="sozDocFile" multiple hidden></label>${isURL(d.AL)?` · <a class="linkbtn" href="${esc(d.AL)}" target="_blank" rel="noopener">Klasörü aç ›</a>`:""}</div>` : "";
}
const isURL = v => /^https?:\/\//i.test(String(v||""));

// ---------- Fotoğraftan kişi (kartvizit, WhatsApp profili, e-posta imzası …) ----------
const TESS_URL="https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js";
function loadScript(src){ return new Promise((ok,no)=>{ if(window.Tesseract) return ok(); const s=document.createElement("script"); s.src=src; s.onload=ok; s.onerror=()=>no({code:"script", message:"OCR kütüphanesi yüklenemedi"}); document.head.appendChild(s); }); }
async function shrink(file){
  const url=URL.createObjectURL(file);
  try{
    const img=await new Promise((ok,no)=>{ const i=new Image(); i.onload=()=>ok(i); i.onerror=no; i.src=url; });
    const k=Math.min(1,2000/Math.max(img.width,img.height));
    const cv=document.createElement("canvas"); cv.width=Math.round(img.width*k); cv.height=Math.round(img.height*k);
    const g=cv.getContext("2d"); g.filter="grayscale(1) contrast(1.25)"; g.drawImage(img,0,0,cv.width,cv.height);
    return cv;
  } finally { URL.revokeObjectURL(url); }
}
const ORG_RE=/(?<![\p{L}])(a\.?\s?ş\.?|ltd|şti|sti|inc|llc|corp|gmbh|holding|grup|group|emlak|gayrimenkul|realty|real estate|keller|williams|remax|re\/max|century|coldwell|inşaat|insaat|yapı|yapi|danışmanlık|consulting|bank|banka|hukuk|law|ofis|office)(?![\p{L}])/iu;
const TITLE_RE=/(?<![\p{L}])(müdür|mudur|müdürü|danışman|danışmanı|danisman|uzman|uzmanı|direktör|direktörü|direktor|başkan|başkanı|baskan|yönetici|yönetici|yonetici|kurucu|ortak|avukat|mimar|mühendis|muhendis|manager|director|ceo|cfo|coo|founder|partner|agent|broker|realtor|consultant|specialist|associate|owner|sales|satış|satis)(?![\p{L}])/iu;
function parseCard(text){
  const lines=text.split(/\n/).map(l=>l.replace(/[|_~`»«•·]+/g," ").replace(/\s+/g," ").trim()).filter(l=>l.length>1);
  const all=lines.join(" \n ");
  const emails=[...new Set((all.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g)||[]).map(e=>e.toLowerCase()))];
  const phones=[]; for(const m of all.matchAll(/(\+?\(?\d[\d\s().\-]{8,}\d)/g)){
    let d=m[1].replace(/\D/g,"");
    if(d.length===10 && /^[2-5]/.test(d)) d="90"+d; else if(d.length===10) d="1"+d; else if(d.length===11 && d[0]==="0") d="9"+d;
    if(d.length<11 || d.length>13) continue;
    const pretty = d.startsWith("90")&&d.length===12 ? `+90 ${d.slice(2,5)} ${d.slice(5,8)} ${d.slice(8,10)} ${d.slice(10)}` : (d.startsWith("1")&&d.length===11 ? `+1 ${d.slice(1,4)} ${d.slice(4,7)} ${d.slice(7)}` : "+"+d);
    if(!phones.some(p=>p.d===d)) phones.push({d,pretty,mobile:/^905/.test(d)||/^1/.test(d)});
  }
  phones.sort((a,b)=>(b.mobile?1:0)-(a.mobile?1:0));
  let name="", org="", title="";
  for(const l of lines){
    if(/@|www\.|https?:|\d{3}/i.test(l)) continue;
    if(!org && ORG_RE.test(l)){ org=l; continue; }
    if(!title && TITLE_RE.test(l) && l.split(" ").length<=6){ title=l; continue; }
    const w=l.split(" ").filter(Boolean);
    if(!name && w.length>=2 && w.length<=4 && w.every(x=>/^[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'.-]+$/.test(x) || /^[A-ZÇĞİÖŞÜ'.-]{2,}$/.test(x))) name=w.map(x=>x.length>2&&x===x.toLocaleUpperCase("tr")?x[0]+x.slice(1).toLocaleLowerCase("tr"):x).join(" ");
  }
  return {name, phone:phones[0]?phones[0].pretty:"", phone2:phones[1]?phones[1].pretty:"", email:emails[0]||"", org, title, raw:lines.join("\n")};
}
async function photoContact(file){
  if(!file) return;
  banner("");
  const lbl=$("photoBtnTxt"), old=lbl.textContent;
  try{
    lbl.textContent="Fotoğraf okunuyor…";
    await loadScript(TESS_URL);
    const cv=await shrink(file);
    const worker=await Tesseract.createWorker(["tur","eng"],1,{logger:m=>{ if(m.status==="recognizing text") lbl.textContent=`Okunuyor %${Math.round(m.progress*100)}`; else if(/load|download/i.test(m.status)) lbl.textContent="İlk kullanım: dil paketi indiriliyor…"; }});
    const {data}=await worker.recognize(cv); await worker.terminate();
    const p=parseCard(data.text||"");
    clearForm(p.name); S.pendingPhoto=file;
    setVal("t",p.phone); setVal("t2",p.phone2||""); setVal("em",p.email); setVal("rf",[p.org,p.title].filter(Boolean).join(" · "));
    S.kt=""; buildChips();
    $("fTitle").textContent=p.name||"Fotoğraftan yeni kişi";
    $("fMeta").innerHTML=`Fotoğraftan okundu; bilgileri kontrol et, görüşme notunu yaz ve kaydet.${p.phone2?` İkinci telefon: <b>${esc(p.phone2)}</b>.`:""} <label class="keep" style="margin-top:6px"><input type="checkbox" id="photoKeep" checked> Fotoğrafı kişinin dokümanlarına ekle</label><details class="extra"><summary>Okunan metin</summary><pre class="ocr">${esc(p.raw)||"(metin bulunamadı)"}</pre></details>`;
    document.body.classList.add("editing"); checkDup();
    setTimeout(()=>$(p.name?"no":"n").focus({preventScroll:true}),50);
    if(!p.name && !p.phone && !p.email) banner("Fotoğrafta isim, telefon ya da e-posta bulunamadı. Daha net, düz çekilmiş bir fotoğraf dene ya da bilgileri elle gir.");
  }catch(e){ console.error(e); S.pendingPhoto=null; banner(errMsg(e,"Fotoğraf okunamadı")); }
  finally{ lbl.textContent=old; }
}
// after a successful save of a photo-created contact, keep the photo in the contact's documents
function afterSavePhoto(name){
  const f=S.pendingPhoto; S.pendingPhoto=null;
  const keep=$("photoKeep"); if(!f || (keep && !keep.checked)) return;
  const ext=(f.name.match(/\.[a-z0-9]+$/i)||[".jpg"])[0];
  ensureFolder(docPath("kisi",name)).then(folder=>uploadFile(folder.id,f,"Kartvizit "+todayISO()+ext)).then(()=>{ toast("Fotoğraf dokümanlara eklendi"); renderDocs(); }).catch(e=>{ console.error(e); banner(errMsg(e,"Fotoğraf dokümanlara eklenemedi")); });
}

function showView(v){
  const map={k:"viewKayit",b:"viewBugun",s:"viewSoz",r:"viewKira",p:"viewFirsat",o:"viewPf"}, tabs={k:"tabKayit",b:"tabBugun",s:"tabSoz",r:"tabKira",p:"tabFirsat",o:"tabPf"};
  for(const [key,id] of Object.entries(map)){ $(id).classList.toggle("hidden",key!==v); $(tabs[key]).setAttribute("aria-selected",key===v); }
  document.body.classList.remove("editing");
  if(v!=="b") S.quickPlan=false;
  if(v==="b"){ renderToday(); if(!S.leadsAt || Date.now()-S.leadsAt>600000){ S.leadsAt=Date.now(); loadLeads(false); } }
  if(v==="s"){ renderModList("soz"); if(!S.modSel.soz) openMod("soz",null), document.body.classList.remove("editing"); }
  if(v==="r"){ renderModList("kira"); if(!S.modSel.kira) openMod("kira",null), document.body.classList.remove("editing"); }
  if(v==="o"){ renderPfList(); if(!S.pfEd) openPf(null); document.body.classList.remove("editing"); }
  if(v==="p"){ renderModList("firsat"); if(!S.modSel.firsat) openMod("firsat",null), document.body.classList.remove("editing"); }
  try{ localStorage.setItem("gc_tab", v); }catch(_){}
  window.scrollTo(0,0);
}

// ================= İhtiyaç profili (Beş Bilgi Grubu · EKM eğitimi) =================
// Görüşme formunda seçeneklerle işaretlenir; sesli/yazılı nottan "Nottan doldur" ile otomatik ön-doldurulur.
// Veri: CRM dosyasında "🧭 İhtiyaç Profili" sayfası (tablo Ihtiyac) · değişiklikler "🕓 Profil Değişiklikleri" (tablo ProfilDegisiklik).
const IH_SHEET="🧭 İhtiyaç Profili", IH_TABLE="Ihtiyac";
const IH_HEAD=["Contact name","Telephone","Güncelleme","Kaynak","Rol","İşlem türü","Bölge","Alternatif bölge","Tip","Oda","Min m²","Min bütçe","Max bütçe","Para birimi","Esneklik","Öz kaynak","Kredi","Satıp alacak","İşlem zamanı","Olmazsa olmaz","Olsa iyi olur","İstenmeyen","Teyit edilecek","Tamlık %"];
const IL_SHEET="🕓 Profil Değişiklikleri", IL_TABLE="ProfilDegisiklik";
const IL_HEAD=["Tarih","Contact name","Telephone","Alan","Eski değer","Yeni değer","Kaynak"];
const IH_OPT={
  rol:["Alıcı","Satıcı","Kiracı","Ev sahibi","Yatırımcı"],
  it:["Satılık","Kiralık"],
  bolge:["Çankaya","Oran","Gölbaşı","İncek","Çayyolu","Ümitköy","Yaşamkent","Konutkent","Bilkent","Beytepe","Dodurga","Alacaatlı","Bağlıca","Kızılcaşar","Etimesgut","Yenimahalle","Kavaklıdere","Gaziosmanpaşa","Çukurambar","Balgat","Dikmen","Keçiören"],
  tip:["Daire","Villa","Müstakil","Arsa","Ticari","Ofis"],
  oda:["1+1","2+1","3+1","4+1","5+1","6+1"],
  pb:["TL","USD","EUR"],
  esn:["Sabit","%10'a kadar","Uygun ev olursa artar"],
  oz:["Tamamı nakit","%50'den fazla","%50'den az","Bilmiyor"],
  kredi:["Evet","Hayır","Henüz bilmiyor"],
  satip:["Evet","Hayır"],
  zaman:["Hemen","1–3 ay","3–6 ay","6 ay+","Sadece bakıyor"],
};
const IH_FEAT=["Site içi","Asansör","Kapalı otopark","Ebeveyn banyosu","Kapalı mutfak","Sosyal donatı","Okula yakın","Giriş kat","Manzara","Bahçe","Havuz","Güvenlik"];
const IH_FKEY={ "Site içi":["site ici","site icinde","sitede"], "Asansör":["asansor"], "Kapalı otopark":["kapali otopark","otopark","garaj"], "Ebeveyn banyosu":["ebeveyn banyo"], "Kapalı mutfak":["kapali mutfak"], "Sosyal donatı":["sosyal donati","sosyal alan"], "Okula yakın":["okula yakin","okul"], "Giriş kat":["giris kat","bahce kati","zemin kat"], "Manzara":["manzara"], "Bahçe":["bahce"], "Havuz":["havuz"], "Güvenlik":["guvenlik"] };
const IH_FST=["","Olmazsa olmaz","Olsa iyi olur","İstemem"];      // tıklama sırası
const IH_KEYS=[["rol","Rol"],["it","İşlem türü"],["bolge","Bölge"],["tip","Tip"],["oda","Oda"],["bmax","Bütçe"],["pb","Para birimi"],["oz","Öz kaynak"],["kredi","Kredi"],["zaman","İşlem zamanı"]];
S.ih={}; S.ihOrig=null; S.ihRows=[]; S.ihMissing=false; S.ihSrc="";

function ihEmpty(){ return {rol:"",it:"",bolge:[],alt:"",tip:"",oda:[],m2:"",bmin:"",bmax:"",pb:"",esn:"",oz:"",kredi:"",satip:"",zaman:"",feat:{},teyit:""}; }
const ihList=v=>String(v||"").split(/\s*[,;]\s*/).map(x=>x.trim()).filter(Boolean);
function ihNum(v){
  if(typeof v==="number") return isFinite(v)?v:"";
  let s=String(v||"").toLocaleLowerCase("tr").replace(/\s+/g," ").trim(); if(!s) return "";
  let mul=1; if(/milyon|\bm\b|mn/.test(s)) mul=1e6; else if(/\bbin\b|\bk\b/.test(s)) mul=1e3;
  s=s.replace(/[^\d.,]/g,"");
  if(/^\d{1,3}([.,]\d{3})+$/.test(s)) s=s.replace(/[.,]/g,""); else s=s.replace(",",".");
  const n=parseFloat(s); return isFinite(n)?Math.round(n*mul):"";
}
const ihFmt=n=>n===""||n==null?"":Number(n).toLocaleString("tr-TR");
function ihFromRow(c,ix){
  const g=h=>{ const v=c[ix(h)]; return v==null?"":(typeof v==="string"?v.replace(/^'/,""):v); };
  const p=ihEmpty();
  p.rol=String(g("Rol")); p.it=String(g("İşlem türü")); p.bolge=ihList(g("Bölge")); p.alt=String(g("Alternatif bölge")); p.tip=String(g("Tip")); p.oda=ihList(g("Oda"));
  p.m2=ihNum(g("Min m²")); p.bmin=ihNum(g("Min bütçe")); p.bmax=ihNum(g("Max bütçe")); p.pb=String(g("Para birimi")); p.esn=String(g("Esneklik"));
  p.oz=String(g("Öz kaynak")); p.kredi=String(g("Kredi")); p.satip=String(g("Satıp alacak")); p.zaman=String(g("İşlem zamanı")); p.teyit=String(g("Teyit edilecek"));
  for(const [h,st] of [["Olmazsa olmaz","Olmazsa olmaz"],["Olsa iyi olur","Olsa iyi olur"],["İstenmeyen","İstemem"]]) for(const f of ihList(g(h))) p.feat[f]=st;
  return p;
}
function ihFlat(p){       // karşılaştırma ve yazma için alan → metin
  const fl=st=>IH_FEAT.concat(Object.keys(p.feat).filter(f=>!IH_FEAT.includes(f))).filter(f=>p.feat[f]===st).join(", ");
  return {"Rol":p.rol,"İşlem türü":p.it,"Bölge":p.bolge.join(", "),"Alternatif bölge":p.alt,"Tip":p.tip,"Oda":p.oda.join(", "),"Min m²":p.m2===""?"":String(p.m2),
    "Min bütçe":p.bmin===""?"":String(p.bmin),"Max bütçe":p.bmax===""?"":String(p.bmax),"Para birimi":p.pb,"Esneklik":p.esn,"Öz kaynak":p.oz,"Kredi":p.kredi,"Satıp alacak":p.satip,
    "İşlem zamanı":p.zaman,"Olmazsa olmaz":fl("Olmazsa olmaz"),"Olsa iyi olur":fl("Olsa iyi olur"),"İstenmeyen":fl("İstemem"),"Teyit edilecek":p.teyit};
}
function ihHasData(p){ return Object.values(ihFlat(p)).some(v=>String(v).trim()); }
function ihMissingList(p){ return IH_KEYS.filter(([k])=>{ const v=p[k]; return Array.isArray(v)?!v.length:(v===""||v==null); }).map(([,l])=>l); }
function ihScore(p){ return (IH_KEYS.length-ihMissingList(p).length)/IH_KEYS.length; }

async function loadIh(){
  try{
    const r=await graph(`${WB}/tables('${IH_TABLE}')/range?$select=values`);
    const head=r.values[0].map(h=>String(h).trim()), ix=h=>head.indexOf(h);
    S.ihRows=r.values.slice(1).map((c,i)=>{ const g=c[ix("Güncelleme")]; return {i, n:String(c[ix("Contact name")]??"").trim(), t:String(c[ix("Telephone")]??""), u:(typeof g==="number"&&g>2000)?fromSerial(g):"", p:ihFromRow(c,ix)}; }).filter(x=>x.n);
    S.ihMissing=false;
  }catch(e){ if(e && e.code==="http_404"){ S.ihRows=[]; S.ihMissing=true; } else throw e; }
  try{ await loadIhLog(); }catch(e){ console.error(e); if(e && e.code==="redirect") throw e; }
}
async function ensureTable(sheet, table, head){
  try{ await graph(`${WB}/tables('${table}')?$select=name`); return; }catch(e){ if(!(e && e.code==="http_404")) throw e; }
  try{ await graph(`${WB}/worksheets/add`,{method:"POST",body:{name:sheet}}); }catch(e){ if(e && e.code==="redirect") throw e; }
  const end=colLetter(head.length);
  await graph(`${wsPath(sheet)}/range(address='A1:${end}1')`,{method:"PATCH",body:{values:[head]}});
  const tb=await graph(`${wsPath(sheet)}/tables/add`,{method:"POST",body:{address:`A1:${end}1`,hasHeaders:true}});
  await graph(`${WB}/tables/${encodeURIComponent(tb.id||tb.name)}`,{method:"PATCH",body:{name:table}});
}
function colLetter(n){ let s=""; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }
function ihFind(name, tel){
  const pk=phoneKey(tel), nk=norm(name);
  return S.ihRows.find(x=>pk && phoneKey(x.t)===pk) || S.ihRows.find(x=>norm(x.n)===nk) || null;
}

// Kayıtlı profil + formda bilinen eski hali (orig) + formdaki yeni hali (cur) → birleşik profil.
// Boş bırakılan alan, yalnızca formda daha önce dolu görünüp kullanıcı tarafından silindiyse silinir.
function ihMerge(base, orig, cur){
  const m=JSON.parse(JSON.stringify(base)), emp=v=>v===""||v==null||(Array.isArray(v)&&!v.length);
  for(const k of ["rol","it","alt","tip","m2","bmin","bmax","pb","esn","oz","kredi","satip","zaman","teyit","bolge","oda"]){
    const c=cur[k], o=orig?orig[k]:"";
    if(!emp(c)) m[k]=Array.isArray(c)?[...c]:c; else if(orig && !emp(o)) m[k]=Array.isArray(c)?[]:"";
  }
  const fs=new Set([...Object.keys(m.feat),...Object.keys(cur.feat),...Object.keys(orig?orig.feat:{})]);
  for(const f of fs){ const c=cur.feat[f], o=orig?orig.feat[f]:undefined; if(c) m.feat[f]=c; else if(o) delete m.feat[f]; }
  return m;
}
// Değişiklik kaydında boşaltılmış görünen alanları geri getir
function ihClearedFor(name, tel){
  const pk=phoneKey(tel), nk=norm(name), last={};
  for(const r of S.ihLog||[]) if((pk && phoneKey(r.t)===pk) || norm(r.n)===nk) last[r.alan]=r;
  return Object.values(last).filter(r=>r.eski && !r.yeni);
}
function ihApplyField(p, alan, v){
  const map={"Rol":"rol","İşlem türü":"it","Alternatif bölge":"alt","Tip":"tip","Para birimi":"pb","Esneklik":"esn","Öz kaynak":"oz","Kredi":"kredi","Satıp alacak":"satip","İşlem zamanı":"zaman","Teyit edilecek":"teyit"};
  if(map[alan]) p[map[alan]]=v;
  else if(alan==="Bölge") p.bolge=ihList(v);
  else if(alan==="Oda") p.oda=ihList(v);
  else if(alan==="Min m²") p.m2=ihNum(v);
  else if(alan==="Min bütçe") p.bmin=ihNum(v);
  else if(alan==="Max bütçe") p.bmax=ihNum(v);
  else if(alan==="Olmazsa olmaz") ihList(v).forEach(f=>p.feat[f]="Olmazsa olmaz");
  else if(alan==="Olsa iyi olur") ihList(v).forEach(f=>p.feat[f]="Olsa iyi olur");
  else if(alan==="İstenmeyen") ihList(v).forEach(f=>p.feat[f]="İstemem");
}
function ihRestore(){
  const cl=ihClearedFor($("n").value.trim(), $("t").value); if(!cl.length) return;
  ihReadInputs(); const p=S.ih; let k=0;
  for(const r of cl){ const before=JSON.stringify(p); const tmp=JSON.parse(JSON.stringify(p)); ihApplyField(tmp, r.alan, r.eski);
    // yalnızca şu an boş olan alanlara yaz
    const cur=ihFlat(p)[r.alan]; if(!String(cur||"").trim()){ Object.assign(p, tmp); if(JSON.stringify(p)!==before) k++; } }
  S.ihDirty=true; S.ihOpen=true; S.ihSrc=`${k} alan değişiklik kaydından geri getirildi — Kaydet'e bas.`; renderIh(); toast(`${k} alan geri getirildi · Kaydet'e bas`);
}
async function loadIhLog(){
  try{
    const r=await graph(`${WB}/tables('${IL_TABLE}')/range?$select=values`);
    const head=r.values[0].map(h=>String(h).trim()), ix=h=>head.indexOf(h), sv=v=>v==null?"":String(v).replace(/^'/,"");
    S.ihLog=r.values.slice(1).map(c=>({n:sv(c[ix("Contact name")]).trim(), t:sv(c[ix("Telephone")]), alan:sv(c[ix("Alan")]), eski:sv(c[ix("Eski değer")]).trim(), yeni:sv(c[ix("Yeni değer")]).trim()})).filter(x=>x.n&&x.alan);
  }catch(e){ if(e && e.code==="http_404") S.ihLog=[]; else throw e; }
}

// ---------- UI ----------
function ihReadInputs(){
  const p=S.ih; if(!$("ih_alt")) return p;
  p.alt=$("ih_alt").value.trim(); p.m2=ihNum($("ih_m2").value); p.bmin=ihNum($("ih_bmin").value); p.bmax=ihNum($("ih_bmax").value); p.teyit=$("ih_teyit").value.trim();
  return p;
}
function ihChips(key, multi){
  const cur=S.ih[key], list=[...IH_OPT[key]];
  const sel=v=>multi?cur.includes(v):cur===v;
  if(multi) cur.forEach(v=>{ if(!list.includes(v)) list.push(v); }); else if(cur && !list.includes(cur)) list.push(cur);
  return `<div class="chips" data-ih="${key}"${multi?' data-multi="1"':""}>${list.map(v=>`<button type="button" class="chip" aria-pressed="${sel(v)}" data-v="${esc(v)}">${esc(v)}</button>`).join("")}${key==="bolge"?`<button type="button" class="chip add" data-other="1">Diğer…</button>`:""}</div>`;
}
function ihFeatChips(){
  const cls={"Olmazsa olmaz":"must","Olsa iyi olur":"nice","İstemem":"no"};
  const all=IH_FEAT.concat(Object.keys(S.ih.feat).filter(f=>!IH_FEAT.includes(f)));
  return `<div class="chips" id="ihFeat">${all.map(f=>{ const st=S.ih.feat[f]||""; return `<button type="button" class="chip feat ${cls[st]||""}" data-f="${esc(f)}" aria-label="${esc(f)}: ${st||"seçili değil"}">${st==="Olmazsa olmaz"?"✓✓ ":st==="Olsa iyi olur"?"✓ ":st==="İstemem"?"✕ ":""}${esc(f)}</button>`; }).join("")}</div>`;
}
function renderIh(){
  const box=$("ihBox"); if(!box) return;
  const p=S.ih, open=!!S.ihOpen;
  const miss=ihMissingList(p), sc=Math.round(ihScore(p)*100);
  box.innerHTML=`<details class="sec ihsec"${open?" open":""}><summary>🧭 İhtiyaç — ne arıyor? <span class="cnt ihpct ${sc>=80?"ok":sc>=50?"mid":"low"}">%${sc}</span></summary>
    ${(()=>{ const cl=ihClearedFor($("n").value.trim(), $("t").value).filter(r=>!String(ihFlat(p)[r.alan]||"").trim()); return cl.length?`<div class="ihrest">⚠️ Bu kişinin ${cl.length} bilgisi daha önce silinmiş görünüyor (${esc(cl.map(r=>r.alan).join(", "))}). <button type="button" class="act ok" id="ihRestore">↺ Geri getir</button></div>`:""; })()}
    <div class="ihbar"><button type="button" class="act ok" id="ihFill">✨ Nottan doldur</button><span class="hint" id="ihSrcLbl">${S.ihSrc?esc(S.ihSrc):"Görüşme notundaki bilgileri seçeneklere işler; sadece boş alanları doldurur."}</span></div>
    <div class="f"><span class="lab">Rol</span>${ihChips("rol")}</div>
    <div class="f"><span class="lab">İşlem</span>${ihChips("it")}</div>
    <div class="f"><span class="lab">Bölge (birden fazla)</span>${ihChips("bolge",true)}</div>
    <div class="f"><label for="ih_alt">Alternatif bölge</label><input id="ih_alt" value="${esc(p.alt)}" placeholder="Olabilir dediği bölge"></div>
    <div class="f"><span class="lab">Tip</span>${ihChips("tip")}</div>
    <div class="f"><span class="lab">Oda (birden fazla)</span>${ihChips("oda",true)}</div>
    <div class="row">
      <div class="f"><label for="ih_bmin">Min bütçe / fiyat</label><input id="ih_bmin" inputmode="decimal" value="${esc(ihFmt(p.bmin))}" placeholder="12 milyon"></div>
      <div class="f"><label for="ih_bmax">Max bütçe / fiyat</label><input id="ih_bmax" inputmode="decimal" value="${esc(ihFmt(p.bmax))}" placeholder="15 milyon"></div>
    </div>
    <div class="f"><span class="lab">Para birimi</span>${ihChips("pb")}</div>
    <div class="f"><span class="lab">Bütçe esnekliği</span>${ihChips("esn")}</div>
    <div class="f"><label for="ih_m2">Min m²</label><input id="ih_m2" inputmode="numeric" value="${esc(p.m2===""?"":String(p.m2))}"></div>
    <div class="f"><span class="lab">Öz kaynak</span>${ihChips("oz")}</div>
    <div class="f"><span class="lab">Kredi</span>${ihChips("kredi")}</div>
    <div class="f"><span class="lab">Mevcut evini satıp mı alacak?</span>${ihChips("satip")}</div>
    <div class="f"><span class="lab">Ne zaman?</span>${ihChips("zaman")}</div>
    <div class="f"><span class="lab">Tercihler <span class="hint" style="margin:0">— her dokunuş: ✓✓ olmazsa olmaz → ✓ olsa iyi olur → ✕ istemem → boş</span></span>${ihFeatChips()}</div>
    <div class="f"><label for="ih_teyit">Teyit edilecek</label><input id="ih_teyit" value="${esc(p.teyit)}" placeholder="Emin olmadıkların: 'galiba kredi', 'bayramdan sonra'…"></div>
    <p class="hint ihmiss">${miss.length?"Eksik: "+esc(miss.join(", ")):"Profil tam ✓"}</p>
    ${portalHtml(p)}
  </details>`;
}
function ihSummaryHtml(p, upd){
  const parts=[];
  const who=[p.rol,p.it].filter(Boolean).join(" · "); if(who) parts.push(who);
  if(p.bolge.length) parts.push("📍 "+p.bolge.join(", ")+(p.alt?" (alt. "+p.alt+")":""));
  const ev=[p.tip,p.oda.join("/"),p.m2!==""?p.m2+" m²+":""].filter(Boolean).join(" · "); if(ev) parts.push("🏠 "+ev);
  const short=n=>n>=1e6?(Math.round(n/1e5)/10).toLocaleString("tr-TR")+" M":n>=1e3?Math.round(n/1e3).toLocaleString("tr-TR")+" B":String(n);
  if(p.bmin!==""||p.bmax!=="") parts.push("💰 "+(p.bmin!==""&&p.bmax!==""?short(p.bmin)+"–"+short(p.bmax):p.bmax!==""?"≤ "+short(p.bmax):"≥ "+short(p.bmin))+" "+(p.pb||"")+(p.esn&&p.esn!=="Sabit"?" ("+p.esn.toLocaleLowerCase("tr")+")":""));
  const fin=[p.oz,p.kredi?"kredi: "+p.kredi.toLocaleLowerCase("tr"):"",p.satip==="Evet"?"evini satıp alacak":""].filter(Boolean).join(" · "); if(fin) parts.push("🏦 "+fin);
  if(p.zaman) parts.push("⏱ "+p.zaman);
  const f=st=>Object.keys(p.feat).filter(k=>p.feat[k]===st);
  const tags=[...f("Olmazsa olmaz").map(x=>`<span class="ihtag must">✓✓ ${esc(x)}</span>`),...f("Olsa iyi olur").map(x=>`<span class="ihtag nice">✓ ${esc(x)}</span>`),...f("İstemem").map(x=>`<span class="ihtag no">✕ ${esc(x)}</span>`)].join("");
  const miss=ihMissingList(p), sc=Math.round(ihScore(p)*100);
  return `<div class="ihsum" role="button" tabindex="0" title="İhtiyaç bölümünü aç">
    <div class="ihsum-h">🧭 Ne arıyor <span class="cnt ihpct ${sc>=80?"ok":sc>=50?"mid":"low"}">%${sc}</span>${upd?`<span class="sub"> · son güncelleme ${esc(upd)}</span>`:""}<span class="ihedit">Düzenle ›</span></div>
    <div class="ihsum-b">${parts.map(esc).join("<br>")||"—"}</div>${tags?`<div class="ihtags">${tags}</div>`:""}
    ${p.teyit?`<div class="ihsum-w">❓ Teyit: ${esc(p.teyit)}</div>`:""}${miss.length?`<div class="ihsum-w">Eksik: ${esc(miss.join(", "))} — bu görüşmede sor</div>`:""}
  </div>`;
}
function renderIhSum(){
  const box=$("ihSumBox"); if(!box) return;
  const hit=S.ihOrig && S.sel ? ihFind($("n").value.trim(), $("t").value) : null;
  if(!S.ihOrig){ box.innerHTML=""; return; }
  const upd=hit && hit.u ? tr(hit.u) : "";
  box.innerHTML=ihSummaryHtml(S.ihOrig, upd).replace(/<\/div>\s*$/, "")+matchListForContact(S.ihOrig)+portalHtml(S.ihOrig)+"</div>";
}
function ihRefreshHead(){
  const p=ihReadInputs(), sc=Math.round(ihScore(p)*100), miss=ihMissingList(p);
  const pc=document.querySelector("#ihBox .ihpct"); if(pc){ pc.textContent="%"+sc; pc.className="cnt ihpct "+(sc>=80?"ok":sc>=50?"mid":"low"); }
  const m=document.querySelector("#ihBox .ihmiss"); if(m) m.textContent=miss.length?"Eksik: "+miss.join(", "):"Profil tam ✓";
}
function ihLoadFor(name, tel){
  const hit=ihFind(name, tel);
  S.ih=hit?JSON.parse(JSON.stringify(hit.p)):ihEmpty(); S.ihOrig=hit?JSON.parse(JSON.stringify(hit.p)):null; S.ihSrc="";
  if(!S.ih.rol){ const kt=norm(S.kt); if(kt==="alici") S.ih.rol="Alıcı"; else if(kt==="satici") S.ih.rol="Satıcı"; }
  S.ihOpen=ihClearedFor(name, tel).length>0; S.ihDirty=false; renderIh(); renderIhSum();
}
function wireIh(){
  const box=$("ihBox"); if(!box) return;
  const openIh=()=>{ const d=box.querySelector("details"); if(d){ d.open=true; S.ihOpen=true; d.scrollIntoView({behavior:"smooth",block:"start"}); } };
  $("ihSumBox").addEventListener("click",e=>{ if(e.target.closest("a")) return; const pm=e.target.closest(".pmrow[data-pf]"); if(pm){ showView("o"); openPf(pm.dataset.pf); return; } if(e.target.closest(".ihsum")) openIh(); });
  $("ihSumBox").addEventListener("keydown",e=>{ if((e.key==="Enter"||e.key===" ") && e.target.closest(".ihsum")){ e.preventDefault(); openIh(); } });
  box.addEventListener("click",e=>{
    if(e.target.id==="ihFill"){ S.ihDirty=true; ihFillFromNote(); return; }
    if(e.target.id==="ihRestore"){ ihRestore(); return; }
    if(e.target.closest(".chip")) S.ihDirty=true;
    const fb=e.target.closest(".chip.feat");
    if(fb){ ihReadInputs(); const f=fb.dataset.f, cur=S.ih.feat[f]||""; const nx=IH_FST[(IH_FST.indexOf(cur)+1)%IH_FST.length]; if(nx) S.ih.feat[f]=nx; else delete S.ih.feat[f]; renderIh(); return; }
    const b=e.target.closest(".chip"); const g=e.target.closest("[data-ih]"); if(!b||!g) return;
    ihReadInputs();
    const key=g.dataset.ih, multi=!!g.dataset.multi;
    if(b.dataset.other){
      const inp=document.createElement("input"); inp.className="otherin"; inp.placeholder="Bölge yaz, Enter"; b.replaceWith(inp); inp.focus();
      const done=()=>{ const v=inp.value.trim(); if(v && !S.ih[key].includes(v)) S.ih[key].push(v); renderIh(); };
      inp.addEventListener("keydown",ev=>{ if(ev.key==="Enter"){ ev.preventDefault(); done(); } }); inp.addEventListener("blur",done); return;
    }
    const v=b.dataset.v;
    if(multi){ const a=S.ih[key]; const i=a.indexOf(v); if(i>=0) a.splice(i,1); else a.push(v); }
    else S.ih[key]=S.ih[key]===v?"":v;
    renderIh();
  });
  box.addEventListener("input",()=>{ S.ihDirty=true; ihRefreshHead(); });
  box.addEventListener("change",e=>{ if(["ih_bmin","ih_bmax"].includes(e.target.id)){ const n=ihNum(e.target.value); if(n!=="") e.target.value=ihFmt(n); ihRefreshHead(); } });
  box.addEventListener("toggle",e=>{ if(e.target.matches(".ihsec")) S.ihOpen=e.target.open; },true);
}

// ---------- Nottan doldur (kural tabanlı; sadece boş alanları doldurur) ----------
function ihParse(text){
  const out={feat:{}, teyit:[]}; const raw=String(text||""); const n=norm(raw); if(!n) return out;
  const has=re=>re.test(n);
  // bölge
  const bol=IH_OPT.bolge.filter(b=>new RegExp("\\b"+norm(b)+"\\b").test(n)); if(bol.length) out.bolge=bol;
  // işlem / rol
  if(has(/\bkiralik\b|\bkira(da|ya)?\b|kiralamak/)) out.it="Kiralık"; else if(has(/satilik|satin al|almak|alacak|\bev ariyor|daire ariyor|villa ariyor/)) out.it="Satılık";
  if(has(/satmak ist|satacak|satiyor|satisa cikar|evini satmak/) && !has(/satip al/)) out.rol="Satıcı";
  else if(out.it==="Kiralık" && has(/ariyor|kiralamak|tutmak/)) out.rol="Kiracı";
  else if(has(/yatirim/)) out.rol="Yatırımcı";
  else if(has(/ariyor|almak|alacak|bakiyor/)) out.rol="Alıcı";
  // tip
  if(has(/\bvilla\b/)) out.tip="Villa"; else if(has(/mustakil/)) out.tip="Müstakil"; else if(has(/\barsa\b|\barazi\b/)) out.tip="Arsa"; else if(has(/\bofis\b/)) out.tip="Ofis"; else if(has(/dukkan|magaza|ticari/)) out.tip="Ticari"; else if(has(/\bdaire\b|\d\s*\+\s*\d/)) out.tip="Daire";
  // oda (konuşma tanıma "3 artı 1" de yazabilir)
  const oda=new Set(); for(const m of raw.toLocaleLowerCase("tr").matchAll(/(\d)\s*(?:\+|artı|arti)\s*(\d)/g)) oda.add(m[1]+"+"+m[2]); if(oda.size) out.oda=[...oda];
  // m²
  const mm=n.match(/(\d{2,4})\s*(m2|metrekare|metre kare|metre)\b/); if(mm) out.m2=Number(mm[1]);
  // para birimi
  if(has(/dolar|\busd\b/) || /\$/.test(raw)) out.pb="USD"; else if(has(/\beuro\b|\beur\b|avro/) || /€/.test(raw)) out.pb="EUR"; else if(has(/\btl\b|lira|milyon|bin\b/)) out.pb="TL";
  // bütçe: "12-15 milyon", "12 ile 15 milyon", "15 milyona kadar", "300 bin dolar"
  const L=raw.toLocaleLowerCase("tr").replace(/(\d)\.(\d{3})/g,"$1$2");
  const unit=u=>/milyon|\bm\b/.test(u||"")?1e6:/bin/.test(u||"")?1e3:1;
  let mr=L.match(/(\d+(?:[.,]\d+)?)\s*(milyon|bin)?\s*(?:-|–|ile|ila|\/|veya|ya da)\s*(\d+(?:[.,]\d+)?)\s*(milyon|bin|m\b)/);
  if(mr){ const u2=unit(mr[4]), u1=mr[2]?unit(mr[2]):u2; out.bmin=Math.round(parseFloat(mr[1].replace(",","."))*u1); out.bmax=Math.round(parseFloat(mr[3].replace(",","."))*u2); }
  else if((mr=L.match(/(\d+(?:[.,]\d+)?)\s*(milyon|bin)\s*(?:a|e|ya|ye)?\s*(kadar|maksimum|max|üst)/))) out.bmax=Math.round(parseFloat(mr[1].replace(",","."))*unit(mr[2]));
  else if((mr=L.match(/(\d+(?:[.,]\d+)?)\s*(milyon|bin)/))) { const v=Math.round(parseFloat(mr[1].replace(",","."))*unit(mr[2])); out.bmax=v; if(has(/civar/)) out.teyit.push("Bütçe aralığı ('"+mr[0].trim()+" civarı')"); }
  if(has(/guzel olursa|uygun olursa|biraz daha cik|esnek|zorlayabilir/)) out.esn="Uygun ev olursa artar"; else if(has(/kesinlikle ustune|bir kurus fazla|sabit butce/)) out.esn="Sabit";
  // finansman
  if(has(/galiba kredi|belki kredi|kredi dusun|kredi kullanabilir/)){ out.kredi="Henüz bilmiyor"; out.teyit.push("Kredi kullanacak mı"); }
  else if(has(/kredisiz|kredi yok|kredi kullanmay|kredi istemiyor/)) out.kredi="Hayır";
  else if(has(/kredi/)) out.kredi="Evet";
  if(has(/nakit|pesin/)){ out.oz="Tamamı nakit"; if(!out.kredi) out.kredi="Hayır"; }
  if(has(/satip al|evini satip|mevcut evini sat|evini sattiktan|evini satinca/)) out.satip="Evet";
  // zaman
  if(has(/\bhemen\b|\bacil\b|bu ay icinde|bir an once/)) out.zaman="Hemen";
  else if(has(/sadece bakiyor|acelesi yok|piyasaya bakiyor|fikir almak/)) out.zaman="Sadece bakıyor";
  else { const za=n.match(/(\d+)\s*ay/); if(za){ const k=Number(za[1]); out.zaman=k<=3?"1–3 ay":k<=6?"3–6 ay":"6 ay+"; } }
  for(const w of ["bayramdan sonra","bayram sonrasi","yilbasindan sonra","yazin","okullar kapaninca","okullar acilmadan"]) if(n.includes(w)) out.teyit.push("Zaman: '"+w+"'");
  // tercihler: aynı cümle içinde olumsuz / zorunlu ifade
  const clauses=raw.split(/[.;\n!?]|,\s*(?=[^\d])/).map(norm).filter(Boolean);
  for(const f of IH_FEAT){
    for(const c of clauses){
      if(!IH_FKEY[f].some(k=>c.includes(k))) continue;
      if(/istemiyor|olmasin|istemez|olmamali|sevmiyor|haric|\byok\b/.test(c) && !(f==="Kapalı otopark" && /otopark yok mu/.test(c))) out.feat[f]="İstemem";
      else if(/olmazsa olmaz|sart|kesin|mutlaka|kesinlikle|onemli/.test(c)) out.feat[f]="Olmazsa olmaz";
      else if(/onemsiz|fark etmez|farketmez/.test(c)) {}
      else out.feat[f]=out.feat[f]||"Olsa iyi olur";
      break;
    }
  }
  return out;
}
function ihFillFromNote(){
  const txt=$("no").value.trim(); if(!txt){ toast("Önce görüşme notunu yaz ya da sesli anlat"); $("no").focus(); return; }
  ihReadInputs(); const p=S.ih, x=ihParse(txt); let k=0;
  const set=(key)=>{ if(x[key]===undefined) return; const cur=p[key]; const empty=Array.isArray(cur)?!cur.length:(cur===""||cur==null); if(empty){ p[key]=Array.isArray(x[key])?[...x[key]]:x[key]; k++; } };
  ["rol","it","bolge","tip","oda","m2","bmin","bmax","pb","esn","oz","kredi","satip","zaman"].forEach(set);
  for(const [f,st] of Object.entries(x.feat)) if(!p.feat[f]){ p.feat[f]=st; k++; }
  if(x.teyit.length){ const add=x.teyit.filter(t=>!p.teyit.includes(t)); if(add.length){ p.teyit=[p.teyit,...add].filter(Boolean).join("; "); k++; } }
  S.ihSrc=k?`Nottan ${k} alan dolduruldu — kontrol et, gerekirse düzelt.`:"Notta seçeneklere çevrilecek yeni bilgi bulunamadı.";
  S.ihOpen=true; renderIh(); toast(k?`${k} alan nottan dolduruldu`:"Yeni bilgi bulunamadı");
  if(k) S.ihFilled=true;
}

// ---------- kaydet ----------
function ihChanged(){ ihReadInputs(); const a=ihFlat(S.ih), b=S.ihOrig?ihFlat(S.ihOrig):ihFlat(ihEmpty()); return Object.keys(a).some(k=>String(a[k])!==String(b[k])); }
async function ihSave(name, tel, src){
  ihReadInputs();
  if(!ihHasData(S.ih) || !ihChanged()) return false;
  await loadIh(); await ensureTable(IH_SHEET, IH_TABLE, IH_HEAD); if(S.ihMissing) await loadIh();
  { const hit0=ihFind(name, tel); if(hit0) S.ih=ihMerge(hit0.p, S.ihOrig, S.ih); }
  const flat=ihFlat(S.ih), now=serial(todayISO());
  const row=IH_HEAD.map(h=>h==="Contact name"?name:h==="Telephone"?cellIn(tel||""):h==="Güncelleme"?now:h==="Kaynak"?src:h==="Tamlık %"?Math.round(ihScore(S.ih)*100)/100:
    (["Min m²","Min bütçe","Max bütçe"].includes(h)?(flat[h]===""?"":Number(flat[h])):cellIn(flat[h])));
  const hit=ihFind(name, tel);
  if(hit){
    await graph(`${WB}/tables('${IH_TABLE}')/rows/itemAt(index=${hit.i})`,{method:"PATCH",body:{values:[row]}});
    const old=ihFlat(hit.p), ch=Object.keys(flat).filter(k=>String(old[k]||"")!==String(flat[k]||"") && String(old[k]||"").trim());
    if(ch.length){
      try{
        await ensureTable(IL_SHEET, IL_TABLE, IL_HEAD);
        await graph(`${WB}/tables('${IL_TABLE}')/rows/add`,{method:"POST",body:{values:ch.map(k=>{ const f=v=>["Min bütçe","Max bütçe","Min m²"].includes(k)&&v!==""?ihFmt(Number(v)):v; return [now, name, cellIn(tel||""), k, cellIn(f(old[k])), cellIn(f(flat[k])), src]; })}});
      }catch(e){ console.error(e); if(e && e.code==="redirect") throw e; }
    }
  } else {
    await graph(`${WB}/tables('${IH_TABLE}')/rows/add`,{method:"POST",body:{values:[row]}});
  }
  await loadIh(); saveIhCache();
  return true;
}
function saveIhCache(){ try{ localStorage.setItem("gc_ih", JSON.stringify(S.ihRows)); }catch(_){} }
function loadIhCache(){ try{ S.ihRows=JSON.parse(localStorage.getItem("gc_ih")||"[]")||[]; }catch(_){ S.ihRows=[]; } }

// ================= Portföy · Eşleştirme · Portal arama =================
// Portföyler CRM dosyasında "🏠 Portföy" sayfasında (tablo Portfoy). Eşleştirme 25/25/15/15/10/10 (Gültekin, 28.10.2026).
const PF_SHEET="🏠 Portföy", PF_TABLE="Portfoy";
const PF_HEAD=["Portföy ID","Durum","Portföy adı","İşlem türü","Bölge","Mahalle / Site","Tip","Oda","m² brüt","Fiyat","Para birimi","Özellikler","İlan URL","Mal sahibi","Mal sahibi tel","Sözleşme ID","Not","Güncelleme"];
const PF_DURUM=["Aktif","Opsiyonlu","Satıldı/Kiralandı","Pasif"];
const MW={lok:25,but:25,tip:15,oda:15,m2:10,zam:10};
const ZAMAN_P={"Hemen":100,"1–3 ay":80,"3–6 ay":50,"6 ay+":30,"Sadece bakıyor":10};
const ESN_P={"Sabit":0,"%10'a kadar":0.10,"Uygun ev olursa artar":0.15};
const SEMT_ILCE={"Oran":"Çankaya","Çayyolu":"Çankaya","Ümitköy":"Çankaya","Yaşamkent":"Çankaya","Konutkent":"Çankaya","Bilkent":"Çankaya","Beytepe":"Çankaya","Dodurga":"Çankaya","Alacaatlı":"Çankaya","Kavaklıdere":"Çankaya","Gaziosmanpaşa":"Çankaya","Çukurambar":"Çankaya","Balgat":"Çankaya","Dikmen":"Çankaya","İncek":"Gölbaşı","Kızılcaşar":"Gölbaşı","Bağlıca":"Etimesgut"};
const ilceOf=b=>SEMT_ILCE[b]||b;
S.pf=[]; S.pfMissing=false; S.pfSel=null; S.pfEd=null;

function rates(){ let r={}; try{ r=JSON.parse(localStorage.getItem("gc_rates")||"{}")||{}; }catch(_){} return {TL:1, USD:Number(r.USD)||49.77, EUR:Number(r.EUR)||55.87}; }
const toTL=(n,pb)=>n===""||n==null?"":Number(n)*(rates()[pb||"TL"]||1);
const roomN=o=>{ const m=String(o||"").match(/^(\d+)/); return m?Number(m[1]):null; };

function pfFromRow(c,ix,i){
  const g=h=>{ const v=c[ix(h)]; return v==null?"":(typeof v==="string"?v.replace(/^'/,""):v); };
  const u=g("Güncelleme");
  return {i, id:String(g("Portföy ID")), durum:String(g("Durum"))||"Aktif", ad:String(g("Portföy adı")), it:String(g("İşlem türü")), bolge:String(g("Bölge")), mah:String(g("Mahalle / Site")),
    tip:String(g("Tip")), oda:String(g("Oda")), m2:ihNum(g("m² brüt")), fiyat:ihNum(g("Fiyat")), pb:String(g("Para birimi"))||"TL", feat:ihList(g("Özellikler")),
    url:String(g("İlan URL")), sahip:String(g("Mal sahibi")), stel:String(g("Mal sahibi tel")), soz:String(g("Sözleşme ID")), not:String(g("Not")), u:(typeof u==="number"&&u>2000)?fromSerial(u):""};
}
async function loadPf(){
  try{
    const r=await graph(`${WB}/tables('${PF_TABLE}')/range?$select=values`);
    const head=r.values[0].map(h=>String(h).trim()), ix=h=>head.indexOf(h);
    S.pf=r.values.slice(1).map((c,i)=>pfFromRow(c,ix,i)).filter(x=>x.ad||x.id); S.pfMissing=false;
  }catch(e){ if(e && e.code==="http_404"){ S.pf=[]; S.pfMissing=true; } else throw e; }
  try{ localStorage.setItem("gc_pf", JSON.stringify(S.pf)); }catch(_){}
}
function loadPfCache(){ try{ S.pf=JSON.parse(localStorage.getItem("gc_pf")||"[]")||[]; }catch(_){ S.pf=[]; } }
const pfActive=f=>f.durum==="Aktif"||f.durum==="Opsiyonlu"||!f.durum;

// ---------- puanlama ----------
function matchScore(p, f){
  if(p.rol && !["Alıcı","Kiracı","Yatırımcı"].includes(p.rol)) return null;
  if(p.it && f.it && p.it!==f.it) return null;
  if(!p.it && !p.rol) return null;
  const must=Object.keys(p.feat).filter(k=>p.feat[k]==="Olmazsa olmaz"), no=Object.keys(p.feat).filter(k=>p.feat[k]==="İstemem");
  const lacking=must.filter(k=>!f.feat.includes(k)), bad=no.filter(k=>f.feat.includes(k));
  if(lacking.length||bad.length) return {elim:[lacking.length?"olmazsa olmaz yok: "+lacking.join(", "):"", bad.length?"istemediği var: "+bad.join(", "):""].filter(Boolean).join(" · ")};
  const pb=p.bolge.map(norm), fb=norm(f.bolge), fm=norm(f.mah), alt=norm(p.alt);
  const inLoc=b=>b && (b===fb || (fm && (" "+fm+" ").includes(" "+b+" ")));
  let lok;
  if(!pb.length && !alt) lok=50;
  else if(pb.some(inLoc)) lok=100;
  else if(alt && inLoc(alt)) lok=70;
  else if(f.bolge && p.bolge.some(b=>ilceOf(b)===ilceOf(f.bolge))) lok=40;
  else lok=0;
  const price=toTL(f.fiyat,f.pb), max=toTL(p.bmax,p.pb), min=toTL(p.bmin,p.pb);
  let but;
  if(price===""||max==="") but=50;
  else if(price<=max) but=(min===""||price>=min)?100:80;
  else if(price<=max*(1+(ESN_P[p.esn]||0))) but=60;
  else but=0;
  const tip=(!p.tip||!f.tip)?50:(p.tip===f.tip?100:0);
  let oda=50; const fr=roomN(f.oda);
  if(p.oda.length && fr!=null){ if(p.oda.includes(f.oda)) oda=100; else { const d=Math.min(...p.oda.map(roomN).filter(x=>x!=null).map(x=>Math.abs(x-fr))); oda=d===0?100:d===1?50:0; } }
  const m2=(p.m2===""||f.m2==="")?50:(f.m2>=p.m2?100:(f.m2>=0.9*p.m2?60:0));
  const zam=ZAMAN_P[p.zaman]??50;
  const parts={lok,but,tip,oda,m2,zam};
  const total=Object.entries(MW).reduce((s,[k,w])=>s+parts[k]*w,0)/100;
  return {total:Math.round(total), parts};
}
const scoreCls=t=>t>=80?"ok":t>=60?"mid":"low";
function partsTxt(x){ const L={lok:"Bölge",but:"Bütçe",tip:"Tip",oda:"Oda",m2:"m²",zam:"Zaman"}; return Object.entries(x.parts).map(([k,v])=>`${L[k]} ${v}`).join(" · "); }
function buyersFor(f){
  const out=[], el=[];
  for(const r of S.ihRows){ const x=matchScore(r.p,f); if(!x) continue; if(x.elim) el.push({r,why:x.elim}); else out.push({r,...x}); }
  out.sort((a,b)=>b.total-a.total); return {out, el};
}
function portfoliosFor(p){
  return S.pf.filter(pfActive).map(f=>({f, x:matchScore(p,f)})).filter(o=>o.x && !o.x.elim).sort((a,b)=>b.x.total-a.x.total);
}

// ---------- portal arama linkleri ----------
const slug=s=>norm(s).replace(/ /g,"-");
function portalLinks(p){
  const op=p.it==="Kiralık"||p.rol==="Kiracı"?"kiralik":"satilik";
  const T={Daire:"daire",Villa:"villa",Müstakil:"mustakil-ev",Arsa:"arsa",Ticari:"isyeri",Ofis:"isyeri"}[p.tip]||"daire";
  const ilceler=[...new Set((p.bolge.length?p.bolge:[]).map(ilceOf))].slice(0,2);
  if(!ilceler.length) ilceler.push("");
  const tl=n=>n===""?"":Math.round(toTL(n,p.pb));
  const out=[];
  for(const il of ilceler){
    const semtler=p.bolge.filter(b=>ilceOf(b)===il && b!==il);
    const q=new URLSearchParams(); if(p.bmin!=="") q.set("price_min",tl(p.bmin)); if(p.bmax!=="") q.set("price_max",tl(p.bmax)); if(semtler.length===1) q.set("query_text",semtler[0]);
    const loc=il?"/ankara-"+slug(il):"/ankara";
    out.push({site:"sahibinden", il, href:`https://www.sahibinden.com/${op}-${T}${loc}${q.toString()?"?"+q:""}`});
    out.push({site:"HepsiEmlak", il, href:`https://www.hepsiemlak.com/${il?slug(il):"ankara"}-${op}/${T}`});
    out.push({site:"Emlakjet", il, href:`https://www.emlakjet.com/${op}-${T}${loc}`});
  }
  return out;
}
function portalHtml(p){
  if(!ihHasData(p)) return "";
  const ls=portalLinks(p);
  return `<div class="portals"><span class="lab">🔎 Portallarda ara</span><div class="chips">${ls.map(l=>`<a class="chip plink" href="${esc(l.href)}" target="_blank" rel="noopener">${esc(l.site)}${l.il?" · "+esc(l.il):""}</a>`).join("")}</div><p class="hint">sahibinden fiyat aralığıyla açılır; HepsiEmlak ve Emlakjet ilçe listesiyle açılır, oda ve fiyatı sayfada daralt.</p></div>`;
}
function matchListForContact(p){
  if(!S.pf.length) return `<div class="ihsum-w">Uygun portföy için önce Portföy sekmesine ilanlarını gir.</div>`;
  const m=portfoliosFor(p).filter(o=>o.x.total>=50).slice(0,5);
  if(!m.length) return `<div class="ihsum-w">Aktif portföylerinde uygun ilan yok.</div>`;
  return `<div class="pmatch"><span class="lab">🏠 Uygun portföylerimiz</span>${m.map(o=>`<button type="button" class="pmrow" data-pf="${esc(o.f.id)}"><span class="sc ${scoreCls(o.x.total)}">${o.x.total}</span><span><b>${esc(o.f.ad)}</b><span class="sub"> · ${esc([o.f.bolge,o.f.oda,o.f.fiyat!==""?ihFmt(o.f.fiyat)+" "+o.f.pb:""].filter(Boolean).join(" · "))}</span><br><span class="sub">${esc(partsTxt(o.x))}</span></span></button>`).join("")}</div>`;
}

// ---------- Portföy sekmesi ----------
function pfEmpty(){ return {i:-1,id:"",durum:"Aktif",ad:"",it:"Satılık",bolge:"",mah:"",tip:"Daire",oda:"",m2:"",fiyat:"",pb:"TL",feat:[],url:"",sahip:"",stel:"",soz:"",not:""}; }
function renderPfList(){
  const q=norm($("pfQ").value), onlyA=$("pfOnlyA").checked;
  let arr=S.pf.filter(f=>!onlyA||pfActive(f));
  if(q) arr=arr.filter(f=>q.split(" ").every(t=>norm([f.id,f.ad,f.bolge,f.mah,f.tip,f.oda,f.sahip].join(" ")).includes(t)));
  arr.sort((a,b)=>(pfActive(b)-pfActive(a))||String(b.id).localeCompare(String(a.id)));
  $("pfHint").textContent=`${arr.length} portföy${S.ihRows.length?` · ${S.ihRows.length} ihtiyaç profili`:""}`;
  $("pfList").innerHTML=arr.map(f=>{ const b=buyersFor(f).out.filter(o=>o.total>=60).length;
    return `<button type="button" class="person" data-pf="${esc(f.id)}" aria-current="${S.pfSel===f.id}"><div class="nm">${esc(f.ad||f.id)}${b?`<span class="pot pot-Low">${b} uygun alıcı</span>`:""}${!pfActive(f)?`<span class="pot pot-Medium">${esc(f.durum)}</span>`:""}</div><div class="sub">${esc([f.id,f.it,f.bolge,f.mah,f.tip,f.oda,f.m2!==""?f.m2+" m²":"",f.fiyat!==""?ihFmt(f.fiyat)+" "+f.pb:""].filter(Boolean).join("  ·  "))}</div></button>`; }).join("")
    || `<div class="empty">${S.pfMissing||!S.pf.length?"Henüz portföy yok. ＋ Yeni portföy ile ekle ya da Sözleşme DB'den aktar.":"Kayıt yok."}</div>`;
}
function pfChips(key, opts, multi){
  const cur=S.pfEd[key]; const list=[...opts]; if(!multi && cur && !list.includes(cur)) list.push(cur);
  return `<div class="chips" data-pfk="${key}"${multi?' data-multi="1"':""}>${list.map(v=>`<button type="button" class="chip" aria-pressed="${multi?cur.includes(v):cur===v}" data-v="${esc(v)}">${esc(v)}</button>`).join("")}${key==="bolge"?`<button type="button" class="chip add" data-other="1">Diğer…</button>`:""}</div>`;
}
function openPf(id){
  const f=id?S.pf.find(x=>x.id===id):null;
  S.pfSel=f?f.id:null; S.pfEd=f?JSON.parse(JSON.stringify(f)):pfEmpty();
  renderPfForm(); renderPfList();
  if(id) document.body.classList.add("editing");
}
function renderPfForm(){
  const d=S.pfEd, isNew=d.i<0;
  let match="";
  if(!isNew){
    const {out,el}=buyersFor(d);
    const rows=out.slice(0,15).map(o=>{ const c=S.contacts.get(norm(o.r.n)); const wa=waNum(o.r.t||(c&&c.t)||"");
      return `<div class="pmrow"><span class="sc ${scoreCls(o.total)}">${o.total}</span><span class="pmmain" data-k="${esc(c?c.key:"")}" data-n="${esc(o.r.n)}" role="button" tabindex="0"><b>${esc(o.r.n)}</b>${c&&c.p?`<span class="${potCls(c.p)}">${esc(c.p)}</span>`:""}<br><span class="sub">${esc(partsTxt(o))}</span>${o.r.p.teyit?`<br><span class="sub warnc">❓ ${esc(o.r.p.teyit)}</span>`:""}</span>${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:""}</div>`; }).join("");
    match=`<details class="sec" open><summary>🎯 Uygun alıcılar <span class="cnt">${out.length}</span></summary>
      ${rows||`<p class="hint">Bu portföye uyan ihtiyaç profili yok.</p>`}
      ${el.length?`<details class="extra"><summary>Elenenler (${el.length})</summary>${el.map(e=>`<div class="sub" style="padding:4px 0">${esc(e.r.n)} — ${esc(e.why)}</div>`).join("")}</details>`:""}
      <p class="hint">Puan: Bölge 25 · Bütçe 25 · Tip 15 · Oda 15 · m² 10 · Zaman 10. Olmazsa olmazı tutmayan ya da istemediği özelliği olan alıcı elenir.</p></details>`;
  }
  $("pfForm").innerHTML=`<button type="button" class="mobile-back" id="pfBack">‹ Listeye dön</button>
    <h2>${isNew?"Yeni portföy":esc(d.ad||d.id)}</h2><div class="meta">${isNew?"İlanın özelliklerini seç; kaydedince uygun alıcılar listelenir.":esc([d.id,d.u?"güncelleme "+tr(d.u):""].filter(Boolean).join(" · "))}</div>
    ${match}
    <div class="f"><label for="pf_ad">Portföy adı</label><input id="pf_ad" value="${esc(d.ad)}" placeholder="Örn. Park Oran B7 4+1"></div>
    <div class="f"><span class="lab">Durum</span>${pfChips("durum",PF_DURUM)}</div>
    <div class="f"><span class="lab">İşlem</span>${pfChips("it",IH_OPT.it)}</div>
    <div class="f"><span class="lab">Bölge</span>${pfChips("bolge",IH_OPT.bolge)}</div>
    <div class="f"><label for="pf_mah">Mahalle / Site</label><input id="pf_mah" value="${esc(d.mah)}" placeholder="Örn. Park Oran, Ahmet Taner Kışlalı"></div>
    <div class="f"><span class="lab">Tip</span>${pfChips("tip",IH_OPT.tip)}</div>
    <div class="f"><span class="lab">Oda</span>${pfChips("oda",IH_OPT.oda.concat(["4+2","5+2"]))}</div>
    <div class="row"><div class="f"><label for="pf_m2">m² brüt</label><input id="pf_m2" inputmode="numeric" value="${esc(d.m2===""?"":String(d.m2))}"></div>
      <div class="f"><label for="pf_fiyat">Fiyat</label><input id="pf_fiyat" inputmode="decimal" value="${esc(ihFmt(d.fiyat))}" placeholder="14,5 milyon"></div></div>
    <div class="f"><span class="lab">Para birimi</span>${pfChips("pb",IH_OPT.pb)}</div>
    <div class="f"><span class="lab">Özellikler (olanları seç)</span>${pfChips("feat",IH_FEAT,true)}</div>
    <div class="f"><label for="pf_url">İlan URL</label><input id="pf_url" value="${esc(d.url)}"></div>
    <div class="row"><div class="f"><label for="pf_sahip">Mal sahibi</label><input id="pf_sahip" value="${esc(d.sahip)}"></div>
      <div class="f"><label for="pf_stel">Mal sahibi tel</label><input id="pf_stel" inputmode="tel" value="${esc(d.stel)}"></div></div>
    <div class="row"><div class="f"><label for="pf_soz">Sözleşme ID</label><input id="pf_soz" value="${esc(d.soz)}" placeholder="2026-101"></div></div>
    <div class="f"><label for="pf_not">Not</label><textarea id="pf_not">${esc(d.not)}</textarea></div>
    <button class="save" id="pfSave" type="submit">Kaydet</button>`;
}
function pfRead(){ const d=S.pfEd; d.ad=$("pf_ad").value.trim(); d.mah=$("pf_mah").value.trim(); d.m2=ihNum($("pf_m2").value); d.fiyat=ihNum($("pf_fiyat").value); d.url=$("pf_url").value.trim(); d.sahip=$("pf_sahip").value.trim(); d.stel=$("pf_stel").value.trim(); d.soz=$("pf_soz").value.trim(); d.not=$("pf_not").value.trim(); return d; }
function pfRow(d){ return PF_HEAD.map(h=>({"Portföy ID":d.id,"Durum":d.durum,"Portföy adı":cellIn(d.ad),"İşlem türü":d.it,"Bölge":d.bolge,"Mahalle / Site":cellIn(d.mah),"Tip":d.tip,"Oda":cellIn(d.oda||""),"m² brüt":d.m2===""?"":Number(d.m2),"Fiyat":d.fiyat===""?"":Number(d.fiyat),"Para birimi":d.pb,"Özellikler":d.feat.join(", "),"İlan URL":cellIn(d.url),"Mal sahibi":cellIn(d.sahip),"Mal sahibi tel":cellIn(d.stel),"Sözleşme ID":cellIn(d.soz),"Not":cellIn(d.not),"Güncelleme":serial(todayISO())})[h]); }
function pfNextId(){ const n=S.pf.map(f=>Number(String(f.id).replace(/\D/g,""))||0); return "PF-"+String((n.length?Math.max(...n):0)+1).padStart(3,"0"); }
async function savePf(e){
  e.preventDefault(); const d=pfRead();
  if(!d.ad){ toast("Portföy adı gerekli"); $("pf_ad").focus(); return; }
  const btn=$("pfSave"); btn.disabled=true; btn.textContent="Excel'e yazılıyor…";
  try{
    await loadPf(); await ensureTable(PF_SHEET, PF_TABLE, PF_HEAD); if(S.pfMissing) await loadPf();
    const hit=d.id?S.pf.find(x=>x.id===d.id):null;
    if(hit) await graph(`${WB}/tables('${PF_TABLE}')/rows/itemAt(index=${hit.i})`,{method:"PATCH",body:{values:[pfRow(d)]}});
    else { d.id=pfNextId(); await graph(`${WB}/tables('${PF_TABLE}')/rows/add`,{method:"POST",body:{values:[pfRow(d)]}}); }
    await loadPf(); toast("Portföy kaydedildi"); openPf(d.id);
  }catch(err){ console.error(err); if(err && err.code==="redirect") return; banner(errMsg(err,"Portföy kaydedilemedi")); btn.disabled=false; btn.textContent="Kaydet"; }
}
async function importFromSoz(){
  const cand=(S.soz||[]).filter(d=>String(d.AK||"").includes("Devam") && !S.pf.some(f=>f.soz===String(d.A)));
  if(!cand.length){ toast("Aktarılacak 'Devam Ediyor' sözleşme yok"); return; }
  const btn=$("pfImport"); btn.disabled=true; btn.textContent=`${cand.length} kayıt aktarılıyor…`;
  try{
    await loadPf(); await ensureTable(PF_SHEET, PF_TABLE, PF_HEAD); if(S.pfMissing) await loadPf();
    const rows=cand.map((d,k)=>{ const t=String(d.D||""); const f=pfEmpty();
      f.id="PF-"+String((Math.max(0,...S.pf.map(x=>Number(String(x.id).replace(/\D/g,""))||0))+1+k)).padStart(3,"0");
      f.ad=String(d.E||d.A||""); f.it=/KİRA|DEVİR/.test(t)?"Kiralık":"Satılık"; f.tip=/VİLLA/.test(t)?"Villa":/ARSA/.test(t)?"Arsa":/TİCARİ|ENDÜSTRİYEL/.test(t)?"Ticari":"Daire";
      f.mah=String(d.F||""); f.oda=String(d.H||""); f.m2=ihNum(d.G); f.fiyat=ihNum(d.L||d.AC); f.url=String(d.K||""); f.sahip=String(d.N||""); f.stel=String(d.O||""); f.soz=String(d.A||"");
      f.not="Sözleşme DB'den aktarıldı; bölge ve özellikleri tamamla."; return pfRow(f); });
    await graph(`${WB}/tables('${PF_TABLE}')/rows/add`,{method:"POST",body:{values:rows}});
    await loadPf(); renderPfList(); toast(`${rows.length} portföy aktarıldı — bölge ve özellikleri tamamla`);
  }catch(err){ console.error(err); if(err && err.code==="redirect") return; banner(errMsg(err,"Aktarılamadı")); }
  finally{ btn.disabled=false; btn.textContent="⇩ Sözleşme DB'den aktar"; }
}
function wirePf(){
  if(!$("pfForm")) return;
  $("pfQ").addEventListener("input",renderPfList); $("pfOnlyA").addEventListener("change",renderPfList);
  $("pfNew").addEventListener("click",()=>{ openPf(null); document.body.classList.add("editing"); $("pf_ad").focus(); });
  $("pfImport").addEventListener("click",importFromSoz);
  $("pfList").addEventListener("click",e=>{ const b=e.target.closest("[data-pf]"); if(b) openPf(b.dataset.pf); });
  $("pfForm").addEventListener("submit",savePf);
  $("pfForm").addEventListener("click",e=>{
    if(e.target.id==="pfBack"){ document.body.classList.remove("editing"); return; }
    const mm=e.target.closest(".pmmain"); if(mm){ const k=mm.dataset.k; showView("k"); if(k && S.contacts.has(k)) selectContact(k); else { $("q").value=mm.dataset.n; renderResults(); } return; }
    const b=e.target.closest(".chip"), g=e.target.closest("[data-pfk]"); if(!b||!g) return;
    pfRead(); const key=g.dataset.pfk, d=S.pfEd;
    if(b.dataset.other){ const inp=document.createElement("input"); inp.className="otherin"; inp.placeholder="Bölge yaz, Enter"; b.replaceWith(inp); inp.focus();
      const done=()=>{ const v=inp.value.trim(); if(v) d.bolge=v; renderPfForm(); }; inp.addEventListener("keydown",ev=>{ if(ev.key==="Enter"){ ev.preventDefault(); done(); } }); inp.addEventListener("blur",done); return; }
    const v=b.dataset.v;
    if(g.dataset.multi){ const i=d[key].indexOf(v); if(i>=0) d[key].splice(i,1); else d[key].push(v); } else d[key]=d[key]===v&&key!=="durum"&&key!=="it"&&key!=="pb"?"":v;
    renderPfForm();
  });
  $("pfForm").addEventListener("change",e=>{ if(e.target.id==="pf_fiyat"){ const n=ihNum(e.target.value); if(n!=="") e.target.value=ihFmt(n); } });
  $("pfRates").addEventListener("click",()=>{ const r=rates(); const u=prompt("1 USD kaç TL?",r.USD); if(u===null) return; const eu=prompt("1 EUR kaç TL?",r.EUR); if(eu===null) return;
    try{ localStorage.setItem("gc_rates",JSON.stringify({USD:Number(String(u).replace(",","."))||r.USD,EUR:Number(String(eu).replace(",","."))||r.EUR})); }catch(_){} renderRatesLbl(); renderPfList(); if(S.pfEd) renderPfForm(); });
  renderRatesLbl();
}
function renderRatesLbl(){ const r=rates(); $("pfRates").textContent=`Kur: USD ${r.USD.toLocaleString("tr-TR")} · EUR ${r.EUR.toLocaleString("tr-TR")} ✎`; }

// ================= Sesli not (konuşma → yazı + ses kaydı OneDrive'a) =================
// Mikrofona bas: Türkçe konuşma canlı olarak "Görüşme notu"na yazılır, aynı anda ses kaydı tutulur.
// Kaydet'e basınca ses dosyası kişinin Dokümanlar klasörüne "Ses notu <tarih saat>" adıyla yüklenir.
// "Profile çevir" eğitimdeki 4 parçalı komutu notla birlikte panoya kopyalar (Selin'e yapıştır).
const V={on:false, rec:null, sr:null, chunks:[], blob:null, mime:"", t0:0, timer:null, base:"", fin:"", srOk:false, srFail:false, stopWait:null};
const VMAX_MS=10*60*1000;
const SRClass=window.SpeechRecognition||window.webkitSpeechRecognition;
function vPickMime(){
  if(!window.MediaRecorder) return "";
  for(const m of ["audio/mp4","audio/webm;codecs=opus","audio/webm","audio/ogg;codecs=opus"]) if(MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(m)) return m;
  return "";
}
function vExt(m){ return /mp4/.test(m)?".m4a":/ogg/.test(m)?".ogg":".webm"; }
function vClock(ms){ const s=Math.floor(ms/1000); return String(Math.floor(s/60)).padStart(2,"0")+":"+String(s%60).padStart(2,"0"); }
function vUi(){
  const b=$("vBtn"); if(!b) return;
  b.classList.toggle("rec",V.on);
  b.innerHTML=V.on?`⏹ Durdur <span id="vClock">${vClock(Date.now()-V.t0)}</span>`:"🎙 Sesli not";
  b.setAttribute("aria-pressed",V.on?"true":"false");
  $("vKeepWrap").hidden=!V.blob;
  const st=$("vState");
  if(V.on) st.textContent=V.srOk?"Dinliyorum… konuştukça nota yazılıyor.":(V.srFail?"Ses kaydediliyor. Bu tarayıcı canlı yazıya çeviremiyor; notu klavyedeki 🎤 ile dikte edebilirsin.":"Mikrofon açılıyor…");
  else if(V.blob) st.textContent=`Ses kaydı hazır (${vClock(V.dur||0)}, ${fmtSize(V.blob.size)}). Kaydet'e basınca OneDrive'a eklenir.${V.fin?" Notdaki bilgiler İhtiyaç seçeneklerine işlendi, kontrol et.":""}`;
  else st.textContent="";
}
function vWrite(interim){
  const ta=$("no"); const txt=(V.fin+(interim?" "+interim:"")).trim();
  ta.value=(V.base?V.base+(txt?"\n":""):"")+txt;
  ta.scrollTop=ta.scrollHeight;
}
function vStartSR(){
  if(!SRClass){ V.srFail=true; vUi(); return; }
  try{
    const sr=new SRClass(); V.sr=sr;
    sr.lang="tr-TR"; sr.continuous=true; sr.interimResults=true; sr.maxAlternatives=1;
    sr.onstart=()=>{ V.srOk=true; vUi(); };
    sr.onresult=e=>{
      // Oturumdaki tüm sonuçlardan metni baştan kur (iOS aynı cümleyi tekrar tekrar ve büyüyerek gönderebiliyor)
      const fins=[]; let interim="";
      for(let i=0;i<e.results.length;i++){
        const r=e.results[i], t=r[0].transcript.trim(); if(!t) continue;
        if(!r.isFinal){ interim=t; continue; }
        const last=fins[fins.length-1];
        if(last && t.startsWith(last)) fins[fins.length-1]=t; else if(last && last.startsWith(t)) {} else fins.push(t);
      }
      const sess=fins.join(" ");
      if(interim && sess && interim.startsWith(sess)) interim=interim.slice(sess.length).trim();
      V.fin=[V.finPrev,sess].filter(Boolean).join(" ");
      vWrite(interim);
    };
    sr.onerror=e=>{ console.warn("SR",e.error); if(["not-allowed","service-not-allowed","audio-capture","language-not-supported","network"].includes(e.error)){ V.srFail=true; V.srOk=false; V.sr=null; vUi(); } };
    sr.onend=()=>{ V.finPrev=V.fin; if(V.srOk||V.fin) vWrite(""); if(V.on && V.sr===sr){ try{ sr.start(); }catch(_){ } } };   // Chrome sessizlikte durur → devam et
    sr.start();
    setTimeout(()=>{ if(V.on && V.sr===sr && !V.srOk){ V.srFail=true; vUi(); } },4000);   // hiç başlamadıysa kullanıcıyı bilgilendir
  }catch(e){ console.warn(e); V.srFail=true; V.sr=null; vUi(); }
}
async function vStart(){
  if(V.on) return;
  banner("");
  V.base=$("no").value.trim(); V.fin=""; V.finPrev=""; V.chunks=[]; V.blob=null; V.dur=0; V.srOk=false; V.srFail=false;
  V.mime=vPickMime();
  let stream=null;
  if(!SRClass && V.mime && navigator.mediaDevices && navigator.mediaDevices.getUserMedia){
    try{ stream=await navigator.mediaDevices.getUserMedia({audio:{echoCancellation:true,noiseSuppression:true}}); }
    catch(e){ console.warn(e); if(!SRClass){ banner("Mikrofon izni verilmedi. Tarayıcı ayarlarından bu siteye mikrofon izni ver."); return; } }
  }
  if(stream){
    try{
      const rec=new MediaRecorder(stream, V.mime?{mimeType:V.mime, audioBitsPerSecond:32000}:undefined); V.rec=rec;
      rec.ondataavailable=e=>{ if(e.data && e.data.size) V.chunks.push(e.data); };
      rec.onstop=()=>{ stream.getTracks().forEach(t=>t.stop()); V.blob=V.chunks.length?new Blob(V.chunks,{type:(rec.mimeType||V.mime||"audio/webm").split(";")[0]}):null; V.rec=null; vUi(); if(V.stopWait){ V.stopWait(); V.stopWait=null; } };
      rec.start(1000);
    }catch(e){ console.warn(e); stream.getTracks().forEach(t=>t.stop()); V.rec=null; }
  }
  if(!V.rec && !SRClass){ banner("Bu tarayıcı ses kaydını desteklemiyor. Notu klavyedeki 🎤 ile dikte edebilirsin."); return; }
  V.on=true; V.t0=Date.now(); V.used=true;
  vStartSR();
  V.timer=setInterval(()=>{ const c=$("vClock"); if(c) c.textContent=vClock(Date.now()-V.t0); if(Date.now()-V.t0>VMAX_MS){ toast("10 dakika doldu, kayıt durduruldu"); vStop(); } },500);
  vUi();
}
function vStop(){
  if(!V.on) return Promise.resolve();
  V.on=false; V.dur=Date.now()-V.t0; clearInterval(V.timer);
  const sr=V.sr; V.sr=null; if(sr){ try{ sr.stop(); }catch(_){ } }
  if(V.srOk||V.fin) vWrite("");
  const p=new Promise(res=>{ if(V.rec && V.rec.state!=="inactive"){ V.stopWait=res; try{ V.rec.stop(); }catch(_){ res(); } } else res(); });
  vUi();
  return p.then(()=>{ vUi(); if(V.fin) ihFillFromNote(); });
}
function vReset(){
  if(V.on){ V.on=false; clearInterval(V.timer); const sr=V.sr; V.sr=null; if(sr){ try{ sr.abort(); }catch(_){ } } if(V.rec && V.rec.state!=="inactive"){ V.stopWait=null; try{ V.rec.stop(); }catch(_){ } } }
  V.blob=null; V.chunks=[]; V.fin=""; V.base=""; V.dur=0; V.used=false;
  if($("vKeep")) $("vKeep").checked=true;
  vUi();
}
async function vAfterSave(name){
  const blob=V.blob, keep=$("vKeep") && $("vKeep").checked, mime=V.mime;
  V.blob=null; vUi();
  if(!blob || !keep) return;
  try{
    const d=new Date(), hm=String(d.getHours()).padStart(2,"0")+String(d.getMinutes()).padStart(2,"0");
    const folder=await ensureFolder(docPath("kisi",name));
    await uploadFile(folder.id, blob, `Ses notu ${todayISO()} ${hm}${vExt(blob.type||mime)}`);
    toast("Ses notu OneDrive'a eklendi"); renderDocs();
  }catch(e){ console.error(e); if(e && e.code==="redirect") return; banner(errMsg(e,"Ses notu yüklenemedi (not kaydedildi)")); }
}
function wireVoice(){
  if(!$("vBtn")) return;
  if(!SRClass && !(window.MediaRecorder && navigator.mediaDevices)){ $("vBtn").disabled=true; $("vState").textContent="Bu tarayıcı sesli notu desteklemiyor; klavyedeki 🎤 ile dikte edebilirsin."; }
  $("vBtn").addEventListener("click",()=>{ V.on?vStop():vStart(); });
  vUi();
}

function wire(){
  $("todayLbl").textContent=new Date().toLocaleDateString("tr-TR",{weekday:"long",day:"numeric",month:"long"});
  setVal("d",todayISO());
  $("q").addEventListener("input",renderResults);
  $("q").addEventListener("keydown",e=>{ if(e.key==="Enter"){ e.preventDefault(); const first=$("results").querySelector(".person"); if(first) selectContact(first.dataset.k); } });
  $("results").addEventListener("click",e=>{ const b=e.target.closest(".person"); if(!b) return; if(b.dataset.lead){ openLead(b.dataset.lead); return; } selectContact(b.dataset.k); });
  $("newBtn").addEventListener("click",()=>{ const q=$("q").value.trim(); clearForm(/\d{4,}/.test(q)?"":q); if(/\d{4,}/.test(q)){ setVal("t",q); checkDup(); } document.body.classList.add("editing"); rehberSuggest(); $(q && !/\d{4,}/.test(q)?"t":"n").focus(); });
  $("backBtn").addEventListener("click",()=>document.body.classList.remove("editing"));
  $("clearBtn").addEventListener("click",()=>clearForm());
  $("t").addEventListener("input",checkDup);
  $("t").addEventListener("input",()=>{ if(S.sel || S.sdTouched) return; const t=parseTf($("tf").value); const sd=isForeignPhone($("t").value)?["Noel","Thanksgiving"]:["Ramazan Bayramı","Kurban Bayramı","Yılbaşı"]; $("tf").value=formatTf(t.plan, sd); renderPlanChips(); });
  $("n").addEventListener("input",()=>{ clearTimeout(S.sugT); S.sugT=setTimeout(rehberSuggest,200); });
  $("rehSug").addEventListener("click",e=>{ const b=e.target.closest("[data-fill]"); if(b) fillFromLead(b.dataset.fill); });
  $("dupWarn").addEventListener("click",e=>{ const b=e.target.closest("button"); if(!b) return;
    if(b.dataset.fill){ fillFromLead(b.dataset.fill); $("dupWarn").classList.add("hidden"); return; }
    selectContact(b.dataset.k); });
  // Android Chrome: pick straight from the phone's contacts (Contact Picker API); iPhone uses the Rehber export
  if("contacts" in navigator && "ContactsManager" in window){
    $("pickBtn").hidden=false;
    $("pickBtn").addEventListener("click",async()=>{
      try{
        const [c]=await navigator.contacts.select(["name","tel","email"],{multiple:false}); if(!c) return;
        const name=(c.name&&c.name[0])||"", tel=(c.tel&&c.tel[0])||"", em=(c.email&&c.email[0])||"";
        const hit=[...S.contacts.values()].find(x=>x.pk && x.pk===phoneKey(tel)) || S.contacts.get(norm(name));
        if(hit){ selectContact(hit.key); toast("Bu kişi CRM'de kayıtlı"); return; }
        clearForm(name); setVal("t",tel); setVal("em",em); $("fTitle").textContent=name||"Yeni kişi"; $("fMeta").textContent="Telefon rehberinden geldi · görüşme notunu yazıp kaydet.";
        document.body.classList.add("editing"); checkDup(); setTimeout(()=>$("no").focus({preventScroll:true}),50);
      }catch(e){ console.error(e); }
    });
  }
  wireOg(); wireDocs(); wireVoice(); wireIh(); wirePf();
  $("tfChips").addEventListener("click",e=>{ const b=e.target.closest(".chip"); if(!b) return; const t=parseTf($("tf").value); $("tf").value=formatTf(t.plan===b.dataset.v?"":b.dataset.v, t.sd); renderPlanChips(); });
  $("sdChips").addEventListener("click",e=>{ const b=e.target.closest(".chip"); if(!b) return; S.sdTouched=true; const t=parseTf($("tf").value); const v=b.dataset.sd; const sd=t.sd.includes(v)?t.sd.filter(x=>x!==v):SD.map(d=>d.id).filter(x=>x===v||t.sd.includes(x)); $("tf").value=formatTf(t.plan, sd); renderPlanChips(); });
  $("photoFile").addEventListener("change",e=>{ const f=e.target.files[0]; e.target.value=""; photoContact(f); });
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
  $("tabKayit").onclick=()=>show("k"); $("tabBugun").onclick=()=>show("b"); $("tabSoz").onclick=()=>show("s"); $("tabKira").onclick=()=>show("r"); $("tabFirsat").onclick=()=>show("p"); $("tabPf").onclick=()=>show("o");
  const openFromToday=e=>{ if(e.target.closest(".wa")) return; if(e.target.closest("summary")) return; if(e.target.closest(".acts")){ if(e.type==="click") onAct(e); return; } const r=e.target.closest(".call"); if(!r) return;
    if(r.dataset.mod){ show(r.dataset.mod==="soz"?"s":"r"); openMod(r.dataset.mod,r.dataset.id); return; }
    if(r.dataset.lead){ show("k"); openLead(r.dataset.lead); return; }
    show("k"); selectContact(r.dataset.k); };
  for(const mod of ["soz","kira","firsat"]){
    buildModForm(mod);
    $(mod+"Q").addEventListener("input",()=>renderModList(mod));
    $(mod+"New").addEventListener("click",()=>openMod(mod,null));
    $(mod+"List").addEventListener("click",e=>{ const b=e.target.closest(".person"); if(b) openMod(mod,b.dataset.id); });
  }
  $("todayBox").addEventListener("click",e=>{
    if(e.target.id==="qpOpen"){ S.quickPlan=true; S.qp={}; renderQuickPlan(); window.scrollTo(0,0); return; }
    if(e.target.id==="qpBack"){ S.quickPlan=false; renderToday(); return; }
    if(e.target.id==="qpSave"){ saveQuickPlan(); return; }
    const q=e.target.closest(".qpchips .chip"); if(q){ const row=q.closest(".qprow"); S.qp[row.dataset.k]=q.dataset.qp; row.querySelectorAll(".qpchips .chip").forEach(x=>x.setAttribute("aria-pressed",x===q)); return; }
    if(e.target.closest(".qprow")) return;
    openFromToday(e);
  });
  $("todayBox").addEventListener("keydown",e=>{ if(e.key==="Enter") openFromToday(e); });
  buildChips(); renderResults(); renderToday();
}


function syncLbl(){ const at=new Date(); const rb = S.leadsState==="ready" ? ` · Rehber ${(S.leads||[]).length.toLocaleString("tr-TR")}` : S.leadsState==="loading" ? " · Rehber yükleniyor" : S.leadsState==="error" ? " · Rehber okunamadı" : ""; $("syncLbl").textContent=`${S.contacts.size} kişi · ${S.rows.length} kayıt${rb} · Excel ${pad(at.getHours())}:${pad(at.getMinutes())}`; $("syncLbl").title="Sürüm "+APP_V; }
const APP_V = (document.querySelector('script[src*="crm.js"]')||{src:""}).src.split("v=")[1] || "?";
async function refresh(manual){
  $("syncLbl").textContent="Excel okunuyor…";
  try{
    await Promise.all([loadCrm(), loadMod("soz"), loadMod("kira"), loadMod("firsat"), loadOg().catch(e=>{ console.error(e); if(e && e.code==="redirect") throw e; }), loadIh().then(saveIhCache).catch(e=>{ console.error(e); if(e && e.code==="redirect") throw e; }), loadPf().catch(e=>{ console.error(e); if(e && e.code==="redirect") throw e; })]);
    S.modsLoaded=true;
    const busy=document.activeElement && ["INPUT","TEXTAREA","SELECT"].includes(document.activeElement.tagName);
    rebuild(); if(!busy){ renderModList("soz"); renderModList("kira"); renderModList("firsat"); }
    saveCache(); syncLbl(); S.lastRefresh=Date.now();
    if(S.sel && !S.ihDirty){ const c=S.contacts.get(S.sel); if(c) ihLoadFor(c.name, c.t); }
    if(!S.ihDirty && $("viewPf") && !$("viewPf").classList.contains("hidden")) renderPfList();
    if(manual){ S.leadsAt=Date.now(); loadLeads(false); toast("Excel'den yenilendi"); }
  }catch(err){
    console.error(err); if(err && err.code==="redirect") return;
    $("syncLbl").textContent="Excel okunamadı"; banner(errMsg(err,"Excel okunamadı"));
  }
}

// ---------- stale-version guard (iPhone home-screen apps keep old pages) ----------
async function ensureFresh(page, scriptName){
  try{
    const cur=((document.querySelector(`script[src*="${scriptName}"]`)||{}).src||"").split("v=")[1]||"";
    const html=await (await fetch(page+"?check="+Date.now(),{cache:"no-store"})).text();
    const m=html.match(new RegExp(scriptName.replace(".","\\.")+"\\?v=(\\d+)"));
    if(m && cur && m[1]!==cur){
      const k="gx_reload_"+m[1]; if(sessionStorage.getItem(k)) return; sessionStorage.setItem(k,"1");
      location.replace(location.pathname+"?v="+m[1]+location.hash);
    }
  }catch(_){}
}
async function boot(){
  ensureFresh("crm.html","crm.js");
  await msalApp.initialize();
  const rr=await msalApp.handleRedirectPromise().catch(e=>{ console.error(e); return null; });
  S.account=(rr && rr.account) || msalApp.getAllAccounts()[0] || null;
  if(!S.account){ $("signin").hidden=false; $("app").hidden=true; $("signinBtn").onclick=()=>msalApp.loginRedirect({scopes:SCOPES, prompt:"select_account", state:"crm"}); return; }
  msalApp.setActiveAccount(S.account);
  $("signin").hidden=true; $("app").hidden=false;
  wire();
  const had=loadCache(); loadIhCache(); loadPfCache();
  if(had){ rebuild(); renderModList("soz"); renderModList("kira"); renderModList("firsat"); $("syncLbl").textContent="Excel okunuyor…"; }
  const qs=new URLSearchParams(location.search);
  let tab="k"; try{ tab=localStorage.getItem("gc_tab")||"k"; }catch(_){}
  if(qs.get("tab")) tab=qs.get("tab");
  showView(["k","b","s","r","p","o"].includes(tab)?tab:"k");
  S.leadsAt=Date.now(); S.leadsState="loading"; const lp=loadLeads(false);
  await refresh(false);
  lp.then(()=>renderToday());
  const c=qs.get("c"); if(c){ const key=norm(c); if(S.contacts.has(key)){ showView("k"); selectContact(key); } else { showView("k"); $("q").value=c; renderResults(); } history.replaceState(null,"",location.pathname); }
  document.addEventListener("visibilitychange",()=>{ if(document.visibilityState==="visible" && Date.now()-(S.lastRefresh||0)>60000){ const busy=document.activeElement && ["INPUT","TEXTAREA","SELECT"].includes(document.activeElement.tagName); if(!busy) refresh(false); } });
}
boot();
})();
