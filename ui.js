// ---- Gophers Tracker (standalone) ----
const CONFIG = window.GT_CONFIG;
const DRIVE_ID = "b!y07j3zs0wUORORx62WJ844NU9RJFZu9JsXYsPHd-SIEociUqeJtSTq6ZENgXtiJP";
const FILE_ID = "01HMEMOXW2WM52PZ5FIVEKHXHYQA3HN4SC";
const CRM_URL = "crm.html";
const BUCKETS = ["This Week (4-1-1)","MREA Actions","Monthly Cycle","Setup & Open Items"];
const PRIOS = ["Urgent","Important","Medium","Low"];
const SRCS = ["4-1-1","MREA Podcast","Monthly Cycle","Setup","Calendar","Client","Other"];
const DOW_TR = ["Paz","Pzt","Sal","Çar","Per","Cum","Cmt"];
const MON_TR = ["Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"];
const GROUPS = [["is","İş · lead üretimi"],["beden","Beden"],["zihin","Zihin ve denge"]];
const S = {account:null, tasks:new Map(), daily:{}, weekly:{}, rows:{daily:{}, weekly:{}, tasks:{}}, TT:20, ST:12, sel:null, week:null, wweek:null, form:{bucket:BUCKETS[0], prio:"Important", src:"4-1-1"}};
const $ = id => document.getElementById(id);
const today = () => isoOf(new Date());
const trShort = iso => { const d = parseISO(iso); return d.getDate()+" "+MON_TR[d.getMonth()]; };
const hm = d => pad2(d.getHours())+":"+pad2(d.getMinutes());
function esc(s){ return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
function toast(m){ const t = $("toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toast._t); toast._t = setTimeout(()=>t.classList.remove("on"), 2400); }
function banner(m){ const b = $("banner"); b.hidden = !m; b.textContent = m || ""; }
function debounce(fn, ms){ let t; return (...a) => { clearTimeout(t); t = setTimeout(()=>fn(...a), ms); }; }

// ---------- rules ----------
function isLodgeThursday(iso){
  const d = parseISO(iso); if(d.getDay() !== 4) return false;
  const md = iso.slice(5); if(md >= "06-15" && md <= "09-15") return false;
  const n = d.getDate(); return n <= 7 || (n >= 15 && n <= 21);
}
function naDefaults(iso){
  const dow = parseISO(iso).getDay(), out = {};
  if(dow === 6){ out.i02 = "N/A"; }
  if(dow === 2 || dow === 5 || dow === 6) out.i11 = "N/A";
  if(isLodgeThursday(iso)){ out.i06 = "N/A"; out.i10 = "N/A"; }
  return out;
}
const naReason = (k, iso) => {
  const dow = parseISO(iso).getDay();
  if(k === "i02" && dow === 6) return "Cumartesi";
  if(k === "i11" && [2,5,6].includes(dow)) return "Bugün yok";
  if((k === "i06" || k === "i10") && isLodgeThursday(iso)) return "Lodge";
  return "";
};

// ---------- Günlük ----------
function weekDates(mon){ return [0,1,2,3,4,5].map(i => addDaysISO(mon, i)); }
function renderDays(){
  const t = today();
  $("days").innerHTML = weekDates(S.week).map(d => {
    const sc = dayScore(S.daily[d], S.TT), cls = sc === "" ? "" : (sc >= S.ST ? "ok" : "lo");
    return `<button type="button" class="day${d===t?" today":""}" data-d="${d}" aria-current="${d===S.sel}"><div class="dn">${DOW_TR[parseISO(d).getDay()]}</div><div class="dd num">${parseISO(d).getDate()}</div><div class="ds num ${cls}">${sc===""?"–":sc}</div></button>`;
  }).join("");
}
function curDay(){ return S.daily[S.sel] || null; }
function shownVal(k){ const r = curDay(); if(r) return r[k]; return naDefaults(S.sel)[k]; }
function renderScore(){
  const r = curDay(), sc = dayScore(r, S.TT), d = parseISO(S.sel);
  $("scBig").innerHTML = `${sc===""?"–":sc}<small>/15</small>`;
  const pct = sc === "" ? 0 : sc/15*100;
  $("scBar").style.width = pct+"%"; $("scBar").className = sc !== "" && sc >= S.ST ? "ok" : "";
  $("scTgt").style.left = (S.ST/15*100)+"%";
  $("scDay").textContent = DOW_TR[d.getDay()]+" "+trShort(S.sel)+(S.sel===today()?" · bugün":"");
  $("scNote").textContent = sc === "" ? "Henüz giriş yok" : (sc >= S.ST ? "Hedef tuttu" : `Hedef ${S.ST}+ · ${S.ST - sc} puan kaldı`);
}
function itemHTML(it){
  const v = shownVal(it.k), reason = naReason(it.k, S.sel);
  const head = `<div><div class="lbl"><span class="no">${it.n}</span>${esc(it.l)}${reason && v==="N/A" ? `<span class="auto">N/A · ${reason}</span>` : ""}</div><div class="hint">${esc(it.h)}</div></div>`;
  if(it.t === "num"){
    const tgt = it.k === "touch" ? S.TT : 1, n = Number(v||0), has = v !== undefined && v !== "" && v !== null;
    return `<div class="item">${head.replace("</div><div class=\"hint\">", `${has?`<span class="target ${n>=tgt?"ok":"lo"}">hedef ${tgt}${it.k==="appt"?"+":""}</span>`:""}</div><div class="hint">`)}<div class="step" data-k="${it.k}"><button type="button" data-d="-1" aria-label="Azalt">−</button><input class="num" type="number" inputmode="numeric" min="0" id="in_${it.k}" value="${has?n:""}" aria-label="${esc(it.l)}">${it.k==="touch"?`<button type="button" class="plus5" data-d="5">+5</button>`:""}<button type="button" data-d="1" aria-label="Artır">+</button></div></div>`;
  }
  if(it.t === "scale"){
    return `<div class="item">${head}<div class="scale" data-k="${it.k}">${[1,2,3,4,5,6,7,8,9,10].map(n=>`<button type="button" class="num" data-v="${n}" aria-pressed="${Number(v)===n}">${n}</button>`).join("")}</div></div>`;
  }
  return `<div class="item">${head}<div class="seg" data-k="${it.k}" role="group" aria-label="${esc(it.l)}">${["Y","N","N/A"].map(x=>`<button type="button" data-v="${x}" aria-pressed="${v===x}">${x}</button>`).join("")}</div></div>`;
}
function renderChecklist(){
  $("checklist").innerHTML = GROUPS.map(([g,name]) => `<div class="grp"><h3>${name}</h3>${ITEMS.filter(i=>i.g===g).map(itemHTML).join("")}</div>`).join("");
  const r = curDay(); $("notes").value = r && r.notes ? r.notes : "";
}
function renderG(){ renderDays(); renderScore(); renderChecklist(); }

function ensureDay(){ if(!S.daily[S.sel]) S.daily[S.sel] = {...naDefaults(S.sel)}; return S.daily[S.sel]; }
const persistDay = (() => { const timers = {}; return iso => { clearTimeout(timers[iso]); $("savedLbl").textContent = "Kaydediliyor…"; timers[iso] = setTimeout(async () => {
  try{ await writeDay(iso); $("savedLbl").textContent = "Excel'e kaydedildi · "+hm(new Date()); saveCache(); }
  catch(e){ console.error(e); $("savedLbl").textContent = "Kaydedilemedi"; showErr(e, "Giriş Excel'e yazılamadı"); }
}, 600); }; })();
function setVal(k, v){
  const r = ensureDay();
  if(v === "" || v === null || v === undefined) delete r[k]; else r[k] = v;
  persistDay(S.sel); renderDays(); renderScore();
}
function wireChecklist(){
  $("checklist").addEventListener("click", e => {
    const b = e.target.closest("button"); if(!b) return;
    const seg = b.closest(".seg"), sc = b.closest(".scale"), st = b.closest(".step");
    if(seg){ const k = seg.dataset.k, cur = shownVal(k), nv = (curDay() && cur === b.dataset.v) ? "" : b.dataset.v; setVal(k, nv); seg.querySelectorAll("button").forEach(x => x.setAttribute("aria-pressed", x.dataset.v === nv)); refreshAutoTags(); }
    else if(sc){ const k = sc.dataset.k, nv = Number(shownVal(k)) === Number(b.dataset.v) ? "" : Number(b.dataset.v); setVal(k, nv); sc.querySelectorAll("button").forEach(x => x.setAttribute("aria-pressed", Number(x.dataset.v) === nv)); }
    else if(st){ const k = st.dataset.k, inp = $("in_"+k), n = Math.max(0, Number(inp.value||0) + Number(b.dataset.d)); inp.value = n; setVal(k, n); updateTarget(k); }
  });
  $("checklist").addEventListener("input", e => {
    const inp = e.target.closest(".step input"); if(!inp) return;
    const k = inp.closest(".step").dataset.k; setVal(k, inp.value === "" ? "" : Math.max(0, Number(inp.value))); updateTarget(k);
  });
  $("notes").addEventListener("input", () => { const r = ensureDay(); r.notes = $("notes").value; persistDay(S.sel); });
  $("fillY").addEventListener("click", () => {
    const r = ensureDay(); let n = 0;
    for(const k of YN_KEYS){ if(!r[k]){ r[k] = "Y"; n++; } }
    persistDay(S.sel); renderG(); toast(n ? `${n} madde Y yapıldı` : "Boş madde yok");
  });
  $("days").addEventListener("click", e => { const b = e.target.closest(".day"); if(b){ S.sel = b.dataset.d; renderG(); } });
  $("wPrev").onclick = () => { S.week = addDaysISO(S.week, -7); S.sel = S.week; renderG(); };
  $("wNext").onclick = () => { S.week = addDaysISO(S.week, 7); S.sel = S.week; renderG(); };
}
function updateTarget(k){
  const it = ITEMS.find(i => i.k === k), row = $("in_"+k).closest(".item"), lbl = row.querySelector(".lbl");
  let tg = lbl.querySelector(".target"); const tgt = k === "touch" ? S.TT : 1, n = Number($("in_"+k).value || 0);
  if(!tg){ tg = document.createElement("span"); lbl.appendChild(tg); }
  tg.className = "target " + (n >= tgt ? "ok" : "lo"); tg.textContent = "hedef " + tgt + (it.k === "appt" ? "+" : "");
}
function refreshAutoTags(){ document.querySelectorAll("#checklist .seg").forEach(seg => { const k = seg.dataset.k, lbl = seg.closest(".item").querySelector(".lbl"), tag = lbl.querySelector(".auto"), reason = naReason(k, S.sel); if(tag && (!reason || shownVal(k) !== "N/A")) tag.remove(); }); }

// ---------- Görevler ----------
function nextTaskId(){ let mx = 0; for(const id of S.tasks.keys()){ const n = Number((id.match(/\d+/)||[0])[0]); if(n > mx) mx = n; } return "T-" + String(mx+1).padStart(4,"0"); }
const isClosed = t => t.Status === "Done" || t.Status === "Cancelled";
function dueBadge(t){
  if(t.Status === "Cancelled") return `<span class="due soon">İptal</span>`;
  if(t.Status === "Done") return `<span class="due done">${t.CompletedDate ? trShort(t.CompletedDate) : "Bitti"}</span>`;
  if(!t.DueDate) return `<span class="due soon">Tarihsiz</span>`;
  const td = today();
  if(t.DueDate < td) return `<span class="due late">Gecikti · ${trShort(t.DueDate)}</span>`;
  if(t.DueDate === td) return `<span class="due now">Bugün</span>`;
  return `<span class="due soon">${trShort(t.DueDate)}</span>`;
}
function syncTag(t){
  if(t.PlannerSynced !== "Yes") return isClosed(t) ? "" : `<span class="tag wait">Planner'a gidiyor</span>`;
  if(t.PlannerUpdate === "Yes") return `<span class="tag wait">Planner güncelleniyor</span>`;
  return isClosed(t) ? "" : `<span class="tag on">Planner'da</span>`;
}
function editorHTML(t){
  const ch = (list, cur, name) => list.map(v => `<button type="button" class="chip" data-f="${name}" data-v="${esc(v)}" aria-pressed="${cur===v}">${esc(v)}</button>`).join("");
  const closed = isClosed(t);
  return `<div class="editor" data-id="${esc(t.TaskID)}">
    <div class="f"><label for="eTitle">Görev</label><input id="eTitle" value="${esc(t.Title)}"></div>
    <div class="f"><span class="lab">Kolon</span><div class="chips">${ch(BUCKETS, t.Bucket, "Bucket")}</div></div>
    <div class="row2"><div class="f"><label for="eDue">Son tarih</label><input id="eDue" type="date" value="${esc(t.DueDate||"")}"></div>
    <div class="f"><span class="lab">Öncelik</span><div class="chips">${ch(PRIOS, t.Priority, "Priority")}</div></div></div>
    <div class="f"><label for="eNotes">Not</label><input id="eNotes" value="${esc(t.Notes||"")}"></div>
    <div class="eact">
      <button type="button" class="save small" data-a="save">Kaydet</button>
      ${closed ? `<button type="button" class="ghost" data-a="reopen">Yeniden aç</button>` : `<button type="button" class="ghost ok" data-a="done">Tamamla</button><button type="button" class="ghost" data-a="cancel">İptal et</button>`}
      <button type="button" class="ghost" data-a="close">Kapat</button>
    </div>
    <div class="confirm" hidden><span>Görev iptal edilsin mi? Planner'da da kapanır.</span><button type="button" class="yes" data-a="cancelYes">Evet, iptal et</button><button type="button" class="ghost" data-a="cancelNo">Vazgeç</button></div>
  </div>`;
}
function taskRow(t){
  const open = S.editId === t.TaskID;
  return `<div class="task${open?" editing":""}" data-id="${esc(t.TaskID)}" tabindex="0" role="button" aria-expanded="${open}">${dueBadge(t)}<div class="t">${esc(t.Title)}</div><div class="m">${esc(t.Bucket)} · ${esc(t.Priority||"")}${syncTag(t)}${t.Notes?`<br>${esc(t.Notes)}`:""}</div></div>${open?editorHTML(t):""}`;
}
async function saveTask(t, patch, msg){
  try{ await writeTask(t, patch); toast(msg); saveCache(); }
  catch(e){ console.error(e); showErr(e, "Görev kaydedilemedi"); }
  renderT();
}
function wireTaskList(){
  const host = $("viewT");
  host.addEventListener("click", e => {
    const ed = e.target.closest(".editor");
    if(ed){
      const t = S.tasks.get(ed.dataset.id); if(!t) return;
      const chip = e.target.closest(".chip");
      if(chip){ ed.querySelectorAll(`.chip[data-f="${chip.dataset.f}"]`).forEach(x => x.setAttribute("aria-pressed", x === chip)); return; }
      const b = e.target.closest("button[data-a]"); if(!b) return;
      const a = b.dataset.a, pick = f => { const c = ed.querySelector(`.chip[data-f="${f}"][aria-pressed="true"]`); return c ? c.dataset.v : t[f]; };
      if(a === "close"){ S.editId = null; renderT(); return; }
      if(a === "cancel"){ ed.querySelector(".confirm").hidden = false; return; }
      if(a === "cancelNo"){ ed.querySelector(".confirm").hidden = true; return; }
      if(a === "save"){ const title = $("eTitle").value.trim(); if(!title){ $("eTitle").focus(); return; } S.editId = null; saveTask(t, {Title:title, Bucket:pick("Bucket"), Priority:pick("Priority"), DueDate:$("eDue").value||"", Notes:$("eNotes").value.trim()}, "Kaydedildi"); return; }
      if(a === "done"){ S.editId = null; saveTask(t, {Status:"Done", CompletedDate:today()}, "Tamamlandı"); return; }
      if(a === "cancelYes"){ S.editId = null; saveTask(t, {Status:"Cancelled", CompletedDate:today()}, "İptal edildi"); return; }
      if(a === "reopen"){ S.editId = null; saveTask(t, {Status:"Open", CompletedDate:""}, "Yeniden açıldı"); return; }
      return;
    }
    const row = e.target.closest(".task[data-id]"); if(!row) return;
    S.editId = S.editId === row.dataset.id ? null : row.dataset.id; renderT();
    if(S.editId) setTimeout(() => { const el = $("eTitle"); if(el) el.scrollIntoView({block:"nearest"}); }, 30);
  });
  host.addEventListener("keydown", e => { if(e.key === "Enter" && e.target.matches(".task[data-id]")){ e.preventDefault(); e.target.click(); } });
}
function renderT(){
  const all = [...S.tasks.values()].filter(t => !/^TEST/i.test(t.Title||""));
  const open = all.filter(t => !isClosed(t)).sort((a,b) => (a.DueDate||"9").localeCompare(b.DueDate||"9"));
  const done = all.filter(isClosed).sort((a,b) => (b.CompletedDate||"").localeCompare(a.CompletedDate||""));
  $("tList").innerHTML = open.length ? open.map(taskRow).join("") : `<div class="empty">Açık görev yok. Üstten ilk görevi ekle.</div>`;
  $("tDone").innerHTML = done.map(taskRow).join("") || `<div class="empty">Henüz tamamlanan görev yok.</div>`;
  $("tDoneSum").textContent = `Tamamlananlar (${done.length})`;
  const late = open.filter(t => t.DueDate && t.DueDate < today()).length;
  const nb = late + (S.crmDue || 0); $("lateBadge").hidden = !nb; $("lateBadge").textContent = nb;
}
function chipRow(id, list, key){
  $(id).innerHTML = list.map(v => `<button type="button" class="chip" data-v="${esc(v)}" aria-pressed="${S.form[key]===v}">${esc(v)}</button>`).join("");
  $(id).onclick = e => { const b = e.target.closest(".chip"); if(!b) return; S.form[key] = b.dataset.v; chipRow(id, list, key); };
}
function nextFriday(){ const d = new Date(); const add = (5 - d.getDay() + 7) % 7 || 7; d.setDate(d.getDate()+add); return isoOf(d); }
function wireTasks(){
  chipRow("tBucket", BUCKETS, "bucket"); chipRow("tPrio", PRIOS, "prio"); chipRow("tSrc", SRCS, "src");
  const q = [["Bugün",()=>today()],["Yarın",()=>addDaysISO(today(),1)],["Cuma",nextFriday],["1 hafta",()=>addDaysISO(today(),7)]];
  $("tDueQ").innerHTML = q.map(([l],i) => `<button type="button" class="chip" data-i="${i}">${l}</button>`).join("");
  $("tDueQ").onclick = e => { const b = e.target.closest(".chip"); if(!b) return; $("tDue").value = q[Number(b.dataset.i)][1](); $("tDueQ").querySelectorAll(".chip").forEach(x => x.setAttribute("aria-pressed", x === b)); };
  $("tDue").value = addDaysISO(today(), 1);
  $("taskForm").addEventListener("submit", async e => {
    e.preventDefault();
    const title = $("tTitle").value.trim(); if(!title){ $("tTitle").focus(); return; }
    $("tSave").disabled = true;
    try{
      await loadTasks();
      const t = {TaskID:nextTaskId(), CreatedDate:today(), Title:title, Bucket:S.form.bucket, DueDate:$("tDue").value||"", Priority:S.form.prio, Source:S.form.src, Notes:$("tNotes").value.trim(), PlannerSynced:"No", PlannerTaskID:"", Status:"Open", CompletedDate:"", PlannerUpdate:"No"};
      await addTask(t);
      $("tTitle").value = ""; $("tNotes").value = ""; toast(t.TaskID+" eklendi"); renderT(); saveCache();
    }catch(err){ console.error(err); showErr(err, "Görev kaydedilemedi"); }
    finally{ $("tSave").disabled = false; }
  });
  $("crmRefresh").onclick = () => loadCrm(true);
  const openCrm = e => { if(e.target.closest(".wa")) return; const r = e.target.closest("[data-crm]"); if(r) location.href = "crm.html?c=" + encodeURIComponent(r.dataset.crm); };
  $("crmList").addEventListener("click", openCrm);
  $("crmList").addEventListener("keydown", e => { if(e.key === "Enter") openCrm(e); });
  $("tRefresh").onclick = async () => { banner(""); try{ await loadTasks(); renderT(); saveCache(); toast("Planner durumları güncellendi"); }catch(e){ showErr(e, "Görevler okunamadı"); } };
}


// ---------- CRM · Bugün ----------
const CRM_FILE = "01HMEMOXUZW5ZBHFP6IBEKYO5KFYJMBAFM";
const POT = {"VIP":1,"Very High":2,"High":3,"Medium":4,"Low":5};
function crmNorm(s){ return String(s||"").toLocaleLowerCase("tr").normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/ı/g,"i").replace(/[^a-z0-9 ]+/g," ").replace(/\s+/g," ").trim(); }
function crmDate(v){ v = String(v||"").trim(); let m; if(/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0,10); if(/^\d{5}(\.\d+)?$/.test(v)) return fromSerial(Math.floor(Number(v))); if((m = v.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/))) return m[3]+"-"+pad2(m[2])+"-"+pad2(m[1]); return ""; }
function crmItems(head, rows){
  const ix = n => head.indexOf(n);
  const I = {d:ix("Date"), n:ix("Contact name"), t:ix("Telephone"), k:ix("Keep (Y/N)"), p:ix("Potential Level"), nd:ix("Next Action Date"), na:ix("Next Action Type"), rf:ix("Reference")};
  const recs = [];
  rows.forEach((c, i) => { const g = j => j >= 0 ? String(c[j] ?? "").trim() : ""; if(!g(I.n)) return; recs.push({d:crmDate(g(I.d)), n:g(I.n), t:g(I.t), k:g(I.k), p:g(I.p), nd:crmDate(g(I.nd)), na:g(I.na), rf:g(I.rf), seq:i}); });
  const map = new Map();
  recs.sort((a,b) => (a.d||"").localeCompare(b.d||"") || a.seq-b.seq).forEach(r => { const k = crmNorm(r.n); if(!k) return; const c = map.get(k) || {recs:[]}; c.recs.push(r); map.set(k, c); });
  const td = today(), out = {a:[], b:[]};
  for(const c of map.values()){
    const r = c.recs[c.recs.length-1], pick = f => { for(let i = c.recs.length-1; i >= 0; i--) if(c.recs[i][f]) return c.recs[i][f]; return ""; };
    if(!r.nd || r.k === "N" || crmNorm(r.na).includes("olumsuz")) continue;
    const dd = Math.round((parseISO(r.nd) - parseISO(td)) / 864e5);
    const item = {name:r.n, t:pick("t"), p:pick("p"), na:r.na, rf:pick("rf"), dd};
    if(dd < -7) out.b.push(item); else if(dd <= 0) out.a.push(item);
  }
  out.a.sort((x,y) => x.dd - y.dd);
  out.b.sort((x,y) => (POT[x.p]||6) - (POT[y.p]||6) || y.dd - x.dd);
  return out;
}
function waNum(s){ let d = String(s||"").replace(/\D/g,""); if(d.length === 10 && d[0] === "5") d = "90"+d; if(d.length === 11 && d.startsWith("05")) d = "9"+d; return (d.length === 12 && d.startsWith("905")) ? d : ""; }
function crmRow(x){
  const lbl = x.dd === 0 ? ["Bugün","now"] : [`${-x.dd} gün geçti`, "late"], wa = waNum(x.t);
  return `<div class="call" data-crm="${esc(x.name)}" role="link" tabindex="0" style="cursor:pointer"><span class="due ${lbl[1]}">${lbl[0]}</span><div><div class="t">${esc(x.name)}${x.p?` <span class="tag ${POT[x.p]<=2?"wait":"on"}">${esc(x.p)}</span>`:""}</div><div class="m">${esc([x.na, x.rf].filter(Boolean).join(" · ")) || "Aksiyon tipi yok"}</div></div>${wa?`<a class="wa" href="https://wa.me/${wa}" target="_blank" rel="noopener" aria-label="WhatsApp">WA</a>`:"<span></span>"}</div>`;
}
let crmAt = 0;
async function loadCrm(force){
  if(!S.account) return;
  if(!force && Date.now() - crmAt < 10*60*1000) return;
  $("crmSub").textContent = "CRM okunuyor…";
  try{
    const base = `/drives/${DRIVE_ID}/items/${CRM_FILE}/workbook`;
    if(!S.crmTable){
      const tl = await graph(`${base}/tables?$select=name`);
      for(const tb of tl.value){ const h = await graph(`${base}/tables/${encodeURIComponent(tb.name)}/headerRowRange?$select=values`); if(h.values[0].includes("Next Action Date") && h.values[0].includes("Contact name")){ S.crmTable = tb.name; S.crmHead = h.values[0].map(x => String(x).trim()); break; } }
      if(!S.crmTable){ $("crmList").innerHTML = `<div class="empty">CRM dosyasında "Next Action Date" sütunlu tablo bulunamadı.</div>`; $("crmSub").textContent = "Okunamadı"; return; }
    }
    const body = await graph(`${base}/tables/${encodeURIComponent(S.crmTable)}/dataBodyRange?$select=values`);
    const res = crmItems(S.crmHead, body.values);
    crmAt = Date.now(); S.crmDue = res.a.length;
    $("crmSub").textContent = `${res.a.length} kişi bugün ya da son 7 günde · ${res.b.length} birikmiş · ${hm(new Date())}`;
    $("crmList").innerHTML = (res.a.length ? res.a.map(crmRow).join("") : `<div class="empty">Bugün için CRM aksiyonu yok.</div>`)
      + (res.b.length ? `<div class="crmh">Öncelikli birikmiş · ilk 5 / ${res.b.length}</div>` + res.b.slice(0,5).map(crmRow).join("") : "");
    renderT();
  }catch(err){ console.error(err); $("crmSub").textContent = "Okunamadı"; $("crmList").innerHTML = `<div class="empty">CRM okunamadı. "Yenile"ye bas.</div>`; }
}

// ---------- Hafta ----------
function weekStats(mon){
  const ds = weekDates(mon), rows = ds.map(d => S.daily[d]).filter(Boolean);
  const scores = ds.map(d => dayScore(S.daily[d], S.TT));
  const filled = scores.filter(x => x !== "");
  return {ds, scores,
    touches: rows.reduce((s,x)=>s+Number(x.touch||0),0), appts: rows.reduce((s,x)=>s+Number(x.appt||0),0),
    avg: filled.length ? Math.round(filled.reduce((a,b)=>a+b,0)/filled.length*10)/10 : "",
    hit: filled.filter(x=>x>=S.ST).length, diet: rows.filter(x=>x.i08==="Y").length, filled: filled.length};
}
function renderW(){
  const mon = S.wweek, st = weekStats(mon), sat = addDaysISO(mon,5);
  $("wTitle").textContent = `${trShort(mon)} – ${trShort(sat)}`;
  $("wSub").textContent = mon === mondayOf(today()) ? "Bu hafta" : `${st.filled} gün girildi`;
  const k = (v,lab,extra="") => `<div class="kpi"><div class="v num">${v}${extra}</div><div class="k">${lab}</div></div>`;
  $("kpis").innerHTML = k(st.touches,"Temas",`<small>/${S.TT*6}</small>`) + k(st.appts,"Randevu") + k(st.avg===""?"–":st.avg,"Ort. skor",`<small>/15</small>`) + k(st.hit,"Hedefte gün",`<small>/6</small>`) + k(st.diet,"Diyet günü",`<small>/6</small>`);
  const H = 120;
  $("wChart").innerHTML = st.ds.map((d,i) => { const sc = st.scores[i], h = sc===""?4:Math.max(4, sc/15*H); return `<div class="col"><span class="cv num">${sc===""?"":sc}</span><div class="stick ${sc===""?"none":(sc>=S.ST?"ok":"")}" style="height:${h}px"></div><span class="cl">${DOW_TR[parseISO(d).getDay()]}</span></div>`; }).join("")
    + `<div class="tline" style="bottom:${S.ST/15*H + 22}px"></div><div class="tlab" style="bottom:${S.ST/15*H + 22}px">hedef ${S.ST}</div>`;
  const w = S.weekly[mon] || {};
  $("wWeight").value = w.weight ?? ""; $("wList").value = w.listings ?? ""; $("wAction").value = w.action || ""; $("wNotes").value = w.notes || "";
  $("wSaved").textContent = "";
  const weeks = []; for(let x = "2026-10-12"; x <= mondayOf(today()); x = addDaysISO(x,7)) weeks.push(x);
  $("wHist").innerHTML = weeks.reverse().map(m => { const s = weekStats(m), ww = S.weekly[m] || {}; return `<tr><td>${trShort(m)}</td><td class="r num">${s.avg===""?"–":s.avg}</td><td class="r num">${s.touches}</td><td class="r num">${s.appts}</td><td class="r num">${s.diet}/6</td><td class="r num">${ww.weight??"–"}</td></tr>`; }).join("");
}
function wireWeek(){
  const save = debounce(async () => {
    const mon = S.wweek, w = S.weekly[mon] = S.weekly[mon] || {};
    const num = v => v === "" ? undefined : Number(v);
    w.weight = num($("wWeight").value); w.listings = num($("wList").value); w.action = $("wAction").value.trim(); w.notes = $("wNotes").value.trim();
    for(const k of Object.keys(w)) if(w[k] === undefined || w[k] === "") delete w[k];
    try{ await writeWeek(mon); $("wSaved").textContent = "Excel'e kaydedildi · "+hm(new Date()); saveCache(); renderWHistOnly(); }
    catch(e){ console.error(e); $("wSaved").textContent = "Kaydedilemedi"; showErr(e, "Haftalık giriş yazılamadı"); }
  }, 700);
  ["wWeight","wList","wAction","wNotes"].forEach(id => $(id).addEventListener("input", () => { $("wSaved").textContent = "Kaydediliyor…"; save(); }));
  $("w2Prev").onclick = () => { S.wweek = addDaysISO(S.wweek,-7); renderW(); };
  $("w2Next").onclick = () => { S.wweek = addDaysISO(S.wweek,7); renderW(); };
}
function renderWHistOnly(){ const keep = [$("wWeight").value,$("wList").value,$("wAction").value,$("wNotes").value]; const saved = $("wSaved").textContent; renderW(); [$("wWeight").value,$("wList").value,$("wAction").value,$("wNotes").value] = keep; $("wSaved").textContent = saved; }