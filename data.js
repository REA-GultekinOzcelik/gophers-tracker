// ---- Microsoft sign-in + Excel (Graph) data layer ----
const SCOPES = ["User.Read", "Files.ReadWrite"];
const msalApp = new msal.PublicClientApplication({
  auth: { clientId: CONFIG.clientId, authority: "https://login.microsoftonline.com/" + CONFIG.tenantId, redirectUri: location.origin + location.pathname },
  cache: { cacheLocation: "localStorage" }
});
const WB = `/drives/${DRIVE_ID}/items/${FILE_ID}/workbook`;
const ROWS_START = 5;

async function token(){
  try{ return (await msalApp.acquireTokenSilent({scopes:SCOPES, account:S.account})).accessToken; }
  catch(e){ await msalApp.acquireTokenRedirect({scopes:SCOPES, account:S.account}); throw {code:"redirect"}; }
}
async function graph(path, opts = {}){
  for(let attempt = 1;; attempt++){
    const res = await fetch("https://graph.microsoft.com/v1.0" + path, {method: opts.method || "GET", headers: {"Authorization": "Bearer " + await token(), "Content-Type": "application/json"}, body: opts.body ? JSON.stringify(opts.body) : undefined});
    if(res.ok) return res.status === 204 ? {} : res.json();
    if(attempt < 4 && [409, 423, 429, 500, 502, 503, 504].includes(res.status)){ await new Promise(r => setTimeout(r, 1200 * attempt)); continue; }
    let msg = ""; try{ msg = (await res.json()).error.message; }catch(_){}
    throw {code: "http_" + res.status, message: msg};
  }
}
function showErr(err, prefix){
  if(err && err.code === "redirect") return;
  const c = err && err.code;
  const msg = c === "http_401" ? "Oturum süresi doldu. Sayfayı yenile ve Microsoft ile tekrar giriş yap."
    : c === "http_403" ? "Bu Microsoft hesabının Tracker dosyasına erişimi yok. Gophers Realty hesabıyla giriş yaptığından emin ol."
    : c === "http_404" ? "Tracker dosyası OneDrive'da bulunamadı. Dosyanın adı ya da yeri değişmiş olabilir; Selin'e haber ver."
    : c === "http_423" || c === "http_409" ? prefix + ": dosya şu an başka bir yerde kilitli. Birkaç saniye sonra tekrar dene."
    : prefix + ". İnternet bağlantını kontrol edip tekrar dene.";
  banner(msg);
}
const cellDate = v => (typeof v === "number" && v > 20000) ? fromSerial(Math.floor(v)) : (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v) ? v.slice(0,10) : "");
const startRow = addr => Number(addr.split("!")[1].match(/\d+/)[0]);

// ---------- loaders ----------
async function loadDaily(){
  const r = await graph(`${WB}/tables/tblDaily/range?$select=address,values`);
  const s0 = startRow(r.address); S.rows.daily = {}; const daily = {};
  r.values.slice(1).forEach((c, i) => {
    const iso = cellDate(c[0]); if(!iso) return; S.rows.daily[iso] = s0 + 1 + i;
    const o = {}; let any = false;
    ITEMS.forEach((it, j) => { const v = c[2+j]; if(v !== "" && v !== null && v !== undefined){ o[it.k] = it.t ? Number(v) : String(v); any = true; } });
    if(c[19] !== "" && c[19] != null){ o.notes = String(c[19]); any = true; }
    if(any) daily[iso] = o;
  });
  S.daily = daily;
}
async function loadWeekly(){
  const r = await graph(`${WB}/tables/tblWeekly/range?$select=address,values`);
  const s0 = startRow(r.address); S.rows.weekly = {}; const weekly = {};
  r.values.slice(1).forEach((c, i) => {
    const iso = cellDate(c[0]); if(!iso) return; S.rows.weekly[iso] = s0 + 1 + i;
    const o = {}; if(c[1] !== "") o.weight = Number(c[1]); if(c[7] !== "") o.listings = Number(c[7]); if(c[8] !== "") o.action = String(c[8]); if(c[9] !== "") o.notes = String(c[9]);
    if(Object.keys(o).length) weekly[iso] = o;
  });
  S.weekly = weekly;
}
async function loadTasks(){
  const r = await graph(`${WB}/tables/tblTasks/range?$select=address,values`);
  const s0 = startRow(r.address), head = r.values[0].map(String); S.rows.tasks = {}; const tasks = new Map();
  r.values.slice(1).forEach((c, i) => {
    const o = {}; head.forEach((h, j) => o[h] = c[j]);
    if(!/^T-\d+/.test(String(o.TaskID || ""))) return;
    for(const k of ["CreatedDate","DueDate","CompletedDate"]) o[k] = cellDate(o[k]);
    for(const k of Object.keys(o)) if(o[k] === null || o[k] === undefined) o[k] = ""; else if(typeof o[k] !== "string" && !["CreatedDate","DueDate","CompletedDate"].includes(k)) o[k] = String(o[k]);
    tasks.set(o.TaskID, o); S.rows.tasks[o.TaskID] = s0 + 1 + i;
  });
  S.taskHead = head; S.tasks = tasks;
}
async function loadTargets(){
  try{ const r = await graph(`${WB}/worksheets('Lists')/range(address='G4:G5')?$select=values`); if(Number(r.values[0][0])) S.TT = Number(r.values[0][0]); if(Number(r.values[1][0])) S.ST = Number(r.values[1][0]); }catch(_){}
}
async function loadAll(){
  $("syncLbl").textContent = "Excel okunuyor…";
  await Promise.all([loadDaily(), loadWeekly(), loadTasks(), loadTargets()]);
  $("syncLbl").textContent = "Excel güncel · " + hm(new Date());
  saveCache();
}

// ---------- writers ----------
const patchRange = (sheet, addr, body) => graph(`${WB}/worksheets('${sheet}')/range(address='${addr}')`, {method:"PATCH", body});
async function writeDay(iso){
  const v = S.daily[iso] || {};
  let r = S.rows.daily[iso];
  if(!r){ await graph(`${WB}/tables/tblDaily/rows/add`, {method:"POST", body:{values:[[serial(iso)].concat(Array(19).fill(""))]}}); await loadDaily(); S.daily[iso] = v; r = S.rows.daily[iso];
    await patchRange("DailyChecklist", `B${r}`, {formulas:[[`=TEXT(A${r},"ddd")`]]});
    await patchRange("DailyChecklist", `S${r}`, {formulas:[[`=IF(COUNTA(C${r}:R${r})=0,"",COUNTIF(C${r}:D${r},"Y")+COUNTIF(C${r}:D${r},"N/A")+COUNTIF(G${r}:M${r},"Y")+COUNTIF(G${r}:M${r},"N/A")+COUNTIF(P${r}:R${r},"Y")+COUNTIF(P${r}:R${r},"N/A")+IF(N(E${r})>=TouchTarget,1,0)+IF(N(F${r})>=1,1,0)+IF(AND(N${r}<>"",O${r}<>""),1,0))`]]}); }
  await patchRange("DailyChecklist", `C${r}:R${r}`, {values:[ITEMS.map(it => v[it.k] === undefined ? "" : v[it.k])]});
  await patchRange("DailyChecklist", `T${r}`, {values:[[v.notes || ""]]});
}
async function writeWeek(mon){
  const w = S.weekly[mon] || {};
  let r = S.rows.weekly[mon];
  if(!r){ await graph(`${WB}/tables/tblWeekly/rows/add`, {method:"POST", body:{values:[[serial(mon)].concat(Array(9).fill(""))]}}); await loadWeekly(); S.weekly[mon] = w; r = S.rows.weekly[mon];
    const D = "tblDaily[Date]", A = `A${r}`;
    await patchRange("Weekly", `C${r}:G${r}`, {formulas:[[`=SUMIFS(tblDaily[03 Touches ('#)],${D},">="&${A},${D},"<"&${A}+7)`, `=SUMIFS(tblDaily[04 Appointments ('#)],${D},">="&${A},${D},"<"&${A}+7)`, `=IFERROR(ROUND(AVERAGEIFS(tblDaily[Score (/15)],${D},">="&${A},${D},"<"&${A}+7),1),"")`, `=COUNTIFS(${D},">="&${A},${D},"<"&${A}+7,tblDaily[Score (/15)],">="&ScoreTarget)`, `=COUNTIFS(${D},">="&${A},${D},"<"&${A}+7,tblDaily[08 Diet & Water],"Y")`]]}); }
  await patchRange("Weekly", `B${r}`, {values:[[w.weight ?? ""]]});
  await patchRange("Weekly", `H${r}:J${r}`, {values:[[w.listings ?? "", w.action || "", w.notes || ""]]});
}
const TASK_ORDER = ["TaskID","CreatedDate","Title","Bucket","DueDate","Priority","Source","Notes","PlannerSynced","PlannerTaskID","Status","CompletedDate","PlannerUpdate"];
const taskCell = (t, k) => ["CreatedDate","DueDate","CompletedDate"].includes(k) ? (t[k] ? serial(t[k]) : "") : (t[k] ?? "");
async function addTask(t){
  const head = S.taskHead || TASK_ORDER;
  await graph(`${WB}/tables/tblTasks/rows/add`, {method:"POST", body:{values:[head.map(k => taskCell(t, k))]}});
  await loadTasks();
}
async function writeTask(t, patch){
  await loadTasks();                                   // fresh row: never overwrite what Planner flows wrote
  const fresh = S.tasks.get(t.TaskID), r = S.rows.tasks[t.TaskID];
  if(!fresh || !r) throw {code:"http_404"};
  Object.assign(fresh, patch);
  if(fresh.PlannerSynced === "Yes") fresh.PlannerUpdate = "Yes";
  const head = S.taskHead, col = k => String.fromCharCode(65 + head.indexOf(k));
  for(const k of Object.keys(patch).concat(["PlannerUpdate"])){
    if(head.indexOf(k) < 0) continue;
    await patchRange("Tasks", `${col(k)}${r}`, {values:[[taskCell(fresh, k)]]});
  }
  S.tasks.set(fresh.TaskID, fresh);
}

// ---------- local cache (instant start) ----------
function saveCache(){ try{ localStorage.setItem("gt_cache", JSON.stringify({daily:S.daily, weekly:S.weekly, tasks:[...S.tasks.values()], TT:S.TT, ST:S.ST, at:Date.now()})); }catch(_){} }
function loadCache(){ try{ const c = JSON.parse(localStorage.getItem("gt_cache") || "null"); if(!c) return false; S.daily = c.daily || {}; S.weekly = c.weekly || {}; S.tasks = new Map((c.tasks || []).map(t => [t.TaskID, t])); S.TT = c.TT || 20; S.ST = c.ST || 12; return true; }catch(_){ return false; } }

// ---------- boot ----------
function show(v){
  for(const [k, id, tab] of [["g","viewG","tabG"],["t","viewT","tabT"],["w","viewW","tabW"]]){ $(id).hidden = k !== v; $(tab).setAttribute("aria-selected", k === v); }
  if(v === "g") renderG(); if(v === "t"){ renderT(); loadCrm(false); } if(v === "w") renderW();
  try{ localStorage.setItem("gt_tab", v); }catch(_){}
  window.scrollTo(0,0);
}
const curTab = () => !$("viewT").hidden ? "t" : (!$("viewW").hidden ? "w" : "g");
async function refresh(){
  try{ await loadAll(); banner(""); const busy = document.activeElement && document.activeElement !== document.body; if(!busy) show(curTab()); else { renderDays(); renderScore(); renderT(); } loadCrm(false); }
  catch(e){ console.error(e); $("syncLbl").textContent = "Excel okunamadı"; showErr(e, "Excel okunamadı"); }
}
async function boot(){
  const t = today(); S.sel = t; S.week = mondayOf(t); if(parseISO(t).getDay() === 0){ S.sel = addDaysISO(t,-1); S.week = mondayOf(S.sel); }
  S.wweek = S.week;
  $("todayLbl").textContent = new Date().toLocaleDateString("tr-TR", {weekday:"long", day:"numeric", month:"long"});
  $("crmLink").href = CRM_URL;
  await msalApp.initialize();
  const rr = await msalApp.handleRedirectPromise().catch(e => { console.error(e); return null; });
  if(rr && rr.state === "crm"){ location.replace("crm.html"); return; }
  S.account = (rr && rr.account) || msalApp.getAllAccounts()[0] || null;
  if(!S.account){ $("signin").hidden = false; $("app").hidden = true; $("signinBtn").onclick = () => msalApp.loginRedirect({scopes:SCOPES, prompt:"select_account"}); return; }
  msalApp.setActiveAccount(S.account);
  $("signin").hidden = true; $("app").hidden = false;
  $("who").textContent = S.account.username;
  wireChecklist(); wireTasks(); wireTaskList(); wireWeek();
  $("tabG").onclick = () => show("g"); $("tabT").onclick = () => show("t"); $("tabW").onclick = () => show("w");
  $("syncBtn").onclick = () => { banner(""); refresh(); loadCrm(true); };
  $("signoutBtn").onclick = () => msalApp.logoutRedirect({account:S.account});
  loadCache();
  let tab = "g"; try{ tab = localStorage.getItem("gt_tab") || "g"; }catch(_){}
  show(["g","t","w"].includes(tab) ? tab : "g");
  await refresh();
  document.addEventListener("visibilitychange", () => { if(document.visibilityState === "visible" && Date.now() - (S.lastRefresh || 0) > 60000){ S.lastRefresh = Date.now(); refresh(); } });
  S.lastRefresh = Date.now();
}
boot();
