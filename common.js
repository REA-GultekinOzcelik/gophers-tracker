// ---- Tracker workbook builder (shared by page + node test) ----
const ITEMS = [
  {k:"i01", col:"C", n:"01", l:"ONE Thing", h:"Günün tek en önemli işi bitti", g:"is"},
  {k:"i02", col:"D", n:"02", l:"Lead Block 09–12", h:"3 saatlik lead üretim bloğu korundu", g:"is"},
  {k:"touch", col:"E", n:"03", l:"Touches", h:"Kaç kişiye dokundun", g:"is", t:"num"},
  {k:"appt", col:"F", n:"04", l:"Appointments", h:"Kaç randevu", g:"is", t:"num"},
  {k:"i05", col:"G", n:"05", l:"Follow-up & CRM", h:"Takipler yapıldı, CRM güncel", g:"is"},
  {k:"i06", col:"H", n:"06", l:"USA Window", h:"18:00–19:00 ABD dokunuş penceresi", g:"is"},
  {k:"i16", col:"U", n:"16", l:"Sahibinden Listings", h:"Sahibinden ilanları kontrol edildi", g:"is"},
  {k:"i17", col:"V", n:"17", l:"Endeksa Alerts", h:"Endeksa alarmları incelendi", g:"is"},
  {k:"i07", col:"I", n:"07", l:"Movement", h:"Yürüyüş ya da gym", g:"beden"},
  {k:"i08", col:"J", n:"08", l:"Diet & Water", h:"Kahvaltı, öğle yok, son öğün 17:00, su", g:"beden"},
  {k:"i09", col:"K", n:"09", l:"Sleep", h:"23:30 yatak / 06:00 kalkış", g:"beden"},
  {k:"i10", col:"L", n:"10", l:"Reading", h:"1–1,5 saat okuma", g:"zihin"},
  {k:"i11", col:"M", n:"11", l:"MN License (30 dk)", h:"20:00–20:30 çalışma", g:"zihin"},
  {k:"energy", col:"N", n:"12a", l:"Energy", h:"1–10", g:"zihin", t:"scale"},
  {k:"mood", col:"O", n:"12b", l:"Mood", h:"1–10", g:"zihin", t:"scale"},
  {k:"i13", col:"P", n:"13", l:"MREA Learning", h:"Podcast ya da özet + bir uygulanabilir fikir", g:"zihin"},
  {k:"i14", col:"Q", n:"14", l:"Gratitude & Service", h:"Bir teşekkür, bir hizmet jesti", g:"zihin"},
  {k:"i15", col:"R", n:"15", l:"White Space", h:"Planlı boş alan korundu", g:"zihin"}
];
const YN_KEYS = ITEMS.filter(i => !i.t).map(i => i.k);
const SCORE_MAX = YN_KEYS.length + 3;               // Y/N items + touches + appointments + energy/mood = 17
const colIdx = c => c.charCodeAt(0) - 65;
const CORE_ITEMS = ITEMS.filter(i => i.col <= "R").sort((a,b) => a.col < b.col ? -1 : 1);   // columns C..R
const EXTRA_ITEMS = ITEMS.filter(i => i.col > "T").sort((a,b) => a.col < b.col ? -1 : 1);   // columns U..V
const DAILY_EXTRA_HEADERS = ["16 Sahibinden Listings", "17 Endeksa Alerts"];
const TASK_COLS = ["TaskID","CreatedDate","Title","Bucket","DueDate","Priority","Source","Notes","PlannerSynced","PlannerTaskID","Status","CompletedDate","PlannerUpdate"];

const pad2 = n => String(n).padStart(2, "0");
function isoOf(d){ return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate()); }
function parseISO(s){ const [y,m,d] = s.split("-").map(Number); return new Date(y, m-1, d); }
function addDaysISO(s, n){ const d = parseISO(s); d.setDate(d.getDate()+n); return isoOf(d); }
function serial(iso){ const [y,m,d] = iso.split("-").map(Number); return Math.round((Date.UTC(y,m-1,d) - Date.UTC(1899,11,30)) / 864e5); }
function fromSerial(n){ const d = new Date(Date.UTC(1899,11,30) + Number(n)*864e5); return d.getUTCFullYear()+"-"+pad2(d.getUTCMonth()+1)+"-"+pad2(d.getUTCDate()); }
function mondayOf(iso){ const d = parseISO(iso); const w = (d.getDay()+6)%7; d.setDate(d.getDate()-w); return isoOf(d); }
const DOW_EN = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];

function dayScore(r, touchTarget){
  if(!r) return "";
  let any = false, s = 0;
  for(const i of ITEMS){ const v = r[i.k]; if(v !== undefined && v !== null && v !== "") any = true; }
  if(!any) return "";
  for(const k of YN_KEYS){ if(r[k] === "Y" || r[k] === "N/A") s++; }
  if(Number(r.touch||0) >= touchTarget) s++;
  if(Number(r.appt||0) >= 1) s++;
  if(r.energy !== undefined && r.energy !== "" && r.energy !== null && r.mood !== undefined && r.mood !== "" && r.mood !== null) s++;
  return s;
}

