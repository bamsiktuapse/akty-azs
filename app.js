"use strict";
/* Акты АЗС: ввод заявки, справочники, база актов, печать актов ВР/Пломбы/ВБРР/ККТ по разметке листов Excel. */

const APP_VERSION = window.APP_VERSION || "";   // из version.js, его пишет tools/build_www.py по AndroidManifest.xml
const REQ_PREFIX = "SC00";
const DEFAULT_WORKS = [
  "Диагностика неисправности, перезагрузка оборудования",
  "Произведена замена отрезчика на {обор}",
  "Замена ФН на {обор}",
  "Прошивка ККТ {обор} на версию 5.8.100",
  "Замена терминала {обор} на ",
  "Отключение земли от ТОПАЗ",
  "Замена патчкорда",
  "Сброс ошибок возникших в результате обрыва связи",
  "Установка и первичная настройка ",
  "Замена АКБ на ИБП "
];
const DEFAULT_SETTINGS = {
  fio: "Этмекчян Павел Борисович",
  pos: "Ведущий специалист",
  dept: "ИТС КРПУ",
  org: "ООО ИК «СИБИНТЕК»",                        // организация исполнителя — во всех актах
  cust: "АО «НК «Роснефть»-Кубаньнефтепродукт»",  // заказчик
  contracts: {},       // договоры по актам: {vr:{num,date}, pl:…, vbrr:…, kkt:…}; пусто — строка не печатается
  mailTo: "",          // кому отправлять акты всегда, кроме почты АЗС
  carModel: "",                // отчёт по ГСМ (заполняется в «Исполнителе»): автомобиль, норма расхода, шапка маршрутных листов
  carPlate: "",
  gsmRate: "12.3",             // л на 100 км
  gsmOrg: "ООО ИК «СИБИНТЕК» филиал «Макрорегион-Юг»",
  gsmDept: "Сектор ИТ-сервисов Инженерно-технической службы Краснодар Краснодарского РПУ",
  gsmBoss: "",                 // ФИО начальника — руководитель СП в «Отчёте»
  gsmRep: "", gsmProxy: "",    // «Акт ТС»: работодатель в лице (ФИО) и доверенность
  gsmPark: "",                 // место парковки ТС — начало и конец маршрута
  gsmStart: {}, gsmEnd: {},    // остаток ГСМ по месяцам: введённый на начало и рассчитанный на конец
  replyTo: '"DV_Melnikov6@sibsoft.rosneft.ru" <dv_melnikov6@sibsoft.rosneft.ru>, paveletmekchyan@yandex.ru',   // куда АЗС отправляет подписанные акты
  autoSeal: true,
  sealDesc: true,
  sign: "",            // подпись исполнителя: PNG data-URL с прозрачным фоном
  signRatio: 0,        // ширина/высота подписи
  signOn: true         // вставлять подпись в акты
};
/* источники обновлений зашиты в приложение, в настройках их нет */
const DEFAULT_UPD_LINK = "https://drive.google.com/drive/folders/1lujs4ePRRpA4c8k-DMMMvDpZQ_EQpQBL";   // основной: Google Диск «Cloude»
const RESERVE_UPD_LINK = "";   // резервный: публичная ссылка на папку Яндекс.Диска (пока нет)
const KKT_PLACES = ["ККТ 1","ККТ 2","ККТ 3","ККТ 4","ККТ ГКС","ККТ ГКС 2"];
const VBRR_PLACES = ["ВБРР 1","ВБРР 2","ВБРР 3"];

/* ---------- хранилище ---------- */
/* Данные хранятся в localStorage WebView и дублируются в файлы приложения (window.Android.saveData),
   чтобы пережить обновление и сбой хранилища WebView; раз в несколько секунд после изменений
   пишется резервная копия «Загрузки/АЗС/akty-backup.json» — её можно восстановить после переустановки. */
const HAS_NATIVE_DATA = typeof window.Android !== "undefined" && !!window.Android.saveData;
const store = {
  get(k, d){
    let v=null;
    try{ v = localStorage.getItem("akt2."+k); }catch(e){}
    if(v==null && HAS_NATIVE_DATA){ try{ v = Android.loadData(k) || null; if(v!=null) localStorage.setItem("akt2."+k, v); }catch(e){} }
    try{ return v==null ? d : JSON.parse(v); }catch(e){ return d; }
  },
  set(k, v){
    const json=JSON.stringify(v); let ok=true;
    try{ localStorage.setItem("akt2."+k, json); }catch(e){ ok=false; }
    if(HAS_NATIVE_DATA){ try{ Android.saveData(k, json); ok=true; }catch(e){} }
    if(!ok) toast("Не удалось сохранить: память заполнена");
    scheduleBackup();
    return ok;
  }
};
let backupTimer=null;
function scheduleBackup(){
  if(!HAS_NATIVE_DATA || !Android.writeBackup) return;
  clearTimeout(backupTimer);
  backupTimer=setTimeout(()=>{ try{ Android.writeBackup(JSON.stringify({ref,settings,acts,saved:new Date().toISOString()})); }catch(e){} }, 4000);
}

/* ---------- ФИО исполнителя: из именительного падежа получаем «Фамилия И.О.» и родительный падеж ---------- */
const FLEETING = {"Павел":"Павла","Лев":"Льва","Пётр":"Петра","Петр":"Петра"};
function genName(w, male){
  if(male && FLEETING[w]) return FLEETING[w];
  const l=w.slice(-1), hard=/[гкхжшчщ]$/i.test(w.slice(0,-1));
  if(/ия$/.test(w)) return w.slice(0,-1)+"и";
  if(l==="а") return w.slice(0,-1)+(hard?"и":"ы");
  if(l==="я") return w.slice(0,-1)+"и";
  if(male){ if(l==="й"||l==="ь") return w.slice(0,-1)+"я"; if(/[бвгджзклмнпрстфхцчшщ]$/i.test(w)) return w+"а"; return w; }
  return l==="ь" ? w.slice(0,-1)+"и" : w;
}
function genPatronymic(w, male){ return male ? (/ич$/.test(w) ? w+"а" : w) : w.replace(/на$/,"ны"); }
function genSurname(w, male){
  if(male){
    if(/(ый|ой|ий)$/.test(w)) return w.slice(0,-2)+"ого";
    if(/(ых|их|ко|аго|яго|ово|е|и|о|у|ю)$/.test(w)) return w;
    if(/[йь]$/.test(w)) return w.slice(0,-1)+"я";
    if(/а$/.test(w)) return w.slice(0,-1)+(/[гкхжшчщ]а$/.test(w)?"и":"ы");
    if(/я$/.test(w)) return w.slice(0,-1)+"и";
    return w+"а";
  }
  if(/(ова|ева|ёва|ина|ына)$/.test(w)) return w.slice(0,-1)+"ой";
  if(/ая$/.test(w)) return w.slice(0,-2)+"ой";
  if(/яя$/.test(w)) return w.slice(0,-2)+"ей";
  return w;
}
function fioForms(full, declineSurname=true){
  const [sur="", name="", pat=""] = String(full||"").trim().split(/\s+/);
  const male = pat ? !/(на|кызы)$/.test(pat) : !/[ая]$/.test(name);
  const ini = [name, pat].filter(Boolean).map(x=>x[0].toUpperCase()+".").join("");
  return {
    full: [sur,name,pat].filter(Boolean).join(" "),
    short: ini ? sur+" "+ini : sur,
    gen: [sur&&(declineSurname?genSurname(sur,male):sur), name&&genName(name,male), pat&&genPatronymic(pat,male)].filter(Boolean).join(" ")
  };
}
const FIO = () => fioForms(settings.fio, settings.declineSurname!==false);

/* «Mikrotik RB2011 sn:C44F…» → {model:"Mikrotik RB2011", sn:"C44F…"} */
function splitModel(s){
  s = String(s||"").trim();
  const m = s.match(/^(.*?)[\s,]+(?:s\/n|sn|IMEI|IMSI)[:\s]*(\S.*)$/i);
  return m ? {model:m[1].trim(), sn:m[2].trim()} : {model:s, sn:""};
}
function uniq(a){ return [...new Set(a.filter(Boolean))]; }

function defaultRef(){
  const D = window.DEFAULT_DATA;
  const inv = D.inventory.map(([k,l]) => Object.assign(splitModel(k), {azs:l}));
  return {
    azs: D.azs.map(a => Object.assign({km:""}, a)),
    kassy: D.kassy,
    models: uniq(inv.map(i=>i.model)),
    places: D.places,
    seals: D.seals,
    works: DEFAULT_WORKS.slice(),
    equipBase: inv.filter(i=>i.sn).map(i=>({azs:i.azs, model:i.model, sn:i.sn, date:"", req:""})),
    routes: {},    // снимок маршрута по номеру АЗС (data-URL JPEG) — для «Маршрутов» отчёта по ГСМ
    refuels: []    // заправки по топливной карте: {id, date, liters, photo}
  };
}
let ref = store.get("ref", null) || defaultRef();
let settings = Object.assign({}, DEFAULT_SETTINGS, store.get("settings", {}));
delete settings.fioShort; delete settings.fioGen; delete settings.fioOv;   // поля прежних версий
["updSrc","updLink","updSrc2","updLink2","updAuto","updChecked"].forEach(k=>delete settings[k]);   // настройки обновлений прежних версий
let acts = store.get("acts", []);
let act = store.get("act", null) || newAct();

function today(){ const d=new Date(); return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
function pad(n){ return String(n).padStart(2,"0"); }
function newAct(azs){
  const num = azs || (ref.azs[0] ? ref.azs[0].num : "");
  return { id: Date.now(), azs: num, date: today(), req:"", desc:"", tArr:"", tDep:"", km: (azsOf(num).km||""),
    equip:[{place:"",off:"",on:""}], works:[""], mats:[],
    forms:{vr:true, pl:true, vbrr:false, kkt:false} };
}
function saveAct(){ store.set("act", act); }
function saveRef(){ store.set("ref", ref); }
function saveActs(){ store.set("acts", acts); }

/* ---------- утилиты ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s==null?"":s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function toast(msg){ const t=document.createElement("div"); t.className="toast"; t.textContent=msg; document.body.appendChild(t); setTimeout(()=>t.remove(), 2600); }
function fmtDate(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return d+"."+m+"."+y; }
const MONTHS=["января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"];
function fmtDateLong(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return d+" "+MONTHS[+m-1]+" "+y+" г."; }
function azsOf(num){ return (typeof ref!=="undefined" && ref ? ref.azs : []).find(a => String(a.num)===String(num)) || {}; }

function parseSeals(text){
  const out=[];
  String(text||"").split(/[\s,;]+/).forEach(p=>{
    p=p.replace(/\.$/,""); if(!p) return;
    const m=p.match(/^(\d+)-(\d+)$/);
    if(m && +m[2]>=+m[1] && +m[2]-+m[1]<500){ for(let i=+m[1];i<=+m[2];i++) out.push(String(i)); }
    else out.push(p);
  });
  return out;
}
function joinSeals(text){ const a=parseSeals(text); return a.length ? a.join(", ")+"." : ""; }

/* Касса по месту: «ККТ 2» / «ВБРР 3» → строка листа Кассы */
function kassaFor(place){
  const m = String(place||"").match(/^(ККТ|ВБРР)\s+(\d)$/i);
  if(!m) return null;
  return ref.kassy.find(k => String(k.azs)===String(act.azs) && String(k.n)===m[2]) || null;
}
function descr(place){
  const k = kassaFor(place); if(!k) return "";
  if(/^ККТ/i.test(place)) return k.zn ? "Атол FPrint 22 ПТК, ЗН:"+k.zn : "";
  return [k.model, k.tid && "ID "+k.tid, k.sn && "s/n:"+k.sn].filter(Boolean).join(" ");
}
function matText(m){ return m.name.trim() + (m.sn && m.sn.trim() ? " s/n:"+m.sn.trim() : ""); }
function matQty(m){ return [m.qty, m.unit].filter(x=>x!=="" && x!=null).join(" "); }

/* ---------- модель актов ---------- */
function model(){
  const a = azsOf(act.azs);
  const eq = act.equip.filter(e=>e.place||e.off||e.on).map(e=>({place:e.place, off:joinSeals(e.off), on:joinSeals(e.on)}));
  return { num: act.azs, type: a.type||"АЗС", addr: a.addr||"", date: act.date, req: act.req, desc: act.desc,
    tArr: act.tArr, tDep: act.tDep, km: act.km,
    works: act.works.map(s=>s.trim()).filter(Boolean),
    mats: act.mats.filter(m=>m.name.trim()),
    eq, sealsBy: p => eq.find(e=>e.place===p) || {off:"",on:""},
    obj: (a.type||"АЗС")+" № "+act.azs };
}

/* ================= Ввод ================= */
function fillAzsSelect(sel, value){
  sel.innerHTML = ref.azs.map(a=>`<option value="${esc(a.num)}">${esc(a.num)}</option>`).join("");
  if(value && !ref.azs.some(a=>String(a.num)===String(value))) sel.insertAdjacentHTML("afterbegin",`<option value="${esc(value)}">${esc(value)}</option>`);
  sel.value = value;
}
function renderAzsInfo(){
  const a=azsOf(act.azs);
  const who=[a.boss, a.phone, a.bossPhone && "упр. "+a.bossPhone, a.email && "✉ "+a.email].filter(Boolean).join(" · ");
  $("#azsInfo").textContent = a.num ? [a.type, a.addr].filter(Boolean).join(", ")+(who ? "\n"+who : "")
    : ref.azs.length ? "" : "Справочник АЗС пуст: загрузите «Справочник АЗС.xlsx» в «Данные → Справочник Excel» или восстановите резервную копию с Android в «Данные → Резервная копия».";
}
function reqSuffix(r){ r=String(r||""); return r.toUpperCase().startsWith(REQ_PREFIX) ? r.slice(REQ_PREFIX.length) : r; }
function renderAct(){
  fillAzsSelect($("#f-azs"), act.azs); renderAzsInfo();
  $("#f-date").value=act.date; $("#f-req").value=reqSuffix(act.req); $("#f-desc").value=act.desc;
  $("#f-tArr").value=act.tArr; $("#f-tDep").value=act.tDep; $("#f-km").value=act.km;
  renderEquip(); renderWorks(); renderMats(); updateTitle();
}
function updateTitle(){ $("#title").textContent = (act.req ? act.azs+" · "+act.req : "Акты АЗС")+(APP_VERSION?"  v"+APP_VERSION:""); }

function sortSeals(a){
  const k=settings.sealDesc?-1:1;
  return a.slice().sort((x,y)=>{ const nx=+x, ny=+y; return k*((isNaN(nx)||isNaN(ny)) ? String(x).localeCompare(String(y),"ru",{numeric:true}) : nx-ny); });
}
function usedSeals(){ const s=new Set(); act.equip.forEach(e=>parseSeals(e.on).forEach(x=>s.add(x))); return s; }
function renderEquip(){
  const used = usedSeals();
  const free = sortSeals(ref.seals.filter(s=>!used.has(s))).slice(0,12);
  $("#eqList").innerHTML = act.equip.map((e,i)=>{
    const d = descr(e.place);
    return `<div class="item" data-i="${i}">
      <div class="top">
        <select data-k="place"><option value="">— место установки —</option>${
          ref.places.concat(e.place && !ref.places.includes(e.place)?[e.place]:[]).map(p=>`<option ${p===e.place?"selected":""}>${esc(p)}</option>`).join("")}</select>
        <button class="btn icon del" data-del="eq" title="Удалить">✕</button>
      </div>
      ${d?`<div class="hint">${esc(d)}</div>`:""}
      <label>Сняты пломбы</label>
      <div class="top"><textarea data-k="off" rows="2" inputmode="numeric" placeholder="Списком: по одной в строке, через запятую или пробел; диапазон 110300-110305">${esc(e.off)}</textarea><button class="btn icon cam" data-cam="off" title="Распознать камерой">📷</button></div>
      ${parseSeals(e.off).length>1?`<div class="hint">Снято пломб: ${parseSeals(e.off).length}</div>`:""}
      <label>Установлены пломбы</label>
      <input data-k="on" value="${esc(e.on)}" inputmode="numeric">
      ${free.length?`<div class="chips">${free.map(s=>`<button class="chip" data-seal="${esc(s)}">${esc(s)}</button>`).join("")}</div>`:""}
    </div>`;
  }).join("") || `<div class="muted">Нет строк</div>`;
  document.querySelectorAll("#eqList textarea").forEach(autoGrow);
}
function renderWorks(){
  if(!act.works.length) act.works=[""];
  $("#workList").innerHTML = act.works.map((w,i)=>`<div class="top" style="margin:6px 0" data-i="${i}">
      <span class="num">${i+1}.</span>
      <div class="combo"><textarea data-w="${i}" rows="1" placeholder="Наименование работы">${esc(w)}</textarea><button class="arrow" data-phrase="${i}" title="Выбрать из справочника">▾</button></div>
      <button class="btn icon" data-scanw="${i}" title="Сканировать штрихкод / QR">${ICON_SCAN}</button>
      ${act.works.length>1?`<button class="btn icon del" data-delw="${i}">✕</button>`:""}
    </div>`).join("");
  document.querySelectorAll("#workList textarea").forEach(autoGrow);
}
function autoGrow(t){ t.style.height="auto"; t.style.height=Math.max(42,t.scrollHeight+2)+"px"; }
function renderMats(){
  $("#matList").innerHTML = act.mats.map((m,i)=>`<div class="item" data-i="${i}">
    <div class="top"><div class="combo"><input data-m="name" value="${esc(m.name)}" placeholder="Модель / материал"><button class="arrow" data-pick="${i}">▾</button></div>
      <button class="btn icon del" data-del="mat">✕</button></div>
    ${m.sn!=null ? `<label>Серийный номер</label><div class="top"><input data-m="sn" value="${esc(m.sn)}" autocapitalize="characters"><button class="btn icon cam" data-camsn="${i}" title="Штрихкод, QR или текст с фото">${ICON_SCAN}</button></div>`
                 : `<button class="btn sm" data-addsn="${i}" style="margin-top:6px">+ серийный номер</button>`}
    <div class="row">
      <div><label>Кол-во</label><input data-m="qty" value="${esc(m.qty)}" inputmode="decimal"></div>
      <div><label>Ед. изм.</label><input data-m="unit" value="${esc(m.unit)}" list="units"></div>
    </div></div>`).join("") + `<datalist id="units"><option>шт</option><option>м</option><option>компл</option></datalist>`;
}

const ICON_SCAN=`<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 7V4h3M18 4h3v3M21 17v3h-3M6 20H3v-3"/><path d="M7 8v8M10 8v8M13 8v8M16 8v8"/></svg>`;
function bindAct(){
  const map={"f-date":"date","f-desc":"desc","f-tArr":"tArr","f-tDep":"tDep"};
  Object.entries(map).forEach(([id,k])=>$("#"+id).addEventListener("input",e=>{act[k]=e.target.value; saveAct();}));
  $("#scanDesc").onclick=()=>scan("text", v=>{
    const t=v[0]; act.desc=[act.desc.trim(), t].filter(Boolean).join("\n"); $("#f-desc").value=act.desc;
    const m=t.match(/SC0*\d{6,}/i);           // номер заявки в коде — сразу в поле «Заявка №»
    if(m && !act.req){ act.req=m[0].toUpperCase().replace(/^SC0*/,"SC00"); $("#f-req").value=reqSuffix(act.req); updateTitle(); }
    saveAct(); });
  $("#f-req").addEventListener("input",e=>{
    let v=e.target.value.trim(); if(v.toUpperCase().startsWith(REQ_PREFIX)){ v=v.slice(REQ_PREFIX.length); e.target.value=v; }
    act.req = v ? REQ_PREFIX+v : ""; saveAct(); updateTitle();
  });
  $("#f-km").addEventListener("change",e=>{
    act.km=e.target.value; saveAct();
    const a=ref.azs.find(x=>String(x.num)===String(act.azs));
    if(a && !String(a.km||"").trim() && act.km){ a.km=act.km; saveRef(); toast("Расстояние сохранено в справочник АЗС"); }
  });
  $("#f-azs").addEventListener("change",e=>{ act.azs=e.target.value; act.km=azsOf(act.azs).km||""; $("#f-km").value=act.km; saveAct(); renderAzsInfo(); renderEquip(); updateTitle(); });

  $("#addEq").onclick=()=>{ act.equip.push({place:"",off:"",on:""}); saveAct(); renderEquip(); };
  $("#addWork").onclick=()=>{ act.works.push(""); saveAct(); renderWorks(); const t=document.querySelectorAll("#workList textarea"); t[t.length-1].focus(); };
  $("#addMat").onclick=()=>{ act.mats.push({name:"",qty:"1",unit:"шт"}); saveAct(); renderMats(); };

  $("#eqList").addEventListener("input",e=>{
    const it=e.target.closest(".item"); const k=e.target.dataset.k; if(!it||!k) return;
    act.equip[+it.dataset.i][k]=e.target.value; saveAct();
    if(e.target.tagName==="TEXTAREA") autoGrow(e.target);
  });
  $("#eqList").addEventListener("change",e=>{ const k=e.target.dataset.k; if(k==="place"){ autoForms(); renderEquip(); } else if(k==="on"||k==="off") renderEquip(); });
  $("#eqList").addEventListener("click",e=>{
    const it=e.target.closest(".item"); if(!it) return; const i=+it.dataset.i;
    if(e.target.dataset.del){ act.equip.splice(i,1); saveAct(); renderEquip(); }
    else if(e.target.dataset.seal){ const cur=parseSeals(act.equip[i].on); cur.push(e.target.dataset.seal); act.equip[i].on=cur.join(", "); saveAct(); renderEquip(); }
    else if(e.target.dataset.cam){ scan("digits", found=>{ const cur=parseSeals(act.equip[i].off); act.equip[i].off=uniq(cur.concat(found)).join(", "); saveAct(); renderEquip(); }); }
  });
  $("#workList").addEventListener("input",e=>{ if(e.target.dataset.w!=null){ act.works[+e.target.dataset.w]=e.target.value; autoGrow(e.target); saveAct(); }});
  $("#workList").addEventListener("click",e=>{
    const b=e.target.closest("button"); if(!b) return;
    if(b.dataset.delw!=null){ act.works.splice(+b.dataset.delw,1); saveAct(); renderWorks(); }
    if(b.dataset.phrase!=null) pickWork(+b.dataset.phrase);
    if(b.dataset.scanw!=null){ const i=+b.dataset.scanw; scan("text", v=>{ act.works[i]=[act.works[i].trim(), v[0]].filter(Boolean).join(" "); saveAct(); renderWorks(); }); }
  });
  $("#matList").addEventListener("input",e=>{
    const it=e.target.closest(".item"); const k=e.target.dataset.m; if(!it||!k) return;
    act.mats[+it.dataset.i][k]=e.target.value; saveAct();
  });
  $("#matList").addEventListener("click",e=>{
    const b=e.target.closest("button"); const it=e.target.closest(".item"); if(!b||!it) return; const i=+it.dataset.i;
    if(b.dataset.del){ act.mats.splice(i,1); saveAct(); renderMats(); }
    if(b.dataset.pick!=null) picker("Модель / материал", ref.models, v=>{ act.mats[i].name=v; saveAct(); renderMats(); });
    if(b.dataset.addsn!=null){ act.mats[i].sn=""; saveAct(); renderMats(); document.querySelectorAll('#matList .item')[i].querySelector('[data-m="sn"]').focus(); }
    if(b.dataset.camsn!=null) scan("serial", found=>{ act.mats[i].sn=found[0]||""; saveAct(); renderMats(); });
  });
  $("#btnSave").onclick=()=>{ if(saveToDb()!==false) toast("Акт сохранён в базу"); go("print"); };
}
/* Выбор актов по составу оборудования */
function autoForms(){
  const places=act.equip.map(e=>e.place);
  if(places.some(p=>/^ККТ/i.test(p))) act.forms.kkt=true;
  if(places.some(p=>/^ВБРР/i.test(p))) act.forms.vbrr=true;
  saveAct();
}

/* ---------- нижняя шторка выбора ---------- */
function sheet(title, bodyHtml){
  const bg=document.createElement("div"); bg.className="sheet-bg";
  bg.innerHTML=`<div class="sheet"><div class="hd"><b>${esc(title)}</b><button class="btn sm" data-x>Закрыть</button></div>${bodyHtml}</div>`;
  document.body.appendChild(bg);
  bg.addEventListener("click",e=>{ if(e.target===bg||e.target.dataset.x!=null) bg.remove(); });
  return bg;
}
function picker(title, items, cb, extraHtml){
  const bg=sheet(title, `<div class="pad">${extraHtml||""}<input placeholder="Поиск…" data-q></div><div class="bd"></div>`);
  const bd=bg.querySelector(".bd"), q=bg.querySelector("[data-q]");
  const text=t=>typeof t==="function"?t():t;
  const draw=()=>{ const f=q.value.toLowerCase(); bd.innerHTML=items.map((t,i)=>({t:text(t),i})).filter(o=>o.t.toLowerCase().includes(f)).map(o=>`<button class="opt" data-i="${o.i}">${esc(o.t)}</button>`).join("")||`<div class="muted pad">Ничего не найдено</div>`; };
  draw(); q.oninput=draw;
  bg.addEventListener("change",draw);
  bg.addEventListener("click",e=>{ const o=e.target.closest(".opt"); if(o){ cb(text(items[+o.dataset.i])); bg.remove(); } });
  return bg;
}
function pickWork(i){
  const eqs=uniq(act.equip.map(e=>e.place));
  let bg;
  const target=()=>{ const s=bg && bg.querySelector("[data-eq]"); const p=s? s.value : (eqs[0]||""); if(!p) return "";
    const d=descr(p); return d ? p+" ("+d+")" : p; };
  const items=ref.works.map(ph=>()=>ph.replace(/\{обор\}/g, target()));
  bg=picker("Работа "+(i+1), items, v=>{ act.works[i]=v; saveAct(); renderWorks(); },
    eqs.length?`<label style="margin-top:0">{обор} =</label><select data-eq style="margin-bottom:8px">${eqs.map(p=>`<option>${esc(p)}</option>`).join("")}</select>`:"");
}

/* ================= Распознавание камерой (Tesseract, офлайн) ================= */
let ocrWorker=null;
function loadScript(src){ return new Promise((ok,fail)=>{ const s=document.createElement("script"); s.src=src; s.onload=ok; s.onerror=()=>fail(new Error("не загружен "+src)); document.head.appendChild(s); }); }
async function getWorker(progress){
  if(ocrWorker) return ocrWorker;
  if(!window.Tesseract) await loadScript("vendor/tesseract/tesseract.min.js");
  const base=new URL("vendor/tesseract/", location.href).href;
  ocrWorker = await Tesseract.createWorker("eng", 1, {
    workerPath: base+"worker.min.js", corePath: base, langPath: base+"lang", gzip: true, workerBlobURL: false,
    logger: m => progress && m.status==="recognizing text" && progress(Math.round(m.progress*100))
  });
  return ocrWorker;
}
/* ---------- подготовка фото ---------- */
function loadImg(file){
  return new Promise((ok,fail)=>{ const i=new Image(); i.onload=()=>ok(i); i.onerror=()=>fail(new Error("не удалось открыть фото")); i.src=URL.createObjectURL(file); });
}
/* фото → canvas: длинная сторона maxSide (мелкие снимки увеличиваются до 2 раз), поворот rot°, crop — доля кадра [x,y,w,h] */
function toCanvas(img, maxSide, rot, crop){
  const [fx,fy,fw,fh]=crop||[0,0,1,1], sx=img.width*fx, sy=img.height*fy, sw=img.width*fw, sh=img.height*fh;
  const k=Math.min(2, maxSide/Math.max(sw,sh)), w=Math.round(sw*k), h=Math.round(sh*k);
  const c=document.createElement("canvas"); const sideways=rot===90||rot===270;
  c.width=sideways?h:w; c.height=sideways?w:h;
  if(rot%90){ const r=rot*Math.PI/180, cs=Math.abs(Math.cos(r)), sn=Math.abs(Math.sin(r)); c.width=Math.round(w*cs+h*sn); c.height=Math.round(w*sn+h*cs); }
  const g=c.getContext("2d"); g.imageSmoothingQuality="high"; g.fillStyle="#fff"; g.fillRect(0,0,c.width,c.height);
  g.translate(c.width/2,c.height/2); g.rotate(rot*Math.PI/180); g.translate(-w/2,-h/2);
  g.drawImage(img,sx,sy,sw,sh,0,0,w,h); return c;
}
/* серое изображение с растянутой яркостью (1%–99%); mode: 0 — только это, 1 — плюс резкость, 2 — плюс адаптивный порог (неравномерный свет, блики) */
function enhance(src, mode){
  const w=src.width, h=src.height, d=src.getContext("2d").getImageData(0,0,w,h), px=d.data, n=w*h;
  const L=new Float32Array(n), hist=new Uint32Array(256);
  for(let i=0;i<n;i++){ const v=0.299*px[i*4]+0.587*px[i*4+1]+0.114*px[i*4+2]; L[i]=v; hist[v|0]++; }
  let lo=0, hi=255, acc=0; for(;lo<255 && (acc+=hist[lo])<n*0.01;lo++); acc=0; for(;hi>0 && (acc+=hist[hi])<n*0.01;hi--);
  const sc=255/Math.max(1,hi-lo); for(let i=0;i<n;i++) L[i]=Math.max(0,Math.min(255,(L[i]-lo)*sc));
  let S=L;
  if(mode===1){ S=new Float32Array(n);   // резкость: 3L − полусумма соседей
    for(let y=0;y<h;y++) for(let x=0;x<w;x++){ const i=y*w+x;
      S[i] = x&&y&&x<w-1&&y<h-1 ? Math.max(0,Math.min(255, 3*L[i]-(L[i-1]+L[i+1]+L[i-w]+L[i+w])/2)) : L[i]; } }
  if(mode===2){   // порог Брэдли: точка темнее среднего по окрестности на 15% → чёрная
    const W=w+1, I=new Float64Array(W*(h+1)), r=Math.max(8,Math.round(Math.max(w,h)/32)), B=new Float32Array(n);
    for(let y=0;y<h;y++){ let row=0; for(let x=0;x<w;x++){ row+=S[y*w+x]; I[(y+1)*W+x+1]=I[y*W+x+1]+row; } }
    for(let y=0;y<h;y++) for(let x=0;x<w;x++){
      const x0=Math.max(0,x-r), x1=Math.min(w,x+r+1), y0=Math.max(0,y-r), y1=Math.min(h,y+r+1);
      const sum=I[y1*W+x1]-I[y0*W+x1]-I[y1*W+x0]+I[y0*W+x0];
      B[y*w+x] = S[y*w+x]*(x1-x0)*(y1-y0) < sum*0.85 ? 0 : 255; }
    S=B;
  }
  for(let i=0;i<n;i++){ px[i*4]=px[i*4+1]=px[i*4+2]=S[i]; px[i*4+3]=255; }
  const c=document.createElement("canvas"); c.width=w; c.height=h; c.getContext("2d").putImageData(d,0,0); return c;
}
const pause=()=>new Promise(r=>setTimeout(r,0));   // дать экрану обновить статус между долгими попытками

/* ---------- текст: Tesseract ---------- */
function extract(text, mode, labeledOut){   // labeledOut (Set) — сюда попадут коды, стоящие после метки S/N
  const t=text.toUpperCase();
  // в кодах почти из одних цифр O и I — ошибки распознавания (0 и 1); другие буквы в серийниках бывают настоящими
  const fix=s=>{ const dg=(s.match(/\d/g)||[]).length; return dg>=s.length*0.75 ? s.replace(/O/g,"0").replace(/I/g,"1") : s; };
  if(mode==="digits") return uniq((t.match(/[0-9OI]{4,}/g)||[]).map(fix).filter(s=>/^\d{4,}$/.test(s)));
  // сначала то, что стоит после метки S/N, SN, SERIAL; затем остальные длинные коды
  const LABEL=/(?:S\/?N|SERIAL(?:\s*NO)?|SER)[\s:.#№]*([A-Z0-9][A-Z0-9\-\/]{4,})/g;
  const labeled=[...t.matchAll(LABEL)].map(m=>m[1]);
  const rest=(t.replace(LABEL," ").match(/[A-Z0-9][A-Z0-9\-\/]{5,}/g)||[]).filter(s=>/\d/.test(s));
  if(labeledOut) labeled.forEach(x=>labeledOut.add(fix(x)));
  return uniq(labeled.concat(rest).map(fix));
}
/* находки из распознанного текста с уверенностью (0..100) слов или строк; коды после метки S/N — первыми */
function makeRanker(mode){
  const score=new Map(), labeled=new Set();
  return {
    add(text, words, base){
      const key=s=>s.replace(/O/g,"0").replace(/I/g,"1");   // находка могла быть исправлена O→0, I→1
      words=words.map(x=>({t:key(x.t.toUpperCase().replace(/[^A-Z0-9\-\/]/g,"")), c:x.c})).filter(x=>x.t.length>=2);
      extract(text, mode, labeled).forEach(f=>{
        const k=key(f), c=Math.max(base||0, ...words.filter(x=>k.includes(x.t)||x.t.includes(k)).map(x=>x.c));
        score.set(f, Math.max(score.get(f)||0, c));
      });
    },
    sure(){ return [...score.values()].some(c=>c>=75); },
    list(){ return [...score.entries()].sort((a,b)=>labeled.has(b[0])-labeled.has(a[0]) || b[1]-a[1]).filter(([f,c],i)=>i<1 || c>=30).map(([f])=>f); }
  };
}
/* запасной вариант без нейросети телефона (браузер): Tesseract, несколько проходов по разным версиям снимка */
async function recognizeText(file, mode, progress){
  const img=await loadImg(file);
  try{
    const w=await getWorker(p=>progress(p));
    await w.setParameters({ tessedit_char_whitelist: mode==="digits" ? "0123456789" : "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-/:.# ",
      tessedit_pageseg_mode: "11", user_defined_dpi: "300" });
    const base=toCanvas(img, 2000, 0);
    const passes=[ ()=>enhance(base,0), ()=>enhance(base,1), ()=>enhance(base,2), ()=>enhance(toCanvas(img,2000,90),0), ()=>enhance(toCanvas(img,2000,270),0) ];
    const rank=makeRanker(mode);
    for(let p=0;p<passes.length;p++){
      progress(0, p+1, passes.length); await pause();
      const {data}=await w.recognize(passes[p]());
      rank.add(data.text, (data.words||[]).map(x=>({t:x.text, c:x.confidence})), data.confidence*0.5);
      if(rank.sure()) break;   // уверенно распознано — дальше не ищем
    }
    return rank.list();
  } finally { URL.revokeObjectURL(img.src); }
}

/* ---------- штрихкоды и QR ---------- */
/* среднее по k строкам: шум и зерно уходят, вертикальные штрихи остаются чёткими (для линейных кодов) */
function smear(src, k){
  const w=src.width, h=src.height, p=src.getContext("2d").getImageData(0,0,w,h).data, c=document.createElement("canvas");
  c.width=w; c.height=h; const g=c.getContext("2d"), o=g.createImageData(w,h), q=o.data, half=k>>1;
  for(let x=0;x<w;x++){
    let s=0; for(let y=-half;y<=half;y++) s+=p[(Math.min(h-1,Math.max(0,y))*w+x)*4+1];
    for(let y=0;y<h;y++){ const i=(y*w+x)*4; q[i]=q[i+1]=q[i+2]=s/k; q[i+3]=255;
      s+=p[(Math.min(h-1,y+half+1)*w+x)*4+1]-p[(Math.max(0,y-half)*w+x)*4+1]; }
  }
  g.putImageData(o,0,0); return c;
}
async function decodeCodes(file, progress){
  const img=await loadImg(file);
  try{
    // встроенный распознаватель системы (есть в Chrome/WebView на многих Android) — быстрый и точный
    if("BarcodeDetector" in window){
      try{ const r=await new BarcodeDetector().detect(img); const v=uniq(r.map(x=>String(x.rawValue||"").trim())); if(v.length) return v; }catch(e){}
    }
    if(!window.ZXing) await loadScript("vendor/zxing.min.js");
    const Z=ZXing, hints=new Map();
    hints.set(Z.DecodeHintType.TRY_HARDER, true);
    hints.set(Z.DecodeHintType.CHARACTER_SET, "UTF-8");
    hints.set(Z.DecodeHintType.POSSIBLE_FORMATS, [Z.BarcodeFormat.QR_CODE, Z.BarcodeFormat.DATA_MATRIX, Z.BarcodeFormat.PDF_417, Z.BarcodeFormat.AZTEC,
      Z.BarcodeFormat.CODE_128, Z.BarcodeFormat.CODE_39, Z.BarcodeFormat.CODE_93, Z.BarcodeFormat.EAN_13, Z.BarcodeFormat.EAN_8, Z.BarcodeFormat.ITF, Z.BarcodeFormat.UPC_A, Z.BarcodeFormat.CODABAR]);
    const reader=new Z.MultiFormatReader(); reader.setHints(hints);
    const read=(cv, both)=>{ for(const Bin of both ? [Z.HybridBinarizer, Z.GlobalHistogramBinarizer] : [Z.HybridBinarizer]){
      try{ const r=reader.decode(new Z.BinaryBitmap(new Bin(new Z.HTMLCanvasElementLuminanceSource(cv)))); if(r && r.getText()) return r.getText().trim(); }catch(e){} } return ""; };
    // кадры по возрастанию затрат: целиком (уменьшенный и крупнее); центр без уменьшения; центр с увеличением (мелкий код); четверти; наклон 30°/60°
    const C=[.2,.2,.6,.6], C2=[.3,.3,.4,.4], Q=[[0,0,.55,.55],[.45,0,.55,.55],[0,.45,.55,.55],[.45,.45,.55,.55]];
    const tries=[ [1600,0], [1600,90], [2600,0], [2600,90], [2400,0,C], [2400,90,C], [3200,0,C2], [3200,90,C2],
      ...Q.flatMap(q=>[[2400,0,q],[2400,90,q]]), [2400,30,C], [2400,330,C], [2400,60,C], [2400,120,C] ];
    const t0=Date.now();
    for(let i=0;i<tries.length && Date.now()-t0<20000;i++){
      const [side,rot,crop]=tries[i]; progress && progress(i+1, tries.length); await pause();
      const c=toCanvas(img, side, rot, crop);
      const v = read(c, true) || read(smear(c, 9)) || read(enhance(c, 2));
      if(v) return [v];
    }
    return [];
  } finally { URL.revokeObjectURL(img.src); }
}
/* ---------- нейросети ML Kit на телефоне (без интернета): штрихкоды, QR и текст ---------- */
const mlWait={};
window.onMlScan=(id,err,res)=>{ const p=mlWait[id]; if(!p) return; delete mlWait[id]; err?p.fail(new Error(err)):p.ok(res); };
async function mlScan(file, mode){
  const b64=await blobToB64(file);
  return new Promise((ok,fail)=>{ const id="m"+Date.now()+Math.random(); mlWait[id]={ok,fail}; Android.mlScan(id, b64, mode); });
}

let camCb=null;
function scan(mode, cb){ camCb={mode,cb}; const inp=$("#camInput"); inp.value=""; inp.click(); }
/* mode: digits — номера пломб (код или текст), serial — серийный номер (код или текст), text — только штрихкод/QR (описание, работы) */
async function onPhoto(file){
  const {mode, cb}=camCb||{}; if(!file||!cb) return;
  const title={digits:"Номера пломб с фото", serial:"Серийный номер с фото", text:"Штрихкод / QR-код"}[mode];
  const bg=sheet(title,
    `<div class="bd"><div class="pad"><img alt=""><p class="muted" data-st>Ищу штрихкод или QR-код…</p><div class="chips" data-ch></div>
     <label>Результат (можно исправить)</label>${mode==="text" ? `<textarea data-res rows="3"></textarea>` : `<input data-res ${mode==="digits"?'inputmode="numeric"':'autocapitalize="characters"'}>`}
     <div class="btns"><button class="btn" data-again>Переснять</button><button class="btn pri" data-ok style="flex:1">Добавить</button></div></div></div>`);
  bg.querySelector("img").src=URL.createObjectURL(file);
  const st=bg.querySelector("[data-st]"), res=bg.querySelector("[data-res]"), ch=bg.querySelector("[data-ch]");
  bg.querySelector("[data-ok]").onclick=()=>{ const v=mode==="digits"?parseSeals(res.value):[res.value.trim()].filter(Boolean); if(v.length) cb(v); bg.remove(); };
  if(mode==="text") bg.querySelector("[data-ok]").textContent="Вставить";
  bg.querySelector("[data-again]").onclick=()=>{ bg.remove(); scan(mode, cb); };
  try{
    let found=[], how="";
    const fromCodes=codes=>mode==="digits" ? extract(codes.join(" "),"digits") : mode==="serial" ? uniq(codes.concat(extract(codes.join(" "),"serial"))) : codes;
    // сначала нейросети телефона; если их нет (браузер) или они не запустились — ZXing и Tesseract
    let ml=null;
    if(NATIVE && Android.mlScan){
      st.textContent="Распознаю нейросетью на телефоне…";
      ml=await mlScan(file, mode).catch(e=>{ console.warn("ML Kit:", e); return null; });
    }
    if(ml){
      if((ml.codes||[]).length){ how="по коду"; found=fromCodes(ml.codes); }
      if(!found.length && mode!=="text"){
        const rank=makeRanker(mode), lines=ml.lines||[];
        rank.add(lines.map(l=>l.t).join("\n"), lines.map(l=>({t:l.t, c:(l.c||0)*100})), 0);
        found=rank.list(); how="по тексту";
      }
      if(!found.length){   // нейросеть ничего не нашла — ещё поиск кода прежним способом (блёклые, мелкие коды)
        const codes=await decodeCodes(file, (i,n)=>st.textContent="Ищу штрихкод или QR-код… "+i+"/"+n).catch(e=>{ console.warn(e); return []; });
        if(codes.length){ how="по коду"; found=fromCodes(codes); }
      }
    } else {
      const codes=await decodeCodes(file, (i,n)=>st.textContent="Ищу штрихкод или QR-код… "+i+"/"+n).catch(e=>{ console.warn(e); return []; });
      if(codes.length){ how="по коду"; found=fromCodes(codes); }
      if(!found.length && mode!=="text"){
        st.textContent="Код не найден, распознаю текст…";
        let pass="";
        found=await recognizeText(file, mode, (p,i,n)=>{ if(n) pass=n>1?" (попытка "+i+" из "+n+")":""; st.textContent="Распознаю текст… "+p+"%"+pass; });
        how="по тексту";
      }
    }
    st.textContent = found.length ? "Найдено "+how+(found.length>1?" — нажмите, чтобы выбрать:":":")
                                  : (mode==="text" ? "Штрихкод/QR-код не найден — снимите ближе и ровнее." : "Ничего не распознано, введите вручную.");
    if(!found.length) keepMiss(file, mode, st);
    let sel=new Set(mode==="digits"?found:found.slice(0,1));
    const draw=()=>{ ch.innerHTML=found.map(f=>`<button class="chip ${sel.has(f)?"on":""}" data-f="${esc(f)}">${esc(f)}</button>`).join(""); res.value=[...sel].join(mode==="digits"?", ":""); };
    ch.onclick=e=>{ const f=e.target.dataset.f; if(!f) return; if(mode==="digits"){ sel.has(f)?sel.delete(f):sel.add(f); } else sel=new Set([f]); draw(); };
    draw();
  }catch(e){ console.error(e); st.textContent="Распознавание недоступно ("+e.message+"). Введите вручную."; keepMiss(file, mode, st); }
}

/* ================= Нераспознанные фото: хранятся в IndexedDB (в localStorage не поместятся), выгружаются в ZIP ================= */
const MISS_MAX=50, MISS_SIDE=2500, MISS_PART=20e6, MISS_MODE={digits:"пломбы", serial:"серийный", text:"код"};
let missDb=null;
function missOpen(){
  return missDb || (missDb=new Promise((ok,fail)=>{ const r=indexedDB.open("akty-photos",1);
    r.onupgradeneeded=()=>r.result.createObjectStore("miss",{keyPath:"id",autoIncrement:true});
    r.onsuccess=()=>ok(r.result); r.onerror=()=>{ missDb=null; fail(r.error); }; }));
}
async function missTx(mode, fn){
  const db=await missOpen();
  return new Promise((ok,fail)=>{ const t=db.transaction("miss",mode), q=fn(t.objectStore("miss")); t.oncomplete=()=>ok(q&&q.result); t.onerror=t.onabort=()=>fail(t.error); });
}
const missAll=()=>missTx("readonly",s=>s.getAll());
/* копия снимка (длинная сторона до MISS_SIDE) и миниатюра для списка; самые старые снимки сверх MISS_MAX удаляются */
async function missSave(file, mode){
  const img=await loadImg(file), k=Math.min(1, MISS_SIDE/Math.max(img.width,img.height)), c=document.createElement("canvas");
  c.width=Math.round(img.width*k); c.height=Math.round(img.height*k); c.getContext("2d").drawImage(img,0,0,c.width,c.height);
  URL.revokeObjectURL(img.src);
  const photo=await new Promise(ok=>c.toBlob(ok,"image/jpeg",.9)), thumb=await photoData(file,200,.6);
  await missTx("readwrite",s=>s.add({date:new Date().toISOString(), mode, azs:act.azs||"", thumb, photo}));
  const keys=await missTx("readonly",s=>s.getAllKeys());
  if(keys.length>MISS_MAX) await missTx("readwrite",s=>{ keys.slice(0,keys.length-MISS_MAX).forEach(id=>s.delete(id)); });
  renderMiss();
}
const keepMiss=(file,mode,st)=>missSave(file,mode).then(()=>{ st.textContent+=" Фото сохранено в «Данные → Нераспознанные фото»."; }).catch(e=>console.warn("Фото не сохранено:", e));
const missDay=r=>{ const d=new Date(r.date); return [d.getFullYear(),pad(d.getMonth()+1),pad(d.getDate())].join("-"); };   // местная дата снимка
function missName(r){
  const d=new Date(r.date), t=missDay(r)+"_"+[pad(d.getHours()),pad(d.getMinutes()),pad(d.getSeconds())].join("-");
  return [t, r.azs&&"АЗС"+r.azs, MISS_MODE[r.mode], r.id].filter(Boolean).join("_")+".jpg";
}
async function renderMiss(){
  const all=await missAll().catch(e=>{ console.warn(e); return []; });
  $("#missCount").textContent=all.length ? "Сохранено: "+all.length+" из "+MISS_MAX : "Пока нет сохранённых фото.";
  $("#missList").innerHTML=all.slice().reverse().map(r=>`<div class="miss"><img src="${esc(r.thumb)}" data-miss="${r.id}" alt="">
    <button class="ico del" data-rmmiss="${r.id}" title="Удалить">${ICON_TRASH}</button>
    <div class="muted">${esc(fmtDate(missDay(r)))} · ${esc(MISS_MODE[r.mode]||"")}${r.azs?" · АЗС "+esc(r.azs):""}</div></div>`).join("");
  $("#missExp").disabled=$("#missClr").disabled=!all.length;
}
/* в APK файл идёт через мост строкой base64, поэтому большая выгрузка делится на несколько ZIP */
async function missExport(){
  const parts=[[]]; let size=0;
  (await missAll()).forEach(r=>{ if(NATIVE && size+r.photo.size>MISS_PART && parts[parts.length-1].length){ parts.push([]); size=0; }
    parts[parts.length-1].push(r); size+=r.photo.size; });
  for(let i=0;i<parts.length;i++){
    const zip=new JSZip(); parts[i].forEach(r=>zip.file(missName(r), r.photo));
    await saveBlob(await zip.generateAsync({type:"blob", mimeType:"application/zip"}), "Нераспознанные_"+today()+(parts.length>1?"_"+(i+1):"")+".zip");
  }
  if(NATIVE) toast("Сохранено в «Загрузки/АЗС»"+(parts.length>1?": архивов "+parts.length:""));
}
function bindMiss(){
  $("#missList").addEventListener("click",async e=>{
    const rm=e.target.closest("[data-rmmiss]");
    if(rm){ if(confirm("Удалить это фото?")){ await missTx("readwrite",s=>s.delete(+rm.dataset.rmmiss)); renderMiss(); } return; }
    const id=e.target.dataset.miss; if(id==null) return;
    const r=await missTx("readonly",s=>s.get(+id)); if(!r) return;
    const img=sheet(missName(r), `<div class="bd"><div class="pad"><img alt="" style="width:100%"></div></div>`).querySelector("img");
    img.onload=()=>URL.revokeObjectURL(img.src); img.src=URL.createObjectURL(r.photo);
  });
  $("#missExp").onclick=()=>missExport().catch(e=>{ console.error(e); alert("Не удалось выгрузить фото: "+e.message); });
  $("#missClr").onclick=async()=>{ if(!confirm("Удалить все нераспознанные фото?")) return; await missTx("readwrite",s=>s.clear()); renderMiss(); };
}

/* ================= Подпись по фото ================= */
function processSignature(file){
  return new Promise((ok,fail)=>{
    const img=new Image();
    img.onload=()=>{
      const k=Math.min(1, 900/Math.max(img.width,img.height));
      const w=Math.round(img.width*k), h=Math.round(img.height*k);
      const c=document.createElement("canvas"); c.width=w; c.height=h; const g=c.getContext("2d");
      g.drawImage(img,0,0,w,h); URL.revokeObjectURL(img.src);
      const d=g.getImageData(0,0,w,h), px=d.data, lum=new Uint8Array(w*h), hist=new Array(256).fill(0);
      for(let i=0;i<w*h;i++){ const L=Math.round(0.299*px[i*4]+0.587*px[i*4+1]+0.114*px[i*4+2]); lum[i]=L; hist[L]++; }
      // порог Оцу отделяет чернила от бумаги
      let sum=0; for(let i=0;i<256;i++) sum+=i*hist[i];
      let sB=0,wB=0,best=0,thr=128; for(let t=0;t<256;t++){ wB+=hist[t]; if(!wB) continue; const wF=w*h-wB; if(!wF) break; sB+=t*hist[t];
        const mB=sB/wB, mF=(sum-sB)/wF, v=wB*wF*(mB-mF)*(mB-mF); if(v>best){ best=v; thr=t; } }
      let x0=w,y0=h,x1=-1,y1=-1;
      for(let y=0;y<h;y++) for(let x=0;x<w;x++){ const i=y*w+x, L=lum[i];
        const a = L>=thr ? 0 : Math.min(255, Math.round((thr-L)/Math.max(thr,1)*255*1.8));
        px[i*4]=20; px[i*4+1]=35; px[i*4+2]=110; px[i*4+3]=a;          // синие «чернила»
        if(a>60){ if(x<x0)x0=x; if(x>x1)x1=x; if(y<y0)y0=y; if(y>y1)y1=y; } }
      if(x1<0) return fail(new Error("на фото не найдена подпись"));
      g.putImageData(d,0,0);
      const m=6; x0=Math.max(0,x0-m); y0=Math.max(0,y0-m); x1=Math.min(w-1,x1+m); y1=Math.min(h-1,y1+m);
      const cw=x1-x0+1, ch=y1-y0+1, s=Math.min(1, 160/ch), out=document.createElement("canvas");
      out.width=Math.round(cw*s); out.height=Math.round(ch*s);
      out.getContext("2d").drawImage(c,x0,y0,cw,ch,0,0,out.width,out.height);
      ok({url:out.toDataURL("image/png"), ratio:cw/ch});
    };
    img.onerror=()=>fail(new Error("не удалось открыть фото"));
    img.src=URL.createObjectURL(file);
  });
}
function renderSign(){
  $("#signBox").innerHTML = settings.sign ? `<img src="${settings.sign}" alt="подпись">` : `<span class="muted">Подпись не загружена</span>`;
  $("#signDel").classList.toggle("hidden", !settings.sign);
  $("#s-signOn").checked=!!settings.signOn;
}

/* ================= Акты по разметке Excel ================= */
const D = iso => ({date: iso});
const T = t => ({time: t});
/* «Договор № 12/34 от 01.02.2026 г.» для акта key, или "" если номер не задан */
function contractText(key, prefix){
  const c=(settings.contracts||{})[key]||{}, num=String(c.num||"").trim();
  return num ? (prefix||"Договор № ")+num+(c.date?" от "+fmtDate(c.date)+" г.":"") : "";
}
function cellText(v, fmt){
  if(v==null) return "";
  if(typeof v==="object"){
    if(v.date!=null) return fmt==="dlong" ? fmtDateLong(v.date) : fmtDate(v.date);
    if(v.time!=null) return v.time||"";
    return "";
  }
  return String(v);
}
/* Подпись исполнителя: где на листе стоит (ячейка) и сдвиг от её левого края; накладывается поверх, разметку не меняет */
function textW(text, font){ const c=textW.c||(textW.c=document.createElement("canvas").getContext("2d")); c.font=font; return c.measureText(text).width; }
if(document.fonts) ["11pt Carlito","italic 10pt Carlito"].forEach(f=>document.fonts.load(f,"Яя_"));   // иначе первый замер textW идёт по запасному шрифту
const SIGN_H=41, SIGN_MAXW=156;   // высота и наибольшая ширина подписи в px листа
function signW(){ return settings.signRatio ? Math.min(SIGN_H*settings.signRatio, SIGN_MAXW) : 108; }
const LEAD=" ".repeat(53);   // отступ перед «/ ФИО» в ВБРР и ККТ
const signMid=(start,w)=>start+(w-signW())/2;   // сдвиг, при котором подпись стоит по центру места шириной w
const lineW=(n,font)=>textW("_".repeat(n),font);
const colsW=(sheet,from,to)=>window.LAYOUTS[sheet].cols.slice(from,to).reduce((a,b)=>a+b,0);
/* ВБРР и ККТ: линии нет — подпись над «(подпись)» (её центр в 108 px от края ячейки), но не заходит на «/ ФИО» */
const signCap=()=>Math.min(108-signW()/2, 2+textW(LEAD,"italic 10pt Carlito")-signW()-2);
const SIGN_SPOTS = {
  vr:   [ {ref:"A31", x:()=>signMid(2+textW("Работу выполнил "+FIO().short+" подпись ", "11pt Carlito"), lineW(19,"11pt Carlito"))},
          {ref:"C36", x:()=>signMid(2, lineW(13,"11pt Carlito"))} ],
  pl:   [ {ref:"BG30", x:()=>signMid(0, colsW("Пломбы",58,70))} ],   // над «(подпись)» в BK31
  vbrr: [ {ref:"AT31", x:signCap} ],
  kkt:  [ {ref:"AT40", x:signCap} ]
};
function signImg(spot){
  return `<img class="sign" src="${settings.sign}" alt="" style="left:${Math.round(spot.x())}px;height:${SIGN_H}px;max-width:${SIGN_MAXW}px${spot.mid?";transform:translateY(-50%)":""}">`;
}

/* Описание каждого акта: лист, постоянные ячейки и блоки строк, которые растягиваются под число записей */
const FORMS = [
  { key:"vr", title:"ВР", sheet:"ВР", extra:{ A12:{to:4, like:"A9"} },
    cells: M => ({
      A4:settings.cust, C4:settings.org, A12:contractText("vr"),
      A8:"СДАЧИ-ПРИЕМКИ РАБОТ на "+M.type+" № "+M.num, A11:D(M.date),
      A13:"Мы, ниже подписавшиеся, представитель «Заказчика», в лице руководителя АЗС, расположенной по адресу "+M.addr+" с одной стороны и представитель «Исполнителя», в лице "+FIO().gen+" с другой стороны, составили настоящий акт в том, что произведены следующие виды работ:",
      B19: M.mats.length ? "Расходные материалы :" : "",
      B27:"Время прибытия на "+M.type, C27:T(M.tArr), B28:"Время убытия с "+M.type, C28:T(M.tDep),
      A31:"Работу выполнил "+FIO().short+" подпись ___________________", C36:"_____________ / "+FIO().short+" " }),
    blocks: [
      { from:15, to:18, items:M=>M.works, fill:(w,i,M)=>({A:String(i+1), B:w, C:"Выполнено", D:i===0?M.req:""}) },
      { from:20, to:24, items:M=>M.mats, fill:(m,i)=>({A:String(i+1), B:matText(m), C:m.qty, D:m.unit}) } ] },
  { key:"pl", title:"Пломбы", sheet:"Пломбы",
    cells: M => ({ A3:settings.org, A4:contractText("pl"), A30:settings.org, A5:D(M.date), A7:M.obj+", "+M.addr, A11:D(M.date), O11:T(M.tArr), X11:T(M.tDep), BW11:M.req, Z30:settings.pos, BT30:FIO().short }),
    blocks: [ { from:16, to:27, items:M=>M.eq.filter(e=>e.off||e.on), fill:(e,i,M)=>({A:e.off, W:e.place, BA:e.on, BW:M.req}) } ] },
  { key:"vbrr", title:"ВБРР", sheet:"ВБРР", extra:{ A4:{to:87, like:"A2"} },
    cells: M => { const c={ A4:contractText("vbrr"), A25:settings.cust, AT25:settings.org, AI3:M.req, AX3:D(M.date), A7:M.desc, AJ14:M.mats.length?"Х":"", A27:M.obj, AT27:settings.dept, AT29:settings.pos,
        AT31:"                                                     / "+FIO().short, AT34:T(M.tArr), BL34:D(M.date), AT35:T(M.tDep), BL35:D(M.date) };
      VBRR_PLACES.forEach((p,i)=>{ const s=M.sealsBy(p); c["X"+(21+i)]=s.off; c["BJ"+(21+i)]=s.on; }); return c; },
    blocks: [
      { from:9, to:12, items:M=>M.works, fill:w=>({A:w}) },
      { from:15, to:19, items:M=>M.mats, fill:m=>({A:matText(m), BE:matQty(m)}) } ] },
  { key:"kkt", title:"ККТ", sheet:"ККТ",
    cells: M => { const c={ A34:settings.cust, AT34:settings.org, AT36:settings.dept,
        AI3:M.req, AX3:D(M.date), A8:M.desc, AJ15:M.mats.length?"Х":"", A31:"", A36:M.obj, AT28:"ККТ ГКС 2", AT38:settings.pos,
        AT40:"                                                     / "+FIO().short, AT42:T(M.tArr), BL42:D(M.date), AT43:T(M.tDep), BL43:D(M.date), AU45:M.km, AU46:M.km };
      let sealed=false;
      KKT_PLACES.forEach((p,i)=>{ const s=M.sealsBy(p); c["H"+(23+i)]=s.off; c["BA"+(23+i)]=s.on; if(s.off||s.on) sealed=true; });
      c.CD19=sealed?"Х":""; const ct=contractText("kkt","Договор: № "); if(ct) c.W4=ct; return c; },
    blocks: [
      { from:10, to:13, items:M=>M.works, fill:w=>({A:w}) },
      { from:16, to:18, items:M=>{ const p=[]; for(let i=0;i<M.mats.length;i+=2) p.push([M.mats[i],M.mats[i+1]]); return p; },
        fill:([l,r])=>({A:l?matText(l)+" "+matQty(l):"", AT:r?matText(r)+" "+matQty(r):""}) } ] }
];

function renderSheet(form, M){
  const L=window.LAYOUTS[form.sheet]; const [r1,c1,r2,c2]=L.area; const names=L.colNames;
  const merge=new Map(), covered=new Set();
  const extra=form.extra||{};   // ячейки, которых нет в шаблоне: {A4:{to:последний столбец, like:ячейка-образец стиля}}
  Object.entries(extra).forEach(([ref,x])=>{ const m=ref.match(/^([A-Z]+)(\d+)$/), c=names.indexOf(m[1])+c1; L.merges.some(g=>g[0]==+m[2]&&g[1]==c) || L.merges.push([+m[2],c,+m[2],x.to]); });
  const rspan=new Map();   // объединения по строкам — только для листов с form.rowspan (отчёт по ГСМ); акты рисуются как раньше
  L.merges.forEach(([mr1,mc1,mr2,mc2])=>{ merge.set(mr1+","+mc1, mc2-mc1+1);
    const rl=form.rowspan ? mr2 : mr1; if(rl>mr1) rspan.set(mr1+","+mc1, rl-mr1+1);
    for(let r=mr1;r<=rl;r++) for(let c=mc1;c<=mc2;c++) if(r!==mr1 || c!==mc1) covered.add(r+","+c); });
  const statics=form.cells(M);
  const spots={}; if(settings.sign && settings.signOn) (SIGN_SPOTS[form.key]||[]).forEach(s=>spots[s.ref]=s);
  const row=(tr, vals)=>{   // vals: значения по букве столбца для блочной строки, иначе null
    let h=`<tr style="height:${L.rows[tr]}px">`;
    for(let c=c1;c<=c2;c++){
      if(covered.has(tr+","+c)) continue;
      const col=names[c-c1], ref=col+tr, cell=L.cells[ref]||(extra[ref]&&L.cells[extra[ref].like]), span=merge.get(tr+","+c);
      let v;
      if(vals) v = vals[col]!=null ? vals[col] : "";
      else v = ref in statics ? statics[ref] : (cell && !cell[3] ? cell[1] : "");
      const t=cellText(v, cell && cell[2]);
      const sp=!vals ? spots[ref] : (vals[col] && vals[col].sign!=null && {x:()=>vals[col].sign, mid:true});
      const rs=rspan.get(tr+","+c);
      h+=`<td${span>1?` colspan="${span}"`:""}${rs?` rowspan="${rs}"`:""}${cell||sp?` style="${cell?L.styles[cell[0]]:""}${sp?";position:relative":""}"`:""}>${t?esc(t):""}${sp?signImg(sp):""}</td>`;
    }
    return h+"</tr>";
  };
  let body="";
  for(let r=r1;r<=r2;r++){
    const b=form.blocks.find(x=>r>=x.from && r<=x.to);
    if(!b){ body+=row(r,null); continue; }
    if(r!==b.from) continue;
    const items=b.items(M), slots=b.to-b.from+1, n=Math.max(slots, items.length);
    for(let i=0;i<n;i++) body+=row(b.from+Math.min(i,slots-1), i<items.length ? b.fill(items[i],i,M) : {});
  }
  const width=L.cols.reduce((a,b)=>a+b,0);
  const [mt,mr,mb,ml]=L.margins.map(x=>x*96);
  const pw=L.landscape ? 1123 : 794;   // ширина листа A4, px
  const zoom=Math.min((L.scale||100)/100, (pw-ml-mr)/width);
  return `<div class="page" style="padding:${mt}px ${mr}px ${mb}px ${ml}px${L.landscape?";width:1123px;min-height:794px":""}"><table class="xl" style="width:${width}px;zoom:${zoom}"><colgroup>${L.cols.map(w=>`<col style="width:${w}px">`).join("")}</colgroup>${body}</table></div>`;
}
function renderPrint(){
  document.querySelectorAll("[data-form]").forEach(c=>c.classList.toggle("on",!!act.forms[c.dataset.form]));
  const M=model();
  const on=FORMS.filter(f=>act.forms[f.key]);
  $("#pages").innerHTML = on.length ? on.map(f=>`<div class="pagehead">${f.title}</div><div class="pagewrap">${renderSheet(f,M)}</div>`).join("")
    : `<p class="muted">Выберите хотя бы один акт.</p>`;
  fitPages();
}
/* масштаб листа A4 под ширину экрана */
function fitPages(){
  document.querySelectorAll(".pagewrap").forEach(w=>{
    const p=w.firstElementChild; const k=Math.min(1, w.clientWidth/794);
    p.style.transform=`scale(${k})`; w.style.height=(p.offsetHeight*k)+"px";
  });
}
if(document.fonts) document.fonts.ready.then(()=>{ if(!$("#tab-print").classList.contains("hidden")) renderPrint(); });
window.addEventListener("resize",()=>{ if(!$("#tab-print").classList.contains("hidden")) fitPages(); });
function actFileName(){ return [fmtDate(act.date), act.azs, act.req].filter(Boolean).join(" ").replace(/[\\/:*?"<>|]/g,"_"); }
function printActs(){
  if(!FORMS.some(f=>act.forms[f.key])) return toast("Выберите хотя бы один акт");
  saveToDb();
  if(NATIVE) Android.print(actFileName()); else window.print();
}
/* PDF выбранных актов сразу в «Загрузки/АЗС/Акты» (в браузере — через окно печати) */
function savePdf(){
  if(!FORMS.some(f=>act.forms[f.key])) return toast("Выберите хотя бы один акт");
  saveToDb();
  if(!NATIVE || !Android.savePdf) return webPdf().then(f=>f && saveBlob(f, f.name)).catch(e=>alert("Не удалось сохранить PDF: "+e.message));
  toast("Сохраняю PDF…");
  Android.savePdf(actFileName()+".pdf");
}
/* PDF в браузере (iPhone): каждый лист рисуется картинкой (2.5× — чётко при печати) на страницу A4 */
let pdfBusy=false;
async function webPdf(){
  if(pdfBusy) return null;
  pdfBusy=true; toast("Готовлю PDF…");
  const box=document.createElement("div"); box.style.cssText="position:fixed;left:-20000px;top:0;width:794px";
  document.body.appendChild(box);
  try{
    if(!window.html2canvas) await loadScript("vendor/html2canvas.min.js");
    if(!window.jspdf) await loadScript("vendor/jspdf.umd.min.js");
    if(document.fonts) await document.fonts.ready;
    const pdf=new jspdf.jsPDF({unit:"mm", format:"a4", compress:true});
    const pages=[...document.querySelectorAll("#pages .page")];
    for(let i=0;i<pages.length;i++){
      const pg=pages[i].cloneNode(true); pg.style.transform="none"; pg.style.boxShadow="none";
      // html2canvas не понимает CSS zoom — тот же масштаб через transform
      pg.querySelectorAll("table.xl").forEach(t=>{ const z=parseFloat(t.style.zoom)||1; t.style.zoom=""; if(z!==1){ t.style.transform="scale("+z+")"; t.style.transformOrigin="0 0"; } });
      box.appendChild(pg);
      const c=await html2canvas(pg, {scale:2.5, backgroundColor:"#ffffff", logging:false});
      pg.remove();
      if(i) pdf.addPage();
      const h=210*c.height/c.width, k=Math.min(1, 297/h);   // лист выше A4 — уменьшается целиком, без искажения
      pdf.addImage(c.toDataURL("image/jpeg", 0.9), "JPEG", (210-210*k)/2, 0, 210*k, h*k);
    }
    return new File([pdf.output("blob")], actFileName()+".pdf", {type:"application/pdf"});
  } finally { box.remove(); pdfBusy=false; }
}
/* «Поделиться» файлом (iPhone: Почта, Файлы, мессенджеры). Safari разрешает его только сразу после нажатия —
   если PDF готовился долго, показываем кнопку для повторного нажатия */
const MOBILE=/iPhone|iPad|iPod|Android/i.test(navigator.userAgent) || (navigator.platform==="MacIntel" && navigator.maxTouchPoints>1);
function canShareFile(file){ return MOBILE && !!navigator.canShare && navigator.canShare({files:[file]}); }
async function shareFile(file, title, text){
  const data={files:[file]}; if(title) data.title=title; if(text) data.text=text;
  try{ await navigator.share(data); }
  catch(e){
    if(e.name==="AbortError") return;
    if(e.name!=="NotAllowedError") throw e;
    await new Promise(ok=>{
      const bg=sheet("Файл готов", `<div class="pad"><p>${esc(file.name)}</p><div class="btns"><button class="btn pri" data-go-share>Поделиться…</button></div></div>`);
      bg.querySelector("[data-go-share]").onclick=()=>{ bg.remove(); navigator.share(data).catch(err=>{ if(err.name!=="AbortError") alert(err.message); }).finally(ok); };
      new MutationObserver((m,o)=>{ if(!bg.isConnected){ o.disconnect(); ok(); } }).observe(document.body,{childList:true});
    });
  }
}
/* письмо с PDF выбранных актов: тема и текст собираются из акта */
/* адреса из строки: «"Имя" <адрес>, адрес; …» → [адрес, …] */
function parseAddrs(s){ return uniq((String(s||"").match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g)||[]).map(a=>a.toLowerCase())); }
const replyAddrs=()=>parseAddrs(settings.replyTo);
function mailText(){
  const subject=actFileName();   // тема письма совпадает с именем PDF-файла (без .pdf)
  const to=replyAddrs();
  const body = to.length
    ? "Просьба подписать акты, поставить печать, отсканировать и отправить по следующему адресу:\n"+
      "mailto:"+to.join(";")+"?subject="+encodeURIComponent(subject)   // «;» между адресами понимает Outlook
    : "Просьба подписать акты, поставить печать, отсканировать и отправить в ответ на это письмо.";
  return {subject, body};
}
function sendMail(){
  if(!FORMS.some(f=>act.forms[f.key])) return toast("Выберите хотя бы один акт");
  saveToDb();
  // кому: почта АЗС из справочника и постоянные адреса из «Данные → Исполнитель»
  const {subject, body}=mailText(), to=uniq(parseAddrs(azsOf(act.azs).email).concat(parseAddrs(settings.mailTo))).join(", ");
  if(!to) toast("Не указано, кому отправлять: адрес можно ввести в письме, добавить почту АЗС в «Данные → АЗС» или адреса в «Данные → Исполнитель»");
  if(NATIVE && Android.emailPdf){ toast("Готовлю письмо…"); Android.emailPdf(actFileName()+".pdf", to, subject, body); return; }
  if(MOBILE && navigator.canShare){
    // адрес в «Поделиться» не передаётся — кладём его в буфер обмена, чтобы вставить в поле «Кому»
    if(to && navigator.clipboard) navigator.clipboard.writeText(to).then(()=>toast("Адрес "+to+" скопирован — вставьте его в поле «Кому»")).catch(()=>{});
    webPdf().then(f=>{ if(!f) return; if(!canShareFile(f)) return saveBlob(f, f.name); return shareFile(f, subject, body); })
      .catch(e=>alert("Не удалось подготовить письмо: "+e.message));
    return;
  }
  location.href="mailto:"+encodeURIComponent(to).replace(/%2C/gi,",")+"?subject="+encodeURIComponent(subject)+"&body="+encodeURIComponent(body);
  toast("В браузере PDF не прикладывается — сохраните его кнопкой «Сохранить PDF» и приложите вручную");
}
window.onMailReady=err=>{ if(err) alert("Не удалось подготовить письмо: "+err); };
window.onPdfSaved=(err, where)=>{ if(err) alert("Не удалось сохранить PDF: "+err); else toast("PDF сохранён: "+where); };

/* ================= База актов ================= */
function actHasContent(a){ return a.req || a.works.some(w=>w.trim()) || a.equip.some(e=>e.off||e.on); }
function saveToDb(){
  if(!actHasContent(act)) return false;   // пустые акты в базу не попадают
  act.saved=new Date().toISOString();
  const i=acts.findIndex(a=>a.id===act.id);
  const snap=JSON.parse(JSON.stringify(act));
  if(i>=0) acts[i]=snap; else acts.unshift(snap);
  saveActs(); saveAct();
  learn();
}
/* пополнение справочников из акта */
function learn(){
  act.works.map(w=>w.trim()).filter(Boolean).forEach(w=>{ if(!ref.works.some(x=>x.trim()===w)) ref.works.push(w); });
  act.mats.forEach(m=>{
    const name=m.name.trim(); if(!name) return;
    if(!ref.models.includes(name)) ref.models.push(name);
    const sn=(m.sn||"").trim();
    if(sn && !ref.equipBase.some(e=>e.sn===sn && String(e.azs)===String(act.azs)))
      ref.equipBase.unshift({azs:act.azs, model:name, sn, date:act.date, req:act.req});
  });
  if(settings.autoSeal){
    const used=usedSeals(), before=ref.seals.length;
    ref.seals=ref.seals.filter(s=>!used.has(s));
    if(before!==ref.seals.length) toast("Со склада списано пломб: "+(before-ref.seals.length));
  }
  saveRef();
}
function repeatAct(src){
  const a=JSON.parse(JSON.stringify(src));
  a.id=Date.now(); a.date=today(); a.req=""; a.tArr=""; a.tDep=""; delete a.saved;
  a.equip.forEach(e=>{ e.off=""; e.on=""; });
  a.mats.forEach(m=>{ if(m.sn!=null) m.sn=""; });
  return a;
}
const ICON_REPEAT=`<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v5h-5"/></svg>`;
const ICON_TRASH=`<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16"/><path d="M9 7V4h6v3"/><path d="M6 7l1 13h10l1-13"/><path d="M10 11v6M14 11v6"/></svg>`;
function renderDb(){
  const q=$("#dbSearch").value.toLowerCase();
  const list=acts.slice().sort((a,b)=>(b.date||"").localeCompare(a.date||"") || b.id-a.id)
    .filter(a=>!q || JSON.stringify([a.azs,a.req,a.desc,a.works,a.equip,a.mats,fmtDate(a.date)]).toLowerCase().includes(q));
  $("#dbCount").textContent=acts.length ? "всего "+acts.length : "";
  $("#dbList").innerHTML = list.map(a=>{ const az=azsOf(a.azs);
    return `<div class="li" data-id="${a.id}"><div><b>${esc(fmtDate(a.date))} · ${esc(a.azs)} ${esc(az.type||"")}</b> ${esc(a.req)}
      <div class="muted">${esc(a.works.filter(Boolean).join("; ")||a.desc||"—")}</div></div>
      <div class="acts"><button class="ico rep" data-rep title="Повторить">${ICON_REPEAT}</button><button class="ico del" data-rm title="Удалить">${ICON_TRASH}</button></div></div>`; }).join("")
    || `<p class="muted">${acts.length?"Ничего не найдено":"Сохранённые акты появятся здесь."}</p>`;
}
function bindDb(){
  $("#dbSearch").oninput=renderDb;
  $("#dbList").addEventListener("click",e=>{
    const li=e.target.closest(".li"); if(!li) return;
    const b=e.target.closest("button") || {dataset:{}};   // нажатие на сам акт — открыть его
    const a=acts.find(x=>String(x.id)===li.dataset.id); if(!a) return;
    if(b.dataset.rm!=null){ if(confirm("Удалить акт "+a.azs+" "+(a.req||"")+" из базы?")){ acts=acts.filter(x=>x!==a); saveActs(); renderDb(); } return; }
    if(act && actHasContent(act) && !acts.some(x=>x.id===act.id)) saveToDb();
    act = b.dataset.rep!=null ? repeatAct(a) : JSON.parse(JSON.stringify(a));
    saveAct(); renderAct(); go("act"); if(b.dataset.rep!=null) toast("Создан новый акт по образцу");
  });
}

/* ================= Отчёт по ГСМ =================
   Как «Шаблон отчётности … .xlsx»: маршрутные листы по неделям («мл №1…5»), «Акт ТС», «Отчёт» и фото чеков — одним PDF;
   отдельно «Маршруты» (.docx): дата поездки и снимок маршрута до АЗС. Разметка листов — LAYOUTS из шаблона (tools/build_www.py).
   Поездка — выезд на АЗС в один день (несколько заявок на одной АЗС — одна строка); пробег — туда и обратно. */
const MONTHS_NOM=["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const fmtN=(x,d=2)=>(Math.round(x*10**d)/10**d).toFixed(d).replace(".",",").replace("-0,00","0,00");
const dIso=(y,m,d)=>y+"-"+pad(m)+"-"+pad(d);
function monthNow(){ return today().slice(0,7); }
function prevMonth(month){ const [y,m]=month.split("-").map(Number); return m===1 ? (y-1)+"-12" : y+"-"+pad(m-1); }
function gsmTrips(month){
  const all=acts.slice(); if(act && actHasContent(act) && !all.some(x=>x.id===act.id)) all.push(act);
  const by=new Map();
  all.filter(a=>(a.date||"").startsWith(month) && a.azs).forEach(a=>{
    const k=a.date+"|"+a.azs, km=Number(String(a.km||azsOf(a.azs).km||"").replace(",","."))||0;
    const t=by.get(k)||{date:a.date, azs:String(a.azs), reqs:[], km:0};
    if(a.req && !t.reqs.includes(a.req)) t.reqs.push(a.req);
    t.km=Math.max(t.km, km); by.set(k,t);
  });
  return [...by.values()].sort((a,b)=>a.date.localeCompare(b.date) || a.azs.localeCompare(b.azs,"ru",{numeric:true}));
}
const tripKm=t=>Math.round(t.km*2);   // туда и обратно, целые км — как в маршрутных листах
/* недели как в листе «Отчёт»: 1-я — с 1-го числа по первое воскресенье, дальше по 7 дней, 5-я — до конца месяца */
function gsmWeeks(month){
  const [y,m]=month.split("-").map(Number), last=new Date(y,m,0).getDate();
  const sun=1+(7-new Date(y,m-1,1).getDay())%7;   // первое воскресенье
  const W=[{w:1, from:1, to:sun}];
  for(let w=2;w<=5;w++){ const from=sun+1+(w-2)*7; W.push({w, from:Math.min(from,last+1), to:w===5 ? last : Math.min(from+6,last)}); }
  return W.map(x=>Object.assign(x,{fromIso:dIso(y,m,Math.min(x.from,last)), toIso:dIso(y,m,x.to), empty:x.from>x.to}));
}
function gsmWeekOf(dateStr, W){ const d=+dateStr.slice(8,10); return (W.find(x=>d>=x.from && d<=x.to)||W[4]).w; }
const refuels=month=>(ref.refuels||[]).filter(r=>(r.date||"").startsWith(month)).sort((a,b)=>a.date.localeCompare(b.date));
function gsmStartOf(month){
  const s=(settings.gsmStart||{})[month]; if(s!=null && s!=="") return Number(String(s).replace(",","."))||0;
  const e=(settings.gsmEnd||{})[prevMonth(month)]; return e!=null ? e : 0;   // остаток на конец прошлого месяца
}
/* все расчёты отчёта за месяц */
function gsmModel(month){
  const W=gsmWeeks(month), trips=gsmTrips(month), rate=Number(String(settings.gsmRate||"").replace(",","."))||0, start=gsmStartOf(month);
  let bal=start;
  const weeks=W.map(x=>{
    const list=trips.filter(t=>gsmWeekOf(t.date,W)===x.w), fuel=refuels(month).filter(r=>gsmWeekOf(r.date,W)===x.w);
    const km=list.reduce((s,t)=>s+tripKm(t),0), liters=fuel.reduce((s,r)=>s+(Number(r.liters)||0),0);
    bal=bal-km*rate/100+liters;
    return Object.assign({}, x, {trips:list, km, liters, end:bal});
  });
  return {month, W, trips, rate, start, weeks, end:bal, total:weeks.reduce((s,x)=>s+x.km,0)};
}
function gsmHead(){
  const F=settings.fio||"";
  return ["Организация: "+(settings.gsmOrg||""), "Подразделение: "+(settings.gsmDept||""), "Сведения о собственнике ТС (ФИО):  "+F,
    "Сведения о ТС (марка, государственный номерной знак):  "+[settings.carModel,settings.carPlate].filter(Boolean).join(", "), "Сведения о водителе (ФИО): "+F].join("\n");
}
/* должность в родительном падеже: «Ведущий специалист» → «ведущего специалиста» (склоняется до первого существительного) */
function genPos(pos){
  let done=false;
  return String(pos||"").trim().toLowerCase().split(/\s+/).map(w=>{
    if(done || !w) return w;
    if(/(ый|ий|ой)$/.test(w)) return w.slice(0,-2)+(/[жшчщ]ий$/.test(w)?"его":"ого");
    done=true;
    if(/[йь]$/.test(w)) return w.slice(0,-1)+"я";
    if(/а$/.test(w)) return w.slice(0,-1)+(/[гкхжшчщ]а$/.test(w)?"и":"ы");
    return /[бвгджзклмнпрстфхцчшщ]$/.test(w) ? w+"а" : w;
  }).join(" ");
}
/* текст «Акта ТС»: организация, представитель и ФИО работника — из «Данные → Исполнитель» */
function gsmActText(){
  const org=settings.org||"", F=fioForms(settings.fio||""), male=!/(на|кызы)$/.test(F.full), rep=String(settings.gsmRep||"").trim();
  return org+", именуемое в дальнейшем «Работодатель», в лице "+(rep ? fioForms(rep).gen : "____________________")+
    ", действующего на основании доверенности: "+(String(settings.gsmProxy||"").trim()||"____________")+" , с одной стороны и работник "+org+", "+F.full+
    ", "+(male?"занимающий":"занимающая")+" должность "+genPos(settings.pos)+", "+(male?"именуемый":"именуемая")+" в дальнейшем «Работник», с другой стороны, "+
    "составили настоящий акт о том, что Работник по поручению и в интересах Работодателя использовал"+(male?"":"а")+" свой личный автотранспорт:";
}
/* подпись работника в «Акте ТС» (над чертой после «Работник», текст по центру объединения G:J) и в «Отчёте» */
SIGN_SPOTS.gsmAct=[{ref:"G23", x:()=>signMid((colsW("Акт ТС",6,10)-textW("Работник"+"_".repeat(23),"12pt Arial"))/2+textW("Работник","12pt Arial"), lineW(23,"12pt Arial"))}];
SIGN_SPOTS.gsmRep=[{ref:"A19", x:()=>signMid(2+textW(FIO().short+"  /","12pt Arial"), lineW(18,"12pt Arial"))}];
/* листы отчёта — через renderSheet, как акты */
function gsmPages(G){
  const pages=[], sign=settings.sign && settings.signOn;
  G.weeks.forEach(x=>{
    pages.push({land:true, html:renderSheet({sheet:"мл №1", key:"gsm", blocks:[{from:6, to:12,
      items:()=>{ const a=x.trips.slice(); while(a.length<7) a.push(null); return a; },
      fill:(t,i)=>t ? {A:i+1, B:{date:t.date}, C:"8:00", D:"0", E:{date:t.date}, F:"17:00", G:"0",
          H:"Место парковки ТС - АЗС "+t.azs+" - Место парковки ТС", K:t.reqs.map(q=>"SC00"+q).join(", ")||"ЗИП",
          L:tripKm(t), M:fmtN(tripKm(t)*G.rate/100), N:sign ? {sign:signMid(0, colsW("мл №1",13,14))} : ""}
        : {A:i+1, C:"8:00", D:"0", F:"17:00", G:"0", L:"0", M:"0,00"}}],
      cells:()=>({A1:"    Маршрутный лист № "+x.w, A2:"на период "+fmtDate(x.fromIso)+" - "+fmtDate(x.toIso), A3:gsmHead(),
        E14:fmtN(G.start), E15:fmtN(x.km), E16:fmtN(x.liters), E17:fmtN(x.end), M15:String(G.rate).replace(".",","), O17:""})}, null)});
  });
  const [y,m]=G.month.split("-").map(Number), last=new Date(y,m,0).getDate(), next=m===12 ? dIso(y+1,1,1) : dIso(y,m+1,1);
  pages.push({html:renderSheet({sheet:"Акт ТС", key:"gsmAct", rowspan:true, blocks:[], cells:()=>({
    A10:gsmActText(), A5:"1 "+MONTHS[m-1]+" "+y+" г.", A13:"Марка/модель: "+String(settings.carModel||"").toUpperCase(), A14:"Гос. рег. Знак:  "+(settings.carPlate||""),
    A15:"В период: "+fmtDate(dIso(y,m,1))+" - "+fmtDate(dIso(y,m,last)), G19:fmtN(G.total)})}, null)});
  const used=G.weeks.filter(x=>x.trips.length).map(x=>x.w);
  pages.push({html:renderSheet({sheet:"Отчёт", key:"gsmRep", rowspan:true, blocks:[], cells:()=>{
    const c={B3:MONTHS_NOM[m-1]+" "+y+" г.", A4:"Ф.И.О. Работника: "+(settings.fio||""), A5:"Марка автомобиля: "+(settings.carModel||""),
      A6:"Регистрационный знак:  "+(settings.carPlate||""), C14:String(G.total),
      A16:"Приложение: маршрутные листы №№ "+(used.length?used:[1]).join(","),
      A19:FIO().short+"  /__________________/   "+fmtDate(next)+"г.",
      A24:"  Руководитель СП  /________________/    "+(settings.gsmBoss||"")+"       / «____» ________ "+y+" год."};
    G.weeks.forEach((x,i)=>{ c["A"+(9+i)]=x.empty ? "" : fmtDate(x.fromIso)+" - "+fmtDate(x.toIso); c["C"+(9+i)]=x.km ? String(x.km) : ""; });
    return c; }}, null)});
  refuels(G.month).forEach(r=>{ if(r.photo) pages.push({photo:true, html:`<div class="page" style="padding:48px 60px;font-family:Arial,sans-serif">
    <div style="font-weight:700;font-size:15pt;margin-bottom:18px">${esc(fmtDate(r.date))}</div>
    <img src="${r.photo}" style="display:block;margin:0 auto;max-width:100%;max-height:960px"></div>`}); });
  return pages;
}
/* страницы → PDF: каждая рисуется картинкой на свою страницу A4 (маршрутные листы — альбомные) */
async function pagesToPdf(pages){
  if(!window.html2canvas) await loadScript("vendor/html2canvas.min.js");
  if(!window.jspdf) await loadScript("vendor/jspdf.umd.min.js");
  if(document.fonts) await document.fonts.ready;
  const box=document.createElement("div"); box.style.cssText="position:fixed;left:-20000px;top:0";
  document.body.appendChild(box);
  try{
    let pdf=null;
    for(const p of pages){
      box.innerHTML=p.html; const pg=box.firstElementChild;
      pg.style.boxShadow="none"; pg.style.transform="none";
      if(p.land){ pg.style.width="1123px"; pg.style.minHeight="794px"; }
      pg.querySelectorAll("table.xl").forEach(t=>{ const z=parseFloat(t.style.zoom)||1; t.style.zoom=""; if(z!==1){ t.style.transform="scale("+z+")"; t.style.transformOrigin="0 0"; } });
      await Promise.all([...pg.querySelectorAll("img")].map(i=>i.complete ? 0 : new Promise(r=>{ i.onload=i.onerror=r; })));
      const c=await html2canvas(pg, {scale:2, backgroundColor:"#ffffff", logging:false});
      const [W,H]=p.land ? [297,210] : [210,297];
      if(!pdf) pdf=new jspdf.jsPDF({unit:"mm", format:"a4", orientation:p.land?"landscape":"portrait", compress:true});
      else pdf.addPage("a4", p.land?"landscape":"portrait");
      const h=W*c.height/c.width, k=Math.min(1, H/h);
      pdf.addImage(c.toDataURL("image/jpeg", 0.9), "JPEG", (W-W*k)/2, 0, W*k, h*k);
    }
    return pdf.output("blob");
  } finally { box.remove(); }
}
const gsmName=(month,prefix)=>{ const [y,m]=month.split("-"); return prefix+"_"+m+"_"+y.slice(2); };
let gsmBusy=false;
async function makeGsmPdf(){
  if(gsmBusy) return; gsmBusy=true;
  try{
    const month=$("#gsmMonth").value||monthNow(), G=gsmModel(month);
    if(!G.trips.length && !confirm("За этот месяц в базе нет поездок. Сформировать пустой отчёт?")) return;
    const over=G.weeks.filter(x=>x.trips.length>7);
    if(over.length) toast("В маршрутном листе № "+over.map(x=>x.w).join(", ")+" больше 7 поездок — строки добавлены");
    if(!G.trips.every(t=>t.km)) toast("Не у всех АЗС указано расстояние — пробег 0");
    toast("Готовлю отчёт…");
    const blob=await pagesToPdf(gsmPages(G));
    settings.gsmEnd=Object.assign({}, settings.gsmEnd, {[month]:Math.round(G.end*100)/100}); store.set("settings",settings);
    const file=new File([blob], gsmName(month, String(settings.fio||"Отчёт").split(/\s+/)[0])+".pdf", {type:"application/pdf"});
    await saveBlob(file, file.name);
    if(NATIVE) toast("Сохранено: Загрузки/АЗС/"+file.name);
  } finally { gsmBusy=false; }
}
/* «Маршруты»: для каждой поездки дата и снимок маршрута её АЗС; .docx, альбомный A4 */
async function makeRoutesDocx(){
  const month=$("#gsmMonth").value||monthNow(), trips=gsmTrips(month), routes=ref.routes||{};
  if(!trips.length) return alert("За этот месяц в базе нет поездок.");
  const miss=uniq(trips.filter(t=>!routes[t.azs]).map(t=>t.azs));
  if(miss.length && !confirm("Нет снимка маршрута для АЗС "+miss.join(", ")+" — эти поездки будут без картинки. Продолжить?")) return;
  const zip=new JSZip(), media={}, X=s=>esc(s);
  let n=0, body="";
  for(const t of trips){
    body+=`<w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>${X(fmtDate(t.date))}</w:t></w:r></w:p>`;
    const src=routes[t.azs];
    if(!src){ body+=`<w:p><w:r><w:t>${X("Маршрут: место парковки ТС — АЗС "+t.azs+" — место парковки ТС")}</w:t></w:r></w:p>`; continue; }
    if(!media[t.azs]){ const img=await loadImg(await (await fetch(src)).blob()); n++;
      media[t.azs]={id:"rIdImg"+n, file:"img"+n+".jpg", w:img.naturalWidth||img.width, h:img.naturalHeight||img.height};
      zip.file("word/media/"+media[t.azs].file, src.split(",")[1], {base64:true}); }
    const M=media[t.azs], cx=7429500, cy=Math.min(6181725, Math.round(cx*M.h/M.w)), cxx=Math.round(cy*M.w/M.h);
    body+=`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cxx}" cy="${cy}"/><wp:docPr id="${body.length}" name="Маршрут АЗС ${X(t.azs)}"/>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:nvPicPr><pic:cNvPr id="0" name="${M.file}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${M.id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Object.values(media).map(M=>`<Relationship Id="${M.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${M.file}"/>`).join("")}</Relationships>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="567" w:right="567" w:bottom="567" w:left="567" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  const blob=await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document", compression:"DEFLATE"});
  const file=new File([blob], gsmName(month,"Маршруты")+".docx", {type:blob.type});
  await saveBlob(file, file.name);
  if(NATIVE) toast("Сохранено: Загрузки/АЗС/"+file.name);
}
/* фото (снимок маршрута, чек) — уменьшенная копия JPEG, чтобы данные приложения не разрастались */
async function photoData(file, maxSide, q){
  const img=await loadImg(file), k=Math.min(1, maxSide/Math.max(img.naturalWidth||img.width, img.naturalHeight||img.height));
  const c=document.createElement("canvas"); c.width=Math.round((img.naturalWidth||img.width)*k); c.height=Math.round((img.naturalHeight||img.height)*k);
  c.getContext("2d").drawImage(img,0,0,c.width,c.height); return c.toDataURL("image/jpeg", q);
}
function routeUrl(t){
  const park=settings.gsmPark||"", a=azsOf(t.azs);
  return "https://yandex.ru/maps/?rtt=auto&rtext="+encodeURIComponent([park, a.addr||("АЗС "+t.azs), park].join("~"));
}
function renderGsm(){
  const mi=$("#gsmMonth"); if(!mi.value) mi.value=monthNow();
  const month=mi.value, G=gsmModel(month), routes=ref.routes||{};
  $("#gsmStart").value = ((settings.gsmStart||{})[month]) ?? "";
  $("#gsmStart").placeholder = fmtN(gsmStartOf(month));
  $("#gsmList").innerHTML = G.trips.map(t=>`<div class="li"><div><b>${esc(fmtDate(t.date))} · АЗС ${esc(t.azs)}</b> ${esc(t.reqs.map(r=>"SC00"+r).join(", ")||"ЗИП")}
      <div class="muted">лист № ${gsmWeekOf(t.date,G.W)} · ${t.km ? tripKm(t)+" км туда и обратно" : "расстояние до АЗС не указано"} · ${routes[t.azs] ? "маршрут есть" : "нет снимка маршрута"}</div></div>
      <div class="acts"><button class="btn sm" data-route="${esc(t.azs)}" title="Снимок маршрута для этой АЗС">${routes[t.azs]?"🗺 заменить":"🗺 снимок"}</button>
      <button class="btn sm" data-map="${esc(t.azs)}" title="Открыть маршрут в Яндекс Картах">↗</button></div></div>`).join("")
    || `<p class="muted">В базе нет актов за этот месяц.</p>`;
  $("#fuelList").innerHTML = refuels(month).map(r=>`<div class="li"><div><b>${esc(fmtDate(r.date))}</b> · ${esc(fmtN(Number(r.liters)||0))} л <span class="muted">${r.photo?"чек есть":"без фото чека"}</span></div>
      <div class="acts"><button class="ico del" data-rmfuel="${r.id}" title="Удалить">${ICON_TRASH}</button></div></div>`).join("") || `<p class="muted">Заправок нет.</p>`;
  $("#gsmInfo").textContent = "Поездок: "+G.trips.length+", пробег "+G.total+" км, расход "+fmtN(G.total*G.rate/100)+" л, заправлено "+fmtN(G.weeks.reduce((s,x)=>s+x.liters,0))+" л. Остаток на конец месяца: "+fmtN(G.end)+" л.";
}
let routeFor=null;
function bindGsm(){
  $("#gsmMonth").onchange=renderGsm;
  $("#gsmStart").oninput=e=>{ const month=$("#gsmMonth").value||monthNow();
    settings.gsmStart=Object.assign({}, settings.gsmStart); const v=e.target.value.trim();
    if(v==="") delete settings.gsmStart[month]; else settings.gsmStart[month]=v; store.set("settings",settings);
    $("#gsmInfo").textContent=""; clearTimeout(bindGsm.t); bindGsm.t=setTimeout(renderGsm,600); };
  $("#gsmList").addEventListener("click",e=>{
    const b=e.target.closest("button"); if(!b) return;
    if(b.dataset.map!=null){ const u=routeUrl({azs:b.dataset.map}); NATIVE && Android.openUrl ? Android.openUrl(u) : window.open(u,"_blank"); return; }
    if(b.dataset.route!=null){ routeFor=b.dataset.route; $("#routeFile").click(); }
  });
  $("#routeFile").onchange=async e=>{ const f=e.target.files[0]; e.target.value=""; if(!f || routeFor==null) return;
    try{ ref.routes=Object.assign({}, ref.routes, {[routeFor]:await photoData(f, 1000, 0.8)}); saveRef(); renderGsm(); toast("Снимок маршрута АЗС "+routeFor+" сохранён"); }
    catch(err){ alert("Не удалось прочитать снимок: "+err.message); } };
  $("#fuelAdd").onclick=async()=>{
    const date=$("#fuelDate").value, liters=Number(String($("#fuelLiters").value).replace(",","."));
    if(!date || !liters) return toast("Укажите дату и литры");
    const f=$("#fuelPhoto").files[0];
    let photo=""; try{ if(f) photo=await photoData(f, 1400, 0.7); }catch(err){ return alert("Не удалось прочитать фото чека: "+err.message); }
    ref.refuels=(ref.refuels||[]).concat({id:Date.now(), date, liters, photo}); saveRef();
    $("#fuelLiters").value=""; $("#fuelPhoto").value=""; $("#gsmMonth").value=date.slice(0,7); renderGsm(); toast("Заправка добавлена");
  };
  $("#fuelList").addEventListener("click",e=>{ const b=e.target.closest("[data-rmfuel]"); if(!b) return;
    if(confirm("Удалить заправку?")){ ref.refuels=(ref.refuels||[]).filter(r=>String(r.id)!==b.dataset.rmfuel); saveRef(); renderGsm(); } });
  $("#gsmPdf").onclick=()=>makeGsmPdf().catch(e=>{ console.error(e); alert("Не удалось сформировать отчёт: "+e.message); });
  $("#gsmDocx").onclick=()=>makeRoutesDocx().catch(e=>{ console.error(e); alert("Не удалось сформировать маршруты: "+e.message); });
}

/* ================= Данные ================= */
function listEditor(key){
  const box=$("#lst-"+key);
  box.innerHTML=ref[key].map((t,i)=>`<div class="li"><div>${esc(t)}</div><button class="btn icon del" data-rmi="${i}" data-key="${key}">✕</button></div>`).join("")||`<p class="muted">Пусто</p>`;
}
function renderEquipBase(){
  const sel=$("#eqbAzs"); const cur=sel.value||act.azs;
  sel.innerHTML=`<option value="*">Все АЗС</option>`+uniq(ref.azs.map(a=>String(a.num)).concat(ref.equipBase.map(e=>String(e.azs)))).map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join("");
  sel.value=cur;
  const rows=ref.equipBase.map((e,i)=>({e,i})).filter(o=>sel.value==="*"||String(o.e.azs)===sel.value);
  $("#eqbList").innerHTML=rows.map(({e,i})=>`<div class="li"><div><b>${esc(e.model)}</b> s/n:${esc(e.sn)}<div class="muted">АЗС ${esc(e.azs)}${e.date?" · "+fmtDate(e.date):""}${e.req?" · "+esc(e.req):""}</div></div><button class="btn icon del" data-rmeq="${i}">✕</button></div>`).join("")||`<p class="muted">Нет записей</p>`;
}
function renderFio(){
  const f=FIO();
  $("#fioForms").textContent = f.full ? "В подписях: "+f.short+"\nВ акте ВР: в лице "+f.gen : "";
  $("#s-declSur").checked = settings.declineSurname!==false;
}
function renderSeals(){
  $("#sealCount").textContent=ref.seals.length;
  $("#sealSort").textContent = settings.sealDesc ? "↓ по убыванию" : "↑ по возрастанию";
  $("#sealList").innerHTML=sortSeals(ref.seals).map(x=>`<span class="chip">${esc(x)}<button class="x" data-rmseal="${esc(x)}" title="Убрать">✕</button></span>`).join("")||`<span class="muted">Склад пуст</span>`;
}
function renderAzsList(){
  $("#azsList").innerHTML=ref.azs.map((a,i)=>`<div class="azsrow"><div><b>${esc(a.num)}</b> ${esc(a.type)} <span class="muted">${esc(a.addr)}</span></div>
    <div class="top"><input style="width:76px;flex:0 0 auto" type="number" step="0.1" inputmode="decimal" placeholder="км" data-km="${i}" value="${esc(a.km)}">
    <input type="email" multiple inputmode="email" autocapitalize="off" placeholder="почта АЗС" data-email="${i}" value="${esc(a.email)}"></div></div>`).join("");
}
function renderData(){
  ["fio","pos","dept","org","cust","mailTo","replyTo","carModel","carPlate","gsmRate","gsmOrg","gsmDept","gsmBoss","gsmRep","gsmProxy","gsmPark"].forEach(k=>$("#s-"+k).value=settings[k]||"");
  const ct=settings.contracts||{};
  $("#contracts").innerHTML=FORMS.map(f=>`<div class="top" style="margin-top:6px"><b style="flex:0 0 58px">${esc(f.title)}</b>
    <input data-cnum="${f.key}" placeholder="№ договора" value="${esc((ct[f.key]||{}).num||"")}">
    <input type="date" data-cdate="${f.key}" style="flex:0 0 150px" value="${esc((ct[f.key]||{}).date||"")}"></div>`).join("");
  renderSign();
  renderFio();
  $("#s-autoSeal").checked=!!settings.autoSeal;
  renderSeals();
  ["works","models","places"].forEach(listEditor);
  renderEquipBase(); renderAzsList(); renderMiss(); renderGsm();
  $("#refCount").textContent="Сейчас: АЗС "+ref.azs.length+", касс "+ref.kassy.length;
}
function bindData(){
  bindMiss(); bindGsm();
  $("#contracts").addEventListener("input",e=>{ const k=e.target.dataset.cnum||e.target.dataset.cdate; if(!k) return;
    const ct=settings.contracts=Object.assign({}, settings.contracts); ct[k]=Object.assign({num:"",date:""}, ct[k]);
    ct[k][e.target.dataset.cnum?"num":"date"]=e.target.value.trim(); store.set("settings",settings); });
  ["fio","pos","dept","org","cust","mailTo","replyTo","carModel","carPlate","gsmRate","gsmOrg","gsmDept","gsmBoss","gsmRep","gsmProxy","gsmPark"].forEach(k=>$("#s-"+k).addEventListener("input",e=>{settings[k]=e.target.value; store.set("settings",settings); if(k==="fio") renderFio();}));
  $("#s-autoSeal").onchange=e=>{settings.autoSeal=e.target.checked; store.set("settings",settings);};
  const addSeals=()=>{ const inp=$("#sealAdd"), add=parseSeals(inp.value); if(!add.length) return;
    ref.seals=uniq(ref.seals.concat(add)); saveRef(); inp.value=""; renderSeals(); renderEquip(); toast("Добавлено пломб: "+add.length); };
  $("#signFile").onchange=async e=>{ const f=e.target.files[0]; e.target.value=""; if(!f) return;
    try{ const r=await processSignature(f); settings.sign=r.url; settings.signRatio=r.ratio; settings.signOn=true; store.set("settings",settings); renderSign(); toast("Подпись сохранена"); }
    catch(err){ alert("Подпись: "+err.message); } };
  $("#signDel").onclick=()=>{ if(!confirm("Удалить сохранённую подпись?")) return; settings.sign=""; settings.signRatio=0; store.set("settings",settings); renderSign(); };
  $("#s-signOn").onchange=e=>{ settings.signOn=e.target.checked; store.set("settings",settings); };
  $("#s-declSur").onchange=e=>{ settings.declineSurname=e.target.checked; store.set("settings",settings); renderFio(); };
  $("#sealAddBtn").onclick=addSeals;
  $("#sealSort").onclick=e=>{ e.preventDefault(); settings.sealDesc=!settings.sealDesc; store.set("settings",settings); renderSeals(); renderEquip(); };
  $("#sealAdd").addEventListener("keydown",e=>{ if(e.key==="Enter"){ e.preventDefault(); addSeals(); } });
  $("#sealList").addEventListener("click",e=>{ const b=e.target.closest("[data-rmseal]"); if(!b) return;
    ref.seals=ref.seals.filter(x=>x!==b.dataset.rmseal); saveRef(); renderSeals(); renderEquip(); });
  $("#tab-data").addEventListener("click",e=>{
    const b=e.target.closest("button"); if(!b) return;
    if(b.dataset.rmi!=null){ ref[b.dataset.key].splice(+b.dataset.rmi,1); saveRef(); listEditor(b.dataset.key); }
    if(b.dataset.addto){ const k=b.dataset.addto, inp=$("#add-"+k), v=inp.value.trim(); if(v && !ref[k].includes(v)){ ref[k].push(v); saveRef(); listEditor(k); } inp.value=""; }
    if(b.dataset.rmeq!=null){ ref.equipBase.splice(+b.dataset.rmeq,1); saveRef(); renderEquipBase(); }
  });
  $("#eqbAzs").onchange=renderEquipBase;
  $("#azsList").addEventListener("change",e=>{
    const em=e.target.dataset.email; if(em!=null){ ref.azs[+em].email=e.target.value.trim(); saveRef(); renderAzsInfo(); return; }
    const i=e.target.dataset.km; if(i==null) return; ref.azs[+i].km=e.target.value; saveRef();
    if(String(ref.azs[+i].num)===String(act.azs)){ act.km=e.target.value; saveAct(); } });
  $("#expXlsx").onclick=exportRef;
  $("#impXlsx").onchange=async e=>{ const f=e.target.files[0]; e.target.value=""; if(!f) return; try{ await importXlsx(f); }catch(err){ console.error(err); alert("Не удалось прочитать файл: "+err.message); } };
  $("#bkSave").onclick=async()=>{ try{ await saveBlob(new Blob([JSON.stringify({ref,settings,acts},null,1)],{type:"application/json"}), "akty-backup-"+today()+".json"); toast(NATIVE?"Копия сохранена в «Загрузки/АЗС»":"Копия сохранена"); }catch(err){ alert(err.message); } };
  $("#bkLoad").onchange=async e=>{ const f=e.target.files[0]; e.target.value=""; if(!f) return;
    try{ const d=JSON.parse(await f.text()); if(!d.ref) throw 0;
      ref=Object.assign(defaultRef(), d.ref); settings=Object.assign({},DEFAULT_SETTINGS,d.settings||{}); acts=d.acts||[];
      saveRef(); store.set("settings",settings); saveActs(); renderData(); renderAct(); toast("Восстановлено"); }catch(err){ alert("Неверный файл копии"); } };
}

/* ---------- загрузка справочников из Excel: простой «Справочник АЗС.xlsx» или старый !!!АКТ.xlsx ---------- */
const NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
function splitRef(r){ const m=r.match(/^([A-Z]+)(\d+)$/); return [m[1], +m[2]]; }
async function importXlsx(file){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const ss=[]; const ssf=zip.file("xl/sharedStrings.xml");
  if(ssf){ const d=new DOMParser().parseFromString(await ssf.async("string"),"application/xml");
    for(const si of d.getElementsByTagNameNS(NS,"si")) ss.push(Array.from(si.getElementsByTagNameNS(NS,"t")).map(t=>t.textContent).join("")); }
  const wb=await zip.file("xl/workbook.xml").async("string"), rels=await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const pathOf=name=>{ const m=wb.match(new RegExp('<sheet\\b[^>]*name="'+name+'"[^>]*r:id="([^"]+)"')); if(!m) return null;
    const t=rels.match(new RegExp('Id="'+m[1]+'"[^>]*Target="([^"]+)"'))||rels.match(new RegExp('Target="([^"]+)"[^>]*Id="'+m[1]+'"')); return "xl/"+t[1].replace(/^\/?xl\//,""); };
  async function read(name){
    const p=pathOf(name); if(!p||!zip.file(p)) return null;
    const d=new DOMParser().parseFromString(await zip.file(p).async("string"),"application/xml"); const out={};
    for(const c of d.getElementsByTagNameNS(NS,"c")){
      const t=c.getAttribute("t"); const v=c.getElementsByTagNameNS(NS,"v")[0]; let val="";
      if(t==="s" && v) val=ss[+v.textContent]||""; else if(t==="inlineStr") val=Array.from(c.getElementsByTagNameNS(NS,"t")).map(x=>x.textContent).join("");
      else if(v){ val=v.textContent; if(/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(val)){ const n=Number(val); val=Number.isInteger(n)?String(n):val; } }
      val=String(val).replace(/ /g," ").trim(); if(val) out[c.getAttribute("r")]=val;
    }
    return out;
  }
  const maxRow=o=>Math.max(0,...Object.keys(o).map(r=>splitRef(r)[1]));
  if(await read("АЗС")) return importSimple(read, maxRow);
  const A=await read("AZS"), K=await read("Кассы"), S=await read("Справочник");
  if(!A||!K) throw new Error("не найден лист «АЗС» (простой справочник) или листы AZS и Кассы (!!!АКТ.xlsx)");
  const azs=[]; for(let r=2;r<=maxRow(A);r++){ if(!A["B"+r]) continue;
    const g=c=>A[c+r]||""; const num=g("B");
    azs.push({ceh:g("A"),num,type:g("C"),addr:g("D"),boss:g("E"),phone:g("F"),bossPhone:g("G"),shop:g("H"),db:g("I"),ipTH:g("J"),gw:g("K"),dns:g("L"),ipFuel:g("M"), km:azsOf(num).km||"", email:azsOf(num).email||""}); }
  const kassy=[]; for(let r=3;r<=maxRow(K);r++){ if(!K["B"+r]) continue;
    const g=c=>K[c+r]||""; kassy.push({azs:g("B"),db:g("C"),n:g("D"),posIp:g("E"),rn:g("H"),zn:g("I"),kktIp:g("J"),kktGw:g("K"),fn:g("L"),model:g("M"),sn:g("N"),tid:g("O"),termIp:g("P"),termGw:g("Q")}); }
  ref.azs=azs; ref.kassy=kassy;
  if(S){
    for(let r=1;r<=80;r++){ const k=S["K"+r]; if(!k || /^Произведена/.test(k)) continue;
      const {model:md, sn}=splitModel(k); if(md && !ref.models.includes(md)) ref.models.push(md);
      if(sn && !ref.equipBase.some(e=>e.sn===sn)) ref.equipBase.push({azs:S["L"+r]||"", model:md, sn, date:"", req:""}); }
    for(let r=1;r<=30;r++){ const p=S["U"+r]; if(p && !ref.places.includes(p)) ref.places.push(p); }
  }
  saveRef(); toast("Загружено: АЗС "+azs.length+", касс "+kassy.length); renderData(); renderAct();
}

/* Простой формат: на каждом листе первая строка — заголовки, столбцы ищутся по названию (порядок любой). */
const SIMPLE = {
  "АЗС":          { num:["№ азс","номер азс","азс"], type:["тип"], addr:["адрес"], boss:["управляющ"], phone:["телефон азс","раб. телефон","телефон"], bossPhone:["телефон управляющ"], km:["км","километр","расстоян"], email:["почт","e-mail","email"] },
  "Кассы":        { azs:["№ азс","азс"], n:["№ кассы","касса"], zn:["зн"], rn:["рн"], model:["модель"], sn:["s/n","серийн"], tid:["id"] },
  "Оборудование": { model:["модель","наименование"], sn:["серийн","s/n"], azs:["азс","место"] },
  "Склад пломб":  { seal:["пломб","номер"] },
  "Работы":       { work:["работ","наименование"] },
  "Места пломб":  { place:["место"] }
};
async function importSimple(read, maxRow){
  async function table(name){
    const S=await read(name); if(!S) return null;
    const spec=SIMPLE[name], cols={}, used=new Set();
    // заголовки строки 1; более длинные варианты проверяются первыми («телефон управляющего» раньше «телефон»)
    const heads=Object.keys(S).filter(r=>splitRef(r)[1]===1).map(r=>({col:splitRef(r)[0], h:S[r].toLowerCase()}));
    const pairs=[]; Object.entries(spec).forEach(([k,al])=>al.forEach(a=>pairs.push([k,a])));
    pairs.sort((x,y)=>y[1].length-x[1].length);
    for(const [k,a] of pairs){ if(cols[k]) continue; const h=heads.find(x=>!used.has(x.col) && x.h.includes(a)); if(h){ cols[k]=h.col; used.add(h.col); } }
    const rows=[]; for(let r=2;r<=maxRow(S);r++){ const o={}; let any=false;
      for(const k in spec){ o[k]=cols[k] ? (S[cols[k]+r]||"") : ""; if(o[k]) any=true; }
      if(any) rows.push(o); }
    return rows;
  }
  const done=[];
  const A=await table("АЗС");
  if(A){ const old=ref.azs; ref.azs=A.filter(a=>a.num).map(a=>{ const p=old.find(x=>String(x.num)===String(a.num))||{};
      return Object.assign({}, p, a, {km: a.km || p.km || "", email: a.email || p.email || ""}); }); done.push("АЗС "+ref.azs.length); }
  const K=await table("Кассы");
  if(K){ ref.kassy=K.filter(k=>k.azs&&k.n); done.push("касс "+ref.kassy.length); }
  const E=await table("Оборудование");
  if(E){ let n=0; E.forEach(e=>{ const md=e.model.trim(); if(!md) return; if(!ref.models.includes(md)) ref.models.push(md);
      if(e.sn && !ref.equipBase.some(x=>x.sn===e.sn)){ ref.equipBase.push({azs:e.azs, model:md, sn:e.sn, date:"", req:""}); n++; } });
    done.push("оборудования "+E.length); }
  const P=await table("Склад пломб");
  if(P){ const add=uniq(P.flatMap(p=>parseSeals(p.seal))).filter(s=>!ref.seals.includes(s)); ref.seals=ref.seals.concat(add); done.push("новых пломб "+add.length); }
  const W=await table("Работы");
  if(W){ W.forEach(w=>{ const t=w.work.trim(); if(t && !ref.works.some(x=>x.trim()===t)) ref.works.push(t); }); done.push("работ "+W.length); }
  const M=await table("Места пломб");
  if(M){ M.forEach(m=>{ const t=m.place.trim(); if(t && !ref.places.includes(t)) ref.places.push(t); }); }
  saveRef(); toast("Загружено: "+done.join(", ")); renderData(); renderAct();
}

/* Справочник в формате «Справочник АЗС.xlsx» из текущих данных приложения (для правки и обратной загрузки) */
const REF_HELP = [
  ["Справочник АЗС для приложения «Акты АЗС»","title"], ["",""],
  ["Как пользоваться","head"],
  ["1. Заполните или исправьте листы. Первую строку (заголовки) не меняйте — по ним приложение находит столбцы.",""],
  ["2. Сохраните файл и перекиньте его на телефон.",""],
  ["3. В приложении: Данные → Загрузка из Excel → выберите этот файл.",""], ["",""],
  ["Что делает загрузка","head"],
  ["АЗС, Кассы — список полностью заменяется данными из файла (удалённые строки пропадут из приложения).",""],
  ["Оборудование, Склад пломб, Работы, Места пломб — новые строки добавляются к уже имеющимся; ничего не удаляется.",""],
  ["База актов не меняется. Если «Км в один конец» или «Почта» не указаны — в приложении сохраняются прежние значения.",""],
  ["Лишний или пустой лист можно удалить — он просто не загрузится.",""], ["",""],
  ["Листы","head"],
  ["АЗС — номер, тип, адрес, контакты, расстояние в один конец (для акта ККТ) и почта для отправки актов (несколько адресов — через запятую).",""],
  ["Кассы — по одной строке на кассу: ЗН и РН ККТ, терминал (модель, S/N, ID). Берутся в акты по месту «ККТ N» / «ВБРР N».",""],
  ["Оборудование — модели для выбора в расходном материале; с серийным номером запись попадёт в «Оборудование на АЗС».",""],
  ["Склад пломб — номер пломбы в строке; можно диапазон: 110300-110320.",""],
  ["Работы — готовые наименования работ; {обор} заменится выбранным оборудованием.",""],
  ["Места пломб — места установки пломб (ККТ 1, ВБРР 2, Топаз…).",""]
];
function refSheets(){
  const withSn=new Set(ref.equipBase.map(e=>e.model));
  return [
    {name:"Инструкция", plain:true, widths:[120], rows:REF_HELP},
    {name:"АЗС", headers:["№ АЗС","Тип","Адрес","Управляющий","Телефон АЗС","Телефон управляющего","Км в один конец","Почта"], widths:[9,8,55,32,18,20,12,30], wrap:[2],
     rows:ref.azs.map(a=>[a.num,a.type,a.addr,a.boss,a.phone,a.bossPhone,a.km,a.email])},
    {name:"Кассы", headers:["№ АЗС","№ кассы","ЗН ККТ","РН ККТ","Модель терминала","S/N терминала","ID терминала"], widths:[9,9,17,17,18,28,13],
     rows:ref.kassy.map(k=>[k.azs,k.n,k.zn,k.rn,k.model,k.sn,k.tid])},
    {name:"Оборудование", headers:["Модель","Серийный номер","№ АЗС / место"], widths:[42,30,16],
     rows:ref.equipBase.map(e=>[e.model,e.sn,e.azs]).concat(ref.models.filter(m=>!withSn.has(m)).map(m=>[m,"",""]))},
    {name:"Склад пломб", headers:["Номер пломбы"], widths:[20], rows:sortSeals(ref.seals).map(s=>[s])},
    {name:"Работы", headers:["Наименование работы"], widths:[80], wrap:[0], rows:ref.works.map(w=>[w])},
    {name:"Места пломб", headers:["Место установки пломбы"], widths:[28], rows:ref.places.map(p=>[p])}
  ];
}
async function exportRef(){
  try{
    const name="Справочник АЗС "+fmtDate(today())+".xlsx";
    await saveBlob(await buildXlsx(refSheets()), name);
    toast(NATIVE ? "Сохранено: Загрузки/АЗС/"+name : "Справочник сохранён");
  }catch(e){ console.error(e); alert("Не удалось сохранить справочник: "+e.message); }
}

/* ---------- файлы: в APK через мост window.Android ---------- */
const NATIVE = typeof window.Android !== "undefined";
function blobToB64(blob){ return new Promise((ok,fail)=>{ const r=new FileReader(); r.onload=()=>ok(String(r.result).split(",")[1]); r.onerror=fail; r.readAsDataURL(blob); }); }
async function saveBlob(blob, name){
  if(NATIVE){ const res=Android.saveFile(name, await blobToB64(blob), blob.type||"application/octet-stream"); if(res) throw new Error(res); return; }
  const file=new File([blob], name, {type:blob.type||"application/octet-stream"});
  if(canShareFile(file)) return shareFile(file);   // iPhone: «Сохранить в Файлы» или отправить
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},4000);
}

/* ================= Обновление по сети (Яндекс.Диск / Google Диск) ================= */
const APK_RE=/AktyAZS_v(\d+(?:\.\d+)*)\.apk$/i;
function cmpVer(a,b){ const x=String(a).split(".").map(Number), y=String(b).split(".").map(Number);
  for(let i=0;i<Math.max(x.length,y.length);i++){ const d=(x[i]||0)-(y[i]||0); if(d) return d; } return 0; }
const httpWait={};
window.onHttp=(id,err,text)=>{ const p=httpWait[id]; if(!p) return; delete httpWait[id]; err?p.fail(new Error(err)):p.ok(text); };
function httpGet(url){
  if(NATIVE && Android.httpGet) return new Promise((ok,fail)=>{ const id="h"+Date.now()+Math.random(); httpWait[id]={ok,fail}; Android.httpGet(id,url); });
  return fetch(url).then(r=>{ if(!r.ok) throw new Error("сервер ответил "+r.status); return r.text(); });
}
/* confirm=t — без страницы «Google не может проверить файл на вирусы», которую Диск показывает для APK */
const GDRIVE_DL="https://drive.usercontent.google.com/download?export=download&confirm=t&id=";
const srcOf=(link, fallback)=> /yandex|yadi\.sk/i.test(link) ? "yandex" : /google\.com/i.test(link) ? "google" : fallback;
/* список APK в публичной папке: [{name, version, url}] */
async function listRemote(src, link){
  if(src==="google"){
    const folder=link.match(/folders\/([\w-]+)/) || link.match(/[?&]id=([\w-]+)/);
    const file=link.match(/\/file\/d\/([\w-]+)/);
    if(file){ const html=await httpGet("https://drive.google.com/file/d/"+file[1]+"/view");
      const name=((html.match(/<title>([^<]+?)(?: - Google[^<]*)?<\/title>/)||[])[1]||"").trim();
      return [{name, url:GDRIVE_DL+file[1]}]; }
    if(!folder) throw new Error("не похоже на ссылку Google Диска на папку или файл");
    const html=await httpGet("https://drive.google.com/embeddedfolderview?id="+folder[1]);
    return [...html.matchAll(/id="entry-([\w-]+)"[\s\S]*?class="flip-entry-title">([^<]+)</g)]
      .map(m=>({name:m[2].trim(), url:GDRIVE_DL+m[1]}));
  }
  const j=JSON.parse(await httpGet("https://cloud-api.yandex.net/v1/disk/public/resources?limit=500&public_key="+encodeURIComponent(link)));
  if(j.error) throw new Error(j.description||j.error);
  const items = j.type==="file" ? [j] : ((j._embedded||{}).items||[]);
  return items.filter(i=>i.type==="file").map(i=>({name:i.name, url:i.file}));
}
function updSources(){
  return [[DEFAULT_UPD_LINK, "основной"], [RESERVE_UPD_LINK, "резервный"]]
    .filter(([l])=>l).map(([l,n])=>({src:srcOf(l, "yandex"), link:l, title:n}));
}
let updQueue=[], updState="";   // updState — последняя строка статуса, для окна «О приложении»   // ссылки на найденную новую версию из разных источников — на случай сбоя скачивания
async function checkUpdate(manual){
  if(!NATIVE) return checkWebUpdate(manual);
  const sources=updSources();
  setUpd("Проверяю…");
  try{
    const res=await Promise.allSettled(sources.map(s=>listRemote(s.src, s.link)));
    const ok=res.filter(r=>r.status==="fulfilled");
    if(!ok.length) throw res[0].reason;
    const failed=sources.filter((s,i)=>res[i].status==="rejected").map(s=>s.title+" источник недоступен ("+res[sources.indexOf(s)].reason.message+")");
    const apks=ok.flatMap(r=>r.value).map(f=>{ const m=f.name.match(APK_RE); return m && Object.assign(f,{version:m[1]}); }).filter(Boolean)
      .sort((a,b)=>cmpVer(b.version,a.version));
    if(!apks.length) throw new Error("в папке нет файлов AktyAZS_vX.Y.apk");
    const top=apks[0];
    const note=failed.length ? " · "+failed.join("; ") : "";
    if(cmpVer(top.version, APP_VERSION)<=0){ setUpd("Установлена последняя версия ("+APP_VERSION+")"+note); if(manual) toast("Обновлений нет"); return; }
    setUpd("Доступна версия "+top.version+note);
    if(!confirm("Доступна версия "+top.version+" (установлена "+APP_VERSION+").\nСкачать и установить?\n\nДанные и база актов сохранятся.")) return;
    updQueue=apks.filter(a=>a.version===top.version);
    if(!NATIVE || !Android.downloadApk){ window.open(top.url); return; }
    downloadNext();
  }catch(e){ console.error(e); setUpd(manual ? "Ошибка: "+e.message : ""); if(manual) alert("Проверка обновлений: "+e.message); }
}
/* веб-версия (iPhone): файлы приложения хранит service worker (sw.js), без сети всё работает из кэша.
   Номер новой версии берётся из version.js на сервере; после согласия новый sw.js загружает файлы и страница перезапускается */
let swReg=null, swReload=false;
if(!NATIVE && "serviceWorker" in navigator && (location.protocol==="https:" || location.hostname==="localhost")){
  navigator.serviceWorker.register("sw.js").then(r=>{ swReg=r; }).catch(e=>console.warn("sw:", e));
  navigator.serviceWorker.addEventListener("controllerchange",()=>{ if(swReload) location.reload(); });
  if(navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(()=>{});
}
function waitSw(reg){
  return new Promise((ok,fail)=>{
    const watch=w=>{ if(w.state==="installed"||w.state==="activated") return ok(w);
      if(w.state==="redundant") return fail(new Error("файлы новой версии не загрузились"));
      w.addEventListener("statechange",()=>watch(w),{once:true}); };
    if(reg.waiting) return ok(reg.waiting);
    if(reg.installing) return watch(reg.installing);
    reg.addEventListener("updatefound",()=>watch(reg.installing),{once:true});
    setTimeout(()=>fail(new Error("сервер не ответил")), 180000);
  });
}
async function checkWebUpdate(manual){
  if(!swReg){ if(manual){ setUpd("Обновление доступно только в установленной веб-версии"); alert("Проверка обновлений работает в приложении, открытом с сайта (https)."); } return; }
  setUpd("Проверяю…");
  try{
    const txt=await fetch("version.js?t="+Date.now(), {cache:"no-store"}).then(r=>{ if(!r.ok) throw new Error("сервер ответил "+r.status); return r.text(); });
    const v=(txt.match(/APP_VERSION\s*=\s*["']([^"']+)/)||[])[1];
    if(!v) throw new Error("на сайте нет номера версии");
    if(cmpVer(v, APP_VERSION)<=0){ setUpd("Установлена последняя версия ("+APP_VERSION+")"); if(manual) toast("Обновлений нет"); return; }
    setUpd("Доступна версия "+v);
    if(!confirm("Доступна версия "+v+" (установлена "+APP_VERSION+").\nОбновить сейчас?\n\nДанные и база актов сохранятся.")) return;
    setUpd("Загружаю версию "+v+"…"); toast("Загружаю версию "+v+"…");
    const wait=waitSw(swReg);
    await swReg.update();
    const w=await wait;
    swReload=true;
    if(w.state==="activated") location.reload(); else w.postMessage("skipWaiting");
  }catch(e){ console.error(e); setUpd(manual ? "Ошибка: "+e.message : ""); if(manual) alert("Проверка обновлений: "+e.message); }
}
/* автопроверка без сети или с ошибкой молчит: статус показывается только в «О приложении» */
function setUpd(t){ updState=t; const st=$("#updStatus"); if(st) st.textContent=t; }
function downloadNext(){ const f=updQueue.shift(); if(!f) return false; toast("Скачиваю "+f.name+"…"); Android.downloadApk(f.url, f.name); return true; }
window.onUpdProgress=p=>{ if(p>=0) setUpd("Скачивание… "+p+"%"); };
window.onUpdDone=err=>{
  if(!err){ setUpd("Открыта установка новой версии"); return; }
  if(!/разрешите установку/.test(err) && updQueue.length){ setUpd("Ошибка ("+err+"), пробую резервный источник…"); downloadNext(); return; }
  setUpd("Ошибка: "+err); alert("Обновление: "+err);
};

/* ================= О приложении ================= */
function showAbout(){
  const vers=(window.CHANGELOG||[]).slice(0,3).map(v=>`<div class="ver"><b>Версия ${esc(v.v)}</b> <span class="muted">${esc(v.date)}</span>
    <ul>${v.items.map(i=>`<li>${esc(i)}</li>`).join("")}</ul></div>`).join("");
  const bg=sheet("О приложении", `<div class="bd"><div class="pad about">
    <h3>Акты АЗС <span class="muted">v${esc(APP_VERSION)}</span></h3>
    <p>Заполнение актов по выездам на АЗС: <b>ВР</b> (сдачи-приёмки работ), <b>Пломбы</b> (снятия/установки пломб), <b>ВБРР</b> (об оказании услуг) и <b>ККТ</b> (выезда на АЗК). Акты выглядят как листы Excel-шаблона.</p>
    <ul>
      <li>заявка, оборудование, пломбы, работы и материалы — в одной форме;</li>
      <li>пломбы и серийные номера — вручную, по штрихкоду/QR или распознаванием текста с фото (без интернета);</li>
      <li>печать, сохранение в PDF${NATIVE?" («Загрузки/АЗС/Акты»)":""} и отправка на почту АЗС;</li>
      <li>база актов с поиском и повтором, справочники АЗС, касс, оборудования, работ и склад пломб;</li>
      <li>загрузка и выгрузка справочника в Excel, подпись исполнителя по фото, автокопия данных;</li>
      <li>отчёт по ГСМ за месяц: маршрутные листы, акт использования ТС, отчёт и чеки (PDF), снимки маршрутов (DOCX).</li>
    </ul>
    <details class="guide"><summary>Инструкция</summary>
      <h4>Перед началом</h4>
      <ol>
        <li>«Данные → Исполнитель»: ФИО, должность, подразделение, организация, заказчик, номера договоров, фото подписи.</li>
        <li>Там же — <b>«Кому отправлять акты»</b> (адреса в поле «Кому» каждого письма, вместе с почтой АЗС) и <b>«Адреса для ответа на письмо»</b> (куда АЗС вернёт подписанные акты). Несколько адресов — через запятую.</li>
        <li>Справочник АЗС, касс и оборудования: «Данные → Справочник Excel → Загрузить .xlsx».${NATIVE?"":" В веб-версии справочник изначально пуст — загрузите его или восстановите резервную копию с Android."}</li>
        <li>Почту и расстояние до каждой АЗС — «Данные → АЗС: расстояния и почта».</li>
      </ol>
      <h4>Акт</h4>
      <ol>
        <li>«Новый» в шапке начинает новый акт (заполненный сохраняется в базу).</li>
        <li>Вкладка «Акт»: АЗС, дата, номер заявки, описание (▦ — считать штрихкод/QR заявки), время прибытия и убытия, километраж.</li>
        <li>«Оборудование и пломбы»: место установки, снятые и установленные пломбы и серийные номера — вручную или кнопкой 📷 (фото пломбы, наклейки, штрихкода или QR распознаётся без интернета).</li>
        <li>Работы и расходный материал — из списков или вручную.</li>
      </ol>
      <h4>Печать и отправка</h4>
      <ol>
        <li>Вкладка «Печать»: отметьте нужные акты — ВР, Пломбы, ВБРР, ККТ.</li>
        <li>«Печать» — на принтер; «Сохранить PDF» — ${NATIVE?"в «Загрузки/АЗС/Акты»":"через меню «Поделиться» (например, «Сохранить в Файлы»)"}.</li>
        <li>«✉ На почту» — письмо с PDF: «Кому» — почта АЗС и адреса из настроек, тема — имя файла, в тексте просьба подписать и ссылка для ответа.${NATIVE?"":" На iPhone выберите «Почта» в меню «Поделиться» и вставьте скопированный адрес в «Кому»."}</li>
      </ol>
      <h4>Отчёт по ГСМ</h4>
      <ol>
        <li>Один раз: «Данные → Исполнитель → Для отчёта по ГСМ» — ФИО начальника, работодатель в лице и доверенность, автомобиль, госномер, норма расхода, место парковки (ФИО, должность и организация — из полей исполнителя); расстояние до каждой АЗС в один конец — «Данные → АЗС».</li>
        <li>«Данные → Отчёт по ГСМ»: выберите месяц — поездки берутся из актов. Для каждой АЗС один раз: ↗ откроет маршрут в Яндекс Картах, сделайте снимок экрана и загрузите его кнопкой 🗺.</li>
        <li>Заправки по топливной карте: дата, литры, фото чека. Остаток на начало месяца переносится с прошлого отчёта, его можно исправить.</li>
        <li>«Отчёт PDF» — маршрутные листы, акт использования ТС, отчёт и чеки с вашей подписью; «Маршруты DOCX» — даты и снимки маршрутов.</li>
      </ol>
      <h4>База и данные</h4>
      <ol>
        <li>«База» — все акты с поиском; акт можно открыть, повторить для новой заявки или удалить.</li>
        <li>«Данные» — справочники (склад пломб, работы, модели, оборудование на АЗС, места установки); разделы раскрываются нажатием на заголовок.</li>
        <li>«Нераспознанные фото» — снимки, с которых не удалось считать номер или код, сохраняются сами (последние ${MISS_MAX}); «Выгрузить ZIP» собирает их в архив.</li>
        <li>«Резервная копия → Сохранить» — все данные и база актов одним файлом; «Восстановить» — на этом или другом телефоне.</li>
      </ol>
      <h4>Обновления</h4>
      <ol>
        <li>Новая версия проверяется при каждом запуске (нужен интернет); перед установкой приложение спросит. Данные сохраняются.</li>
        ${NATIVE?"":"<li>Установка на iPhone: открыть сайт в Safari → «Поделиться» → «На экран „Домой“».</li>"}
      </ol>
    </details>
    <div class="btns"><button class="btn pri" id="updCheck">Проверить обновления</button></div>
    <div class="hint" id="updStatus">${esc(updState)}</div>
    <h4>Последние версии</h4>${vers||'<p class="muted">—</p>'}
    <p class="author">Автор: <b>B@m$ik</b></p>
  </div></div>`);
  bg.querySelector("#updCheck").onclick=()=>checkUpdate(true);
}

/* ================= Навигация ================= */
function go(tab){
  ["act","print","db","data"].forEach(t=>$("#tab-"+t).classList.toggle("hidden",t!==tab));
  document.querySelectorAll("nav button").forEach(b=>b.classList.toggle("on",b.dataset.go===tab));
  window.scrollTo(0,0);
  if(tab==="act") renderEquip(); if(tab==="print") renderPrint(); if(tab==="data") renderData(); if(tab==="db") renderDb();
}
document.querySelectorAll("[data-go]").forEach(b=>b.addEventListener("click",()=>go(b.dataset.go)));
/* кнопка «Назад» Android (MainActivity.onBackPressed): закрывает открытое окно, с других вкладок ведёт на «Акт»;
   false — уже на «Акте», приложение спросит о выходе */
window.onBack=()=>{
  const s=document.querySelectorAll(".sheet-bg"); if(s.length){ s[s.length-1].remove(); return true; }
  if($("#tab-act").classList.contains("hidden")){ go("act"); return true; }
  return false;
};
document.querySelectorAll("[data-form]").forEach(c=>c.addEventListener("click",()=>{ act.forms[c.dataset.form]=!act.forms[c.dataset.form]; saveAct(); renderPrint(); }));
$("#btnAbout").onclick=showAbout;
$("#btnPrint").onclick=printActs; $("#btnPdf").onclick=savePdf; $("#btnMail").onclick=sendMail;
$("#btnNew").onclick=()=>{
  if(actHasContent(act)) saveToDb();
  act=newAct(act.azs); saveAct(); renderAct(); go("act"); toast("Новый акт");
};
$("#camInput").onchange=e=>onPhoto(e.target.files[0]);

/* при каждом запуске — новый акт с той же АЗС; прошлый (если заполнен) сохраняется в базу */
(function startNewAct(){
  const last = act;
  if(actHasContent(last)) saveToDb();
  act = newAct(last && last.azs); saveAct();
})();

/* разделы «Данных» сворачиваются нажатием на заголовок; по умолчанию свёрнуты, открытые запоминаются на этом телефоне */
(function foldCards(){
  let open=[]; try{ open=JSON.parse(localStorage.getItem("akt2ui.open")||"[]"); }catch(e){}
  document.querySelectorAll("#tab-data > .card").forEach(c=>{
    const h=c.querySelector(":scope > h2"), key=h.textContent.trim();
    c.classList.toggle("fold", !open.includes(key));
    h.addEventListener("click",()=>{
      c.classList.toggle("fold");
      open=open.filter(k=>k!==key); if(!c.classList.contains("fold")) open.push(key);
      try{ localStorage.setItem("akt2ui.open", JSON.stringify(open)); }catch(e){}
    });
  });
})();
bindAct(); bindDb(); bindData(); renderAct();
setTimeout(()=>checkUpdate(false), 3000);   // проверка обновлений при каждом запуске, без настроек
