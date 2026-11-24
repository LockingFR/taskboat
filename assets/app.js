(function(){
"use strict";
/* Renseigner après l'inscription de l'application Azure (voir le plan en tête de fichier). */
const CLIENT_ID = "abcc1493-94ee-4992-99c8-749ed8a2bbd1";
const CONFIGURED = /^[0-9a-f-]{36}$/i.test(CLIENT_ID);
const MSAL_URL = "/assets/vendor/msal-browser.min.js"; // copie locale de @azure/msal-browser 2.38.3 (lib/msal-browser.min.js)
const GRAPH = "https://graph.microsoft.com/v1.0";

const $ = s => document.querySelector(s);
let lang = "fr";
let data = { source: "sample", account: null, lists: [] };
let cur = 0, lastOut = null, lastKey = "todoist";

const T = {
  fr:{ sample:"Exemple", sampleTxt:"Données d'exemple affichées. Connectez-vous ou ouvrez un fichier pour voir vos propres listes.",
       loaded:(l,t)=>`${l} liste${l>1?"s":""}, ${t} tâche${t>1?"s chargées":" chargée"}`, from:"depuis", noTasks:"Aucune tâche dans cette liste.",
       search:"Rechercher…", due:"échéance", high:"importante", steps:"étapes", needId:"Connexion directe pas encore configurée sur cette page. Utilisez l'option 2 (fichier JSON de Graph Explorer) en attendant.",
       loading:"Chargement…", list:"Liste", badJson:"Ce fichier n'est pas une réponse Microsoft To Do reconnue. Attendu : un objet avec « value » (listes) ou une sauvegarde Taskboat.",
       copied:"Copié.", copyFail:"Copie refusée par le navigateur : sélectionnez le texte et copiez-le.",
       todoistHint:"Les étapes cochées sont notées ✓ dans la description ; les répétitions sont à recréer. Dans Todoist : créez un projet, menu ⋯ › Importer depuis un modèle CSV. 300 tâches max par projet. Pas encore de compte ? https://todoist.com",
       skipped:n=>n?`${n} tâche${n>1?"s terminées non incluses":" terminée non incluse"} : Todoist ${n>1?"les":"la"} rouvrirait. Gardez la sauvegarde JSON. `:"",
       tooMany:n=>`Attention : ${n} lignes, au-delà des 300 acceptées par Todoist. Découpez ce fichier ou décochez les étapes. `,
       icsHint:"Importez ce fichier dans un calendrier de tâches (Thunderbird, Nextcloud Tasks…).",
       mdHint:"Collez dans une note Obsidian. Les dates suivent le format du plugin Tasks.",
       csvHint:"S'ouvre dans Excel, Google Sheets ou Numbers (UTF-8).", jsonHint:"Gardez ce fichier : il contient tout et se rouvre ici.",
       dlFail:"Le téléchargement est bloqué ici : utilisez Copier.", err:"Erreur", throttled:"Microsoft limite le débit, nouvelle tentative…" },
  en:{ sample:"Sample", sampleTxt:"Sample data shown. Sign in or open a file to see your own lists.",
       loaded:(l,t)=>`${l} list${l===1?"":"s"}, ${t} task${t===1?"":"s"} loaded`, from:"from", noTasks:"No tasks in this list.",
       search:"Search…", due:"due", high:"important", steps:"steps", needId:"Direct sign-in is not configured on this page yet. Use option 2 (Graph Explorer JSON file) for now.",
       loading:"Loading…", list:"List", badJson:"This file is not a recognised Microsoft To Do response. Expected an object with \"value\" (lists) or a Taskboat backup.",
       copied:"Copied.", copyFail:"The browser refused to copy: select the text and copy it.",
       todoistHint:"Checked steps are kept as ✓ in the description; recurrences must be set again. In Todoist: create a project, ⋯ menu › Import from CSV template. 300 tasks max per project. No account yet? https://todoist.com",
       skipped:n=>n?`${n} completed task${n>1?"s":""} left out: Todoist would reopen ${n>1?"them":"it"}. Keep the JSON backup. `:"",
       tooMany:n=>`Warning: ${n} rows, above Todoist's limit of 300. Split this file or untick steps. `,
       icsHint:"Import this file into a task calendar (Thunderbird, Nextcloud Tasks…).",
       mdHint:"Paste into an Obsidian note. Dates follow the Tasks plugin format.",
       csvHint:"Opens in Excel, Google Sheets or Numbers (UTF-8).", jsonHint:"Keep this file: it holds everything and reopens here.",
       dlFail:"Download is blocked here: use Copy.", err:"Error", throttled:"Microsoft is throttling, retrying…" }
};
const t = k => T[lang][k];

/* ---------- Données d'exemple ---------- */
function sample(){
  const d = (n)=>{ const x=new Date(); x.setDate(x.getDate()+n); return x.toISOString().slice(0,10); };
  return { source:"sample", account:null, lists:[
    { id:"l1", name:"Tâches", wellknown:"defaultList", tasks:[
      { id:"a1", title:"Renouveler la carte d'identité", status:"notStarted", importance:"high", due:d(12), notes:"Prendre rendez-vous en mairie, photo d'identité récente.", steps:[{title:"Photo d'identité",done:true},{title:"Pré-demande en ligne ANTS",done:false},{title:"Rendez-vous en mairie",done:false}], categories:[], links:[] },
      { id:"a2", title:"Résilier l'ancienne box internet", status:"completed", importance:"normal", due:d(-20), completed:d(-22), notes:"", steps:[], categories:[], links:[] },
      { id:"a3", title:"Payer la taxe foncière", status:"notStarted", importance:"normal", due:d(30), notes:"", steps:[], categories:["Finances"], links:[] }
    ]},
    { id:"l2", name:"Projet Atlas", wellknown:"none", tasks:[
      { id:"b1", title:"Relire la maquette de la page d'accueil", status:"notStarted", importance:"high", due:d(3), notes:"Vérifier les textes en anglais et le contraste du bouton principal.", steps:[{title:"Version mobile",done:false},{title:"Version bureau",done:true}], categories:["Travail"], links:[{name:"Maquette",url:"https://example.com/maquette"}] },
      { id:"b2", title:"Préparer la démo client de jeudi", status:"inProgress", importance:"normal", due:d(6), notes:"", steps:[{title:"Jeu de données de démo",done:false}], categories:["Travail"], links:[] },
      { id:"b3", title:"Envoyer le compte rendu de réunion", status:"completed", importance:"normal", due:"", completed:d(-1), notes:"", steps:[], categories:[], links:[] }
    ]},
    { id:"l3", name:"Courses", wellknown:"none", tasks:[
      { id:"c1", title:"Pain de campagne", status:"notStarted", importance:"normal", due:"", notes:"", steps:[], categories:[], links:[] },
      { id:"c2", title:"Café en grains", status:"notStarted", importance:"normal", due:"", notes:"Torréfaction moyenne", steps:[], categories:[], links:[] },
      { id:"c3", title:"Lessive", status:"completed", importance:"normal", due:"", notes:"", steps:[], categories:[], links:[] }
    ]}
  ]};
}

/* ---------- Normalisation des réponses Graph ---------- */
function strip(html){ if(!html) return ""; /* DOMParser : aucun script ni image chargés (notes de listes partagées = contenu non fiable) */ const d=new DOMParser().parseFromString(html.replace(/<br\s*\/?>/gi,"\n").replace(/<\/p>/gi,"\n"),"text/html"); return ((d.body&&d.body.textContent)||"").replace(/\n{3,}/g,"\n\n").trim(); }
function dt(o){ if(!o) return ""; const s = typeof o==="string" ? o : o.dateTime; if(!s) return "";
  /* minuit local est parfois renvoyé en UTC (veille 22 h ou 23 h) : on revient au jour local */
  if(o.timeZone==="UTC" && !/T00:00:00/.test(s)){ const d=new Date(s.slice(0,19)+"Z"); if(!isNaN(d)) return [d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-"); }
  return s.slice(0,10); }
function normTask(x){
  if(x && x.title!==undefined && Array.isArray(x.steps)) return x; // déjà normalisée (sauvegarde Taskboat)
  const body = x.body ? (x.body.contentType==="html" ? strip(x.body.content) : (x.body.content||"").trim()) : "";
  return {
    id:x.id||"", title:x.title||"(sans titre)", status:x.status||"notStarted", importance:x.importance||"normal",
    due:dt(x.dueDateTime), reminder:x.isReminderOn ? (x.reminderDateTime ? x.reminderDateTime.dateTime : "") : "",
    created:x.createdDateTime||"", completed:dt(x.completedDateTime), notes:body,
    steps:(x.checklistItems||[]).map(c=>({title:c.displayName||"",done:!!c.isChecked})),
    categories:x.categories||[],
    recurrence:x.recurrence && x.recurrence.pattern ? (x.recurrence.pattern.type+(x.recurrence.pattern.interval>1?" /"+x.recurrence.pattern.interval:"")) : "",
    links:(x.linkedResources||[]).filter(r=>r.webUrl).map(r=>({name:r.displayName||r.applicationName||"",url:r.webUrl}))
  };
}
function normList(l){
  return { id:l.id||"", name:l.displayName||l.name||"(liste)", wellknown:l.wellknownListName||l.wellknown||"none",
           tasks:(l.tasks||[]).map(normTask) };
}
function parseAny(obj){
  if(obj && obj.taskboat && Array.isArray(obj.lists)) return { source:"file", account:obj.account||null, lists:obj.lists.map(normList) };
  if(obj && Array.isArray(obj.value)){
    if(obj.value.length && obj.value[0].title!==undefined && obj.value[0].displayName===undefined){
      return { source:"file", account:null, lists:[{ id:"x", name:t("list"), wellknown:"none", tasks:obj.value.map(normTask) }] };
    }
    return { source:"file", account:null, lists:obj.value.map(normList) };
  }
  if(Array.isArray(obj)) return parseAny({value:obj});
  if(obj && Array.isArray(obj.lists)) return { source:"file", account:null, lists:obj.lists.map(normList) };
  throw new Error(t("badJson"));
}

/* ---------- Rendu ---------- */
function fmtDate(s){ const d=new Date(s+"T12:00:00"); return isNaN(d) ? s : d.toLocaleDateString(lang,{day:"numeric",month:"short",year:"numeric"}); }
function esc(s){ return String(s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }
function totals(){ return data.lists.reduce((n,l)=>n+l.tasks.length,0); }
function setStatus(msg){ $("#status").innerHTML = msg; }
function statusDefault(){
  if(data.source==="sample") setStatus(`<span class="tag">${t("sample")}</span>${esc(t("sampleTxt"))}`);
  else setStatus(`<b>${esc(T[lang].loaded(data.lists.length, totals()))}</b>${data.account?` · ${t("from")} ${esc(data.account)}`:""}`);
}
function renderLists(){
  const nav=$("#lists"); nav.innerHTML="";
  data.lists.forEach((l,i)=>{
    const b=document.createElement("button"); b.type="button";
    b.setAttribute("aria-current", i===cur ? "true":"false");
    const open=l.tasks.filter(x=>x.status!=="completed").length;
    b.innerHTML=`<span>${esc(l.name)}</span><span class="c">${open}/${l.tasks.length}</span>`;
    b.addEventListener("click",()=>{ cur=i; renderLists(); renderTasks(); });
    nav.appendChild(b);
  });
}
function renderTasks(){
  const l=data.lists[cur]; const box=$("#task-list");
  if(!l){ $("#list-title").textContent=""; box.innerHTML=""; return; }
  $("#list-title").textContent=l.name;
  const q=$("#q").value.trim().toLowerCase();
  const items=l.tasks.filter(x=>!q || (x.title+" "+x.notes+" "+x.steps.map(s=>s.title).join(" ")).toLowerCase().includes(q));
  if(!items.length){ box.innerHTML=`<p class="empty">${esc(t("noTasks"))}</p>`; return; }
  const sorted=items.slice().sort((a,b)=>(a.status==="completed")-(b.status==="completed"));
  box.innerHTML=sorted.map(x=>{
    const meta=[];
    if(x.importance==="high") meta.push(`<span class="hi">${t("high")}</span>`);
    if(x.due) meta.push(`<span>${t("due")} ${esc(fmtDate(x.due))}</span>`);
    if(x.recurrence) meta.push(`<span>↻ ${esc(x.recurrence)}</span>`);
    if(x.steps.length) meta.push(`<span>${x.steps.filter(s=>s.done).length}/${x.steps.length} ${t("steps")}</span>`);
    x.categories.forEach(c=>meta.push(`<span>${esc(c)}</span>`));
    return `<div class="task ${x.status==="completed"?"done":""}"><div class="box" aria-hidden="true"></div><div>
      <div class="t">${esc(x.title)}</div>
      ${meta.length?`<div class="meta">${meta.join("")}</div>`:""}
      ${x.notes?`<div class="note">${esc(x.notes)}</div>`:""}
      ${x.steps.length?`<ul class="steps">${x.steps.map(s=>`<li class="${s.done?"ok":""}">${esc(s.title)}</li>`).join("")}</ul>`:""}
    </div></div>`;
  }).join("");
}
function renderAll(){ if(cur>=data.lists.length) cur=0; statusDefault(); renderLists(); renderTasks(); }

/* ---------- Exports ---------- */
const optDone=()=>$("#opt-done").checked, optSteps=()=>$("#opt-steps").checked;
const keep=x=>optDone() || x.status!=="completed";
function csvCell(v){ v=String(v==null?"":v); return /[",\n\r;]/.test(v) ? '"'+v.replace(/"/g,'""')+'"' : v; }
function csvRow(a){ return a.map(csvCell).join(","); }
function slug(s){ return (s||"liste").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-zA-Z0-9]+/g,"-").replace(/^-|-$/g,"").toLowerCase()||"liste"; }
const stamp=()=>new Date().toISOString().slice(0,10);

function exJSON(){
  return { name:`microsoft-to-do-${stamp()}.json`, mime:"application/json",
    text:JSON.stringify({ taskboat:1, exported:new Date().toISOString(), account:data.account, lists:data.lists },null,2), hint:t("jsonHint") };
}
function exCSV(){
  const rows=[csvRow(["list","title","status","importance","due","completed","notes","steps","categories","recurrence","links"])];
  data.lists.forEach(l=>l.tasks.filter(keep).forEach(x=>rows.push(csvRow([
    l.name,x.title,x.status,x.importance,x.due,x.completed||"",x.notes,
    optSteps()?x.steps.map(s=>(s.done?"[x] ":"[ ] ")+s.title).join(" | "):"",
    x.categories.join(" | "),x.recurrence||"",x.links.map(k=>k.url).join(" | ")]))));
  return { name:`microsoft-to-do-${stamp()}.csv`, mime:"text/csv", text:"\uFEFF"+rows.join("\r\n"), hint:t("csvHint") };
}
function exTodoist(){
  const l=data.lists[cur];
  const head=["TYPE","CONTENT","DESCRIPTION","PRIORITY","INDENT","AUTHOR","RESPONSIBLE","DATE","DATE_LANG","TIMEZONE","DURATION","DURATION_UNIT","DEADLINE","DEADLINE_LANG"];
  const rows=[csvRow(head)];
  /* Le CSV Todoist n'a pas d'état « terminé » : tout ce qui est fini y redeviendrait à faire. */
  l.tasks.filter(x=>x.status!=="completed").forEach(x=>{
    const content=x.title+(x.categories.length?" "+x.categories.map(c=>"@"+c.replace(/\s+/g,"_")).join(" "):"");
    const done=optSteps()&&optDone() ? x.steps.filter(s=>s.done).map(s=>"✓ "+s.title).join("\n") : "";
    rows.push(csvRow(["task",content,[x.notes,done].filter(Boolean).join("\n\n"),x.importance==="high"?1:4,1,"","",x.due,x.due?"en":"","","","","",""]));
    if(optSteps()) x.steps.filter(s=>!s.done).forEach(s=>rows.push(csvRow(["task",s.title,"",4,2,"","","","","","","","",""])));
  });
  return { name:`todoist-${slug(l.name)}.csv`, mime:"text/csv", text:rows.join("\r\n"), hint:t("skipped")(l.tasks.filter(x=>x.status==="completed").length)+(rows.length>301?t("tooMany")(rows.length-1):"")+t("todoistHint") };
}
function recText(r){ /* format du plugin Tasks : « every 2 weeks » */
  const m=/^(\w+?)(?: \/(\d+))?$/.exec(r); if(!m) return "";
  const u={daily:"day",weekly:"week",absoluteMonthly:"month",relativeMonthly:"month",absoluteYearly:"year",relativeYearly:"year"}[m[1]];
  return u ? (m[2] ? `every ${m[2]} ${u}s` : `every ${u}`) : "";
}
function exMD(){
  const out=[];
  data.lists.forEach(l=>{
    const ts=l.tasks.filter(keep); if(!ts.length) return;
    out.push(`## ${l.name}`,"");
    ts.forEach(x=>{
      let line=`- [${x.status==="completed"?"x":" "}] ${x.title}`;
      if(x.importance==="high") line+=" ⏫";
      if(recText(x.recurrence)) line+=` 🔁 ${recText(x.recurrence)}`;
      if(x.due) line+=` 📅 ${x.due}`;
      if(x.completed) line+=` ✅ ${x.completed}`;
      x.categories.forEach(c=>line+=` #${c.replace(/\s+/g,"_")}`);
      out.push(line);
      if(optSteps()) x.steps.forEach(s=>out.push(`    - [${s.done?"x":" "}] ${s.title}`));
      if(x.notes) x.notes.split("\n").forEach(n=>out.push(`    > ${n}`));
      x.links.forEach(k=>out.push(`    - [${k.name||k.url}](${k.url})`));
    });
    out.push("");
  });
  return { name:`microsoft-to-do-${stamp()}.md`, mime:"text/markdown", text:`# Microsoft To Do · ${stamp()}\n\n`+out.join("\n"), hint:t("mdHint") };
}
function icsEsc(s){ return String(s).replace(/\\/g,"\\\\").replace(/;/g,"\\;").replace(/,/g,"\\,").replace(/\r?\n/g,"\\n"); }
function fold(line){ const parts=[]; while(line.length>73){ parts.push(line.slice(0,73)); line=" "+line.slice(73); } parts.push(line); return parts.join("\r\n"); }
function exICS(){
  const now=new Date().toISOString().replace(/[-:]/g,"").replace(/\.\d+/,"");
  const L=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Taskboat//Microsoft To Do export//FR","CALSCALE:GREGORIAN"];
  data.lists.forEach(l=>l.tasks.filter(keep).forEach((x,i)=>{
    let desc=x.notes||"";
    if(optSteps() && x.steps.length) desc+=(desc?"\n\n":"")+x.steps.map(s=>(s.done?"[x] ":"[ ] ")+s.title).join("\n");
    L.push("BEGIN:VTODO",fold(`UID:${(x.id||l.id+"-"+i).replace(/[^A-Za-z0-9-_=]/g,"")}@taskboat`),`DTSTAMP:${now}`,
      fold("SUMMARY:"+icsEsc(x.title)),fold("CATEGORIES:"+[l.name].concat(x.categories).map(icsEsc).join(",")));
    if(desc) L.push(fold("DESCRIPTION:"+icsEsc(desc)));
    if(x.due) L.push("DUE;VALUE=DATE:"+x.due.replace(/-/g,""));
    L.push("PRIORITY:"+(x.importance==="high"?1:x.importance==="low"?9:5));
    L.push("STATUS:"+(x.status==="completed"?"COMPLETED":x.status==="inProgress"?"IN-PROCESS":"NEEDS-ACTION"));
    if(x.completed) L.push("COMPLETED:"+x.completed.replace(/-/g,"")+"T000000Z");
    L.push("END:VTODO");
  }));
  L.push("END:VCALENDAR");
  return { name:`microsoft-to-do-${stamp()}.ics`, mime:"text/calendar", text:L.join("\r\n")+"\r\n", hint:t("icsHint") };
}
const EX={json:exJSON,csv:exCSV,todoist:exTodoist,md:exMD,ics:exICS};
function showOut(o){
  document.querySelectorAll("#fmt button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.f===lastKey));
  lastOut=o; $("#out").hidden=false; $("#fname").textContent=o.name;
  $("#outtext").value=o.text.length>200000 ? o.text.slice(0,200000)+"\n…" : o.text; $("#outhint").innerHTML=esc(o.hint).replace(/(https:\/\/[^\s<]+)/g,'<a href="$1" rel="noopener" target="_blank">$1</a>');
}
function download(){
  if(!lastOut) return;
  try{
    const url=URL.createObjectURL(new Blob([lastOut.text],{type:lastOut.mime+";charset=utf-8"}));
    const a=document.createElement("a"); a.href=url; a.download=lastOut.name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),4000);
  }catch(e){ $("#outhint").textContent=t("dlFail"); }
}
function copyOut(){
  if(!lastOut) return;
  const ta=$("#outtext");
  try{
    navigator.clipboard.writeText(lastOut.text).then(()=>{ $("#outhint").textContent=t("copied"); },()=>{ $("#preview").open=true; ta.select(); $("#outhint").textContent=t("copyFail"); });
  }catch(e){ $("#preview").open=true; ta.select(); $("#outhint").textContent=t("copyFail"); }
}

/* ---------- Fichier JSON ---------- */
function loadText(txt){
  try{ data=parseAny(JSON.parse(txt)); cur=0; renderAll(); $("#out").hidden=true; }
  catch(e){ setStatus(`<b>${t("err")}</b> · ${esc(e.message||String(e))}`); }
}
$("#file").addEventListener("change",e=>{ const f=e.target.files[0]; if(f) f.text().then(loadText); e.target.value=""; });
const drop=$("#drop");
["dragenter","dragover"].forEach(ev=>document.addEventListener(ev,e=>{ e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(ev=>document.addEventListener(ev,e=>{ if(ev==="drop") e.preventDefault(); drop.classList.remove("over"); }));
document.addEventListener("drop",e=>{ const f=e.dataTransfer && e.dataTransfer.files[0]; if(f) f.text().then(loadText); });
document.addEventListener("paste",e=>{
  if(e.target && (e.target.tagName==="INPUT"||e.target.tagName==="TEXTAREA")) return;
  const s=(e.clipboardData||window.clipboardData).getData("text"); if(s && /^\s*[\[{]/.test(s)) loadText(s);
});

/* ---------- Connexion Microsoft (MSAL, PKCE, côté client) ---------- */
let msalApp=null;
function loadScript(src){ return new Promise((ok,ko)=>{ const s=document.createElement("script"); s.src=src; s.onload=ok; s.onerror=()=>ko(new Error("MSAL")); document.head.appendChild(s); }); }
async function initMsal(){
  if(!window.msal) await loadScript(MSAL_URL);
  if(!msalApp){
    msalApp=new msal.PublicClientApplication({ auth:{ clientId:CLIENT_ID, authority:"https://login.microsoftonline.com/common", redirectUri:location.origin+"/" },
      cache:{ cacheLocation:"sessionStorage" } });
    if(msalApp.initialize) await msalApp.initialize();
  }
}
async function getToken(){
  await initMsal();
  const req={ scopes:["Tasks.Read","User.Read"] };
  let acc=msalApp.getAllAccounts()[0];
  if(!acc){ const r=await msalApp.loginPopup(req); acc=r.account; }
  data.account=acc ? acc.username : null;
  try{ return (await msalApp.acquireTokenSilent({...req,account:acc})).accessToken; }
  catch(e){ return (await msalApp.acquireTokenPopup(req)).accessToken; }
}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function gget(url,token){
  const out=[];
  while(url){
    const r=await fetch(url,{headers:{Authorization:"Bearer "+token}});
    if(r.status===429||r.status===503){ setStatus(esc(t("throttled"))); await sleep((+r.headers.get("Retry-After")||5)*1000); continue; }
    if(!r.ok){ const e=new Error(`Graph ${r.status}`); e.status=r.status; throw e; }
    const j=await r.json(); out.push(...(j.value||[])); url=j["@odata.nextLink"]||null;
  }
  return out;
}
/* $expand n'est pas documenté sur /tasks : si Graph le refuse (400), lecture des étapes et liens tâche par tâche. */
let canExpand=true;
async function getTasks(base,token){
  if(canExpand){
    try{ return await gget(`${base}?$top=100&$expand=checklistItems,linkedResources`,token); }
    catch(e){ if(e.status!==400) throw e; canExpand=false; }
  }
  const tasks=await gget(`${base}?$top=100`,token);
  for(const x of tasks){
    const u=`${base}/${encodeURIComponent(x.id)}`;
    x.checklistItems=await gget(u+"/checklistItems",token);
    x.linkedResources=await gget(u+"/linkedResources",token);
  }
  return tasks;
}
$("#btn-ms").addEventListener("click",async()=>{
  if(!CONFIGURED){ setStatus(`<b>${esc(t("needId"))}</b>`); return; }
  const btn=$("#btn-ms"); btn.disabled=true;
  try{
    setStatus(esc(t("loading")));
    const token=await getToken();
    const lists=await gget(GRAPH+"/me/todo/lists",token);
    const res=[];
    for(let i=0;i<lists.length;i++){
      setStatus(`${esc(t("loading"))} ${i+1}/${lists.length} · ${esc(lists[i].displayName||"")}`);
      const tasks=await getTasks(`${GRAPH}/me/todo/lists/${encodeURIComponent(lists[i].id)}/tasks`,token);
      res.push(normList({...lists[i],tasks}));
    }
    data={ source:"graph", account:data.account, lists:res }; cur=0; renderAll(); $("#out").hidden=true;
  }catch(e){ setStatus(`<b>${t("err")}</b> · ${esc(e.message||String(e))}`); }
  finally{ btn.disabled=false; }
});

/* ---------- Divers ---------- */
$("#fmt").addEventListener("click",e=>{ const b=e.target.closest("button[data-f]"); if(b && data.lists.length){ lastKey=b.dataset.f; showOut(EX[lastKey]()); } });
$("#btn-dl").addEventListener("click",download);
$("#btn-copy").addEventListener("click",copyOut);
$("#q").addEventListener("input",renderTasks);
["#opt-done","#opt-steps"].forEach(s=>$(s).addEventListener("change",()=>{ if(lastKey && !$("#out").hidden) showOut(EX[lastKey]()); }));

function setLang(l){
  lang=l; document.documentElement.lang=l;
  $("#lang-fr").setAttribute("aria-pressed",l==="fr"); $("#lang-en").setAttribute("aria-pressed",l==="en");
  $("#q").placeholder=t("search");
  try{ localStorage.setItem("taskboat-lang",l); }catch(e){}
  renderAll();
}
$("#lang-fr").addEventListener("click",()=>setLang("fr"));
$("#lang-en").addEventListener("click",()=>setLang("en"));

function countdown(){
  const end=new Date("2027-06-30T23:59:59"), start=new Date("2026-11-18T00:00:00"), now=new Date();
  const days=Math.max(0,Math.ceil((end-now)/86400000));
  $("#days").textContent=days;
  $("#daysarc").setAttribute("stroke-dasharray",Math.min(100,Math.max(0,(end-now)/(end-start)*100)).toFixed(1)+" 100");
}

data=sample();
let saved=null; try{ const qp=new URLSearchParams(location.search).get("lang"); saved=(qp==="en"||qp==="fr")?qp:localStorage.getItem("taskboat-lang"); }catch(e){}
const initial = saved || ((navigator.language||"fr").toLowerCase().startsWith("fr") ? "fr" : "en");
countdown(); setLang(initial);
showOut(exTodoist());
if(CONFIGURED) initMsal().catch(()=>{}); // prêt avant le clic, sinon Safari bloque la fenêtre
})();
