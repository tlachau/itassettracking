// @ts-nocheck
/* eslint-disable */
// IT Asset Tracker: the app UI and budget logic, reading and writing the SharePoint "IT" lists.
import { APP_CSS } from './appStyles';

export interface IAppOptions {
  webUrl: string;
  spHttpClient: any;
  spConfig: any;
  listPrefix: string;
  theme: string;
  aiUrl?: string;
  getAiClient?: () => Promise<any>;
  aiConfig?: any;
  version?: string;
  userName?: string;
  title?: string;
  orgName?: string;
}
export interface IAppHandle { destroy(): void; }

export function mountApp(host: HTMLElement, opts: IAppOptions): IAppHandle {
  if (!document.getElementById('itat-css')) {
    const st = document.createElement('style'); st.id = 'itat-css'; st.textContent = APP_CSS; document.head.appendChild(st);
    const fl = document.createElement('link'); fl.rel = 'stylesheet';
    fl.href = 'https://fonts.googleapis.com/css2?family=Archivo:wdth,wght@62..125,400..800&display=swap'; document.head.appendChild(fl);
  }
  const hEsc = (v: any) => String(v ?? "").replace(/[&<>"']/g, (ch: string) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" } as any)[ch]);
  host.innerHTML = `<div class="itat itat-page">
    <header class="topbar"><div class="brand">${hEsc(opts.title || "IT Asset Tracker")}${opts.orgName ? `<small>${hEsc(opts.orgName)}</small>` : ""}</div>
      <nav class="tabs" id="nav" aria-label="Sections"></nav>
      <div class="sync"><span id="sync">Loading…</span><button class="btn sm" data-act="refresh" title="Reload from SharePoint">Refresh</button></div></header>
    <main id="main"><div class="empty">Loading assets…</div></main></div>`;
  const ROOT = host.firstElementChild as HTMLElement;
  const OVL = document.createElement('div');
  OVL.className = 'itat';
  OVL.innerHTML = '<div class="scrim" id="scrim"></div><aside class="drawer" id="drawer" aria-hidden="true" role="dialog" aria-modal="true"></aside><div class="bulkbar" id="bulkbar" role="region" aria-label="Selected assets" hidden></div><div class="toast" id="toast" role="status"></div>';
  document.body.appendChild(OVL);
  const listeners = [];
  function on(type, fn, capture) { for (const el of [ROOT, OVL]) { el.addEventListener(type, fn, !!capture); listeners.push([el, type, fn, !!capture]); } }
  let mq = null; const applyTheme = () => {
    const t = opts.theme === 'dark' || (opts.theme === 'system' && mq && mq.matches) ? 'dark' : 'light';
    ROOT.setAttribute('data-theme', t); OVL.setAttribute('data-theme', t);
  };
  if (opts.theme === 'system' && window.matchMedia) { mq = window.matchMedia('(prefers-color-scheme: dark)'); mq.addEventListener('change', applyTheme); }
  applyTheme();

  /* ---------------- constants ---------------- */
  const CATS=["Computers","Servers","Network","Printers","Telephony","Connectivity","Other"];
  const STATUSES=["In Use","Spare","In Repair","Review","Return to Vendor","Retired"];
  const PSTATUS=["Employed","Terminated","Leave","Unknown"];
  const SITE_ORDER=[];
  const NOW=new Date(); const THIS_YEAR=NOW.getFullYear(); const Y0=THIS_YEAR+1; const YEARS=[Y0,Y0+1,Y0+2,Y0+3,Y0+4];
  const VIEWS=[["overview","Overview"],["assets","Assets"],["sites","Locations"],["spares","Spares"],["people","People"],["budget","Budget"],["contracts","Contracts"],["actions","Actions"],["settings","Settings"]];
  const hasContracts=()=>!!(S.ext&&S.ext.contractsAvailable);
  const visibleViews=()=>VIEWS.filter(([k])=>k!=="contracts"||hasContracts());
  const CCATS=["Connectivity","Support Renewal","Maintenance","Software Subscription","Licensing","Service","Lease","Other"];
  const BILLING=["Monthly","Annual","Multi-year"];
  const CSTATUS=["Active","Renewing","Cancelled"];
  const DEFAULT_CLASSES={none:{label:"Not budgeted",standard:"",cost:0,life:0}};

  let SAVED_SORT={};try{SAVED_SORT=JSON.parse(localStorage.getItem("itat-sort")||"{}")||{}}catch(e){}
  const S={assets:new Map(),people:new Map(),actions:new Map(),contracts:new Map(),cq:"",ccat:"All",cshow:"Active",classes:null,loaded:{assets:false,people:false,actions:false,classes:false},
    view:"overview",sample:null,downloads:null,canWrite:true,loadError:null,
    f:{q:"",cat:"All",loc:"All",status:"All",attn:false},
    site:null,pq:"",ploc:"All",pstat:"All",personId:null,
    byear:Y0,bloc:"All",open:new Set(),aiReview:null,aiBusy:false,
    aloc:"All",ashow:"Open",sploc:"All",spcat:"All",sel:new Set(),sort:SAVED_SORT,
    drawer:null};

  /* ---------------- helpers ---------------- */
  const $=s=>ROOT.querySelector(s)||OVL.querySelector(s);
  const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const money=n=>n==null||isNaN(n)?"—":"$"+Math.round(n).toLocaleString("en-US");
  const moneyK=n=>n>=1000?"$"+(n/1000).toFixed(n>=100000?0:1).replace(/\.0$/,"")+"k":money(n);
  const plural=(n,w,p)=>`${n.toLocaleString()} ${n===1?w:(p||w+"s")}`;
  const cls=()=>S.classes||DEFAULT_CLASSES;
  const klass=a=>cls()[a.rclass]||cls().none||DEFAULT_CLASSES.none;
  const assets=()=>[...S.assets.values()];
  const person=a=>a.personId?S.people.get(a.personId):null;
  function locations(){
    const set=new Set();
    for(const a of S.assets.values()) if(a.location) set.add(a.location);
    for(const p of S.people.values()) if(p.location) set.add(p.location);
    return [...set].sort((a,b)=>{const ia=SITE_ORDER.indexOf(a),ib=SITE_ORDER.indexOf(b);return (ia<0?99:ia)-(ib<0?99:ib)||a.localeCompare(b)});
  }
  function age(a){return a.year?Math.max(0,THIS_YEAR-a.year):null}
  function lifeBar(a){
    const k=klass(a),ag=age(a);
    if(ag==null) return `<span class="agecell muted small">Unknown</span>`;
    if(!k.life) return `<span class="agecell">${ag} yr</span>`;
    const p=ag/k.life, tone=p>=1?"bad":p>=.8?"warn":"";
    return `<span class="agecell" title="${ag} of ${k.life} years"><span class="life"><i class="${tone}" style="width:${Math.min(p,1)*100}%"></i></span>${ag} yr</span>`;
  }
  const needsSerial=a=>a.category==="Computers"&&!inactive(a)&&!String(a.serial||"").trim();
  const isTerm=a=>{const p=person(a);return !!p&&p.status==="Terminated"};
  const inactive=a=>a.status==="Retired"||a.status==="Return to Vendor";
  function flags(a){
    const out=[],k=klass(a),ag=age(a);
    if(isTerm(a)&&!inactive(a)) out.push(["Terminated user","warn"]);
    if(a.status==="Return to Vendor") out.push(["Return to Vendor","info"]);
    if(a.rclass&&a.rclass!=="none"&&!a.year&&!inactive(a)) out.push(["Age unknown","warn"]);
    if(k.life&&ag!=null&&ag>=k.life&&!inactive(a)&&a.plan!=="exclude") out.push(["Past lifecycle","bad"]);
    if(a.warrantyEnd){const d=(new Date(a.warrantyEnd)-NOW)/864e5; if(d<0) out.push(["Warranty ended","warn"]); else if(d<=90) out.push(["Warranty ends soon","warn"]);}
    return out;
  }
  const flagTags=a=>flags(a).map(([t,c])=>`<span class="tag ${c}">${esc(t)}</span>`).join(" ");
  function statusTag(s){const c={"In Use":"ok","Spare":"acc","In Repair":"warn","Review":"warn","Return to Vendor":"info","Retired":""}[s]||"";return `<span class="tag ${c}">${esc(s||"—")}</span>`}

  /* budget logic: which year (if any) and how an asset is funded */
  function budget(a){
    const k=klass(a);
    if(!a.rclass||a.rclass==="none"||a.status==="Retired") return null;
    const unit=a.quote!=null&&a.quote!==""?Number(a.quote):k.cost;
    const cost=(unit||0)*(a.qty||1);
    if(a.plan==="exclude") return {kind:"skip",year:null,cost,why:"Skipped"};
    if(a.plan==="replace"||a.plan==="reserve") return {kind:a.plan==="replace"?"base":"reserve",year:Number(a.planYear)||Y0,cost,why:a.plan==="replace"?"Planned replacement":"Conditional reserve"};
    if(a.status==="Return to Vendor") return {kind:"skip",year:null,cost,why:"Returning"};
    if(isTerm(a)) return {kind:"review",year:Y0,cost,why:"Assigned to terminated user"};
    if(!cost) return null;
    if(!a.year) return {kind:"reserve",year:Y0,cost,why:"Age unknown — validate"};
    const due=a.year+(k.life||99);
    return {kind:"base",year:Math.max(due,Y0),cost,why:due<Y0?`Past lifecycle since ${due}`:`Reaches ${k.life}-year lifecycle`};
  }
  function yearTotals(loc){
    const t={};YEARS.forEach(y=>t[y]={base:0,reserve:0,n:0,review:0});
    for(const a of S.assets.values()){
      if(loc&&loc!=="All"&&a.location!==loc) continue;
      const b=budget(a); if(!b||!t[b.year]) continue;
      if(b.kind==="base"){t[b.year].base+=b.cost;t[b.year].n++}
      else if(b.kind==="reserve"){t[b.year].reserve+=b.cost;t[b.year].n++}
      else if(b.kind==="review"){t[b.year].review++}
    }
    return t;
  }

  /* ---------------- toast ---------------- */
  let toastT;function toast(m){const t=$("#toast");t.textContent=m;t.classList.add("show");clearTimeout(toastT);toastT=setTimeout(()=>t.classList.remove("show"),2600)}

  /* ---------------- writes ---------------- */
  /* write() lives in the SharePoint data layer below */
  const stamp=()=>({updatedAt:new Date().toISOString()});

  /* ---------------- nav ---------------- */
  function renderNav(){
    const counts={assets:S.assets.size,spares:assets().filter(a=>a.status==="Spare").reduce((n,a)=>n+(Number(a.qty)||1),0)||undefined,people:S.people.size,contracts:hasContracts()?[...S.contracts.values()].filter(c=>c.status!=="Cancelled").length:undefined,actions:[...S.actions.values()].filter(t=>t.status!=="Done").length};
    $("#nav").innerHTML=visibleViews().map(([k,l])=>`<button data-act="nav" data-v="${k}" ${S.view===k?'aria-current="page"':""}><span>${l}</span>${counts[k]!=null&&S.loaded.assets?`<span class="count">${counts[k]}</span>`:""}</button>`).join("");
  }

  /* ---------------- views ---------------- */
  function horizonHTML(loc,selected){
    const t=yearTotals(loc);const max=Math.max(1,...YEARS.map(y=>t[y].base+t[y].reserve));
    return `<div class="horizon" role="group" aria-label="Replacement spend by year">${YEARS.map(y=>{
      const v=t[y],tot=v.base+v.reserve,h=tot/max*150;
      return `<div class="hcol" data-act="year" data-y="${y}" ${selected===y?'aria-pressed="true"':""} role="button" tabindex="0" aria-label="${y}: ${money(tot)}">
        <div class="hamt">${moneyK(tot)}<small>${plural(v.n,"device")}</small></div>
        <div class="hbar" style="height:${Math.max(h,3)}px">${v.reserve?`<div class="r" style="height:${v.reserve/tot*100}%"></div>`:""}${v.base?`<div class="b" style="height:${v.base/tot*100}%"></div>`:""}</div>
        <div class="hyear">${y}${y===Y0?' <span class="muted small pd">+ past due</span>':""}</div></div>`}).join("")}</div>
      <div class="legend" style="margin-top:12px"><span><i class="lb"></i>Base budget</span><span><i class="lr"></i>Conditional reserve</span></div>`;
  }

  function vOverview(){
    if(!S.assets.size&&!S.people.size) return `<h1>Welcome</h1><p class="lede">Your tracker is set up and empty. Here's how to get started.</p>
    <section class="panel"><ol class="small" style="line-height:1.9;margin:0;padding-left:20px">
      <li><b>Add your first device</b> with Add asset. Type a location name and it's created for you; add the rest of your sites the same way.</li>
      <li><b>Add people</b> with Add person, and add their devices in the same step.</li>
      <li><b>Check the replacement classes</b> in Settings. Set each class's cost and lifecycle to match what you actually buy; the budget is built from them.</li>
      <li><b>Importing a spreadsheet?</b> Paste rows into the SharePoint lists with Edit in grid view, or use Import-Assets.ps1 from the release.</li>
    </ol>
    <div class="row" style="margin-top:14px">${S.canWrite?`<button class="btn primary" data-act="newAsset">Add asset</button><button class="btn" data-act="newPerson">Add person</button>`:""}<button class="btn" data-act="nav" data-v="settings">Open Settings</button></div></section>`;
    return vOverviewMain();
  }
  function vOverviewMain(){
    const A=assets(),locs=locations();
    const live=A.filter(a=>!inactive(a));
    const term=A.filter(a=>isTerm(a)&&!inactive(a)).length;
    const unk=A.filter(a=>a.rclass&&a.rclass!=="none"&&!a.year&&!inactive(a)).length;
    const rtv=A.filter(a=>a.status==="Return to Vendor").length;
    const past=A.filter(a=>flags(a).some(f=>f[0]==="Past lifecycle")).length;
    const warr=A.filter(a=>flags(a).some(f=>f[0].startsWith("Warranty"))).length;
    const noser=A.filter(needsSerial).length;
    const openActs=[...S.actions.values()].filter(t=>t.status!=="Done");
    const crit=openActs.filter(t=>t.priority==="Critical"||t.priority==="High").length;
    const cOver=hasContracts()?contracts().filter(c=>{const r=renewalState(c);return r&&r.d<0}).length:0;
    const cSoon=hasContracts()?contracts().filter(c=>{const r=renewalState(c);return r&&r.d>=0&&r.d<=30}).length:0;
    const t=yearTotals();
    const siteRows=locs.map(l=>{
      const la=A.filter(a=>a.location===l&&!inactive(a));
      const lp=[...S.people.values()].filter(p=>p.location===l&&p.status==="Employed").length;
      const lpast=la.filter(a=>flags(a).some(f=>f[0]==="Past lifecycle")).length;
      const ly=yearTotals(l)[Y0];
      return `<div class="siterow" data-act="site" data-l="${esc(l)}" role="button" tabindex="0"><span class="nm">${esc(l)}</span><span class="r">${la.length}</span><span class="r hm">${lp}</span><span class="r hm">${lpast}</span><span class="r">${money(ly.base+ly.reserve)}</span></div>`}).join("");
    const byCat=CATS.map(c=>[c,live.filter(a=>a.category===c).length]).filter(x=>x[1]);
    const att=[
      [past,"devices are past their replacement lifecycle","assets",{attn:true,cat:"All"}],
      [term,"devices are assigned to terminated users","assets",{attn:"term"}],
      [unk,"devices have no age information","assets",{attn:"unknown"}],
      [rtv,"items are marked Return to Vendor","assets",{status:"Return to Vendor"}],
      [warr,"warranties ended or end within 90 days","assets",{attn:"warranty"}],
      [noser,"computers have no serial number or service tag","assets",{attn:"noserial",cat:"Computers"}],
      [cOver,"contract renewals are past due","contracts",null],
      [cSoon,"contracts renew within 30 days","contracts",null],
      [crit,"critical or high-priority actions are open","actions",null]].filter(x=>x[0]>0);
    return `
    <div class="spread"><div><h1>IT assets</h1>
    <p class="lede">${plural(live.length,"active asset")} across ${plural(locs.length,"location")}, assigned to ${plural([...S.people.values()].filter(p=>p.status==="Employed").length,"current employee")}. ${money(t[Y0].base)} base and ${money(t[Y0].reserve)} in reserves are planned for ${Y0}.${hasContracts()?` Contracts add ${money(recurringTotal(Y0,"All"))} per year.`:""}</p></div>
    ${S.canWrite?`<div class="row"><button class="btn primary" data-act="newAsset">Add asset</button><button class="btn" data-act="newPerson">Add person</button></div>`:""}</div>
    <section class="panel"><div class="spread"><div><h2 style="margin-bottom:2px">Replacement horizon</h2><p class="muted small" style="margin:0">Hardware spend by year, based on device age and lifecycle. Select a year to open its budget.</p></div></div>
    <div style="margin-top:14px">${horizonHTML("All",null)}</div></section>
    <div class="grid2" style="margin-top:16px">
      <section class="panel"><h2>Locations</h2><div class="sitelist">
        <div class="siterow head"><span>Site</span><span class="r">Assets</span><span class="r hm">Staff</span><span class="r hm">Past life</span><span class="r">${Y0} plan</span></div>${siteRows}</div></section>
      <section class="panel"><h2>Needs attention</h2>${att.length?`<ul class="attn">${att.map(([n,txt,v,f])=>`<li><span><b>${n}</b> ${txt}</span><button data-act="goattn" data-v="${v}" data-f='${esc(JSON.stringify(f||{}))}'>View</button></li>`).join("")}</ul>`:`<p class="muted">Nothing needs attention right now.</p>`}
      <h3 style="margin-top:18px">Active assets by class</h3><div class="chips">${byCat.map(([c,n])=>`<button class="chip" data-act="gocat" data-c="${c}">${c}<span class="n">${n}</span></button>`).join("")}</div></section>
    </div>`;
  }

  function filteredAssets(){
    const f=S.f,q=f.q.trim().toLowerCase();
    return assets().filter(a=>{
      if(f.cat!=="All"&&a.category!==f.cat) return false;
      if(f.loc!=="All"&&a.location!==f.loc) return false;
      if(f.status!=="All"&&a.status!==f.status) return false;
      if(f.attn===true&&!flags(a).length) return false;
      if(f.attn==="term"&&!(isTerm(a)&&!inactive(a))) return false;
      if(f.attn==="unknown"&&!(a.rclass&&a.rclass!=="none"&&!a.year&&!inactive(a))) return false;
      if(f.attn==="warranty"&&!flags(a).some(x=>x[0].startsWith("Warranty"))) return false;
      if(f.attn==="noserial"&&!needsSerial(a)) return false;
      if(q){const p=person(a);const hay=[a.name,a.model,a.make,a.type,a.assignedTo,p&&p.login,p&&p.title,a.serial,a.ip,a.location,a.notes].join(" ").toLowerCase();if(!q.split(/\s+/).every(w=>hay.includes(w)))return false}
      return true;
    }).sort((a,b)=>(a.location||"").localeCompare(b.location||"")||(a.name||"").localeCompare(b.name||"",undefined,{numeric:true}));
  }
  /* ---- sortable tables ---- */
  const PRI={Critical:0,High:1,Medium:2,Low:3};
  const STATUS_ORDER={"In Use":0,"Spare":1,"In Repair":2,"Review":3,"Return to Vendor":4,"Retired":5};
  const ipKey=v=>{const m=String(v||"").match(/(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})/);return m?m.slice(1).map(n=>n.padStart(3,"0")).join("."):null};
  const SORTERS={
    assets:{name:a=>a.name,type:a=>a.type||a.category,location:a=>a.location,assignedTo:a=>a.assignedTo,model:a=>a.model,age:a=>age(a),ip:a=>ipKey(a.ip)||a.ip,status:a=>STATUS_ORDER[a.status]??9,flags:a=>flags(a).length||null},
    contracts:{name:c=>c.name,vendor:c=>c.vendor,category:c=>c.category,location:c=>c.location,billing:c=>({Monthly:0,Annual:1,"Multi-year":2}[c.billing]??9),annual:c=>contractAnnual(c,THIS_YEAR),next:c=>c.next?daysUntil(c.next):null,quote:c=>c.quote==null||c.quote===""?null:Number(c.quote)},
    budget:{name:r=>r.a.name,model:r=>r.a.model,age:r=>age(r.a),why:r=>r.b.why,plan:r=>r.a.plan||"auto",cost:r=>r.b.cost},
    actions:{priority:t=>PRI[t.priority]??9,title:t=>t.title,location:t=>t.location,area:t=>t.area,status:t=>t.status},
    classes:{spareMin:k=>Number(cls()[k].spareMin)||null,label:k=>cls()[k].label,standard:k=>cls()[k].standard,cost:k=>cls()[k].cost,life:k=>cls()[k].life,n:k=>assets().filter(a=>a.rclass===k).length},
  };
  function cmpVal(x,y){
    if(typeof x==="number"&&typeof y==="number") return x-y;
    return String(x).localeCompare(String(y),undefined,{numeric:true,sensitivity:"base"});
  }
  function sortList(list,table){
    const st=S.sort[table]; if(!st||!st.key||!SORTERS[table][st.key]) return list;
    const get=SORTERS[table][st.key], dir=st.dir==="desc"?-1:1;
    return list.map((v,i)=>[v,get(v),i]).sort((p,q)=>{
      const a=p[1],b=q[1],an=a==null||a==="",bn=b==null||b==="";
      if(an||bn) return an&&bn?p[2]-q[2]:an?1:-1;   // blanks always last
      return cmpVal(a,b)*dir||p[2]-q[2];
    }).map(r=>r[0]);
  }
  function th(table,key,label,extra){
    const st=S.sort[table]||{}, on=st.key===key, dir=on?st.dir:null;
    const cls=extra?` class="${extra}"`:"";
    return `<th${cls}${on?` aria-sort="${dir==="desc"?"descending":"ascending"}"`:""}><button class="sorth" data-act="sort" data-t="${table}" data-k="${key}" title="Sort by ${esc(label)}">${label}<span class="ar" aria-hidden="true">${on?(dir==="desc"?"▼":"▲"):"▲"}</span></button></th>`;
  }
  /* ---- multi-select ---- */
  function selHead(ids){
    if(!S.canWrite||!ids.length) return "";
    const all=ids.every(id=>S.sel.has(id)), some=!all&&ids.some(id=>S.sel.has(id));
    return `<th class="selc"><input type="checkbox" data-act="selall" data-ids="${ids.join(",")}" ${all?"checked":""} ${some?'data-some="1"':""} aria-label="Select all ${ids.length} rows"></th>`;
  }
  function selCell(a){
    if(!S.canWrite) return "";
    return `<td class="selc"><input type="checkbox" data-act="selrow" data-id="${a._id}" ${S.sel.has(a._id)?"checked":""} aria-label="Select ${esc(a.name)}"></td>`;
  }
  function renderBulkBar(){
    const bar=$("#bulkbar"); if(!bar) return;
    for(const id of [...S.sel]) if(!S.assets.has(id)) S.sel.delete(id);
    const n=S.sel.size;
    if(!n||!S.canWrite||S.drawer||S.view==="settings"||S.view==="actions"||S.view==="overview"){bar.hidden=true;bar.innerHTML="";return}
    bar.hidden=false;
    bar.innerHTML=`<span><b>${n}</b> selected</span>
      <button class="btn sm primary" data-act="bulkEdit">Edit ${n===1?"item":"all "+n}</button>
      <button class="btn sm danger" data-act="bulkDelete">Delete</button>
      <button class="btn sm" data-act="selclear">Clear</button>`;
  }
  function assetTable(list,opts={}){
    if(!list.length) return `<div class="tablewrap"><div class="empty">No assets match these filters.${S.canWrite&&!opts.noAdd?` <button class="btn sm" data-act="newAsset">Add asset</button>`:""}</div></div>`;
    const showIP=list.some(a=>a.ip);
    list=sortList(list,"assets");
    return `<div class="tablewrap"><table><thead><tr>${selHead(list.map(a=>a._id))}${th("assets","name","Asset")}${opts.noLoc?"":th("assets","location","Location")}${th("assets","assignedTo","Assigned to")}${th("assets","model","Model")}${th("assets","age","Age")}${showIP&&opts.ip?th("assets","ip","IP"):""}${th("assets","status","Status")}${th("assets","flags","Flags")}${opts.unassign&&S.canWrite?"<th></th>":""}</tr></thead><tbody>
    ${list.map(a=>`<tr class="click${S.sel.has(a._id)?" sel":""}" data-act="open" data-id="${a._id}">${selCell(a)}<td><span class="nm">${esc(a.name)}</span><span class="sub">${esc(a.type||a.category)}</span></td>
    ${opts.noLoc?"":`<td>${esc(a.location||"—")}</td>`}
    <td>${a.personId&&S.people.has(a.personId)?`<button class="plink" data-act="gotoPerson" data-id="${a.personId}" title="Open ${esc(a.assignedTo||"")}'s page">${esc(a.assignedTo||"")}</button>`:a.assignedTo?esc(a.assignedTo):'<span class="muted">—</span>'}${(()=>{const p=person(a);return p&&p.status!=="Employed"?`<span class="sub">${esc(p.status)}</span>`:""})()}</td>
    <td>${esc(a.model||"—")}</td><td>${lifeBar(a)}</td>${showIP&&opts.ip?`<td class="small">${esc(a.ip||"")}</td>`:""}
    <td>${statusTag(a.status)}</td><td class="small">${flagTags(a)}</td>${opts.unassign&&S.canWrite?`<td><button class="btn sm" data-act="unassign" data-id="${a._id}">Unassign</button></td>`:""}</tr>`).join("")}</tbody></table></div>`;
  }
  function vAssets(){
    const A=assets(),f=S.f;
    const counts=Object.fromEntries(CATS.map(c=>[c,A.filter(a=>a.category===c).length]));
    const list=filteredAssets();
    const attnLabel={true:"Needs attention",term:"Terminated users",unknown:"Age unknown",warranty:"Warranty",noserial:"No serial number"}[f.attn];
    return `<div class="spread"><div><h1>Assets</h1><p class="lede">Every device, server, circuit, and system the IT team tracks.</p></div>
    <div class="row">${S.canWrite?`<button class="btn primary" data-act="newAsset">Add asset</button>`:""}<button class="btn" data-act="export">Export CSV</button><button class="btn" data-act="print" data-r="assets">Print</button></div></div>
    <div class="chips" style="margin-bottom:12px"><button class="chip" data-act="cat" data-c="All" aria-pressed="${f.cat==="All"}">All<span class="n">${A.length}</span></button>
    ${CATS.filter(c=>counts[c]).map(c=>`<button class="chip" data-act="cat" data-c="${c}" aria-pressed="${f.cat===c}">${c}<span class="n">${counts[c]}</span></button>`).join("")}</div>
    <div class="row" style="margin-bottom:14px">
      <input type="search" class="search" id="q" placeholder="Search name, model, user, serial, IP…" value="${esc(f.q)}" data-in="q" aria-label="Search assets">
      <select data-ch="loc" aria-label="Location"><option value="All">All locations</option>${locations().map(l=>`<option ${f.loc===l?"selected":""}>${esc(l)}</option>`).join("")}</select>
      <select data-ch="status" aria-label="Status"><option value="All">All statuses</option>${STATUSES.map(s=>`<option ${f.status===s?"selected":""}>${s}</option>`).join("")}</select>
      ${f.attn&&f.attn!==true?`<span class="tag warn">${attnLabel} <button class="btn ghost sm" data-act="clearattn" aria-label="Clear filter">✕</button></span>`:`<label class="row small" style="gap:6px"><input type="checkbox" data-ch="attn" ${f.attn?"checked":""}> Needs attention</label>`}
      <span class="muted small" style="margin-left:auto">${plural(list.length,"result")}</span>
    </div>
    ${assetTable(list,{ip:true})}`;
  }

  function vSites(){
    const locs=locations(); if(!S.site||!locs.includes(S.site)) S.site=locs[0];
    const l=S.site, A=assets().filter(a=>a.location===l);
    const live=A.filter(a=>!inactive(a));
    const staff=[...S.people.values()].filter(p=>p.location===l);
    const yt=yearTotals(l);
    const circuits=A.filter(a=>a.category==="Connectivity");
    const siteContracts=hasContracts()?contracts().filter(c=>activeContract(c)&&c.location===l):[];
    const monthly=hasContracts()?siteContracts.reduce((s,c)=>s+contractAnnual(c,THIS_YEAR),0)/12:circuits.reduce((s,a)=>s+(Number(a.monthlyCost)||0),0);
    const acts=[...S.actions.values()].filter(t=>t.location===l&&t.status!=="Done");
    const groups=CATS.map(c=>[c,A.filter(a=>a.category===c)]).filter(g=>g[1].length);
    return `<div class="spread"><div><h1>Locations</h1><p class="lede">Pick a site to see everything that lives there, including shared and spare equipment.</p></div><button class="btn" data-act="print" data-r="site">Print</button></div>
    <div class="chips" style="margin-bottom:18px">${locs.map(x=>`<button class="chip" data-act="site" data-l="${esc(x)}" aria-pressed="${x===l}">${esc(x)}</button>`).join("")}</div>
    <section class="panel"><div class="metrics">
      <div><div class="metric">${live.length}</div><div class="metric-l">Active assets</div></div>
      <div><div class="metric">${staff.filter(p=>p.status==="Employed").length}</div><div class="metric-l">Current staff</div></div>
      <div><div class="metric">${live.filter(a=>flags(a).some(f=>f[0]==="Past lifecycle")).length}</div><div class="metric-l">Past lifecycle</div></div>
      <div><div class="metric">${money(yt[Y0].base+yt[Y0].reserve)}</div><div class="metric-l">${Y0} plan, incl. reserve</div></div>
      ${(hasContracts()?siteContracts.length:circuits.length)?`<div><div class="metric">${money(monthly)}</div><div class="metric-l">${hasContracts()?"Site contracts per month":"Circuits per month"}</div></div>`:""}
    </div></section>
    <section class="panel"><h2>Replacement horizon for ${esc(l)}</h2>${horizonHTML(l,null)}</section>
    ${acts.length?`<section class="panel"><h2>Open actions</h2><ul class="attn">${acts.map(t=>`<li><span><span class="tag ${t.priority==="Critical"?"bad":t.priority==="High"?"warn":""}">${esc(t.priority)}</span> ${esc(t.title)}</span><button data-act="nav" data-v="actions">Open</button></li>`).join("")}</ul></section>`:""}
    ${siteContracts.length?`<section style="margin-top:22px"><h2>Contracts <span class="muted small">${siteContracts.length}</span></h2>${contractTable(siteContracts,{noLoc:true,noAdd:true})}</section>`:""}
    ${groups.map(([c,list])=>`<section style="margin-top:22px"><h2>${c} <span class="muted small">${list.length}</span></h2>${assetTable(list,{noLoc:true,ip:true,noAdd:true})}</section>`).join("")}`;
  }

  function vPeople(){
    const q=S.pq.trim().toLowerCase();
    const P=[...S.people.entries()].map(([id,p])=>({...p,_id:id})).filter(p=>{
      if(S.ploc!=="All"&&(p.location||"Unassigned")!==S.ploc) return false;
      if(S.pstat!=="All"&&p.status!==S.pstat) return false;
      if(q){const hay=[p.name,p.aka,p.login,p.title].join(" ").toLowerCase();return q.split(/\s+/).every(w=>hay.includes(w))}
      return true}).sort((a,b)=>a.name.localeCompare(b.name));
    if(S.personId&&!S.people.has(S.personId)) S.personId=null;
    const devCount=id=>assets().filter(a=>a.personId===id).length;
    const sel=S.personId?{...S.people.get(S.personId),_id:S.personId}:null;
    return `<div class="spread"><div><h1>People</h1><p class="lede">Search anyone to see the equipment assigned to them.</p></div>${S.canWrite?`<button class="btn primary" data-act="newPerson">Add person</button>`:""}</div>
    <div class="people">
      <section class="panel" style="padding:12px">
        <input type="search" id="pq" data-in="pq" placeholder="Search name, login, or title" value="${esc(S.pq)}" style="width:100%" aria-label="Search people">
        <div class="row" style="margin:8px 0 6px"><select data-ch="ploc" aria-label="Location" style="flex:1"><option value="All">All locations</option>${[...locations(),"Unassigned"].map(l=>`<option ${S.ploc===l?"selected":""}>${esc(l)}</option>`).join("")}</select>
        <select data-ch="pstat" aria-label="Status" style="flex:1"><option value="All">All statuses</option>${PSTATUS.map(s=>`<option ${S.pstat===s?"selected":""}>${s}</option>`).join("")}</select></div>
        <div class="muted tiny" style="padding:2px 4px 6px">${plural(P.length,"person","people")}</div>
        <div class="plist">${P.map(p=>`<button class="pitem" data-act="person" data-id="${p._id}" ${S.personId===p._id?'aria-current="true"':""}><span><span class="nm">${esc(p.name)}</span><span class="sub">${esc(p.title||"No title")} · ${esc(p.location||"Unassigned")}</span></span><span class="small muted" style="text-align:right">${devCount(p._id)||""}${p.status==="Terminated"?'<span class="sub" style="color:var(--warn)">Terminated</span>':""}</span></button>`).join("")||'<div class="empty small">No one matches that search.</div>'}</div>
      </section>
      <section>${sel?personDetail(sel):`<div class="panel empty">Select a person to see their devices and account status.</div>`}</section>
    </div>`;
  }
  function personDetail(p){
    const devs=assets().filter(a=>a.personId===p._id);
    const spare=assets().filter(a=>!a.personId&&a.category==="Computers"&&(a.status==="Spare"||a.status==="In Use")&&!a.assignedTo).sort((a,b)=>a.name.localeCompare(b.name));
    const b=(ok,l)=>`<span class="tag ${ok?"ok":"bad"}">${l} ${ok?"✓":"✕"}</span>`;
    return `<div class="panel"><div class="spread" style="align-items:flex-start"><div><h2 style="margin:0">${esc(p.name)}</h2><div class="muted">${esc(p.title||"No title")}${p.aka?` · also listed as ${esc(p.aka)}`:""}</div></div>
      <div class="row"><button class="btn sm" data-act="print" data-r="person" data-id="${p._id}">Print</button>${S.canWrite?`<button class="btn sm" data-act="editPerson" data-id="${p._id}">Edit</button>`:""}</div></div>
      <dl class="kv" style="margin-top:14px"><dt>Login</dt><dd>${esc(p.login||"—")}</dd><dt>Location</dt><dd>${esc(p.location||"Unassigned")}</dd>
      <dt>Status</dt><dd>${p.status==="Terminated"?`<span class="tag warn">Terminated</span>`:esc(p.status)}</dd></dl>
      <div class="badges" style="margin-top:12px">${b(p.o365,"O365")}${b(p.mfa,"MFA")}${b(p.vpn,"VPN")}${b(p.av,"AV")}</div>
      ${p.status==="Terminated"&&devs.length?`<div class="banner" style="margin:16px 0 0">This person is terminated but still has ${plural(devs.length,"device")} assigned. Review whether to collect, reassign, or mark them spare.${S.canWrite?` <button class="btn sm" data-act="unassignAll" data-id="${p._id}" style="margin-left:6px">Unassign all and mark spare</button>`:""}</div>`:""}
    </div>
    <div style="margin-top:16px"><h3>Assigned devices (${devs.length})</h3>${devs.length?assetTable(devs,{noAdd:true,unassign:true}):`<div class="panel muted">No devices assigned.</div>`}</div>
    ${(()=>{if(!S.ext||!S.ext.pastDevices)return "";const past=S.ext.pastDevices(p._id).filter(x=>!devs.some(d=>d._id===x.assetId));
      if(!past.length)return "";
      return `<div style="margin-top:16px"><h3>Previously assigned (${past.length})</h3><div class="tablewrap"><table><thead><tr><th>Device then</th><th>Device now</th><th>Unassigned</th><th>By</th></tr></thead><tbody>
      ${past.map(x=>{const now=S.assets.get(x.assetId);return `<tr ${now?`class="click" data-act="open" data-id="${x.assetId}"`:""}><td class="nm">${esc(x.assetName||"—")}</td><td>${now?`${esc(now.name)}<span class="sub">${esc(now.assignedTo||now.status||"")}</span>`:'<span class="muted">Deleted</span>'}</td><td>${esc(fmtDate(x.when))}</td><td>${esc(x.who||"")}</td></tr>`}).join("")}</tbody></table></div></div>`})()}
    ${S.canWrite?`<div class="row" style="margin-top:12px"><select id="assignSel" aria-label="Unassigned computer" style="flex:1;max-width:360px"><option value="">Assign an unassigned computer…</option>${spare.map(a=>`<option value="${a._id}">${esc(a.name)} — ${esc(a.model||a.type)} (${esc(a.location||"")})</option>`).join("")}</select><button class="btn" data-act="assign" data-id="${p._id}">Assign</button></div>`:""}`;
  }

  function vBudget(){
    const yr=S.byear, loc=S.bloc;
    const rows=[];
    for(const a of S.assets.values()){
      if(loc!=="All"&&a.location!==loc) continue;
      const b=budget(a); if(!b) continue;
      if(yr==="all"? (b.year&&YEARS.includes(b.year)) : b.year===yr) rows.push({a,b});
    }
    const funded=rows.filter(r=>r.b.kind==="base"||r.b.kind==="reserve");
    const review=rows.filter(r=>r.b.kind==="review");
    const base=funded.filter(r=>r.b.kind==="base").reduce((s,r)=>s+r.b.cost,0);
    const res=funded.filter(r=>r.b.kind==="reserve").reduce((s,r)=>s+r.b.cost,0);
    const groups={};for(const r of funded){(groups[r.a.rclass]=groups[r.a.rclass]||[]).push(r)}
    const gl=Object.entries(groups).sort((x,y)=>y[1].reduce((s,r)=>s+r.b.cost,0)-x[1].reduce((s,r)=>s+r.b.cost,0));
    const label=yr==="all"?`${Y0}–${YEARS[4]}`:yr;
    const recYears=yr==="all"?YEARS:[yr];
    return `<div class="spread"><div><h1>Budget</h1><p class="lede">Replacement needs are calculated from each device's age and its class lifecycle. Override any device's plan, or enter a real quote to replace the standard price.</p></div>
      <div class="row"><button class="btn" data-act="print" data-r="budget">Print</button><button class="btn" data-act="nav" data-v="settings">Edit costs and lifecycles</button></div></div>
    <section class="panel">${horizonHTML(loc,yr)}</section>
    <div class="row" style="margin:18px 0 12px">
      <div class="chips">${YEARS.map(y=>`<button class="chip" data-act="year" data-y="${y}" aria-pressed="${yr===y}">${y}</button>`).join("")}<button class="chip" data-act="year" data-y="all" aria-pressed="${yr==="all"}">All 5 years</button></div>
      <select data-ch="bloc" aria-label="Location"><option value="All">All locations</option>${locations().map(l=>`<option ${loc===l?"selected":""}>${esc(l)}</option>`).join("")}</select>
    </div>
    <section class="panel"><div class="metrics">
      <div><div class="metric">${money(base)}</div><div class="metric-l">Base budget, ${label}</div></div>
      <div><div class="metric">${money(res)}</div><div class="metric-l">Conditional reserve</div></div>
      <div><div class="metric">${money(base+res)}</div><div class="metric-l">Full planning envelope</div></div>
      <div><div class="metric">${review.length}</div><div class="metric-l">Held for review</div></div>
      ${hasContracts()?`<div><div class="metric">${money(recYears.reduce((s,y)=>s+recurringTotal(y,loc),0))}</div><div class="metric-l">Recurring contracts, ${label}</div></div>
      <div><div class="metric">${money(base+res+recYears.reduce((s,y)=>s+recurringTotal(y,loc),0))}</div><div class="metric-l">Hardware and contracts</div></div>`:""}
    </div>
    ${review.length?`<p class="small muted" style="margin:12px 0 0">${plural(review.length,"device")} assigned to terminated users ${review.length===1?"isn't":"aren't"} budgeted until reviewed. <button class="btn ghost sm" data-act="goattn" data-v="assets" data-f='{"attn":"term"}'>Review them</button></p>`:""}
    </section>
    <section class="panel">
      <div class="ghead"><span>Replacement class</span><span class="r">Qty</span><span class="r">Base</span><span class="r">Reserve</span></div>
      ${gl.length?gl.map(([rc,list])=>{const k=cls()[rc]||{label:rc};const b=list.filter(r=>r.b.kind==="base").reduce((s,r)=>s+r.b.cost,0),rv=list.filter(r=>r.b.kind==="reserve").reduce((s,r)=>s+r.b.cost,0);
        const key=rc+"|"+yr+"|"+loc;
        return `<details class="grp" data-key="${esc(key)}" ${S.open.has(key)?"open":""}><summary><span><span class="tw">▸</span> <b>${esc(k.label)}</b><span class="sub" style="padding-left:20px">${esc(k.standard||"")}${k.cost?` · ${money(k.cost)} each`:""}</span></span><span class="r">${list.length}</span><span class="r">${b?money(b):"—"}</span><span class="r">${rv?money(rv):"—"}</span></summary>
        <div class="gbody">${budgetRows(list)}</div></details>`}).join(""):`<div class="empty">Nothing is due for replacement in ${label}${loc!=="All"?" at "+esc(loc):""}.</div>`}
    </section>
    ${hasContracts()?recurringPanel(recYears,loc,label):""}
    ${S.sample?`<section class="panel"><div class="spread"><div><h2 style="margin-bottom:2px">AI planning review</h2><p class="muted small" style="margin:0">Claude reviews this plan and suggests what to fund first, what to stagger, and what to verify.</p></div>
    <div class="row">${S.aiBusy?`<button class="btn" data-act="stopai">Stop</button>`:""}<button class="btn primary" data-act="aireview" ${S.aiBusy||!S.sample?"disabled":""}>${S.aiReview?"Run again":"Review this plan"}</button></div></div>
    ${!S.sample?`<p class="muted small">AI features aren't connected yet. A site owner can add the AI service in the web part settings.</p>`:""}
    <div id="aiout" class="prose" style="margin-top:12px">${S.aiReview?mdLite(S.aiReview):""}</div></section>`:copilotPanel(yr,loc,label)}`;
  }
  function recurringPanel(years,loc,label){
    const list=contracts().filter(c=>activeContract(c)&&(loc==="All"||c.location===loc));
    const byCat={};for(const c of list){const v=years.reduce((s,y)=>s+contractAnnual(c,y),0);(byCat[c.category||"Other"]=byCat[c.category||"Other"]||[]).push({c,v})}
    const cats=Object.entries(byCat).map(([k,rows])=>[k,rows,rows.reduce((s,r)=>s+r.v,0)]).sort((a,b)=>b[2]-a[2]);
    const total=cats.reduce((s,x)=>s+x[2],0);
    return `<section class="panel"><div class="spread"><div><h2 style="margin-bottom:2px">Recurring contracts</h2><p class="muted small" style="margin:0">${label}${loc!=="All"?` · ${esc(loc)} only (company-wide contracts show under All locations)`:""}. Multi-year costs are spread evenly; quotes apply after the renewal year.</p></div><button class="btn" data-act="nav" data-v="contracts">Open contracts</button></div>
    ${cats.length?`<div class="ghead" style="margin-top:12px"><span>Category</span><span class="r">Contracts</span><span class="r hide-m"></span><span class="r">${years.length>1?"Total":"Per year"}</span></div>
    ${cats.map(([k,rows,v])=>{const key="rec|"+k+"|"+label+"|"+loc;return `<details class="grp" data-key="${esc(key)}" ${S.open.has(key)?"open":""}><summary><span><span class="tw">▸</span> <b>${esc(k)}</b></span><span class="r">${rows.length}</span><span class="r hide-m"></span><span class="r">${money(v)}</span></summary>
      <div class="gbody"><div class="tablewrap" style="margin-top:4px"><table><thead><tr><th>Contract</th><th>Vendor</th><th>Billing</th><th>Next renewal</th><th class="num">${years.length>1?"Total":"Per year"}</th></tr></thead><tbody>
      ${rows.sort((a,b)=>b.v-a.v).map(({c,v})=>`<tr class="click" data-act="openContract" data-id="${c._id}"><td class="nm">${esc(c.name)}</td><td>${esc(c.vendor||"")}</td><td>${c.cost!=null?billingText(c):"—"}</td><td>${c.next?esc(fmtDate(c.next)):"—"}</td><td class="num">${money(v)}</td></tr>`).join("")}</tbody></table></div></div></details>`}).join("")}
    <p class="small" style="margin:10px 6px 0;text-align:right"><b>${money(total)}</b> ${years.length>1?"over "+years.length+" years":"for "+label}</p>`:`<div class="empty">No contracts${loc!=="All"?" for "+esc(loc):""}.</div>`}</section>`;
  }
  function budgetRows(list){
    list.sort((x,y)=>(x.a.year||0)-(y.a.year||0));
    list=sortList(list,"budget");
    return `<div class="tablewrap" style="margin-top:4px"><table><thead><tr>${selHead(list.map(r=>r.a._id))}${th("budget","name","Asset")}${th("budget","model","Model","hide-m")}${th("budget","age","Age")}${th("budget","why","Why")}${th("budget","plan","Plan")}${th("budget","cost","Cost","num")}<th></th></tr></thead><tbody>
    ${list.map(({a,b})=>`<tr${S.sel.has(a._id)?' class="sel"':""}>${selCell(a)}<td><button class="btn ghost sm" style="padding:0" data-act="open" data-id="${a._id}"><span class="nm">${esc(a.name)}</span></button><span class="sub">${esc(a.location||"")}${a.assignedTo?" · "+esc(a.assignedTo):""}</span></td>
    <td class="hide-m">${esc(a.model||"—")}</td><td>${lifeBar(a)}</td><td class="small">${esc(b.why)}</td>
    <td>${S.canWrite?`<select data-plan="${a._id}" aria-label="Plan for ${esc(a.name)}"><option value="auto" ${a.plan==="auto"||!a.plan?"selected":""}>Auto</option><option value="replace" ${a.plan==="replace"?"selected":""}>Replace</option><option value="reserve" ${a.plan==="reserve"?"selected":""}>Reserve</option><option value="exclude">Skip</option></select>
     ${a.plan==="replace"||a.plan==="reserve"?`<select data-planyear="${a._id}" aria-label="Year">${YEARS.map(y=>`<option ${Number(a.planYear||Y0)===y?"selected":""}>${y}</option>`).join("")}</select>`:""}`:esc(a.plan||"auto")}</td>
    <td class="num">${money(b.cost)}${a.quote!=null&&a.quote!==""?'<span class="sub">Quote</span>':""}</td>
    <td>${a.rclass!=="none"?replButton(a,"Find replacement"):""}</td></tr>`).join("")}</tbody></table></div>`;
  }

  function vActions(){
    const all=[...S.actions.entries()].map(([id,t])=>({...t,_id:id}));
    const pr={Critical:0,High:1,Medium:2,Low:3};
    const list=all.filter(t=>(S.aloc==="All"||t.location===S.aloc)&&(S.ashow==="All"||(S.ashow==="Open"?t.status!=="Done":t.status==="Done")))
      .sort((a,b)=>(pr[a.priority]??9)-(pr[b.priority]??9)||(a.location||"").localeCompare(b.location||""));
    const slist=sortList(list,"actions");
    return `<div class="spread"><div><h1>Actions</h1><p class="lede">Follow-ups for validation, transitions, and data gaps.</p></div>${S.canWrite?`<button class="btn primary" data-act="newAction">Add action</button>`:""}</div>
    <div class="row" style="margin-bottom:14px"><div class="chips">${["Open","Done","All"].map(s=>`<button class="chip" data-act="ashow" data-s="${s}" aria-pressed="${S.ashow===s}">${s}</button>`).join("")}</div>
    <select data-ch="aloc" aria-label="Location"><option value="All">All locations</option>${locations().map(l=>`<option ${S.aloc===l?"selected":""}>${esc(l)}</option>`).join("")}</select></div>
    ${list.length?`<div class="tablewrap"><table><thead><tr><th></th>${th("actions","priority","Priority")}${th("actions","title","Action")}${th("actions","location","Location")}${th("actions","area","Area")}<th></th></tr></thead><tbody>
    ${slist.map(t=>`<tr><td>${S.canWrite?`<input type="checkbox" data-done="${t._id}" ${t.status==="Done"?"checked":""} aria-label="Mark done">`:""}</td>
    <td><span class="tag ${t.priority==="Critical"?"bad":t.priority==="High"?"warn":""}">${esc(t.priority||"—")}</span></td>
    <td style="min-width:280px"><span class="nm" ${t.status==="Done"?'style="text-decoration:line-through;color:var(--ink-3)"':""}>${esc(t.title)}</span><span class="sub">${esc(t.detail||"")}</span></td>
    <td>${esc(t.location||"—")}</td><td>${esc(t.area||"")}</td><td>${S.canWrite?`<button class="btn ghost sm" data-act="editAction" data-id="${t._id}">Edit</button>`:""}</td></tr>`).join("")}</tbody></table></div>`:`<div class="panel empty">No actions here.</div>`}`;
  }

  /* ---------------- contracts ---------------- */
  const TODAY=new Date(NOW.getFullYear(),NOW.getMonth(),NOW.getDate());
  const contracts=()=>[...S.contracts.values()];
  const activeContract=c=>c.status!=="Cancelled";
  function yearsOf(months){return Math.max(1,Number(months)||12)/12}
  /* Annual cost, spread evenly across the term. From the renewal year on, a quote replaces the current price. */
  function contractAnnual(c,year){
    if(!activeContract(c)) return 0;
    const useQuote=c.quote!=null&&c.quote!==""&&c.next&&year!=null&&year>Number(String(c.next).slice(0,4));
    const cost=Number(useQuote?c.quote:c.cost)||0;
    const term=useQuote?(c.quoteTerm||c.term):c.term;
    if(c.billing==="Monthly") return cost*12;
    if(c.billing==="Multi-year") return cost/yearsOf(term||36);
    return cost;   // Annual
  }
  function daysUntil(d){if(!d)return null;const x=new Date(String(d).slice(0,10)+"T00:00:00");return Math.round((x-TODAY)/864e5)}
  function renewalState(c){
    if(!activeContract(c)||!c.next) return null;
    const d=daysUntil(c.next);
    if(d<0) return {d,cls:"bad",label:`Overdue ${-d} day${d===-1?"":"s"}`};
    if(d===0) return {d,cls:"bad",label:"Renews today"};
    if(d<=30) return {d,cls:"warn",label:`In ${d} day${d===1?"":"s"}`};
    if(d<=90) return {d,cls:"info",label:`In ${d} days`};
    return {d,cls:"",label:""};
  }
  function recurringTotal(year,loc){
    return contracts().filter(c=>loc==null||loc==="All"||c.location===loc).reduce((s,c)=>s+contractAnnual(c,year),0);
  }
  function billingText(c){
    if(c.billing==="Monthly") return `${money(c.cost)}/mo`;
    if(c.billing==="Multi-year") return `${money(c.cost)} / ${Math.round((c.term||36)/12)} yr`;
    return `${money(c.cost)}/yr`;
  }
  function addMonths(iso,m){
    const d=new Date(String(iso).slice(0,10)+"T00:00:00"); if(isNaN(d)) return null;
    const day=d.getDate(); d.setDate(1); d.setMonth(d.getMonth()+Number(m||12));
    d.setDate(Math.min(day,new Date(d.getFullYear(),d.getMonth()+1,0).getDate()));
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  }
  function contractTable(list,opts={}){
    if(!list.length) return `<div class="tablewrap"><div class="empty">No contracts here.${S.canWrite&&!opts.noAdd?` <button class="btn sm" data-act="newContract">Add contract</button>`:""}</div></div>`;
    list=sortList(list,"contracts");
    return `<div class="tablewrap"><table><thead><tr>${th("contracts","name","Contract")}${th("contracts","vendor","Vendor")}${th("contracts","category","Category")}${opts.noLoc?"":th("contracts","location","Location")}${th("contracts","billing","Cost")}${th("contracts","annual","Per year","num")}${th("contracts","next","Next renewal")}${th("contracts","quote","Quote","num")}</tr></thead><tbody>
    ${list.map(c=>{const r=renewalState(c);return `<tr class="click" data-act="openContract" data-id="${c._id}"><td><span class="nm">${esc(c.name)}</span><span class="sub">${esc([c.account?"Acct "+c.account:"",c.serviceId||""].filter(Boolean).join(" · "))}</span></td>
      <td>${esc(c.vendor||"—")}</td><td>${esc(c.category||"—")}</td>${opts.noLoc?"":`<td>${esc(c.location||"All sites")}</td>`}
      <td>${c.cost!=null?billingText(c):"—"}${c.status!=="Active"?`<span class="sub">${esc(c.status)}</span>`:""}</td>
      <td class="num">${money(contractAnnual(c,THIS_YEAR))}</td>
      <td>${c.next?esc(fmtDate(c.next)):'<span class="muted">—</span>'}${r&&r.label?` <span class="tag ${r.cls}">${esc(r.label)}</span>`:""}${c.auto?'<span class="sub">Auto-renews</span>':""}</td>
      <td class="num">${c.quote!=null&&c.quote!==""?`${money(c.quote)}${c.quoteTerm?`<span class="sub">${Math.round(c.quoteTerm/12*10)/10} yr</span>`:""}`:'<span class="muted">—</span>'}</td></tr>`}).join("")}</tbody></table></div>`;
  }
  function vContracts(){
    if(!hasContracts()) return `<h1>Contracts</h1><div class="panel">Run Install-AssetTracker.ps1 again on this site (it adds any missing lists) to turn on contract tracking.</div>`;
    const all=contracts(), q=S.cq.trim().toLowerCase();
    const counts=Object.fromEntries(CCATS.map(c=>[c,all.filter(x=>x.category===c&&(S.cshow==="All"||activeContract(x))).length]));
    const list=all.filter(c=>(S.cshow==="All"||activeContract(c))&&(S.ccat==="All"||c.category===S.ccat)&&
      (!q||[c.name,c.vendor,c.account,c.serviceId,c.contact,c.notes,c.location,c.category].join(" ").toLowerCase().includes(q)));
    const active=all.filter(activeContract);
    const yearTotal=active.reduce((s,c)=>s+contractAnnual(c,THIS_YEAR),0);
    const soon=active.filter(c=>{const r=renewalState(c);return r&&r.d>=0&&r.d<=90});
    const overdue=active.filter(c=>{const r=renewalState(c);return r&&r.d<0});
    const upcoming=active.filter(c=>{const r=renewalState(c);return r&&r.d<=90}).sort((a,b)=>daysUntil(a.next)-daysUntil(b.next));
    return `<div class="spread"><div><h1>Contracts</h1><p class="lede">Services, subscriptions, licensing, support renewals, maintenance, and circuits. Costs are spread evenly across each term for the budget.</p></div>
    <div class="row">${S.canWrite?`<button class="btn primary" data-act="newContract">Add contract</button>`:""}<button class="btn" data-act="exportContracts">Export CSV</button><button class="btn" data-act="print" data-r="contracts">Print</button></div></div>
    <section class="panel"><div class="metrics">
      <div><div class="metric">${money(yearTotal)}</div><div class="metric-l">Per year, current pricing</div></div>
      <div><div class="metric">${money(yearTotal/12)}</div><div class="metric-l">Monthly equivalent</div></div>
      <div><div class="metric">${soon.length}</div><div class="metric-l">Renewing in 90 days</div></div>
      <div><div class="metric">${overdue.length}</div><div class="metric-l">Overdue</div></div>
    </div></section>
    ${upcoming.length?`<section class="panel"><h2>Coming up</h2><ul class="attn">${upcoming.map(c=>{const r=renewalState(c);return `<li><span><span class="tag ${r.cls}">${esc(r.label)}</span> <b>${esc(c.name)}</b> <span class="muted">${esc(c.vendor||"")} · ${esc(fmtDate(c.next))}${c.quote!=null&&c.quote!==""?` · quote ${money(c.quote)} vs ${money(c.cost)}`:""}</span></span><button data-act="openContract" data-id="${c._id}">Open</button></li>`}).join("")}</ul></section>`:""}
    <div class="chips" style="margin:18px 0 12px"><button class="chip" data-act="ccat" data-c="All" aria-pressed="${S.ccat==="All"}">All<span class="n">${Object.values(counts).reduce((a,b)=>a+b,0)}</span></button>
    ${CCATS.filter(c=>counts[c]).map(c=>`<button class="chip" data-act="ccat" data-c="${c}" aria-pressed="${S.ccat===c}">${c}<span class="n">${counts[c]}</span></button>`).join("")}</div>
    <div class="row" style="margin-bottom:14px"><input type="search" class="search" id="cq" data-in="cq" placeholder="Search contract, vendor, account #, circuit ID…" value="${esc(S.cq)}" aria-label="Search contracts">
      <div class="chips">${["Active","All"].map(x=>`<button class="chip" data-act="cshow" data-s="${x}" aria-pressed="${S.cshow===x}">${x==="All"?"Include cancelled":"Active"}</button>`).join("")}</div>
      <span class="muted small" style="margin-left:auto">${plural(list.length,"contract")}</span></div>
    ${contractTable(list)}`;
  }
  function contractForm(c,isNew){
    const ro=!S.canWrite?"disabled":"";
    const r=isNew?null:renewalState(c);
    const hasQuote=c.quote!=null&&c.quote!=="";
    return `<div class="dhead"><div><h2 style="margin:0">${isNew?"Add contract":esc(c.name)}</h2>${isNew?"":`<div class="muted small">${esc(c.vendor||"")}${c.category?" · "+esc(c.category):""}</div>${r&&r.label?`<div style="margin-top:6px"><span class="tag ${r.cls}">${esc(r.label)}</span></div>`:""}`}</div><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><form id="cform" class="stack" onsubmit="return false">
    <div class="sect"><h3>Contract</h3><div class="form">
      <label class="full">Name<input type="text" name="name" value="${esc(c.name)}" required ${ro}></label>
      <label>Vendor<input type="text" name="vendor" list="vendlist" value="${esc(c.vendor)}" ${ro}><datalist id="vendlist">${[...new Set(contracts().map(x=>x.vendor).filter(Boolean))].sort().map(v=>`<option value="${esc(v)}">`).join("")}</datalist></label>
      <label>Category<select name="category" ${ro}>${opt(CCATS,c.category||"Support Renewal")}</select></label>
      <label>Account / contract #<input type="text" name="account" value="${esc(c.account)}" ${ro}></label>
      <label>Circuit / service ID<input type="text" name="serviceId" value="${esc(c.serviceId)}" ${ro}></label>
      <label>Location<input type="text" name="location" list="loclistc" value="${esc(c.location)}" placeholder="All sites" ${ro}><datalist id="loclistc">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
      <label>Status<select name="status" ${ro}>${opt(CSTATUS,c.status||"Active")}</select></label>
      <label class="full">Contact<input type="text" name="contact" value="${esc(c.contact)}" placeholder="Account rep, phone, email, support portal" ${ro}></label>
    </div></div>
    <div class="sect"><h3>Cost and renewal</h3><div class="form">
      <label>Cost<input type="number" step="0.01" min="0" name="cost" value="${esc(c.cost)}" ${ro}></label>
      <label>Billing<select name="billing" ${ro}>${opt(BILLING,c.billing||"Annual")}</select></label>
      <label>Term (months)<input type="number" min="1" name="term" value="${esc(c.term)}" placeholder="12" ${ro}></label>
      <label>Auto-renews<select name="auto" ${ro}>${opt([["no","No"],["yes","Yes"]],c.auto?"yes":"no")}</select></label>
      <label>Last renewal<input type="date" name="last" value="${esc(c.last)}" ${ro}></label>
      <label>Next renewal<input type="date" name="next" value="${esc(c.next)}" ${ro}></label>
    </div>${isNew?"":`<p class="small" style="margin:10px 0 0">Counts as <b>${money(contractAnnual(c,THIS_YEAR))}</b> per year in the budget${hasQuote&&c.next?`, then <b>${money(contractAnnual(c,Number(String(c.next).slice(0,4))+1))}</b> per year after renewal at the quoted price`:""}.</p>`}</div>
    <div class="sect"><h3>Renewal quote</h3><div class="form">
      <label>Quote amount<input type="number" step="0.01" min="0" name="quote" value="${esc(c.quote)}" ${ro}></label>
      <label>Quote term (months)<input type="number" min="1" name="quoteTerm" value="${esc(c.quoteTerm)}" placeholder="Same as term" ${ro}></label>
      <label class="full">Quote notes<input type="text" name="quoteNotes" value="${esc(c.quoteNotes)}" placeholder="e.g. Quote 3 year 2026-2029, from CDW" ${ro}></label>
    </div></div>
    <div class="sect"><h3>Notes</h3><textarea name="notes" ${ro}>${esc(c.notes)}</textarea></div>
    </form></div>
    ${S.canWrite?`<div class="dfoot">${isNew?"<span></span>":`<div class="row"><button class="btn danger" data-act="delContract" data-id="${c._id}">Delete</button><button class="btn" data-act="renewContract" data-id="${c._id}" title="Roll the dates forward one term${hasQuote?" and apply the quote":""}">Mark renewed</button></div>`}<div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="saveContract" data-id="${isNew?"":c._id}">${isNew?"Add contract":"Save changes"}</button></div></div>`:""}`;
  }
  function readContractForm(){
    const f=$("#cform"),v=n=>f.elements[n].value.trim(),num=n=>v(n)===""?null:Number(v(n));
    return {name:v("name"),vendor:v("vendor")||null,category:v("category"),account:v("account")||null,serviceId:v("serviceId")||null,location:v("location")||null,
      status:v("status"),contact:v("contact")||null,cost:num("cost"),billing:v("billing"),term:num("term"),auto:v("auto")==="yes",last:v("last")||null,next:v("next")||null,
      quote:num("quote"),quoteTerm:num("quoteTerm"),quoteNotes:v("quoteNotes")||null,notes:v("notes")||null};
  }
  async function exportContracts(){
    const cols=[["name","Contract"],["vendor","Vendor"],["category","Category"],["account","Account / contract #"],["serviceId","Circuit / service ID"],["location","Location"],["status","Status"],["cost","Cost"],["billing","Billing"],["term","Term (months)"],["annual","Per year"],["last","Last renewal"],["next","Next renewal"],["auto","Auto-renews"],["quote","Quote"],["quoteTerm","Quote term (months)"],["quoteNotes","Quote notes"],["contact","Contact"],["notes","Notes"]];
    const q=v=>{v=v==null?"":String(v);return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};
    const rows=contracts().map(c=>({...c,annual:Math.round(contractAnnual(c,THIS_YEAR)*100)/100,auto:c.auto?"Yes":"No"}));
    const csv=[cols.map(x=>x[1]).join(","),...rows.map(r=>cols.map(([k])=>q(r[k])).join(","))].join("\n");
    const name=`it-contracts-${NOW.toISOString().slice(0,10)}.csv`;
    if(S.downloads){try{await S.downloads.save({filename:name,data:new Blob([csv],{type:"text/csv"})});return}catch(e){}}
    try{await navigator.clipboard.writeText(csv);toast("Copied the CSV to your clipboard.")}catch(e){toast("Export isn't available in this view.")}
  }
  function vSettings(){
    const C=cls();
    const keys=sortList(Object.keys(C).filter(k=>k!=="none").sort((a,b)=>C[a].label.localeCompare(C[b].label)),"classes");
    return `<h1>Settings</h1><p class="lede">Replacement classes drive the budget. Each device belongs to one class, which sets its standard replacement, unit cost, and lifecycle in years.</p>
    <section class="panel"><h2>Replacement classes</h2>
    <div class="tablewrap"><table><thead><tr>${th("classes","label","Class")}${th("classes","standard","Standard replacement")}${th("classes","cost","Unit cost","num")}${th("classes","life","Lifecycle (yrs)","num")}${hasSpareMin()?th("classes","spareMin","Keep on hand per site","num"):""}${th("classes","n","Devices","num")}</tr></thead><tbody>
    ${keys.map(k=>{const c=C[k];const n=assets().filter(a=>a.rclass===k).length;return `<tr><td><b>${esc(c.label)}</b>${c.from?`<span class="sub">${esc(c.from)}</span>`:""}</td>
    <td>${S.canWrite?`<input type="text" data-cls="${k}" data-fld="standard" value="${esc(c.standard)}" style="width:100%;min-width:200px">`:esc(c.standard)}</td>
    <td class="num">${S.canWrite?`<input type="number" step="1" min="0" data-cls="${k}" data-fld="cost" value="${c.cost}" style="width:110px;text-align:right">`:money(c.cost)}</td>
    <td class="num">${S.canWrite?`<input type="number" step="1" min="1" max="30" data-cls="${k}" data-fld="life" value="${c.life}" style="width:70px;text-align:right">`:c.life}</td>${hasSpareMin()?`<td class="num">${S.canWrite?`<input type="number" step="1" min="0" max="99" data-cls="${k}" data-fld="spareMin" value="${Number(c.spareMin)||""}" placeholder="0" style="width:70px;text-align:right" aria-label="Spares to keep on hand per site for ${esc(c.label)}">`:(Number(c.spareMin)||"—")}</td>`:""}<td class="num">${n}</td></tr>`}).join("")}
    </tbody></table></div><p class="muted small">The starting costs are examples; replace them with your own. Classes marked "Default estimate" are starting points; replace them with your own numbers.${hasSpareMin()?` "Keep on hand per site" sets how many spares of that class each site should have; the Spares page warns when a site that uses the class falls below it.`:""}</p></section>
    <section class="panel"><h2>Data</h2><p class="small">Export every asset with its age, budget year, and planned cost.</p><button class="btn" data-act="export">Export assets as CSV</button>
    ${S.appInfo?`<p class="small" style="margin-top:14px">${S.appInfo.missing&&S.appInfo.missing.length?`<span class="tag warn">Missing columns</span> These list columns weren't found, so their values can't be saved: ${S.appInfo.missing.map(esc).join("; ")}. Add them to the list or tell IT.`:`<span class="tag ok">All list columns found</span>`}${S.appInfo.version?` <span class="muted">· Web part version ${esc(S.appInfo.version)}</span>`:""}</p>${S.appInfo.logMissing?`<p class="small"><span class="tag warn">Asset log not set up</span> Run Install-AssetTracker.ps1 again on this site (it adds any missing lists) to turn on dated notes and assignment history. Device change history works without it.</p>`:`<p class="small"><span class="tag ok">Asset log connected</span></p>`}${hasContracts()?`<p class="small"><span class="tag ok">Contracts connected</span></p>`:`<p class="small"><span class="tag warn">Contracts not set up</span> Run Install-AssetTracker.ps1 again on this site (it adds any missing lists) to track renewals and add them to the budget.</p>`}`:""}
    <p class="small muted" style="margin-top:14px">Device ages are estimated from model generation unless you enter a purchase date. Changes save to the SharePoint lists right away. Other people see them the next time their page refreshes, within a minute.</p></section>`;
  }

  /* ---------------- spares ---------------- */
  const hasSpareMin=()=>!S.ext||S.ext.spareMinAvailable!==false;
  const isSpare=a=>a.status==="Spare";
  const spareQty=list=>list.reduce((s,a)=>s+(Number(a.qty)||1),0);
  function spareSince(a){
    const logs=(S.log||[]).filter(l=>l.assetId===a._id&&(l.type==="Unassigned"||/spare/i.test(l.details||""))).map(l=>String(l.when||"").slice(0,10)).filter(Boolean).sort();
    if(logs.length) return {date:logs[logs.length-1],src:"Unassigned"};
    if(a.purchaseDate) return {date:a.purchaseDate,src:"Purchased"};
    if(a.created) return {date:String(a.created).slice(0,10),src:"Added"};
    return null;
  }
  function shelfDays(a){const s=spareSince(a);if(!s)return null;const d=Math.round((TODAY-new Date(s.date+"T00:00:00"))/864e5);return d<0?0:d}
  const shelfText=d=>d==null?"—":d<45?`${d} days`:d<730?`${Math.round(d/30.4)} mo`:`${(d/365).toFixed(1).replace(/\.0$/,"")} yr`;
  function spareKind(a){const k=klass(a);return a.rclass&&a.rclass!=="none"?k.label:(a.type||a.category||"Other")}
  function spareFlags(a){
    const out=flags(a).filter(f=>f[0]!=="Terminated user");
    const d=shelfDays(a); if(d!=null&&d>180) out.push(["On shelf "+shelfText(d),"warn"]);
    return out;
  }
  /* sites below the "keep on hand" minimum set per replacement class */
  function spareShortfalls(){
    const out=[]; if(!hasSpareMin()) return out;
    const C=cls();
    for(const [k,c] of Object.entries(C)){
      const min=Number(c.spareMin)||0; if(!min||k==="none") continue;
      for(const l of locations()){
        const inUse=assets().some(a=>a.location===l&&a.rclass===k&&!inactive(a)&&!isSpare(a)); if(!inUse) continue;
        const have=spareQty(assets().filter(a=>a.location===l&&a.rclass===k&&isSpare(a)));
        if(have<min) out.push({loc:l,key:k,label:c.label,have,min});
      }
    }
    return out;
  }
  function spareFiltered(){
    return assets().filter(a=>isSpare(a)&&(S.sploc==="All"||a.location===S.sploc)&&(S.spcat==="All"||a.category===S.spcat))
      .sort((a,b)=>(a.location||"").localeCompare(b.location||"")||spareKind(a).localeCompare(spareKind(b))||(a.name||"").localeCompare(b.name||"",undefined,{numeric:true}));
  }
  function spareGrid(list){
    const kinds=[...new Set(list.map(spareKind))].sort();
    const locs=[...new Set(list.map(a=>a.location||"No location"))].sort((a,b)=>{const ia=SITE_ORDER.indexOf(a),ib=SITE_ORDER.indexOf(b);return (ia<0?99:ia)-(ib<0?99:ib)||a.localeCompare(b)});
    const cnt=(l,k)=>spareQty(list.filter(a=>(a.location||"No location")===l&&(k==null||spareKind(a)===k)));
    return {kinds,locs,cnt};
  }
  function vSpares(){
    const all=assets().filter(isSpare), list=spareFiltered();
    const catCounts=Object.fromEntries(CATS.map(c=>[c,spareQty(all.filter(a=>a.category===c))]));
    const g=spareGrid(list), short=spareShortfalls();
    const aging=list.filter(a=>{const d=shelfDays(a);return d!=null&&d>180}).length;
    const pastLife=list.filter(a=>flags(a).some(f=>f[0]==="Past lifecycle")).length;
    const warr=list.filter(a=>flags(a).some(f=>f[0]==="Warranty ends soon")).length;
    const locsAll=locations();
    return `<div class="spread"><div><h1>Spares</h1><p class="lede">Equipment on hand that isn't assigned: anything with the status Spare. Unassigning a device or adding one with no user puts it here.</p></div>
    <div class="row">${S.canWrite?`<button class="btn primary" data-act="newSpare">Add spare</button>`:""}<button class="btn" data-act="print" data-r="spares">Print</button></div></div>
    <section class="panel"><div class="metrics">
      <div><div class="metric">${spareQty(list)}</div><div class="metric-l">Spares${S.sploc!=="All"?" at "+esc(S.sploc):""}</div></div>
      <div><div class="metric">${aging}</div><div class="metric-l">On the shelf 6+ months</div></div>
      <div><div class="metric">${pastLife}</div><div class="metric-l">Past lifecycle</div></div>
      <div><div class="metric">${warr}</div><div class="metric-l">Warranty ending in 90 days</div></div>
      ${hasSpareMin()?`<div><div class="metric">${short.length}</div><div class="metric-l">Below minimum</div></div>`:""}
    </div></section>
    ${short.length?`<section class="panel"><h2>Running low</h2><ul class="attn">${short.map(s=>`<li><span><span class="tag warn">${s.have} of ${s.min}</span> <b>${esc(s.label)}</b> <span class="muted">at ${esc(s.loc)}</span></span><button data-act="sploc" data-l="${esc(s.loc)}">Show</button></li>`).join("")}</ul>
      <p class="muted small" style="margin:8px 0 0">Minimums are set per replacement class in Settings ("Keep on hand per site").</p></section>`:""}
    ${list.length?`<section class="panel"><h2>On hand by location</h2><div class="tablewrap" style="margin-top:6px"><table><thead><tr><th>Location</th>${g.kinds.map(k=>`<th class="num">${esc(k)}</th>`).join("")}<th class="num">Total</th></tr></thead><tbody>
      ${g.locs.map(l=>`<tr class="click" data-act="sploc" data-l="${esc(l)}"><td class="nm">${esc(l)}</td>${g.kinds.map(k=>{const n=g.cnt(l,k);return `<td class="num">${n||'<span class="muted">—</span>'}</td>`}).join("")}<td class="num"><b>${g.cnt(l)}</b></td></tr>`).join("")}
      </tbody></table></div></section>`:""}
    <div class="chips" style="margin:18px 0 12px"><button class="chip" data-act="sploc" data-l="All" aria-pressed="${S.sploc==="All"}">All locations<span class="n">${spareQty(all)}</span></button>
      ${locsAll.map(l=>{const n=spareQty(all.filter(a=>a.location===l));return `<button class="chip" data-act="sploc" data-l="${esc(l)}" aria-pressed="${S.sploc===l}">${esc(l)}<span class="n">${n}</span></button>`}).join("")}</div>
    <div class="row" style="margin-bottom:14px"><select data-ch="spcat" aria-label="Category"><option value="All">All categories</option>${CATS.filter(c=>catCounts[c]).map(c=>`<option ${S.spcat===c?"selected":""}>${c}</option>`).join("")}</select>
      <span class="muted small" style="margin-left:auto">${plural(list.length,"record")}</span></div>
    ${list.length?spareTable(list):`<div class="tablewrap"><div class="empty">No spares${S.sploc!=="All"?" at "+esc(S.sploc):""}. Set a device's status to Spare to add it here.${S.canWrite?` <button class="btn sm" data-act="newSpare">Add spare</button>`:""}</div></div>`}`;
  }
  function spareTable(list){
    return `<div class="tablewrap"><table><thead><tr>${selHead(list.map(a=>a._id))}<th>Asset</th>${S.sploc==="All"?"<th>Location</th>":""}<th>Model</th><th>Serial</th><th class="num">Qty</th><th>Age</th><th>Spare since</th><th>Flags</th></tr></thead><tbody>
    ${list.map(a=>{const s=spareSince(a),d=shelfDays(a);return `<tr class="click${S.sel.has(a._id)?" sel":""}" data-act="open" data-id="${a._id}">${selCell(a)}<td><span class="nm">${esc(a.name)}</span><span class="sub">${esc(spareKind(a))}</span></td>
    ${S.sploc==="All"?`<td>${esc(a.location||"—")}</td>`:""}<td>${esc(a.model||"—")}</td><td class="small">${esc(a.serial||"")||'<span class="muted">—</span>'}</td><td class="num">${Number(a.qty)||1}</td><td>${lifeBar(a)}</td>
    <td>${s?`${esc(fmtDate(s.date))}<span class="sub">${esc(s.src)} · ${shelfText(d)}</span>`:'<span class="muted">—</span>'}</td><td class="small">${spareFlags(a).map(([t,c])=>`<span class="tag ${c}">${esc(t)}</span>`).join(" ")}</td></tr>`}).join("")}</tbody></table></div>`;
  }

  /* ---------------- printing ---------------- */
  const PRINT_CSS=`*{box-sizing:border-box}body{font:10pt/1.35 "Segoe UI",Arial,sans-serif;color:#111;margin:0;padding:24px 28px;background:#fff}
  .bar{position:sticky;top:0;display:flex;gap:8px;justify-content:flex-end;padding:8px 0 14px;background:#fff}
  .bar button{font:inherit;padding:7px 14px;border:1px solid #888;border-radius:6px;background:#fff;cursor:pointer}.bar button.p{background:#0E7480;border-color:#0E7480;color:#fff}
  header{border-bottom:2px solid #111;padding-bottom:8px;margin-bottom:14px}h1{font-size:16pt;margin:0}.meta{color:#555;font-size:9pt;margin-top:3px}
  h2{font-size:12pt;margin:18px 0 6px;padding-bottom:3px;border-bottom:1px solid #bbb}h2 .n{color:#666;font-weight:normal;font-size:9.5pt}
  h3{font-size:10.5pt;margin:12px 0 4px}
  .sum{display:flex;flex-wrap:wrap;gap:6px 28px;margin:4px 0 10px}.sum div b{display:block;font-size:13pt}.sum div span{color:#555;font-size:8.5pt}
  table{width:100%;border-collapse:collapse;margin:4px 0 8px}th,td{text-align:left;padding:4px 6px;border-bottom:1px solid #ddd;vertical-align:top}
  th{font-size:8.5pt;text-transform:uppercase;letter-spacing:.03em;color:#444;border-bottom:1.5px solid #555}
  td.num,th.num{text-align:right;white-space:nowrap}.sub{display:block;color:#666;font-size:8.5pt}tr.tot td{font-weight:bold;border-top:1.5px solid #555;border-bottom:0}
  tr.gh td{background:#f1f1f1;font-weight:bold}.muted{color:#777}.flag{display:inline-block;border:1px solid #999;border-radius:3px;padding:0 4px;font-size:8pt;margin:1px 2px 1px 0}
  .sign{display:grid;grid-template-columns:1fr 1fr;gap:28px 40px;margin-top:40px}.sign div{border-top:1px solid #111;padding-top:4px;font-size:9pt;color:#444}
  .note{color:#555;font-size:9pt}footer{margin-top:16px;color:#777;font-size:8.5pt}
  .pb{break-before:page}thead{display:table-header-group}tr{break-inside:avoid}h2,h3{break-after:avoid}
  @page{margin:0.5in}@media print{body{padding:0}.bar{display:none}}`;
  function printDoc(title,meta,body){
    const w=window.open("","_blank");
    if(!w){toast("Allow pop-ups for this site to print.");return}
    const when=new Date().toLocaleString("en-US",{month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"});
    w.document.open();
    w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>${PRINT_CSS}</style></head><body>
    <div class="bar"><button class="p" id="pp">Print or save as PDF</button><button id="pc">Close</button></div>
    <header><h1>${esc(title)}</h1><div class="meta">${meta?meta+" · ":""}Printed ${esc(when)}</div></header>${body}<footer>IT Asset Tracker</footer></body></html>`);
    w.document.close();
    const pp=w.document.getElementById("pp"),pc=w.document.getElementById("pc");
    if(pp) pp.addEventListener("click",()=>w.print());
    if(pc) pc.addEventListener("click",()=>w.close());
    w.focus();
  }
  /* cols: [label, a=>html, "num"?] */
  function ptable(cols,rows,foot){
    return `<table><thead><tr>${cols.map(c=>`<th${c[2]?' class="num"':""}>${c[0]}</th>`).join("")}</tr></thead><tbody>
    ${rows.map(r=>`<tr>${cols.map(c=>`<td${c[2]?' class="num"':""}>${c[1](r)}</td>`).join("")}</tr>`).join("")}${foot||""}</tbody></table>`;
  }
  const pflags=list=>list.map(([t])=>`<span class="flag">${esc(t)}</span>`).join("");
  const pAge=a=>{const g=age(a);return g==null?'<span class="muted">—</span>':g+" yr"};
  const pName=a=>`<b>${esc(a.name)}</b><span class="sub">${esc(a.type||a.category||"")}</span>`;
  const pUser=a=>{const p=person(a);return a.assignedTo?esc(a.assignedTo)+(p&&p.status!=="Employed"?`<span class="sub">${esc(p.status)}</span>`:""):'<span class="muted">—</span>'};
  const pModel=a=>esc([a.make&&!String(a.model||"").toLowerCase().includes(String(a.make).toLowerCase())?a.make:"",a.model].filter(Boolean).join(" ")||"—");
  function assetCols(list,opts={}){
    const c=[["Asset",pName]];
    if(!opts.noLoc) c.push(["Location",a=>esc(a.location||"—")]);
    if(!opts.noUser) c.push(["Assigned to",pUser]);
    c.push(["Model",pModel],["Serial",a=>esc(a.serial||"")]);
    if(list.some(a=>a.ip)) c.push(["IP",a=>esc(a.ip||"")]);
    c.push(["Age",pAge,"num"],["Status",a=>esc(a.status||"")]);
    if(!opts.noFlags&&list.some(a=>flags(a).length)) c.push(["Flags",a=>pflags(flags(a))]);
    return c;
  }
  function groupBy(list,fn,order){
    const m=new Map();for(const x of list){const k=fn(x)||"—";if(!m.has(k))m.set(k,[]);m.get(k).push(x)}
    const keys=[...m.keys()];if(order)keys.sort(order);return keys.map(k=>[k,m.get(k)]);
  }
  const siteOrder=(a,b)=>{const ia=SITE_ORDER.indexOf(a),ib=SITE_ORDER.indexOf(b);return (ia<0?99:ia)-(ib<0?99:ib)||a.localeCompare(b)};

  function printAssets(){
    const list=sortList(filteredAssets(),"assets"), f=S.f;
    const parts=[f.cat!=="All"?f.cat:"All categories",f.loc!=="All"?f.loc:"all locations",f.status!=="All"?f.status:"",f.q?`matching "${f.q}"`:"",
      f.attn?({true:"needs attention",term:"terminated users",unknown:"age unknown",warranty:"warranty",noserial:"no serial number"}[f.attn]):""].filter(Boolean);
    let body;
    if(f.loc==="All"&&new Set(list.map(a=>a.location)).size>1){
      body=groupBy(list,a=>a.location,siteOrder).map(([l,g],i)=>`<h2${i?' class="pb"':""}>${esc(l)} <span class="n">${plural(g.length,"asset")}</span></h2>${ptable(assetCols(g,{noLoc:true}),g)}`).join("");
    } else body=ptable(assetCols(list,{noLoc:f.loc!=="All"}),list);
    printDoc("Assets",esc(parts.join(" · "))+` · ${plural(list.length,"asset")}`,list.length?body:`<p class="note">No assets match these filters.</p>`);
  }
  function printSite(){
    const l=S.site; if(!l) return;
    const A=assets().filter(a=>a.location===l), live=A.filter(a=>!inactive(a)), yt=yearTotals(l);
    const staff=[...S.people.values()].filter(p=>p.location===l&&p.status==="Employed").length;
    const sc=hasContracts()?contracts().filter(c=>activeContract(c)&&c.location===l):[];
    let body=`<div class="sum"><div><b>${live.length}</b><span>Active assets</span></div><div><b>${staff}</b><span>Current staff</span></div>
      <div><b>${live.filter(a=>flags(a).some(f=>f[0]==="Past lifecycle")).length}</b><span>Past lifecycle</span></div>
      <div><b>${spareQty(A.filter(isSpare))}</b><span>Spares</span></div><div><b>${money(yt[Y0].base+yt[Y0].reserve)}</b><span>${Y0} plan, incl. reserve</span></div></div>`;
    for(const c of CATS){const g=sortList(A.filter(a=>a.category===c),"assets");if(!g.length)continue;
      body+=`<h2>${c} <span class="n">${g.length}</span></h2>${ptable(assetCols(g,{noLoc:true}),g)}`}
    if(sc.length) body+=`<h2>Contracts <span class="n">${sc.length}</span></h2>`+contractPTable(sc,{noLoc:true});
    printDoc(l,"Location report",body);
  }
  function printPerson(id){
    const p=S.people.get(id); if(!p) return;
    const devs=sortList(assets().filter(a=>a.personId===id),"assets");
    const body=`<div class="sum"><div><b>${esc(p.title||"—")}</b><span>Title</span></div><div><b>${esc(p.login||"—")}</b><span>Login</span></div>
      <div><b>${esc(p.location||"—")}</b><span>Location</span></div><div><b>${esc(p.status||"—")}</b><span>Status</span></div></div>
      <h2>Assigned equipment <span class="n">${devs.length}</span></h2>
      ${devs.length?ptable([["Asset",pName],["Make / model",pModel],["Serial / service tag",a=>esc(a.serial||"")],["Qty",a=>Number(a.qty)||1,"num"],["Location",a=>esc(a.location||"")],["Status",a=>esc(a.status||"")]],devs):`<p class="note">No equipment assigned.</p>`}
      <p class="note" style="margin-top:14px">I acknowledge receipt of the equipment listed above. It remains company property and will be returned to IT when no longer needed or on leaving the company.</p>
      <div class="sign"><div>Employee signature</div><div>Date</div><div>Issued by (IT)</div><div>Date</div></div>`;
    printDoc(p.name,"Equipment assignment",body);
  }
  function printSpares(){
    const list=spareFiltered(), g=spareGrid(list), short=spareShortfalls();
    let body="";
    if(short.length) body+=`<h2>Running low</h2>`+ptable([["Location",s=>esc(s.loc)],["Class",s=>esc(s.label)],["On hand",s=>s.have,"num"],["Minimum",s=>s.min,"num"]],short);
    if(list.length){
      body+=`<h2>On hand by location</h2><table><thead><tr><th>Location</th>${g.kinds.map(k=>`<th class="num">${esc(k)}</th>`).join("")}<th class="num">Total</th></tr></thead><tbody>
      ${g.locs.map(l=>`<tr><td><b>${esc(l)}</b></td>${g.kinds.map(k=>`<td class="num">${g.cnt(l,k)||"—"}</td>`).join("")}<td class="num"><b>${g.cnt(l)}</b></td></tr>`).join("")}
      ${g.locs.length>1?`<tr class="tot"><td>Total</td>${g.kinds.map(k=>`<td class="num">${spareQty(list.filter(a=>spareKind(a)===k))}</td>`).join("")}<td class="num">${spareQty(list)}</td></tr>`:""}</tbody></table>`;
      const cols=[["Asset",a=>`<b>${esc(a.name)}</b><span class="sub">${esc(spareKind(a))}</span>`],["Model",pModel],["Serial",a=>esc(a.serial||"")],["Qty",a=>Number(a.qty)||1,"num"],["Age",pAge,"num"],
        ["Spare since",a=>{const s=spareSince(a);return s?`${esc(fmtDate(s.date))}<span class="sub">${esc(s.src)} · ${shelfText(shelfDays(a))}</span>`:"—"}],["Flags",a=>pflags(spareFlags(a))]];
      body+=groupBy(list,a=>a.location||"No location",siteOrder).map(([l,gl],i)=>`<h2${i?' class="pb"':""}>${esc(l)} <span class="n">${plural(spareQty(gl),"item")}</span></h2>${ptable(cols,gl)}`).join("");
    } else body+=`<p class="note">No spares${S.sploc!=="All"?" at "+esc(S.sploc):""}.</p>`;
    printDoc("Spares",[S.sploc!=="All"?esc(S.sploc):"All locations",S.spcat!=="All"?esc(S.spcat):""].filter(Boolean).join(" · ")+` · ${plural(spareQty(list),"item")}`,body);
  }
  function printBudget(){
    const yr=S.byear, loc=S.bloc, label=yr==="all"?`${Y0}–${YEARS[4]}`:String(yr), recYears=yr==="all"?YEARS:[yr];
    const rows=[];for(const a of S.assets.values()){if(loc!=="All"&&a.location!==loc)continue;const b=budget(a);if(!b)continue;if(yr==="all"?(b.year&&YEARS.includes(b.year)):b.year===yr)rows.push({a,b})}
    const funded=rows.filter(r=>r.b.kind==="base"||r.b.kind==="reserve"), review=rows.filter(r=>r.b.kind==="review");
    const base=funded.filter(r=>r.b.kind==="base").reduce((s,r)=>s+r.b.cost,0), res=funded.filter(r=>r.b.kind==="reserve").reduce((s,r)=>s+r.b.cost,0);
    const rec=hasContracts()?recYears.reduce((s,y)=>s+recurringTotal(y,loc),0):0;
    const planTxt=a=>({auto:"Auto",replace:"Replace",reserve:"Reserve",exclude:"Skip"}[a.plan||"auto"])+(a.plan==="replace"||a.plan==="reserve"?" "+(a.planYear||Y0):"");
    let body=`<div class="sum"><div><b>${money(base)}</b><span>Base budget</span></div><div><b>${money(res)}</b><span>Conditional reserve</span></div><div><b>${money(base+res)}</b><span>Hardware envelope</span></div>
      ${hasContracts()?`<div><b>${money(rec)}</b><span>Recurring contracts</span></div><div><b>${money(base+res+rec)}</b><span>Hardware and contracts</span></div>`:""}<div><b>${review.length}</b><span>Held for review</span></div></div>`;
    if(yr==="all"){const t=yearTotals(loc);body+=ptable([["Year",y=>y],["Devices",y=>t[y].n,"num"],["Base",y=>money(t[y].base),"num"],["Reserve",y=>money(t[y].reserve),"num"],
      ...(hasContracts()?[["Contracts",y=>money(recurringTotal(y,loc)),"num"]]:[]),["Total",y=>`<b>${money(t[y].base+t[y].reserve+(hasContracts()?recurringTotal(y,loc):0))}</b>`,"num"]],YEARS)}
    const groups=groupBy(funded,r=>r.a.rclass).map(([k,g])=>[k,g,g.reduce((s,r)=>s+r.b.cost,0)]).sort((x,y)=>y[2]-x[2]);
    if(groups.length){
      body+=`<h2>Hardware replacements <span class="n">${plural(funded.length,"device")}</span></h2><table><thead><tr><th>Asset</th>${yr==="all"?"<th>Year</th>":""}<th>Model</th><th class="num">Age</th><th>Why</th><th>Plan</th><th class="num">Cost</th></tr></thead><tbody>`;
      for(const [k,g,tot] of groups){const c=cls()[k]||{label:k};
        body+=`<tr class="gh"><td colspan="${yr==="all"?6:5}">${esc(c.label)}${c.standard?` <span class="sub" style="display:inline">· ${esc(c.standard)}${c.cost?` · ${money(c.cost)} each`:""}</span>`:""}</td><td class="num">${money(tot)}</td></tr>`;
        body+=g.sort((x,y)=>(x.b.year-y.b.year)||((x.a.year||0)-(y.a.year||0))).map(({a,b})=>`<tr><td><b>${esc(a.name)}</b><span class="sub">${esc([a.location,a.assignedTo].filter(Boolean).join(" · "))}</span></td>${yr==="all"?`<td>${b.year}</td>`:""}
          <td>${pModel(a)}</td><td class="num">${pAge(a)}</td><td>${esc(b.why)}${b.kind==="reserve"?'<span class="sub">Reserve</span>':""}</td><td>${planTxt(a)}</td><td class="num">${money(b.cost)}${a.quote!=null&&a.quote!==""?'<span class="sub">Quote</span>':""}</td></tr>`).join("")}
      body+=`<tr class="tot"><td colspan="${yr==="all"?6:5}">Hardware total</td><td class="num">${money(base+res)}</td></tr></tbody></table>`;
    } else body+=`<p class="note">Nothing is due for replacement in ${label}.</p>`;
    if(review.length) body+=`<h3>Held for review (not budgeted)</h3>`+ptable([["Asset",r=>`<b>${esc(r.a.name)}</b>`],["Assigned to",r=>pUser(r.a)],["Model",r=>pModel(r.a)],["Why",r=>esc(r.b.why)],["Cost if replaced",r=>money(r.b.cost),"num"]],review);
    if(hasContracts()){
      const list=contracts().filter(c=>activeContract(c)&&(loc==="All"||c.location===loc)).map(c=>({c,v:recYears.reduce((s,y)=>s+contractAnnual(c,y),0)})).filter(x=>x.v);
      if(list.length){
        body+=`<h2 class="pb">Recurring contracts <span class="n">${plural(list.length,"contract")}</span></h2><table><thead><tr><th>Contract</th><th>Vendor</th><th>Billing</th><th>Next renewal</th><th class="num">${recYears.length>1?"Total":"Per year"}</th></tr></thead><tbody>`;
        for(const [cat,g] of groupBy(list,x=>x.c.category||"Other").map(([k,g])=>[k,g,g.reduce((s,x)=>s+x.v,0)]).sort((a,b)=>b[2]-a[2]).map(x=>[x[0],x[1]])){
          body+=`<tr class="gh"><td colspan="4">${esc(cat)}</td><td class="num">${money(g.reduce((s,x)=>s+x.v,0))}</td></tr>`+g.sort((a,b)=>b.v-a.v).map(({c,v})=>`<tr><td>${esc(c.name)}</td><td>${esc(c.vendor||"")}</td><td>${c.cost!=null?esc(billingText(c)):"—"}</td><td>${c.next?esc(fmtDate(c.next)):"—"}</td><td class="num">${money(v)}</td></tr>`).join("")}
        body+=`<tr class="tot"><td colspan="4">Contracts total</td><td class="num">${money(rec)}</td></tr></tbody></table>`;
      }
    }
    printDoc(`Budget ${label}`,loc!=="All"?esc(loc):"All locations",body);
  }
  function contractPTable(list,opts={}){
    const total=list.reduce((s,c)=>s+contractAnnual(c,THIS_YEAR),0);
    const cols=[["Contract",c=>`<b>${esc(c.name)}</b><span class="sub">${esc([c.account?"Acct "+c.account:"",c.serviceId||""].filter(Boolean).join(" · "))}</span>`],["Vendor",c=>esc(c.vendor||"")]];
    if(!opts.noCat) cols.push(["Category",c=>esc(c.category||"")]);
    if(!opts.noLoc) cols.push(["Location",c=>esc(c.location||"All sites")]);
    cols.push(["Cost",c=>c.cost!=null?esc(billingText(c)):"—"],["Per year",c=>money(contractAnnual(c,THIS_YEAR)),"num"],
      ["Next renewal",c=>c.next?esc(fmtDate(c.next))+(c.auto?'<span class="sub">Auto-renews</span>':""):"—"],["Quote",c=>c.quote!=null&&c.quote!==""?money(c.quote):"—","num"]);
    const n=cols.length;
    return ptable(cols,sortList(list,"contracts"),`<tr class="tot"><td colspan="${n-3}">Total per year</td><td class="num">${money(total)}</td><td colspan="2"></td></tr>`);
  }
  function printContracts(){
    if(!hasContracts()) return;
    const q=S.cq.trim().toLowerCase();
    const list=contracts().filter(c=>(S.cshow==="All"||activeContract(c))&&(S.ccat==="All"||c.category===S.ccat)&&
      (!q||[c.name,c.vendor,c.account,c.serviceId,c.contact,c.notes,c.location,c.category].join(" ").toLowerCase().includes(q)));
    const total=list.filter(activeContract).reduce((s,c)=>s+contractAnnual(c,THIS_YEAR),0);
    let body=`<div class="sum"><div><b>${money(total)}</b><span>Per year, current pricing</span></div><div><b>${money(total/12)}</b><span>Monthly equivalent</span></div>
      <div><b>${list.filter(c=>{const r=renewalState(c);return r&&r.d>=0&&r.d<=90}).length}</b><span>Renewing in 90 days</span></div></div>`;
    body+=S.ccat!=="All"?contractPTable(list,{noCat:true}):groupBy(list,c=>c.category||"Other",(a,b)=>CCATS.indexOf(a)-CCATS.indexOf(b)).map(([k,g])=>`<h2>${esc(k)} <span class="n">${g.length}</span></h2>${contractPTable(g,{noCat:true})}`).join("");
    printDoc("Contracts and renewals",[S.ccat!=="All"?esc(S.ccat):"All categories",S.cshow==="All"?"including cancelled":"active only",q?`matching "${esc(S.cq)}"`:""].filter(Boolean).join(" · "),list.length?body:`<p class="note">No contracts match.</p>`);
  }
  function printReport(kind,id){
    try{({assets:printAssets,site:printSite,spares:printSpares,budget:printBudget,contracts:printContracts}[kind]||(()=>printPerson(id)))()}
    catch(e){console.error(e);toast("Couldn't build the printable view.")}
  }

  /* ---------------- render ---------------- */
  let raf;function render(){cancelAnimationFrame(raf);raf=requestAnimationFrame(doRender)}
  function doRender(){
    renderNav();
    const main=$("#main");
    if(S.loadError){main.innerHTML=`<div class="banner"><b>The asset lists could not be loaded.</b><br>${esc(S.loadError)}</div>`;return}
    if(!(S.loaded.assets&&S.loaded.classes)){main.innerHTML=`<div class="empty">Loading assets…</div>`;return}
    const ae=document.activeElement, fid=ae&&ae.id, pos=ae&&ae.selectionStart;
    if(S.view==="contracts"&&!hasContracts()&&S.loaded.assets&&S.ext!==undefined) S.view="overview";
    const html={overview:vOverview,assets:vAssets,sites:vSites,spares:vSpares,people:vPeople,budget:vBudget,contracts:vContracts,actions:vActions,settings:vSettings}[S.view]();
    main.innerHTML=(S.canWrite===false?`<div class="banner small">You have view-only access. Ask a site owner for Edit permission on the IT lists to make changes.</div>`:"")+html;
    renderBulkBar();
    document.querySelectorAll('input[data-some="1"]').forEach(el=>{el.indeterminate=true});
    if(fid&&/^[A-Za-z][\w-]*$/.test(fid)){const el=ROOT.querySelector("#"+fid);if(el){el.focus();try{el.setSelectionRange(pos,pos)}catch(e){}}}
  }
  function go(v){S.view=v;render();try{if(ROOT.getBoundingClientRect().top<0)ROOT.scrollIntoView({block:"start"})}catch(e){}}

  /* ---------------- drawer: asset ---------------- */
  function openDrawer(html){const bb=$("#bulkbar");if(bb)bb.hidden=true;const d=$("#drawer");d.innerHTML=html;d.classList.add("open");d.setAttribute("aria-hidden","false");$("#scrim").classList.add("open");setTimeout(()=>{const f=d.querySelector("input,select,textarea,button");f&&f.focus()},50)}
  function closeDrawer(){const d=$("#drawer");d.classList.remove("open");d.setAttribute("aria-hidden","true");$("#scrim").classList.remove("open");S.drawer=null;renderBulkBar();if(S.aiCtl){S.aiCtl.abort();S.aiCtl=null}}
  const opt=(list,v)=>list.map(x=>{const [val,lab]=Array.isArray(x)?x:[x,x];return `<option value="${esc(val)}" ${String(v??"")===String(val)?"selected":""}>${esc(lab)}</option>`}).join("");
  function assetForm(a,isNew){
    const C=cls(); const ro=!S.canWrite?"disabled":"";
    const pplOpts=[["","— Not assigned —"],...[...S.people.entries()].sort((x,y)=>x[1].name.localeCompare(y[1].name)).map(([id,p])=>[id,`${p.name}${p.status!=="Employed"?" ("+p.status+")":""}`])];
    const b=isNew?null:budget(a);
    const sug=a.suggestion;
    return `<div class="dhead"><div><h2 style="margin:0">${isNew?"Add asset":esc(a.name)}</h2>${isNew?"":`<div class="muted small">${esc(a.type||a.category)} · ${esc(a.location||"")}</div><div style="margin-top:6px">${flagTags(a)}</div>`}</div><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><form id="aform" class="stack" onsubmit="return false">
    <div class="sect"><h3>Details</h3><div class="form">
      <label>Name / ID<input type="text" name="name" value="${esc(a.name)}" required ${ro}></label>
      <label>Category<select name="category" ${ro}>${opt(CATS,a.category||"Computers")}</select></label>
      <label>Type<input type="text" name="type" value="${esc(a.type)}" placeholder="Laptop, Switch, UPS…" ${ro}></label>
      <label>Location<input type="text" name="location" list="loclist" value="${esc(a.location)}" ${ro}><datalist id="loclist">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
      <label>Make<input type="text" name="make" value="${esc(a.make)}" ${ro}></label>
      <label>Model<input type="text" name="model" value="${esc(a.model)}" ${ro}></label>
      <label>Serial / service tag<input type="text" name="serial" value="${esc(a.serial)}" ${ro}></label>
      <label>IP address<input type="text" name="ip" value="${esc(a.ip)}" ${ro}></label>
      <label>Assigned person<select name="personId" ${ro}>${opt(pplOpts,a.personId||"")}</select>${!isNew&&a.personId&&S.people.has(a.personId)?`<span class="fieldlinks"><button type="button" data-act="gotoPerson" data-id="${a.personId}">View person</button>${S.canWrite?`<button type="button" data-act="unassign" data-id="${a._id}">Unassign</button>`:""}</span>`:""}</label>
      <label>Or shared use / role<input type="text" name="assignedToText" value="${a.personId?"":esc(a.assignedTo)}" placeholder="Forklift, Front office…" ${ro}></label>
      <label>Status<select name="status" ${ro}>${opt(STATUSES,a.status||"In Use")}</select></label>
      <label>Quantity<input type="number" name="qty" min="1" value="${a.qty||1}" ${ro}></label>
      ${a.category==="Connectivity"||a.monthlyCost!=null?`<label>Monthly cost<input type="number" step="0.01" name="monthlyCost" value="${esc(a.monthlyCost)}" ${ro}></label><label>Contract end<input type="text" name="contractEnd" value="${esc(a.contractEnd)}" ${ro}></label>`:""}
      ${a.transition?`<label class="full">Transition plan<input type="text" name="transition" value="${esc(a.transition)}" ${ro}></label>`:""}
    </div></div>
    <div class="sect"><h3>Lifecycle and budget</h3><div class="form">
      <label>Replacement class<select name="rclass" ${ro}>${opt([["","Auto-detect"],...Object.keys(C).sort((x,y)=>C[x].label.localeCompare(C[y].label)).map(k=>[k,C[k].label])],a.rclass||"")}</select></label>
      <label>In-service year<input type="number" name="year" min="1995" max="${THIS_YEAR+1}" value="${esc(a.year)}" ${ro}><span class="tiny muted">${esc(a.yearSource||"Enter a purchase date or year")}</span></label>
      <label>Purchase date<input type="date" name="purchaseDate" value="${esc(a.purchaseDate)}" ${ro}></label>
      <label>Warranty end<input type="date" name="warrantyEnd" value="${esc(a.warrantyEnd)}" ${ro}></label>
      <label>Plan<select name="plan" ${ro}>${opt([["auto","Auto (by age)"],["replace","Replace in base budget"],["reserve","Hold as reserve"],["exclude","Skip — don't budget"]],a.plan||"auto")}</select></label>
      <label>Plan year<select name="planYear" ${ro}>${opt([["","—"],...YEARS.map(String)],a.planYear||"")}</select></label>
      <label>Quoted unit price<input type="number" step="0.01" name="quote" value="${esc(a.quote)}" placeholder="${C[a.rclass]?.cost?"Standard "+money(C[a.rclass].cost):""}" ${ro}></label>
      <label>Quote source<input type="text" name="quoteSource" value="${esc(a.quoteSource)}" placeholder="Provantage, 10/2026" ${ro}></label>
    </div>${b?`<p class="small" style="margin:10px 0 0">${b.kind==="skip"?"Not in the budget.":b.kind==="review"?`Held for review: ${esc(b.why)}.`:`Budgeted in <b>${b.year}</b> as ${b.kind==="base"?"base":"reserve"} at ${money(b.cost)}. ${esc(b.why)}.`}</p>`:""}</div>
    ${!isNew&&a.rclass&&a.rclass!=="none"?`<div class="sect" id="replsect"><h3>Replacement${replButton(a,S.sample&&sug?"Ask again":"Find replacement")}</h3>
      <div id="replout">${sug?suggestionHTML(sug,a):`<p class="small muted">${S.sample?"Ask Claude for a current replacement model, then check live prices with the shopping buttons.":`Find replacement asks Google to search B&amp;H, Provantage, or CDW for a replacement for ${esc(replModelText(a))}.${C[a.rclass]?.standard?` To price your standard, ${esc(C[a.rclass].standard)}, directly:`:""}`}</p>${vendorButtons(C[a.rclass]?.standard||a.model)}`}</div></div>`:""}
    ${!isNew&&S.ext?`<div class="sect" id="histsect"><h3>History${S.canWrite&&S.ext.logAvailable?`<button class="btn sm" data-act="noteOpen" data-id="${a._id}">Add note</button>`:""}</h3>
      <div id="noteform"></div><div id="histout"><p class="small muted">Loading history…</p></div>
      ${S.ext.logAvailable?"":`<p class="tiny muted">Dated notes and assignment history turn on once the IT Asset Log list is added to the site.</p>`}</div>`:""}
    <div class="sect"><h3>Notes</h3><textarea name="notes" ${ro} placeholder="${!isNew&&S.ext&&S.ext.logAvailable?"Standing notes about this device. Use Add note in History for dated entries.":""}">${esc(a.notes)}</textarea>${a.action?`<p class="small muted">Imported action: ${esc(a.action)}</p>`:""}</div>
    </form></div>
    ${S.canWrite?`<div class="dfoot">${isNew?"<span></span>":`<div class="row"><button class="btn danger" data-act="delAsset" data-id="${a._id}">Delete</button>${a.category==="Computers"?`<button class="btn" data-act="repurpose" data-id="${a._id}" title="Reassign and rename this device">Repurpose</button>`:""}</div>`}<div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="saveAsset" data-id="${isNew?"":a._id}">${isNew?"Add asset":"Save changes"}</button></div></div>`:""}`;
  }
  /* Google search of the preferred resellers for a replacement of this device */
  function replModelText(a){
    const model=String(a.model||"").trim(), make=String(a.make||"").replace(/\//g," ").trim();
    if(model){
      const first=make.split(/\s+/)[0]||"";
      return first&&!model.toLowerCase().startsWith(first.toLowerCase())&&!/^(vm|generic)$/i.test(make)?`${make} ${model}`:model;
    }
    const what=(klass(a).label||a.type||"").toLowerCase();
    return make&&!/^(vm|generic)$/i.test(make)?`${make} ${what}`:what;
  }
  function replSearchUrl(a){
    const q=`search bhphotovideo provantage or cdw for a replacement for ${replModelText(a)}`;
    return "https://www.google.com/search?q="+encodeURIComponent(q);
  }
  function replButton(a,label,extraCls){
    if(S.sample) return `<button class="btn sm${extraCls||""}" data-act="findrepl" data-id="${a._id}">${label}</button>`;
    return `<a class="btn sm${extraCls||""}" target="_blank" rel="noopener" href="${replSearchUrl(a)}" title="Search Google for a recommended replacement">${label}</a>`;
  }
  function vendorButtons(q){
    if(!q) return "";
    const e=encodeURIComponent(q);
    return `<div class="vendors"><a class="btn sm" target="_blank" rel="noopener" href="https://www.provantage.com/service/searchsvcs?QUERY=${e}">Provantage</a>
    <a class="btn sm" target="_blank" rel="noopener" href="https://www.bhphotovideo.com/c/search?q=${e}">B&amp;H</a>
    <a class="btn sm" target="_blank" rel="noopener" href="https://www.cdw.com/search/?key=${e}">CDW</a>
    <button class="btn sm ghost" data-act="copy" data-t="${esc(q)}">Copy model</button></div>`;
  }
  function suggestionHTML(s,a){
    return `<div class="ai"><div class="model">${esc(s.model)}</div>
    ${s.specs?`<div class="small">${esc(s.specs)}</div>`:""}
    <p class="small" style="margin:8px 0">${esc(s.reason||"")}</p>
    <div class="row small"><span>Estimated price <b>${money(s.estPrice)}</b></span><span class="muted">· suggested ${esc((s.at||"").slice(0,10))}</span>
    ${S.canWrite&&s.estPrice?`<button class="btn sm" data-act="usePrice" data-id="${a._id}">Use as budget price</button>`:""}</div>
    ${s.alternatives&&s.alternatives.length?`<p class="tiny muted" style="margin:8px 0 0">Alternatives: ${s.alternatives.map(esc).join(", ")}</p>`:""}
    ${vendorButtons(s.searchQuery||s.model)}
    <p class="tiny muted" style="margin:8px 0 0">AI estimates can lag current pricing. Check a reseller and enter your quote above.</p></div>`;
  }
  function readForm(){
    const fd=new FormData($("#aform")),o={};for(const [k,v] of fd.entries()) o[k]=typeof v==="string"?v.trim():v;
    const num=v=>v===""||v==null?null:Number(v);
    const d={name:o.name,category:o.category,type:o.type||null,location:o.location||null,make:o.make||null,model:o.model||null,serial:o.serial||null,ip:o.ip||null,
      status:o.status,qty:num(o.qty)||1,rclass:o.rclass,purchaseDate:o.purchaseDate||null,warrantyEnd:o.warrantyEnd||null,plan:o.plan,
      planYear:num(o.planYear),quote:num(o.quote),quoteSource:o.quoteSource||null,notes:o.notes||null};
    if("monthlyCost" in o){d.monthlyCost=num(o.monthlyCost);d.contractEnd=o.contractEnd||null}
    if("transition" in o) d.transition=o.transition||null;
    d.personId=o.personId||null;
    d.assignedTo=d.personId?(S.people.get(d.personId)?.name||null):(o.assignedToText||null);
    if(d.purchaseDate){d.year=Number(d.purchaseDate.slice(0,4));d.yearSource="Purchase date"}
    else if(o.year){d.year=Number(o.year);const prev=S.drawer&&S.assets.get(S.drawer.id);d.yearSource=prev&&prev.year===d.year&&prev.yearSource?prev.yearSource:"Entered manually"}
    else {d.year=null;d.yearSource=null}
    if((d.plan==="replace"||d.plan==="reserve")&&!d.planYear) d.planYear=Y0;
    return d;
  }
  function guessClass(d){
    const m=(d.model||"").toLowerCase(),t=(d.type||"").toLowerCase();
    if(d.category==="Computers"){if(/ipad|tablet/.test(m+t))return"tablet";if(/zbook|precision/.test(m))return"workstation";if(/laptop|book|probook|latitude|thinkpad/.test(m+t))return"laptop";return"desktop"}
    if(/switch/.test(t))return"access_switch";if(/access point|\bap\b/.test(t))return"ap";if(/host|server/.test(t)&&!/virtual|vm/.test(t))return"standalone_host";
    if(/label/.test(t))return"label_printer";if(/printer/.test(t))return"printer";if(/scan/.test(t))return"scanner";if(/ups/.test(t))return"ups";if(/phone/.test(t))return"phone";
    return"none";
  }

  /* ---------------- drawer: person, action ---------------- */
  function bulkForm(){
    const ids=[...S.sel], C=cls(), K="__keep";
    const keepOpt=`<option value="${K}" selected>Leave unchanged</option>`;
    const pplOpts=[...S.people.entries()].sort((x,y)=>x[1].name.localeCompare(y[1].name)).map(([id,p])=>`<option value="${id}">${esc(p.name)}${p.status!=="Employed"?" ("+esc(p.status)+")":""}</option>`).join("");
    const names=ids.slice(0,6).map(id=>esc(S.assets.get(id)?.name||"")).join(", ")+(ids.length>6?` and ${ids.length-6} more`:"");
    return `<div class="dhead"><div><h2 style="margin:0">Edit ${plural(ids.length,"asset")}</h2><div class="muted small" style="margin-top:4px">${names}</div></div><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><p class="bulkhint">Only fields you change are applied. Everything left as "Leave unchanged" or blank stays as it is on each asset.</p>
    <form id="bform" class="stack" onsubmit="return false">
    <div class="sect"><h3>Details</h3><div class="form">
      <label>Location<input type="text" name="location" list="loclistb" placeholder="Leave unchanged"><datalist id="loclistb">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
      <label>Status<select name="status">${keepOpt}${STATUSES.map(x=>`<option>${x}</option>`).join("")}</select></label>
      <label>Category<select name="category">${keepOpt}${CATS.map(x=>`<option>${x}</option>`).join("")}</select></label>
      <label>Type<input type="text" name="type" placeholder="Leave unchanged"></label>
      <label class="full">Assigned person<select name="personId">${keepOpt}<option value="">Unassign (no person)</option>${pplOpts}</select></label>
    </div></div>
    <div class="sect"><h3>Lifecycle and budget</h3><div class="form">
      <label>Replacement class<select name="rclass">${keepOpt}${Object.keys(C).sort((x,y)=>C[x].label.localeCompare(C[y].label)).map(k=>`<option value="${k}">${esc(C[k].label)}</option>`).join("")}</select></label>
      <label>In-service year<input type="number" name="year" min="1995" max="${THIS_YEAR+1}" placeholder="Leave unchanged"></label>
      <label>Plan<select name="plan">${keepOpt}<option value="auto">Auto (by age)</option><option value="replace">Replace in base budget</option><option value="reserve">Hold as reserve</option><option value="exclude">Skip — don't budget</option></select></label>
      <label>Plan year<select name="planYear">${keepOpt}${YEARS.map(y=>`<option>${y}</option>`).join("")}</select></label>
    </div></div>
    </form></div>
    <div class="dfoot"><span class="muted small" id="bprog"></span><div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="bulkApply">Apply to ${plural(ids.length,"asset")}</button></div></div>`;
  }
  async function runBulk(ids,fn,label){
    let ok=0,fail=0;const prog=$("#bprog");
    for(let i=0;i<ids.length;i++){
      if(prog) prog.textContent=`${label} ${i+1} of ${ids.length}…`;
      if(await fn(ids[i])) ok++; else fail++;
    }
    return {ok,fail};
  }
  /* ---- history, notes, repurpose ---- */
  const fmtDate=v=>{if(!v)return "";const d=/^\d{4}-\d{2}-\d{2}$/.test(String(v))?new Date(v+"T00:00:00"):new Date(v);return isNaN(d)?String(v):d.toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"})};
  const fmtWhen=v=>{if(!v)return "";const d=new Date(v);return isNaN(d)?String(v):d.toLocaleString("en-US",{month:"short",day:"numeric",year:"numeric",hour:"numeric",minute:"2-digit"})};
  function timelineHTML(items){
    if(!items.length) return `<p class="small muted">No history yet.</p>`;
    return `<ul class="tl">${items.map(it=>`<li class="${it.kind==="note"?"note":it.key?"key":""}"><div class="when">${esc(fmtWhen(it.when))}${it.who?` · ${esc(it.who)}`:""}</div>
      <div class="what">${it.title?`<b>${esc(it.title)}</b>`:""}${it.text?`${it.title?" ":""}${esc(it.text)}`:""}${it.changes&&it.changes.length?`<ul>${it.changes.map(c=>`<li>${esc(c.label)}: ${c.from==null?'<span class="muted">blank</span>':esc(c.from)} <span class="arrow">→</span> ${c.to==null?'<span class="muted">blank</span>':esc(c.to)}</li>`).join("")}</ul>`:""}</div></li>`).join("")}</ul>`;
  }
  async function loadHistory(id){
    const out=$("#histout"); if(!out||!S.ext||!S.ext.history) return;
    try{
      const items=await S.ext.history(id);
      if(!S.drawer||S.drawer.id!==id) return;
      const o=$("#histout"); if(o) o.innerHTML=timelineHTML(items);
    }catch(e){console.error(e);const o=$("#histout");if(o)o.innerHTML=`<p class="small muted">History couldn't be loaded: ${esc(e.message||"unknown error")}</p>`}
  }
  function noteFormHTML(id){
    return `<div class="notebox"><div class="form">
      <label>Entry type<select id="noteType">${["Note","Repair","Reimage","Other"].map(x=>`<option>${x}</option>`).join("")}</select></label><span></span>
      <label class="full">Details<textarea id="noteText" placeholder="What happened? e.g. Replaced battery under warranty, case #12345"></textarea></label></div>
      <div class="row" style="justify-content:flex-end;margin-top:8px"><button class="btn sm" data-act="noteCancel">Cancel</button><button class="btn sm primary" data-act="noteSave" data-id="${id}">Save note</button></div></div>`;
  }
  function siteCode(loc,fallbackName){
    const counts={};
    for(const x of S.assets.values()){
      if(x.category!=="Computers"||x.location!==loc) continue;
      const m=String(x.name||"").match(/^([A-Za-z0-9]{2,4})-/); if(m){const k=m[1].toLowerCase();counts[k]=(counts[k]||0)+1}
    }
    const best=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0];
    if(best) return best[0];
    const m=String(fallbackName||"").match(/^([A-Za-z0-9]{2,4})-/);
    return m?m[1].toLowerCase():"pc";
  }
  function suggestName(a,p,loc){
    if(!p) return a.name;
    const sm=String(a.name||"").match(/-?(\d{2})$/), suffix=sm?sm[1]:"11";
    const parts=String(p.name||"").replace(/\s+[A-Z]\.\s+/," ").trim().split(/\s+/);
    const first=((parts[0]||"")[0]||"x").toLowerCase(), last=(parts[parts.length-1]||"").toLowerCase().replace(/[^a-z]/g,"")||"user";
    const head=`${siteCode(loc,a.name)}-${first}-`, tail=`-${suffix}`, room=15-head.length-tail.length;   // Windows names max 15 characters
    return head+last.slice(0,Math.max(3,room))+tail;
  }
  function repurposeForm(a){
    const old=person(a);
    const pplOpts=[...S.people.entries()].filter(([id,p])=>p.status!=="Terminated"||id===a.personId).sort((x,y)=>x[1].name.localeCompare(y[1].name))
      .map(([id,p])=>`<option value="${id}">${esc(p.name)}${p.location?" · "+esc(p.location):""}</option>`).join("");
    return `<div class="dhead"><div><h2 style="margin:0">Repurpose ${esc(a.name)}</h2><div class="muted small" style="margin-top:4px">${esc(a.model||a.type||"")} · currently ${a.assignedTo?esc(a.assignedTo):"unassigned"}</div></div><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><p class="bulkhint">Reassigns and renames the device in one step. The old name and user stay in the device's History.</p>
    <form id="rform" class="form" onsubmit="return false">
      <label class="full">New user<select name="personId"><option value="">No user (put in spare pool)</option>${pplOpts}</select></label>
      <label>New computer name<input type="text" name="name" value="${esc(a.name)}" maxlength="64"><span class="tiny muted" id="rnameHint">Suggested from your naming convention. Edit if needed.</span></label>
      <label>Location<input type="text" name="location" list="loclistr" value="${esc(a.location)}"><datalist id="loclistr">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
      <label>Status<select name="status">${opt(STATUSES,"In Use")}</select></label><span></span>
      <label class="full">Note for the history<textarea name="note" placeholder="e.g. Reimaged, new user account set up">${old&&old.status==="Terminated"?esc(old.name+" left the company."):""}</textarea></label>
      ${old&&old.status!=="Terminated"?`<label class="full" style="flex-direction:row;align-items:center;gap:8px;color:var(--ink)"><input type="checkbox" name="markTerm"> Mark ${esc(old.name)} as Terminated</label>`:""}
    </form></div>
    <div class="dfoot"><button class="btn" data-act="open" data-id="${a._id}">Back</button><div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="repurposeApply" data-id="${a._id}">Repurpose device</button></div></div>`;
  }
  function personForm(p,isNew){
    return `<div class="dhead"><h2 style="margin:0">${isNew?"Add person":"Edit "+esc(p.name)}</h2><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><form id="pform" class="form" onsubmit="return false">
    <label class="full">Full name<input type="text" name="name" value="${esc(p.name)}" required></label>
    <label>Job title<input type="text" name="title" value="${esc(p.title)}"></label>
    <label>Login / username<input type="text" name="login" value="${esc(p.login)}"></label>
    <label>Location<input type="text" name="location" list="loclist2" value="${esc(p.location)}"><datalist id="loclist2">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
    <label>Status<select name="status">${opt(PSTATUS,p.status||"Employed")}</select></label>
    <label class="full" style="flex-direction:row;gap:16px;flex-wrap:wrap">${["o365","mfa","vpn","av"].map(k=>`<span class="row" style="gap:6px;color:var(--ink)"><input type="checkbox" name="${k}" ${p[k]?"checked":""}> ${k==="o365"?"O365":k.toUpperCase()}</span>`).join("")}</label>
    </form>
    <div class="sect" style="margin-top:22px"><h3>Devices<button class="btn sm" data-act="addDevRow">Add device</button></h3>
      ${!isNew?(()=>{const d=assets().filter(a=>a.personId===p._id);return d.length?`<p class="small" style="margin:0 0 10px">Currently assigned: ${d.map(a=>esc(a.name)).join(", ")}</p>`:`<p class="small muted" style="margin:0 0 10px">No devices assigned yet.</p>`})():""}
      <div id="devrows"></div>
      <label class="small muted" style="display:flex;flex-direction:column;gap:4px;font-weight:600;margin-top:6px">Or assign an existing unassigned computer
        <select id="pAssignExisting"><option value="">— None —</option>${assets().filter(a=>!a.personId&&a.category==="Computers"&&!a.assignedTo&&!inactive(a)).sort((x,y)=>x.name.localeCompare(y.name)).map(a=>`<option value="${a._id}">${esc(a.name)} — ${esc(a.model||a.type||"")} (${esc(a.location||"")})</option>`).join("")}</select></label>
    </div></div>
    <div class="dfoot">${isNew?"<span></span>":`<button class="btn danger" data-act="delPerson" data-id="${p._id}">Delete</button>`}<div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="savePerson" data-id="${isNew?"":p._id}">${isNew?"Add person":"Save changes"}</button></div></div>`;
  }
  let devRowN=0;
  function devRowHTML(){
    const n=++devRowN;
    return `<div class="devrow panel" data-row="${n}" style="padding:12px 14px;margin-bottom:10px">
    <div class="form">
      <label>Computer name / ID<input type="text" data-f="name" placeholder="nr-j-smith-11"></label>
      <label>Device type<select data-f="kind">${opt([["Laptop","Laptop"],["Desktop","Desktop"],["Tablet","Tablet"],["Other","Other"]],"Laptop")}</select></label>
      <label>Model<input type="text" data-f="model" placeholder="HP EliteBook 8 G1a 14"></label>
      <label>Serial / service tag<input type="text" data-f="serial"></label>
      <label>Setup<input type="text" data-f="detail" placeholder="Laptop / Dock / Monitors"></label>
      <label>Purchase date<input type="date" data-f="purchaseDate"></label>
    </div>
    <div style="text-align:right;margin-top:8px"><button class="btn ghost sm danger" data-act="rmDevRow">Remove</button></div></div>`;
  }
  function readDevRows(){
    return [...OVL.querySelectorAll("#devrows .devrow")].map(r=>{const o={};r.querySelectorAll("[data-f]").forEach(el=>o[el.dataset.f]=el.value.trim());return o})
      .filter(o=>o.name||o.model||o.serial);
  }
  function actionForm(t,isNew){
    return `<div class="dhead"><h2 style="margin:0">${isNew?"Add action":"Edit action"}</h2><button class="btn ghost" data-act="close" aria-label="Close">✕</button></div>
    <div class="dbody"><form id="tform" class="form" onsubmit="return false">
    <label class="full">Action<input type="text" name="title" value="${esc(t.title)}" required></label>
    <label class="full">Details<textarea name="detail">${esc(t.detail)}</textarea></label>
    <label>Priority<select name="priority">${opt(["Critical","High","Medium","Low"],t.priority||"Medium")}</select></label>
    <label>Location<input type="text" name="location" list="loclist3" value="${esc(t.location)}"><datalist id="loclist3">${locations().map(l=>`<option value="${esc(l)}">`).join("")}</datalist></label>
    <label>Area<input type="text" name="area" value="${esc(t.area)}" placeholder="Network, Servers…"></label>
    <label>Status<select name="status">${opt(["Open","Done"],t.status||"Open")}</select></label>
    </form></div>
    <div class="dfoot">${isNew?"<span></span>":`<button class="btn danger" data-act="delAction" data-id="${t._id}">Delete</button>`}<div class="row"><button class="btn" data-act="close">Cancel</button><button class="btn primary" data-act="saveAction" data-id="${isNew?"":t._id}">${isNew?"Add action":"Save changes"}</button></div></div>`;
  }
  const formObj=sel=>{const f=$(sel),o={};for(const el of f.elements){if(!el.name)continue;o[el.name]=el.type==="checkbox"?el.checked:el.value.trim()}return o};

  /* ---------------- AI ---------------- */
  function sampleErr(e){
    const c=e&&e.code;
    if(c==="not_granted") return "AI access wasn't allowed. Reload the page to be asked again.";
    if(c==="rate_limited") return "Too many requests right now. Wait a minute and try again.";
    if(c==="cancelled") return "Stopped.";
    if(c==="invalid_json") return "Claude's answer couldn't be read. Try again.";
    return "Claude couldn't answer right now. Try again in a moment.";
  }
  async function findReplacement(id){
    const a=S.assets.get(id); if(!a||!S.sample) return;
    if(!S.drawer||S.drawer.id!==id){S.drawer={type:"asset",id};openDrawer(assetForm({...a,_id:id},false))}
    const out=$("#replout"); if(!out) return;
    const sec=$("#replsect");sec&&sec.scrollIntoView({block:"start"});
    out.innerHTML=`<div class="ai"><span class="muted">Thinking… this can take up to a minute.</span></div>`;
    const k=klass(a), p=person(a);
    const prompt=`You help the IT team at ${opts.orgName||"an organization"} plan hardware replacements. Today's date is ${NOW.toISOString().slice(0,10)}.
  Device to replace:
  - Type: ${a.type||a.category} (${a.category})
  - Make/model: ${[a.make,a.model].filter(Boolean).join(" ")||"unknown"}
  - In service since: ${a.year||"unknown"} (${a.yearSource||"no date"})
  - Role/usage: ${p&&p.title?p.title:(a.assignedTo||a.detail||"shared")}${a.detail?"; setup: "+a.detail:""}
  - Notes: ${(a.notes||"").slice(0,300)}
  Company standard for this class ("${k.label}"): ${k.standard||"none set"}, budgeted at $${k.cost} each.
  Recommend ONE replacement model that is currently sold through US business resellers such as Provantage, B&H, and CDW. Prefer the company standard when it fits this role; if the role needs something different (for example CAD work needs a workstation GPU), say why. Use the most recent model generation you know of.
  Respond with JSON only, no other text:
  {"model":"exact product name","searchQuery":"short search string a reseller site would match","estPrice":number (typical USD street price, no symbols),"specs":"key specs on one line","reason":"one or two plain sentences","alternatives":["model","model"]}`;
    const ctl=new AbortController();S.aiCtl=ctl;
    try{
      const r=await S.sample.json(prompt,{signal:ctl.signal,modelTier:"default"});
      if(!r||!r.model) throw {code:"invalid_json"};
      const sug={model:String(r.model),searchQuery:String(r.searchQuery||r.model),estPrice:Number(r.estPrice)||null,specs:r.specs?String(r.specs):null,reason:r.reason?String(r.reason):null,alternatives:Array.isArray(r.alternatives)?r.alternatives.slice(0,3).map(String):[],at:new Date().toISOString()};
      if(S.canWrite) await write("assets",id,{suggestion:sug,...stamp()});
      const o=$("#replout"); if(o&&S.drawer&&S.drawer.id===id) o.innerHTML=suggestionHTML(sug,{...a,_id:id});
    }catch(e){const o=$("#replout");if(o)o.innerHTML=`<div class="ai"><p class="small">${esc(sampleErr(e))}</p>${vendorButtons(k.standard||a.model)}</div>`}
    finally{if(S.aiCtl===ctl)S.aiCtl=null}
  }
  function mdLite(t){
    const lines=esc(t).split(/\n/);let h="",inl=false;
    for(let l of lines){
      l=l.replace(/\*\*(.+?)\*\*/g,"<b>$1</b>");
      if(/^\s*[-*•]\s+/.test(l)){if(!inl){h+="<ul>";inl=true}h+="<li>"+l.replace(/^\s*[-*•]\s+/,"")+"</li>";continue}
      if(inl){h+="</ul>";inl=false}
      if(/^#{1,4}\s+/.test(l)) h+="<h4>"+l.replace(/^#+\s+/,"")+"</h4>"; else if(l.trim()) h+="<p>"+l+"</p>";
    }
    return h+(inl?"</ul>":"");
  }
  const COPILOT_URL="https://m365.cloud.microsoft/chat";
  function buildPlanPrompt(yr,loc){
    const yrs=yr==="all"?YEARS:[yr];
    const scope=`${yr==="all"?`${YEARS[0]} to ${YEARS[4]}`:yr}${loc!=="All"?`, ${loc} only`:", all locations"}`;
    const groups={};let review=0,unknown=0;
    for(const a of S.assets.values()){
      if(loc!=="All"&&a.location!==loc) continue;
      const b=budget(a);if(!b||!yrs.includes(b.year)) continue;
      if(b.kind==="review"){review++;continue}
      if(b.kind!=="base"&&b.kind!=="reserve") continue;
      if(b.why.startsWith("Age unknown")) unknown++;
      const key=[b.year,klass(a).label,b.kind].join("|");
      const g=groups[key]=groups[key]||{year:b.year,label:klass(a).label,kind:b.kind,count:0,cost:0,sites:{},models:{},past:0};
      g.count++;g.cost+=b.cost;g.sites[a.location||"?"]=(g.sites[a.location||"?"]||0)+1;if(a.model)g.models[a.model]=(g.models[a.model]||0)+1;if(b.why.startsWith("Past"))g.past++;
    }
    const top=(o,n)=>Object.entries(o).sort((x,y)=>y[1]-x[1]).slice(0,n).map(([k,v])=>`${k} ${v}`).join(", ");
    const lines=Object.values(groups).sort((a,b)=>a.year-b.year||(a.kind<b.kind?-1:1)||b.cost-a.cost)
      .map(g=>`- ${g.year}, ${g.label}, ${g.kind==="base"?"base budget":"conditional reserve"}: ${plural(g.count,"device")}, ${money(g.cost)}${g.past?` (${g.past} already past lifecycle)`:""}. Sites: ${top(g.sites,5)}. Models: ${top(g.models,4)||"unknown"}.`);
    const tot=yearTotals(loc);
    const hc=hasContracts();
    const totals=YEARS.map(y=>`- ${y}: hardware base ${money(tot[y].base)}, reserve ${money(tot[y].reserve)}${hc?`, contracts ${money(recurringTotal(y,loc))}`:""}`);
    let contractLines=[];
    if(hc){
      contractLines=contracts().filter(c=>activeContract(c)&&(loc==="All"||c.location===loc)).sort((a,b)=>(a.next||"9999").localeCompare(b.next||"9999")).map(c=>{
        const r=renewalState(c);
        return `- ${c.name} (${c.category}${c.vendor?", "+c.vendor:""}${c.location?", "+c.location:""}): ${c.cost!=null?billingText(c):"cost unknown"}, about ${money(contractAnnual(c,THIS_YEAR))} per year`+
          `${c.next?`, next renewal ${fmtDate(c.next)}${r&&r.label?` (${r.label.toLowerCase()})`:""}`:""}${c.auto?", auto-renews":""}`+
          `${c.quote!=null&&c.quote!==""?`, renewal quote ${money(c.quote)}${c.quoteTerm?` for ${c.quoteTerm} months`:""}`:""}.`;
      });
    }
    const sections=['"Fund first" (critical risks such as legacy core switches or end-of-life hosts)','"Stagger" (what can move to later years to smooth spending, with rough dollar shifts)','"Verify before buying" (data to confirm first)'];
    if(hc) sections.push('"Contracts and renewals" (overdue or upcoming renewals, quote increases worth negotiating, consolidation)');
    sections.push('"Savings ideas" (for example, redeploying devices from former employees as spares)');
    return [
  `Act as an IT budget advisor. Review the IT replacement${hc?" and contracts":""} plan below for ${opts.orgName||"our organization"}${SITE_ORDER.length?`, which has ${SITE_ORDER.length} site${SITE_ORDER.length===1?"":"s"}: ${SITE_ORDER.join(", ")}`:""}. Today is ${fmtDate(TODAY)}.`,
  ``,
  `Scope: ${scope}.`,
  ``,
  `Totals by year:`,...totals,
  ``,
  `Hardware replacements due in scope (device ages are mostly estimated from model generation; base = planned, reserve = needs validation first):`,
  ...(lines.length?lines:["- None in scope."]),
  `- Held for review because they are assigned to former employees: ${plural(review,"device")}.`,
  `- Carried as reserve because their age is unknown: ${plural(unknown,"device")}.`,
  ...(hc?[``,`Contracts and recurring costs (multi-year costs are spread evenly per year):`,...(contractLines.length?contractLines:["- None in scope."])]:[]),
  ``,
  `Write a practical review in under 400 words with these short sections: ${sections.join("; ")}. Use plain sentences and short bullets, no tables. Base the advice only on the data above, and say when something needs to be confirmed.`
    ].join("\n");
  }
  function copilotPanel(yr,loc,label){
    return `<section class="panel"><div class="spread"><div><h2 style="margin-bottom:2px">Copilot planning review</h2><p class="muted small" style="margin:0">Builds a planning prompt from this budget${hasContracts()?" and your contracts":""} (${esc(String(label))}${loc!=="All"?", "+esc(loc):""}), copies it, and opens Microsoft 365 Copilot. Paste it there with Ctrl+V and send.</p></div>
    <div class="row"><button class="btn" data-act="copyPrompt">Copy prompt</button><button class="btn primary" data-act="openCopilot">Open in Copilot</button></div></div>
    <details style="margin-top:12px"><summary class="small" style="cursor:pointer;color:var(--accent);font-weight:600">Preview the prompt</summary><textarea id="cpPrompt" readonly style="margin-top:8px;min-height:240px;font-size:13px">${esc(buildPlanPrompt(yr,loc))}</textarea></details></section>`;
  }
  async function copyPrompt(opening){
    const txt=buildPlanPrompt(S.byear,S.bloc);
    try{
      await navigator.clipboard.writeText(txt);
      if(opening) window.open(COPILOT_URL,"_blank","noopener");
      toast(opening?"Prompt copied. In Copilot, paste with Ctrl+V and send.":"Prompt copied");
    }catch(e){
      const d=$("#cpPrompt");
      if(d){d.closest("details").open=true;d.focus();d.select()}
      if(opening) window.open(COPILOT_URL,"_blank","noopener");
      toast("Press Ctrl+C to copy the selected prompt"+(opening?", then paste it into Copilot.":"."));
    }
  }
  async function aiReview(){
    if(!S.sample||S.aiBusy) return;
    const yr=S.byear, loc=S.bloc;
    const yrs=yr==="all"?YEARS:[yr];
    const summary={};let review=0,unknown=0;
    for(const a of S.assets.values()){
      if(loc!=="All"&&a.location!==loc) continue;
      const b=budget(a);if(!b||!yrs.includes(b.year)) continue;
      if(b.kind==="review"){review++;continue}
      if(b.kind!=="base"&&b.kind!=="reserve") continue;
      if(b.why.startsWith("Age unknown")) unknown++;
      const key=`${b.year} | ${klass(a).label} | ${b.kind}`;
      const s=summary[key]=summary[key]||{count:0,cost:0,locations:{},models:{},overdue:0};
      s.count++;s.cost+=b.cost;s.locations[a.location]=(s.locations[a.location]||0)+1;if(a.model)s.models[a.model]=(s.models[a.model]||0)+1;if(b.why.startsWith("Past"))s.overdue++;
    }
    const tot=yearTotals(loc);
    const prompt=buildPlanPrompt(yr,loc);
    const _unusedPrompt=`You are advising the network administrator of ${opts.orgName||"an organization"} on an IT hardware replacement budget.
  Scope: ${yr==="all"?`${YEARS[0]}-${YEARS[4]}`:yr}${loc!=="All"?`, location ${loc}`:", all locations"}.
  Totals per year (base/reserve): ${YEARS.map(y=>`${y}: $${Math.round(tot[y].base)} / $${Math.round(tot[y].reserve)}`).join("; ")}.
  Line items (year | class | base or reserve): ${JSON.stringify(summary)}
  Devices held for review because they are assigned to terminated employees: ${review}. Devices with unknown age carried as reserve: ${unknown}.
  Ages are mostly estimated from model generation. Base = planned; reserve = conditional on validation.
  Write a practical review under 300 words with these short sections: "Fund first" (critical risks such as legacy core switches, end-of-life hosts), "Stagger" (what can move to later years to smooth spend, with rough dollar shifts), "Verify before buying" (data to collect), and "Savings ideas" (e.g., redeploy devices from terminated users as spares). Use plain sentences and "- " bullets. No tables.`;
    S.aiBusy=true;S.aiReview="";render();
    const ctl=new AbortController();S.aiRevCtl=ctl;
    setTimeout(()=>{const o=$("#aiout");if(o&&!S.aiReview)o.innerHTML='<p class="muted">Thinking… the first words can take up to a minute.</p>'},30);
    try{
      const r=await S.sample(prompt,{signal:ctl.signal,modelTier:"default",onText:({text})=>{S.aiReview=text;const o=$("#aiout");if(o)o.innerHTML=mdLite(text)}});
      S.aiReview=r.text;
    }catch(e){S.aiReview=(e&&e.text?e.text+"\n\n":"")+sampleErr(e)}
    finally{S.aiBusy=false;S.aiRevCtl=null;render()}
  }

  /* ---------------- CSV ---------------- */
  async function exportCSV(){
    const cols=["name","category","type","location","assignedTo","personStatus","make","model","serial","ip","status","year","yearSource","age","replacementClass","lifecycle","plan","budgetYear","budgetKind","budgetCost","quote","warrantyEnd","monthlyCost","contractEnd","notes"];
    const rows=assets().map(a=>{const b=budget(a)||{};const k=klass(a);const p=person(a);return {...a,personStatus:p?p.status:"",age:age(a)??"",replacementClass:k.label,lifecycle:k.life||"",budgetYear:b.year||"",budgetKind:b.kind||"",budgetCost:b.cost!=null?Math.round(b.cost):""}});
    const q=v=>{v=v==null?"":String(v);return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v};
    const csv=[cols.join(","),...rows.map(r=>cols.map(c=>q(r[c])).join(","))].join("\n");
    const name=`it-assets-${NOW.toISOString().slice(0,10)}.csv`;
    if(S.downloads){try{await S.downloads.save({filename:name,data:new Blob([csv],{type:"text/csv"})});return}catch(e){if(e&&e.code==="declined")return}}
    try{await navigator.clipboard.writeText(csv);toast("Download isn't available here, so the CSV was copied to your clipboard.")}catch(e){toast("Export isn't available in this view.")}
  }

  /* ---------------- events ---------------- */
  on("click",async e=>{
    const t=e.target.closest("[data-act]"); if(!t) return;
    const act=t.dataset.act,id=t.dataset.id;
    switch(act){
      case "nav": go(t.dataset.v); break;
      case "selrow": { e.stopPropagation(); const sid=t.dataset.id; S.sel.has(sid)?S.sel.delete(sid):S.sel.add(sid); render(); break; }
      case "selall": { const ids=t.dataset.ids.split(",").filter(Boolean); const all=ids.every(x=>S.sel.has(x));
        ids.forEach(x=>all?S.sel.delete(x):S.sel.add(x)); render(); break; }
      case "selclear": S.sel.clear(); render(); break;
      case "ccat": S.ccat=t.dataset.c; render(); break;
      case "cshow": S.cshow=t.dataset.s; render(); break;
      case "exportContracts": exportContracts(); break;
      case "newContract": S.drawer={type:"contract",id:null}; openDrawer(contractForm({category:S.ccat!=="All"?S.ccat:"Support Renewal",billing:"Annual",status:"Active"},true)); break;
      case "openContract": { const c=S.contracts.get(id); if(!c) return; S.drawer={type:"contract",id}; openDrawer(contractForm({...c,_id:id},false)); break; }
      case "saveContract": { const d=readContractForm(); if(!d.name){toast("Give the contract a name.");return}
        if(!d.term) d.term=d.billing==="Monthly"?1:d.billing==="Multi-year"?36:12;
        t.disabled=true;
        const ok=id?await write("contracts",id,{...d,...stamp()}):await write("contracts","c-"+Date.now().toString(36),{...d,...stamp()},"set");
        t.disabled=false;
        if(ok){toast(id?"Saved":"Contract added");closeDrawer()} break; }
      case "delContract": if(confirm("Delete this contract? To keep its history, set Status to Cancelled instead.")){ if(await write("contracts",id,null,"delete")){toast("Deleted");closeDrawer()} } break;
      case "renewContract": { const c=S.contracts.get(id); if(!c) return;
        const hasQuote=c.quote!=null&&c.quote!=="";
        const term=Number((hasQuote&&c.quoteTerm)||c.term||12);
        const last=c.next||TODAY.toISOString().slice(0,10);
        const next=c.billing==="Monthly"&&!c.next?null:addMonths(last,term);
        const cost=hasQuote?Number(c.quote):c.cost;
        if(!confirm(`Mark ${c.name} renewed?\n\nLast renewal: ${fmtDate(last)}\nNext renewal: ${next?fmtDate(next):"none"}\nCost: ${money(cost)}${hasQuote?" (from the quote)":""}`)) break;
        const d={last,next,cost,term,status:"Active",quote:null,quoteTerm:null,quoteNotes:null,
          billing:c.billing==="Monthly"?"Monthly":term>=24?"Multi-year":"Annual",
          notes:[c.notes,`Renewed ${fmtDate(TODAY)}: ${money(cost)} for ${term} months${hasQuote&&c.quoteNotes?" ("+c.quoteNotes+")":""}.`].filter(Boolean).join("\n"),...stamp()};
        t.disabled=true; const ok=await write("contracts",id,d); t.disabled=false;
        if(ok){toast(`${c.name} renewed through ${next?fmtDate(next):"next term"}`);openDrawer(contractForm({...S.contracts.get(id),_id:id},false))}
        break; }
      case "gotoPerson": { e.stopPropagation(); if(!S.people.has(id)) break; closeDrawer(); S.personId=id; S.pq=""; S.ploc="All"; S.pstat="All"; go("people"); break; }
      case "unassign": { e.stopPropagation(); const a=S.assets.get(id); if(!a||!a.personId) break;
        const who=a.assignedTo||"this person", wasInUse=a.status==="In Use", nm=a.name;
        t.disabled=true;
        const ok=await write("assets",id,{personId:null,assignedTo:null,status:wasInUse?"Spare":a.status,...stamp()});
        t.disabled=false;
        if(ok){ toast(`${nm} unassigned from ${who}${wasInUse?" and set to Spare":""}`);
          if(S.drawer&&S.drawer.type==="asset"&&S.drawer.id===id){ openDrawer(assetForm({...S.assets.get(id),_id:id},false)); loadHistory(id); } }
        break; }
      case "unassignAll": { const ids=assets().filter(a=>a.personId===id).map(a=>a._id); if(!ids.length) break;
        const pn=(S.people.get(id)||{}).name||"this person";
        if(!confirm(`Unassign ${plural(ids.length,"device")} from ${pn} and set them to Spare?`)) break;
        t.disabled=true; let n=0;
        for(const aid of ids){ const a=S.assets.get(aid); if(a&&await write("assets",aid,{personId:null,assignedTo:null,status:a.status==="In Use"?"Spare":a.status,...stamp()})) n++; }
        toast(`Unassigned ${plural(n,"device")} from ${pn}`); break; }
      case "bulkEdit": if(S.sel.size){ S.drawer={type:"bulk",id:null}; openDrawer(bulkForm()); } break;
      case "bulkApply": {
        const f=$("#bform"), K="__keep", ids=[...S.sel].filter(x=>S.assets.has(x));
        const v=n=>f.elements[n].value.trim();
        const base={};
        if(v("location")) base.location=v("location");
        if(v("status")!==K) base.status=v("status");
        if(v("category")!==K) base.category=v("category");
        if(v("type")) base.type=v("type");
        if(v("rclass")!==K) base.rclass=v("rclass");
        if(v("year")){ const y=Number(v("year")); if(y<1995||y>THIS_YEAR+1){toast("Enter an in-service year between 1995 and "+(THIS_YEAR+1)+".");return} base.year=y; base.yearSource="Entered manually"; }
        if(v("plan")!==K) base.plan=v("plan");
        if(v("planYear")!==K) base.planYear=Number(v("planYear"));
        const pv=v("personId");
        if(pv!==K){ if(pv){const p=S.people.get(pv);base.personId=pv;base.assignedTo=p?p.name:null;} else {base.personId=null;base.assignedTo=null;} }
        if(!Object.keys(base).length){toast("Change at least one field to apply.");return}
        t.disabled=true;
        const r=await runBulk(ids,async id=>{
          const a=S.assets.get(id); if(!a) return false;
          const d={...base,...stamp()};
          if((d.plan==="replace"||d.plan==="reserve")&&!("planYear" in d)) d.planYear=a.planYear||(budget(a)||{}).year||Y0;
          return !!(await write("assets",id,d));
        },"Saving");
        t.disabled=false;
        toast(r.fail?`Updated ${r.ok}; ${r.fail} couldn't be saved.`:`Updated ${plural(r.ok,"asset")}`);
        if(!r.fail){S.sel.clear();closeDrawer();} render(); break; }
      case "bulkDelete": {
        const ids=[...S.sel].filter(x=>S.assets.has(x)); if(!ids.length) break;
        if(!confirm(`Permanently delete ${plural(ids.length,"asset")}? This can't be undone here.\n\nTo keep their history, cancel and use Edit to set Status to Retired instead.`)) break;
        t.disabled=true;
        const r=await runBulk(ids,async id=>!!(await write("assets",id,null,"delete")),"Deleting");
        t.disabled=false; S.sel.clear();
        toast(r.fail?`Deleted ${r.ok}; ${r.fail} couldn't be deleted.`:`Deleted ${plural(r.ok,"asset")}`); render(); break; }
      case "sort": { const tb=t.dataset.t,k=t.dataset.k,cur=S.sort[tb]||{};
        S.sort[tb]=cur.key!==k?{key:k,dir:"asc"}:cur.dir==="asc"?{key:k,dir:"desc"}:{};
        try{localStorage.setItem("itat-sort",JSON.stringify(S.sort))}catch(e){}
        render(); break; }
      case "site": S.site=t.dataset.l; go("sites"); break;
      case "cat": S.f.cat=t.dataset.c; render(); break;
      case "gocat": S.f={q:"",cat:t.dataset.c,loc:"All",status:"All",attn:false}; go("assets"); break;
      case "goattn": { if(t.dataset.v==="actions"){S.ashow="Open";go("actions");break}
        if(t.dataset.v==="contracts"){S.ccat="All";S.cshow="Active";S.cq="";go("contracts");break}
        const f=JSON.parse(t.dataset.f||"{}"); S.f={q:"",cat:"All",loc:"All",status:"All",attn:false,...f}; go("assets"); break;}
      case "clearattn": S.f.attn=false; render(); break;
      case "year": S.byear=t.dataset.y==="all"?"all":Number(t.dataset.y); if(S.view!=="budget") go("budget"); else render(); break;
      case "person": S.personId=id; render(); if(innerWidth<900) setTimeout(()=>ROOT.querySelector(".people>section:last-child")?.scrollIntoView({behavior:"smooth"}),60); break;
      case "ashow": S.ashow=t.dataset.s; render(); break;
      case "open": { const a=S.assets.get(id); if(!a) return; S.drawer={type:"asset",id}; openDrawer(assetForm({...a,_id:id},false)); loadHistory(id); break;}
      case "noteOpen": { const nf=$("#noteform"); if(nf){nf.innerHTML=noteFormHTML(id); const ta=$("#noteText"); ta&&ta.focus();} break; }
      case "noteCancel": { const nf=$("#noteform"); if(nf) nf.innerHTML=""; break; }
      case "noteSave": { const txt=($("#noteText")||{}).value||"", typ=($("#noteType")||{}).value||"Note";
        if(!txt.trim()){toast("Write something for the note.");return}
        t.disabled=true; const a=S.assets.get(id);
        const ok=S.ext&&S.ext.addLog?await S.ext.addLog({assetId:id,assetName:a?a.name:"",type:typ,details:txt.trim()}):false;
        t.disabled=false;
        if(ok){toast("Note added");const nf=$("#noteform");if(nf)nf.innerHTML="";loadHistory(id)} break; }
      case "repurpose": { const a=S.assets.get(id); if(!a) return; S.drawer={type:"repurpose",id}; openDrawer(repurposeForm({...a,_id:id})); break; }
      case "repurposeApply": {
        const a=S.assets.get(id); if(!a) return; const f=$("#rform");
        const pid=f.elements.personId.value, p=pid?S.people.get(pid):null;
        const name=f.elements.name.value.trim(), loc=f.elements.location.value.trim(), status=f.elements.status.value, note=f.elements.note.value.trim();
        const markTerm=f.elements.markTerm&&f.elements.markTerm.checked;
        if(!name){toast("Enter the new computer name.");return}
        if(pid&&pid===a.personId&&name===a.name){toast("Pick a different user or name to repurpose.");return}
        const old=person(a), oldName=a.name, oldUser=a.assignedTo;
        t.disabled=true;
        const ok=await write("assets",id,{name,personId:pid||null,assignedTo:p?p.name:null,location:loc||null,status,...stamp()});
        if(ok){
          if(S.ext&&S.ext.addLog){
            const parts=[];
            if(oldName!==name) parts.push(`Renamed ${oldName} → ${name}.`);
            parts.push(`${oldUser?"From "+oldUser:"From no user"} to ${p?p.name:"spare pool"}.`);
            if(note) parts.push(note);
            await S.ext.addLog({assetId:id,assetName:name,type:"Repurposed",details:parts.join(" "),personId:old?old._id||a.personId:null});
          }
          if(markTerm&&a.personId) await write("people",a.personId,{status:"Terminated",...stamp()});
          toast(`Repurposed as ${name}`);
          S.drawer={type:"asset",id}; openDrawer(assetForm({...S.assets.get(id),_id:id},false)); loadHistory(id);
        }
        t.disabled=false; break; }
      case "close": closeDrawer(); break;
      case "newAsset": S.drawer={type:"asset",id:null}; openDrawer(assetForm({category:S.f.cat!=="All"?S.f.cat:"Computers",location:S.view==="sites"?S.site:(S.f.loc!=="All"?S.f.loc:""),status:"In Use",plan:"auto",qty:1,rclass:""},true)); break;
      case "saveAsset": {
        const d=readForm(); if(!d.name){toast("Give the asset a name or ID.");return}
        if(!d.rclass||d.rclass==="") d.rclass=guessClass(d);
        t.disabled=true;
        if(id){ if(await write("assets",id,{...d,...stamp()})){toast("Saved");closeDrawer()} }
        else { if(!d.year&&d.rclass!=="none"){d.year=THIS_YEAR;d.yearSource="Added "+THIS_YEAR}
          const nid="a-"+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
          if(await write("assets",nid,{...d,key:nid,createdAt:new Date().toISOString(),...stamp()},"set")){toast("Asset added");closeDrawer()} }
        t.disabled=false; break;}
      case "delAsset": if(confirm("Delete this asset permanently? Consider marking it Retired instead to keep its history.")){ if(await write("assets",id,null,"delete")){toast("Deleted");closeDrawer()} } break;
      case "findrepl": findReplacement(id); break;
      case "usePrice": { const a=S.assets.get(id); if(a&&a.suggestion){ if(await write("assets",id,{quote:a.suggestion.estPrice,quoteSource:"AI estimate "+NOW.toISOString().slice(0,10),...stamp()})){toast("Budget price updated");const f=$("#aform");if(f){f.quote.value=a.suggestion.estPrice;f.quoteSource.value="AI estimate "+NOW.toISOString().slice(0,10)}}} break;}
      case "copy": try{await navigator.clipboard.writeText(t.dataset.t);toast("Copied")}catch(e){toast(t.dataset.t)} break;
      case "newPerson": S.drawer={type:"person",id:null}; openDrawer(personForm({status:"Employed",location:S.view==="sites"?S.site:""},true)); break;
      case "editPerson": { const p=S.people.get(id); S.drawer={type:"person",id}; openDrawer(personForm({...p,_id:id},false)); break;}
      case "savePerson": { const o=formObj("#pform"); if(!o.name){toast("Enter a name.");return}
        const devs=readDevRows();
        if(devs.some(dv=>!dv.name)){toast("Give each new device a computer name or ID.");return}
        const existing=$("#pAssignExisting")?.value||"";
        const d={name:o.name,title:o.title||null,login:o.login||null,location:o.location||null,status:o.status,o365:o.o365,mfa:o.mfa,vpn:o.vpn,av:o.av,...stamp()};
        t.disabled=true;
        let pid=id, ok;
        if(id){ ok=await write("people",id,d);
          if(ok) for(const [aid,a] of S.assets) if(a.personId===id&&a.assignedTo!==d.name) await write("assets",aid,{assignedTo:d.name}); }
        else { pid="p-"+(o.login||o.name).toLowerCase().replace(/[^a-z0-9]+/g,"-").slice(0,40)+"-"+Math.random().toString(36).slice(2,5);
          ok=await write("people",pid,{...d,key:pid},"set"); if(ok) pid=ok; }
        if(!ok){t.disabled=false;break}
        let added=0;
        for(const dv of devs){
          const typ=dv.kind==="Other"?"Device":dv.kind;
          const detail=dv.detail||(dv.kind==="Laptop"?"Laptop / Dock / Monitors":dv.kind==="Tablet"?"Tablet":dv.kind==="Desktop"?"Desktop":null);
          const a={name:dv.name,category:"Computers",type:"Assigned endpoint",location:d.location,make:null,model:dv.model||null,serial:dv.serial||null,
            status:"In Use",qty:1,personId:pid,assignedTo:d.name,detail,plan:"auto",purchaseDate:dv.purchaseDate||null,
            year:dv.purchaseDate?Number(dv.purchaseDate.slice(0,4)):THIS_YEAR,yearSource:dv.purchaseDate?"Purchase date":"Added "+THIS_YEAR,
            createdAt:new Date().toISOString(),...stamp()};
          a.rclass=guessClass({category:"Computers",model:a.model,type:typ+" "+(detail||"")});
          const aid="a-"+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
          if(await write("assets",aid,{...a,key:aid},"set")) added++;
        }
        if(existing){ const ex=S.assets.get(existing); if(ex&&await write("assets",existing,{personId:pid,assignedTo:d.name,status:"In Use",location:ex.location||d.location||null,...stamp()})) added++; }
        toast((id?"Saved":"Person added")+(added?` with ${plural(added,"device")}`:""));
        S.personId=pid; closeDrawer(); if(!id) go("people"); else render();
        break;}
      case "addDevRow": { const box=$("#devrows"); if(box){box.insertAdjacentHTML("beforeend",devRowHTML());const rows=box.querySelectorAll(".devrow");rows[rows.length-1].querySelector("input").focus()} break;}
      case "rmDevRow": t.closest(".devrow")?.remove(); break;
      case "delPerson": { const n=assets().filter(a=>a.personId===id).length;
        if(confirm(n?`Delete this person? Their ${plural(n,"device")} will become unassigned.`:"Delete this person?")){
          for(const [aid,a] of S.assets) if(a.personId===id) await write("assets",aid,{personId:null,assignedTo:null,...stamp()});
          if(await write("people",id,null,"delete")){toast("Deleted");S.personId=null;closeDrawer()} } break;}
      case "assign": { const sel=$("#assignSel"); if(!sel||!sel.value){toast("Pick a computer to assign.");return}
        const p=S.people.get(id); if(await write("assets",sel.value,{personId:id,assignedTo:p.name,status:"In Use",location:S.assets.get(sel.value).location||p.location||null,...stamp()})) toast("Assigned to "+p.name); break;}
      case "newAction": S.drawer={type:"action",id:null}; openDrawer(actionForm({},true)); break;
      case "editAction": { const a=S.actions.get(id); S.drawer={type:"action",id}; openDrawer(actionForm({...a,_id:id},false)); break;}
      case "saveAction": { const o=formObj("#tform"); if(!o.title){toast("Describe the action.");return}
        const d={title:o.title,detail:o.detail||null,priority:o.priority,location:o.location||null,area:o.area||null,status:o.status,...stamp()};
        if(id){ if(await write("actions",id,d)){toast("Saved");closeDrawer()} }
        else { if(await write("actions","t-"+Date.now().toString(36),d,"set")){toast("Action added");closeDrawer()} } break;}
      case "delAction": if(confirm("Delete this action?")){ if(await write("actions",id,null,"delete")){toast("Deleted");closeDrawer()} } break;
      case "aireview": aiReview(); break;
      case "copyPrompt": copyPrompt(false); break;
      case "openCopilot": copyPrompt(true); break;
      case "stopai": S.aiRevCtl&&S.aiRevCtl.abort(); break;
      case "export": exportCSV(); break;
      case "print": printReport(t.dataset.r,id); break;
      case "sploc": S.sploc=t.dataset.l; if(S.view==="spares") render(); else go("spares"); break;
      case "newSpare": S.drawer={type:"asset",id:null}; openDrawer(assetForm({category:S.spcat!=="All"?S.spcat:"Computers",location:S.sploc!=="All"?S.sploc:"",status:"Spare",plan:"auto",qty:1,rclass:""},true)); break;
    }
  });
  on("keydown",e=>{
    if(e.key==="Escape"&&S.drawer) closeDrawer();
    if((e.key==="Enter"||e.key===" ")&&e.target.matches('[role="button"][data-act]')){e.preventDefault();e.target.click()}
  });
  $("#scrim").addEventListener("click",closeDrawer);
  on("toggle",e=>{const d=e.target;if(d.matches&&d.matches("details.grp")){d.open?S.open.add(d.dataset.key):S.open.delete(d.dataset.key)}},true);
  let inT;
  on("input",e=>{
    const k=e.target.dataset.in; if(!k) return;
    clearTimeout(inT); inT=setTimeout(()=>{ if(k==="q") S.f.q=e.target.value; else S[k]=e.target.value; render(); },140);
  });
  let clsT={};
  on("change",async e=>{
    const el=e.target;
    if(el.dataset.ch){const k=el.dataset.ch;
      if(["loc","status"].includes(k)) S.f[k]=el.value; else if(k==="attn") S.f.attn=el.checked; else S[k]=el.value; render(); return}
    if(el.dataset.plan){const id=el.dataset.plan,v=el.value;const a=S.assets.get(id);await write("assets",id,{plan:v,planYear:(v==="replace"||v==="reserve")?(a.planYear||(budget(a)||{}).year||Y0):a.planYear??null,...stamp()});return}
    if(el.dataset.planyear){await write("assets",el.dataset.planyear,{planYear:Number(el.value),...stamp()});return}
    if(el.dataset.done){await write("actions",el.dataset.done,{status:el.checked?"Done":"Open",...stamp()});return}
    if(el.dataset.cls){const k=el.dataset.cls,f=el.dataset.fld;let v=el.value;if(f!=="standard"){v=Number(v);if(isNaN(v)||v<0){toast("Enter a positive number.");return}}
      await write("config","classes",{classes:{[k]:{[f]:v}}});toast("Saved");return}
    if(el.form&&el.form.id==="rform"&&(el.name==="personId"||el.name==="location")&&S.drawer){
      const a=S.assets.get(S.drawer.id); if(a){
        const p=el.form.elements.personId.value?S.people.get(el.form.elements.personId.value):null;
        if(el.name==="personId"&&p&&p.location) el.form.elements.location.value=p.location;
        if(!el.form.elements.name.dataset.touched) el.form.elements.name.value=suggestName(a,p,el.form.elements.location.value.trim());
        el.form.elements.status.value=p?"In Use":"Spare";
      }
    }
    if(el.form&&el.form.id==="rform"&&el.name==="name"){el.dataset.touched="1"}
    if(el.form&&el.form.id==="aform"&&el.name==="purchaseDate"&&el.value){el.form.year.value=el.value.slice(0,4)}
    if(el.form&&el.form.id==="aform"&&(el.name==="category"||el.name==="model"||el.name==="type")&&S.drawer&&!S.drawer.id){
      const d=readForm(); el.form.rclass.value=guessClass(d);}
  });


  /* ================= SharePoint data layer ================= */
  const P=opts.listPrefix==null?"IT ":opts.listPrefix;
  const LISTS={locations:P+"Locations",classes:P+"Replacement Classes",people:P+"People",assets:P+"Assets",actions:P+"Actions",log:P+"Asset Log",contracts:P+"Contracts"};
  const ACCEPT="application/json;odata=nometadata";
  const listUrl=t=>`${opts.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(t.replace(/'/g,"''"))}')`;
  async function spErr(r){
    let m=r.statusText||("HTTP "+r.status);
    try{const j=await r.json();m=(j["odata.error"]&&j["odata.error"].message&&j["odata.error"].message.value)||(j.error&&j.error.message&&(j.error.message.value||j.error.message))||m}catch(e){}
    const e=new Error(m);e.status=r.status;return e;
  }
  async function spGet(url){const r=await opts.spHttpClient.get(url,opts.spConfig,{headers:{Accept:ACCEPT}});if(!r.ok)throw await spErr(r);return r.json()}
  async function spPost(url,body,extra){
    const r=await opts.spHttpClient.post(url,opts.spConfig,{headers:{Accept:ACCEPT,"Content-Type":ACCEPT,...(extra||{})},body:body?JSON.stringify(body):undefined});
    if(!r.ok)throw await spErr(r);
    const t=await r.text();return t?JSON.parse(t):null;
  }
  async function getAll(title,extra){
    let url=listUrl(title)+"/items?$top=5000"+(extra||"");const out=[];
    while(url){const j=await spGet(url);out.push(...(j.value||[]));url=j["odata.nextLink"]||null}
    return out;
  }

  /* ---- field mapping between the app's model and the list columns ---- */
  const locById=new Map(),locByName=new Map(),classKeyById=new Map(),classIdByKey=new Map();
  const PLAN_IN={Auto:"auto",Replace:"replace",Reserve:"reserve",Skip:"exclude"};
  const PLAN_OUT={auto:"Auto",replace:"Replace",reserve:"Reserve",exclude:"Skip"};
  const dIn=v=>{if(!v)return null;const d=new Date(v);if(isNaN(d))return null;return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`};
  const dOut=v=>v?`${v}T12:00:00Z`:null;
  const nz=v=>v===""||v===undefined?null:v;
  const json=v=>{if(!v)return null;try{return JSON.parse(v)}catch(e){return null}};
  const lookupId=v=>v==null?null:String(v);

  const IN={
    locations:i=>({_id:String(i.Id),name:i.Title}),
    classes:i=>({_id:String(i.Id),key:i.ClassKey,label:i.Title,standard:i.StandardModel||"",cost:Number(i.UnitCost)||0,life:Number(i.LifecycleYears)||0,from:i.CostSource||"",spareMin:i.SpareMin??null}),
    people:i=>({_id:String(i.Id),key:i.PersonKey,name:i.Title||"(no name)",aka:i.AlsoKnownAs,title:i.JobTitle,login:i.Login,
      location:locById.get(i.PersonLocationId)||null,status:i.EmploymentStatus||"Unknown",o365:!!i.O365,mfa:!!i.MFA,vpn:!!i.VPN,av:!!i.AV}),
    assets:i=>({_id:String(i.Id),key:i.AssetKey,name:i.Title||"(no name)",category:i.Category||"Other",type:i.AssetType,location:locById.get(i.AssetLocationId)||null,
      make:i.Make,model:i.Model,serial:i.SerialNumber,ip:i.IPAddress,qty:i.Quantity||1,status:i.AssetStatus||"In Use",transition:i.TransitionPlan,
      personId:lookupId(i.AssignedPersonId),sharedUse:i.SharedUse,assignedTo:null,detail:i.Setup,year:i.InServiceYear??null,yearSource:i.YearSource,
      purchaseDate:dIn(i.PurchaseDate),warrantyEnd:dIn(i.WarrantyEnd),rclass:classKeyById.get(i.ReplacementClassId)||"none",
      plan:PLAN_IN[i.BudgetPlan]||"auto",planYear:i.PlanYear??null,quote:i.QuotePrice??null,quoteSource:i.QuoteSource,suggestion:json(i.AISuggestion),
      monthlyCost:i.MonthlyCost??null,contractEnd:i.ContractEnd,notes:i.Notes,action:i.ImportedAction,created:dIn(i.Created)}),
    contracts:i=>({_id:String(i.Id),name:i.Title||"(no name)",vendor:i.Vendor,account:i.AccountNumber,serviceId:i.ServiceId,category:i.ContractCategory||"Other",
      cost:i.Cost??null,billing:i.Billing||"Annual",term:i.TermMonths??null,last:dIn(i.LastRenewal),next:dIn(i.NextRenewal),auto:!!i.AutoRenew,
      quote:i.QuoteAmount??null,quoteTerm:i.QuoteTermMonths??null,quoteNotes:i.QuoteNotes,location:locById.get(i.ContractLocationId)||null,
      contact:i.Contact,status:i.ContractStatus||"Active",notes:i.ContractNotes}),
    actions:i=>({_id:String(i.Id),title:i.Title,detail:i.Details,priority:i.Priority||"Medium",location:locById.get(i.ActionLocationId)||null,area:i.Area,status:i.ActionStatus||"Open"}),
  };
  const MAPS={
    assets:{name:"Title",key:"AssetKey",category:"Category",type:"AssetType",make:"Make",model:"Model",serial:"SerialNumber",ip:"IPAddress",qty:"Quantity",
      status:"AssetStatus",transition:"TransitionPlan",detail:"Setup",year:"InServiceYear",yearSource:"YearSource",planYear:"PlanYear",quote:"QuotePrice",
      quoteSource:"QuoteSource",monthlyCost:"MonthlyCost",contractEnd:"ContractEnd",notes:"Notes"},
    people:{name:"Title",key:"PersonKey",title:"JobTitle",login:"Login",status:"EmploymentStatus",o365:"O365",mfa:"MFA",vpn:"VPN",av:"AV",aka:"AlsoKnownAs"},
    contracts:{name:"Title",vendor:"Vendor",account:"AccountNumber",serviceId:"ServiceId",category:"ContractCategory",cost:"Cost",billing:"Billing",term:"TermMonths",
      auto:"AutoRenew",quote:"QuoteAmount",quoteTerm:"QuoteTermMonths",quoteNotes:"QuoteNotes",contact:"Contact",status:"ContractStatus",notes:"ContractNotes"},
    actions:{title:"Title",detail:"Details",priority:"Priority",area:"Area",status:"ActionStatus"},
  };
  const LOCFIELD={assets:"AssetLocationId",people:"PersonLocationId",actions:"ActionLocationId",contracts:"ContractLocationId"};
  async function ensureLoc(name){
    name=(name||"").trim();if(!name)return null;
    const hit=locByName.get(name.toLowerCase());if(hit)return hit;
    const j=await spPost(listUrl(LISTS.locations)+"/items",forSP("locations",{Title:name}));
    locById.set(j.Id,name);locByName.set(name.toLowerCase(),j.Id);return j.Id;
  }
  async function toSP(col,d){
    const o={},m=MAPS[col];
    for(const k in d) if(k in m) o[m[k]]=nz(d[k]);
    if("location" in d) o[LOCFIELD[col]]=await ensureLoc(d.location);
    if(col==="assets"){
      if("personId" in d){o.AssignedPersonId=d.personId?Number(d.personId):null;o.SharedUse=d.personId?null:nz(d.assignedTo)}
      if("rclass" in d) o.ReplacementClassId=classIdByKey.get(d.rclass)||null;
      if("plan" in d) o.BudgetPlan=PLAN_OUT[d.plan]||"Auto";
      if("purchaseDate" in d) o.PurchaseDate=dOut(d.purchaseDate);
      if("warrantyEnd" in d) o.WarrantyEnd=dOut(d.warrantyEnd);
      if("suggestion" in d) o.AISuggestion=d.suggestion?JSON.stringify(d.suggestion):null;
      if(o.Quantity!=null) o.Quantity=Number(o.Quantity)||1;
    }
    if(col==="contracts"){
      if("last" in d) o.LastRenewal=dOut(d.last);
      if("next" in d) o.NextRenewal=dOut(d.next);
      if("auto" in d) o.AutoRenew=!!d.auto;
    }
    return o;
  }
  function relink(){
    for(const a of S.assets.values()){
      const p=a.personId?S.people.get(a.personId):null;
      a.assignedTo=p?p.name:(a.sharedUse||null);
    }
  }

  /* ---- column names ----
     The app refers to each column by the name the setup script requested ("logical"
     name). SharePoint may store a column under a different internal name (for
     example when a requested name is reserved or escaped) and may expose it to REST
     under yet another name. Each list's real columns are read once; every logical
     name is resolved by internal name first, then by display name. Columns that
     can't be found are left out of saves (and listed in Settings) instead of
     failing the whole save. */
  const LOGICAL={
    locations:[],
    classes:["ClassKey","StandardModel","UnitCost","LifecycleYears","CostSource","SpareMin"],
    people:["PersonKey","JobTitle","Login","PersonLocation","EmploymentStatus","O365","MFA","VPN","AV","AlsoKnownAs"],
    assets:["AssetKey","Category","AssetType","AssetLocation","AssignedPerson","SharedUse","Make","Model","SerialNumber","IPAddress","Quantity",
      "AssetStatus","TransitionPlan","Setup","InServiceYear","YearSource","PurchaseDate","WarrantyEnd","ReplacementClass","BudgetPlan","PlanYear",
      "QuotePrice","QuoteSource","AISuggestion","MonthlyCost","ContractEnd","Notes","ImportedAction"],
    actions:["Details","Priority","ActionLocation","Area","ActionStatus"],
    log:["AssetRef","PersonRef","EntryType","Details"],
    contracts:["Vendor","AccountNumber","ServiceId","ContractCategory","Cost","Billing","TermMonths","LastRenewal","NextRenewal","AutoRenew",
      "QuoteAmount","QuoteTermMonths","QuoteNotes","ContractLocation","Contact","ContractStatus","ContractNotes"],
  };
  const FIELDS={}, MISSING={}, INTERNAL={};
  const norm=v=>String(v||"").toLowerCase().replace(/[^a-z0-9]/g,"");
  async function loadFields(col){
    const map={}, missing=[];
    let fields;
    try{
      const j=await spGet(listUrl(LISTS[col])+"/fields?$select=Title,InternalName,EntityPropertyName,TypeAsString,Hidden,ReadOnlyField&$top=500");
      fields=(j.value||[]).filter(f=>!f.Hidden);
    }catch(e){console.warn("Couldn't read columns for "+LISTS[col],e);FIELDS[col]=null;MISSING[col]=[];return}
    const byInternal=new Map(fields.map(f=>[f.InternalName,f]));
    const internal={Title:"Title"};
    for(const name of LOGICAL[col]){
      let f=byInternal.get(name);
      if(!f) f=fields.find(x=>!x.ReadOnlyField&&norm(x.Title)===norm(name))||fields.find(x=>!x.ReadOnlyField&&norm(x.InternalName)===norm(name));
      if(!f){missing.push(name);continue}
      const ent=f.EntityPropertyName||f.InternalName;
      map[name]=ent; internal[name]=f.InternalName;
      if(/^Lookup/.test(f.TypeAsString||"")) map[name+"Id"]=ent+"Id";
    }
    // If most columns look missing, the column list itself is unreliable; send names unchanged instead.
    if(LOGICAL[col].length&&missing.length>LOGICAL[col].length/2){
      console.warn("Column list for "+LISTS[col]+" looks incomplete; using requested names as-is.");
      FIELDS[col]=null; MISSING[col]=[]; return;
    }
    FIELDS[col]=map; MISSING[col]=missing; INTERNAL[col]=internal;
    if(missing.length) console.warn(LISTS[col]+" is missing columns:",missing);
  }
  function fromSP(col,item){
    const map=FIELDS[col];if(!map)return item;
    for(const k in map){const ent=map[k];if(ent!==k&&ent in item)item[k]=item[ent]}
    return item;
  }
  function forSP(col,body){
    const map=FIELDS[col];if(!map||!body)return body;
    const miss=new Set(MISSING[col]||[]);
    const o={};
    for(const k in body){
      const base=k.endsWith("Id")&&!(k in map)&&(k.slice(0,-2) in map||miss.has(k.slice(0,-2)))?k.slice(0,-2):null;
      if(miss.has(k)||(base&&miss.has(base))) continue;          // column doesn't exist: skip it
      o[map[k]||(base?map[base]+"Id":k)]=body[k];
    }
    return o;
  }
  function columnReport(){
    const out=[];
    const OPTIONAL=new Set(["SpareMin"]);
    for(const col in MISSING){const m=(MISSING[col]||[]).filter(x=>!OPTIONAL.has(x));if(m.length) out.push(`${LISTS[col]}: ${m.join(", ")}`)}
    return out;
  }

  /* ---- loading ---- */
  async function load(col){
    if(!FIELDS[col]) await loadFields(col);
    const items=(await getAll(LISTS[col])).map(i=>fromSP(col,i));
    if(col==="locations"){locById.clear();locByName.clear();SITE_ORDER.length=0;for(const i of items.slice().sort((x,y)=>x.Id-y.Id)){locById.set(i.Id,i.Title);locByName.set(String(i.Title||"").toLowerCase(),i.Id);if(i.Title)SITE_ORDER.push(i.Title)}return}
    if(col==="classes"){
      classKeyById.clear();classIdByKey.clear();const c={};
      for(const i of items){const x=IN.classes(i);if(!x.key)continue;c[x.key]=x;classKeyById.set(i.Id,x.key);classIdByKey.set(x.key,i.Id)}
      if(!c.none) c.none={label:"Not budgeted",standard:"",cost:0,life:0};
      S.classes=c;S.loaded.classes=true;return;
    }
    const m=new Map();for(const i of items){const x=IN[col](i);m.set(x._id,x)}
    S[col]=m;S.loaded[col]=true;
  }
  function stampSync(){const el=ROOT.querySelector("#sync");if(el)el.textContent="Updated "+new Date().toLocaleTimeString([],{hour:"numeric",minute:"2-digit"})}
  /* ---- asset log (dated notes + assignment history) ---- */
  let LOG_OK=false;
  S.log=[];
  async function loadLog(){
    try{
      if(!FIELDS.log&&FIELDS.log!==null) await loadFields("log");
      if(FIELDS.log===null&&!MISSING.log){LOG_OK=false;return}
      const items=(await getAll(LISTS.log,"&$select=*,Author/Title&$expand=Author")).map(i=>fromSP("log",i));
      S.log=items.map(i=>({id:String(i.Id),assetId:i.AssetRefId!=null?String(i.AssetRefId):null,assetName:i.Title||"",
        personId:i.PersonRefId!=null?String(i.PersonRefId):null,type:i.EntryType||"Note",details:i.Details||"",when:i.Created,who:(i.Author&&i.Author.Title)||""}));
      LOG_OK=true;
    }catch(e){
      LOG_OK=false; S.log=[];
      if(e.status!==404) console.warn("Couldn't load "+LISTS.log,e);
    }
  }
  async function addLog({assetId,assetName,type,details,personId}){
    if(!LOG_OK) return false;
    try{
      const body=forSP("log",{Title:String(assetName||"").slice(0,255),AssetRefId:assetId?Number(assetId):null,PersonRefId:personId?Number(personId):null,EntryType:type||"Note",Details:details||null});
      const j=await spPost(listUrl(LISTS.log)+"/items",body);
      S.log.push({id:String(j.Id),assetId:assetId?String(assetId):null,assetName:assetName||"",personId:personId?String(personId):null,type:type||"Note",details:details||"",when:new Date().toISOString(),who:opts.userName||"You"});
      return true;
    }catch(e){console.error(e);toast("Couldn't save the note: "+(e.message||"check your connection"));return false}
  }
  /* SharePoint keeps a version of each item on every save. History compares consecutive versions of the fields below. */
  const TRACK=[["Title","Name"],["AssignedPerson","Assigned to"],["SharedUse","Shared use"],["AssetLocation","Location"],["AssetStatus","Status"],
    ["Category","Category"],["AssetType","Type"],["Model","Model"],["SerialNumber","Serial / service tag"],["ReplacementClass","Replacement class"],
    ["BudgetPlan","Budget plan"],["PlanYear","Plan year"],["InServiceYear","In-service year"],["PurchaseDate","Purchase date"],["WarrantyEnd","Warranty end"],
    ["QuotePrice","Quoted price"],["Notes","Notes"]];
  const KEY_FIELDS=new Set(["Title","AssignedPerson","SharedUse","AssetLocation","AssetStatus"]);
  function versionValue(v,name){
    const keys=[];
    if(FIELDS.assets&&FIELDS.assets[name]) keys.push(FIELDS.assets[name]);
    if(INTERNAL.assets&&INTERNAL.assets[name]) keys.push(INTERNAL.assets[name]);
    keys.push(name);
    for(const k of keys){
      if(!(k in v)) continue;
      let x=v[k];
      if(Array.isArray(x)) x=x.map(y=>y&&typeof y==="object"?(y.LookupValue??y.Title):y).filter(Boolean).join(", ");
      else if(x&&typeof x==="object") x=x.LookupValue??x.Title??x.Label??null;
      return x===""||x===undefined?null:x;
    }
    for(const k of keys){
      const idv=v[k+"Id"]; if(idv==null) continue;
      if(name==="AssignedPerson") return (S.people.get(String(idv))||{}).name||("#"+idv);
      if(name==="AssetLocation") return locById.get(Number(idv))||("#"+idv);
      if(name==="ReplacementClass"){const k2=classKeyById.get(Number(idv));return k2&&S.classes[k2]?S.classes[k2].label:("#"+idv)}
    }
    return null;
  }
  function showVal(name,x){
    if(x==null) return null;
    if(name==="PurchaseDate"||name==="WarrantyEnd") return fmtDate(x);
    if(name==="QuotePrice") return money(Number(x));
    return String(x);
  }
  async function assetHistory(id){
    const j=await spGet(listUrl(LISTS.assets)+`/items(${Number(id)})/versions`);
    const vs=(j.value||[]).slice().sort((a,b)=>(Number(a.VersionId)||0)-(Number(b.VersionId)||0)||new Date(a.Modified)-new Date(b.Modified));
    const out=[];let prev=null;
    for(const v of vs){
      const who=(v.Editor&&(v.Editor.LookupValue||v.Editor.Title))||(typeof v.Editor==="string"?v.Editor:"")||(v.CreatedBy&&v.CreatedBy.Title)||"";
      const when=v.Modified||v.Created;
      if(!prev){
        const nm=versionValue(v,"Title"), u=versionValue(v,"AssignedPerson")||versionValue(v,"SharedUse");
        out.push({kind:"created",when,who,key:true,title:"Added",text:`as ${nm||"(no name)"}${u?`, assigned to ${u}`:""}`});
      }else{
        const changes=[];let key=false;
        for(const [name,label] of TRACK){
          const a=versionValue(prev,name), b=versionValue(v,name);
          if(String(a??"")===String(b??"")) continue;
          if(name==="Notes"){changes.push({label:"Notes",from:null,to:"edited"});continue}
          changes.push({label,from:showVal(name,a),to:showVal(name,b)});
          if(KEY_FIELDS.has(name)) key=true;
        }
        if(changes.length) out.push({kind:"change",when,who,key,title:key?"Changed":"Updated",changes});
      }
      prev=v;
    }
    for(const l of S.log){
      if(l.assetId!==String(id)||l.type==="Assigned"||l.type==="Unassigned") continue;
      out.push({kind:"note",when:l.when,who:l.who,title:l.type==="Note"?"Note:":l.type+":",text:l.details});
    }
    return out.sort((a,b)=>new Date(b.when)-new Date(a.when));
  }
  function pastDevices(pid){
    const seen=new Map();
    for(const l of S.log){
      if(l.personId!==String(pid)||!l.assetId||l.type!=="Unassigned") continue;
      const cur=seen.get(l.assetId);
      if(!cur||new Date(l.when)>new Date(cur.when)) seen.set(l.assetId,l);
    }
    return [...seen.values()].sort((a,b)=>new Date(b.when)-new Date(a.when));
  }
  /* Record assignment changes so a person's page can list devices they used to have. */
  async function logAssignment(assetId,oldName,newName,oldPid,newPid){
    if(!LOG_OK||String(oldPid||"")===String(newPid||"")) return;
    const pname=x=>(S.people.get(String(x))||{}).name||"someone";
    if(oldPid) await addLog({assetId,assetName:oldName||newName,type:"Unassigned",personId:oldPid,details:`Unassigned from ${pname(oldPid)}`+(newPid?` (reassigned to ${pname(newPid)})`:"")});
    if(newPid) await addLog({assetId,assetName:newName||oldName,type:"Assigned",personId:newPid,details:`Assigned to ${pname(newPid)}`});
  }
  let CONTRACTS_OK=false;
  async function loadContracts(){
    try{ await load("contracts"); CONTRACTS_OK=true; }
    catch(e){ CONTRACTS_OK=false; S.contracts=new Map(); if(e.status!==404) console.warn("Couldn't load "+LISTS.contracts,e); }
  }
  S.ext={history:assetHistory,addLog,pastDevices,get logAvailable(){return LOG_OK},get contractsAvailable(){return CONTRACTS_OK},get spareMinAvailable(){return !!FIELDS.classes&&!(MISSING.classes||[]).includes("SpareMin")}};

  async function loadAll(fresh){
    if(fresh) for(const k in FIELDS) delete FIELDS[k];
    await Promise.all([load("locations"),load("classes")]);
    await Promise.all([load("people"),load("assets"),load("actions"),loadLog(),loadContracts()]);
    relink();stampSync();
  }
  let refreshTimer=null;
  function scheduleRefresh(){clearTimeout(refreshTimer);refreshTimer=setTimeout(async()=>{try{await loadAll();render()}catch(e){console.error(e)}},1500)}

  /* ---- writes (same signature the UI already uses) ---- */
  async function write(col,id,data,mode){
    if(S.canWrite===false){toast("You have view-only access to the IT lists.");return false}
    try{
      if(col==="config"){
        for(const [k,f] of Object.entries((data&&data.classes)||{})){
          const cid=classIdByKey.get(k);if(!cid)continue;
          const o={};if("standard" in f)o.StandardModel=f.standard;if("cost" in f)o.UnitCost=Number(f.cost);if("life" in f)o.LifecycleYears=Number(f.life);if("label" in f)o.Title=f.label;if("spareMin" in f)o.SpareMin=Number(f.spareMin)||0;
          await spPost(listUrl(LISTS.classes)+`/items(${cid})`,forSP("classes",o),{"X-HTTP-Method":"MERGE","IF-MATCH":"*"});
          Object.assign(S.classes[k],f);
        }
        render();return true;
      }
      const url=listUrl(LISTS[col])+"/items";
      if(mode==="delete"){
        await spPost(`${url}(${Number(id)})`,null,{"X-HTTP-Method":"DELETE","IF-MATCH":"*"});
        S[col].delete(id);relink();render();return true;
      }
      const raw=await toSP(col,data||{});
      const body=forSP(col,raw);
      if(Object.keys(raw).length&&!Object.keys(body).length){
        toast("Couldn't save: none of those fields exist on the "+LISTS[col]+" list. Check Settings for missing columns.");
        return false;
      }
      if(mode==="set"){
        const j=await spPost(url,body);
        const x=IN[col](fromSP(col,j));S[col].set(x._id,x);relink();render();
        if(col==="assets"&&x.personId) logAssignment(x._id,x.name,x.name,null,x.personId);
        return x._id;
      }
      const before=S[col].get(id);
      const oldPid=before?before.personId:null, oldName=before?before.name:"";
      if(Object.keys(body).length){
        await spPost(`${url}(${Number(id)})`,body,{"X-HTTP-Method":"MERGE","IF-MATCH":"*"});
      }
      if(col==="assets"&&"personId" in data&&String(oldPid||"")!==String(data.personId||"")){
        await logAssignment(id,oldName,data.name||oldName,oldPid,data.personId);
      }
      const cur=S[col].get(id);
      if(cur){
        for(const k in data){if(k!=="updatedAt"&&k!=="createdAt")cur[k]=data[k]}
        if(col==="assets"&&"personId" in data)cur.sharedUse=data.personId?null:(data.assignedTo||null);
        relink();render();
      }
      return true;
    }catch(e){
      console.error(e);
      if(e.status===401||e.status===403){S.canWrite=false;render();toast("You don't have permission to edit the IT lists.")}
      else if(e.status===404){toast("That item no longer exists. Reloading…");scheduleRefresh()}
      else toast("Couldn't save to SharePoint: "+(e.message||"check your connection"));
      return false;
    }
  }

  /* ---- AI (optional) ---- */
  if(opts.aiUrl&&opts.getAiClient){
    const call=async(prompt,format,signal)=>{
      let r;
      try{const c=await opts.getAiClient();r=await c.post(opts.aiUrl,opts.aiConfig,{headers:{"Content-Type":"application/json",Accept:"application/json"},body:JSON.stringify({prompt,format}),signal})}
      catch(e){throw {code:e&&e.name==="AbortError"?"cancelled":"error"}}
      if(!r.ok)throw {code:r.status===429?"rate_limited":r.status===401||r.status===403?"not_granted":"error"};
      return r.json();
    };
    S.sample=Object.assign(async(p,o)=>{const j=await call(p,"text",o&&o.signal);const text=String(j.text||"");o&&o.onText&&o.onText({text});return {text}},{
      json:async(p,o)=>{const j=await call(p,"json",o&&o.signal);if(j.json)return j.json;try{return JSON.parse(String(j.text||"").replace(/^\s*```(json)?/,"").replace(/```\s*$/,""))}catch(e){throw {code:"invalid_json"}}}
    });
  }

  /* ---- CSV download ---- */
  S.downloads={save:async({filename,data})=>{
    const u=URL.createObjectURL(data);const a=document.createElement("a");a.href=u;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),3000);
  }};

  /* ---- boot ---- */
  on("click",e=>{const t=e.target.closest&&e.target.closest('[data-act="refresh"]');if(!t)return;t.disabled=true;
    loadAll(true).then(()=>{S.appInfo={version:opts.version||"",missing:columnReport(),logMissing:!LOG_OK};render()}).catch(err=>toast("Couldn't reload: "+err.message)).finally(()=>{t.disabled=false})});
  let pollTimer=null;
  function onVisible(){if(document.visibilityState==="visible"&&S.loaded.assets)scheduleRefresh()}
  render();
  (async function boot(){
    try{
      const perm=await spGet(listUrl(LISTS.assets)+"/EffectiveBasePermissions");
      S.canWrite=(Number(perm.Low)&4)!==0;   // EditListItems
    }catch(e){
      S.loadError=e.status===404
        ?`This site doesn't have a list named "${LISTS.assets}". Run Install-AssetTracker.ps1 on this site, or check the list name prefix in the web part settings (currently "${P}").`
        :"SharePoint returned: "+e.message;
      render();ROOT.querySelector("#sync").textContent="Not connected";return;
    }
    try{await loadAll()}
    catch(e){S.loadError=(e.status===404?"One of the IT lists is missing. ":"")+"SharePoint returned: "+e.message;render();return}
    S.appInfo={version:opts.version||"",missing:columnReport(),logMissing:!LOG_OK};
    render();
    pollTimer=setInterval(()=>{if(document.visibilityState==="visible"&&!S.drawer)scheduleRefresh()},60000);
    document.addEventListener("visibilitychange",onVisible);
  })();

  return {
    destroy() {
      clearInterval(pollTimer); clearTimeout(refreshTimer);
      document.removeEventListener('visibilitychange', onVisible);
      if (mq) mq.removeEventListener('change', applyTheme);
      for (const [el, type, fn, cap] of listeners) el.removeEventListener(type, fn, cap);
      if (S.aiCtl) S.aiCtl.abort(); if (S.aiRevCtl) S.aiRevCtl.abort();
      OVL.remove(); host.innerHTML = '';
    }
  };
}
