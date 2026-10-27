(()=>{
const C=window.CMMS_CONFIG,$=s=>document.querySelector(s),G="https://graph.microsoft.com/v1.0",SC=["Files.ReadWrite"],CS="Closing Maintenance Orders";
const FREQ={INSP_DAILY:"Daily",INSP_WEEK:"Weekly",MAINT_15D:"Fortnightly",MAINT_1M:"Monthly",MAINT_3M:"Quarterly",MAINT_6M:"Half-Yearly",MAINT_1Y:"Yearly"};
let acct={username:"Google Sheets"},orders=[],lib=[],today=0,items=[],cur=null,saved=new Set();
const esc=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const ser=d=>Math.floor(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())/864e5)+25569;
const iso=n=>new Date((n-25569)*864e5).toISOString().slice(0,10);
const toSer=s=>{const[y,m,d]=s.split("-").map(Number);return Math.floor(Date.UTC(y,m-1,d)/864e5)+25569};
const dueOf=no=>{const m=/-(\d\d)(\d\d)(\d\d)$/.exec(no);return m?toSer(`20${m[1]}-${m[2]}-${m[3]}`):today};
const user=()=>localStorage.cmmsUser||C.defaultUser||"";
const toast=(m,bad)=>{const t=$("#toast");t.textContent=m;t.className=bad?"bad show":"show";clearTimeout(toast.t);toast.t=setTimeout(()=>t.className="",7000)};
const enc=encodeURIComponent;

async function call(a,body){const r=body?await fetch(C.url,{method:"POST",body:JSON.stringify({key:C.key,action:a,...body})}):await fetch(`${C.url}?action=${a}&key=${enc(C.key)}`);
 const j=await r.json();if(j.error)throw new Error(j.error);return j}
const setLib=r=>{const h=r[0];lib=r.slice(1).map(v=>Object.fromEntries(h.map((k,i)=>[k,v[i]])))};

async function load(){
 const d=await call("load"),cl=d.cl,cr=d.cr;today=d.today;setLib(d.lib);const crit={};cr.forEach(r=>{if(r[0])crit[r[0]]=r[5]});
 orders=[];cl.forEach((r,i)=>{if(r[0])orders.push({row:9+i,no:r[0],eq:r[1],name:r[2],type:r[3],loc:r[4],status:r[5],cl:r[11],crit:crit[r[0]]||"",
  freq:FREQ[String(r[0]).split("-")[0]]||"",act:r[6]||"",onD:typeof r[7]==="number"?iso(r[7]):iso(today),byU:r[8]||user(),remk:r[9]||""})});
 $("#sub").textContent=`${acct.username}  |  Operational date ${iso(today).split("-").reverse().join("/")}`;
 drawClosing();fillPick()}

const chip=o=>{const s=saved.has(o.no)?"SAVED":o.status||"OPEN";return`<span class="chip ${s.replace(" ","_")}">${s}</span><small>Checklist: ${esc(o.cl)}</small>`};
function drawClosing(){const n=orders.length,c=orders.filter(o=>o.status==="CLOSED").length;
 $("#sum").innerHTML=`<div><b>${n}</b>orders due today</div><div><b>${c}</b>closed</div><div><b>${n-c}</b>open or deferred</div>`;
 $("#rows").innerHTML=orders.map((o,i)=>`<tr data-i="${i}"><td>${i+1}</td><td class="mono">${esc(o.no)}</td><td>${esc(o.name)}<small>${esc(o.eq)}, ${esc(o.type)}, ${esc(o.loc)}</small></td><td>${chip(o)}</td>
 <td><select data-k="act"><option></option>${["Closed","Deferred"].map(a=>`<option${o.act===a?" selected":""}>${a}</option>`).join("")}</select></td>
 <td><input type="date" data-k="onD" value="${o.onD}"></td><td><input data-k="byU" size="6" value="${esc(o.byU)}"></td><td><input data-k="remk" value="${esc(o.remk)}"></td>
 <td><button class="link" data-open="${i}">Checklist</button></td></tr>`).join("")||`<tr><td colspan="9" class="empty">No orders are due today.</td></tr>`}
$("#rows").addEventListener("input",e=>{const tr=e.target.closest("tr"),k=e.target.dataset.k;if(tr&&k){orders[tr.dataset.i][k]=e.target.value;if(k==="byU")localStorage.cmmsUser=e.target.value}});
$("#rows").addEventListener("click",e=>{const i=e.target.dataset.open;if(i!==undefined){tab("checklist");$("#order").value=orders[i].no;pick(orders[i].no)}});

async function saveClosures(){
 const todo=orders.filter(o=>o.act&&!saved.has(o.no));if(!todo.length)return toast("Choose Closed or Deferred for at least one order.",1);
 const bad=todo.filter(o=>!o.byU.trim()||!o.onD||(o.act==="Closed"&&o.cl!=="COMPLETE"));
 if(bad.length)return toast("Fix before saving: "+bad.map(o=>o.no+(o.cl!=="COMPLETE"&&o.act==="Closed"?" (checklist not complete)":" (needs date and closed by)")).join("; "),1);
 const left=orders.length-todo.length-saved.size;if(left&&!confirm(`${left} order(s) still have no action. Save the other ${todo.length} anyway?`))return;
 $("#saveClose").disabled=true;
 try{const res=await call("closures",{rows:todo.map(o=>({no:o.no,due:dueOf(o.no),eq:o.eq,name:o.name,freq:o.freq,act:o.act,on:toSer(o.onD),by:o.byU.trim(),rem:o.remk,row:o.row}))});
  todo.forEach(o=>saved.add(o.no));toast(`Saved ${res.added} order(s) to Maintenance Closure History.`+(res.added<todo.length?" Others were already in history.":""));await load()}
 catch(e){toast("Save failed: "+e.message,1)}$("#saveClose").disabled=false}
$("#saveClose").onclick=saveClosures;

const res=i=>!i.ok?"":i.ok==="OK"?"PASS":i.notes.trim().length>20?"ATTENTION":"FAIL";
function fillPick(){const v=$("#order").value;$("#order").innerHTML=`<option value="">Select the order number</option>`+orders.map(o=>`<option>${esc(o.no)}</option>`).join("");$("#order").value=orders.some(o=>o.no===v)?v:""}
function pick(no){cur=orders.find(o=>o.no===no)||null;
 items=cur?lib.filter(r=>r["Equipment Type"]===cur.type&&r.Frequency===cur.freq&&r.Active==="Yes").sort((a,b)=>a.Sequence-b.Sequence).map(r=>({r,reading:"",ok:"",notes:""})):[];
 $("#by").value=user();
 $("#ckHead").innerHTML=cur?[["EQ ID",cur.eq],["Equipment",cur.name],["Type",cur.type],["Frequency",cur.freq],["Location",cur.loc],["Criticality",cur.crit],["Checklist in Excel",cur.cl]].map(([a,b])=>`<div><span>${a}</span>${esc(b)}</div>`).join(""):"";
 $("#ckTable").hidden=!cur||!items.length;$("#ckBar").hidden=!cur||!items.length;
 if(cur&&!items.length)$("#ckHead").insertAdjacentHTML("beforeend",`<p class="empty">No active checklist items for ${esc(cur.type)} / ${esc(cur.freq)} in Checklist Library.</p>`);
 $("#items").innerHTML=items.map(({r},i)=>`<tr data-i="${i}"><td>${r.Sequence}</td><td>${esc(r["Check Item"])}<small>${esc(r["Checklist ID"])}</small></td>
 <td>${r["Reading Required"]==="Yes"?`<input data-k="reading" size="7"> ${esc(r.Unit)}`:"Not required"}</td><td>${esc(r.Mandatory)}</td><td>${esc(r["Acceptance Criteria"])}</td>
 <td><select data-k="ok"><option></option><option>OK</option><option>NOT OK</option></select></td><td><input data-k="notes" placeholder="Required if NOT OK (more than 20 characters)"></td><td class="r"></td></tr>`).join("");upd()}
function upd(){items.forEach((it,i)=>{const t=res(it),c=$("#items").rows[i]?.querySelector(".r");if(c)c.innerHTML=t?`<span class="chip ${t}">${t}</span>`:""});
 const rs=items.map(res),ok=items.length&&rs.every(x=>x==="PASS"||x==="ATTENTION")&&$("#by").value.trim();
 $("#ckState").textContent=ok?(rs.includes("ATTENTION")?"Complete, with attention items":"Complete, all pass"):`${rs.filter(x=>x==="PASS"||x==="ATTENTION").length} of ${items.length} items done`;$("#saveCk").disabled=!ok}
$("#items").addEventListener("input",e=>{const tr=e.target.closest("tr"),k=e.target.dataset.k;if(tr&&k){items[tr.dataset.i][k]=e.target.value;upd()}});
$("#by").oninput=e=>{localStorage.cmmsUser=e.target.value;upd()};$("#order").onchange=e=>pick(e.target.value);

$("#saveCk").onclick=async()=>{$("#saveCk").disabled=true;
 try{
  const by=$("#by").value.trim(),rec=ser(new Date()),due=dueOf(cur.no);
  const r=await call("measurements",{no:cur.no,rows:items.map(it=>[cur.no,due,cur.eq,cur.name,cur.type,cur.freq,cur.loc,cur.crit,it.r["Checklist ID"],it.r["Check Item"],it.r["Reading Required"],it.r.Unit||"",it.reading===""?"":isNaN(it.reading)?it.reading:Number(it.reading),it.r.Mandatory,it.r["Acceptance Criteria"],it.ok,res(it),it.notes.trim(),by,rec])});
  if(r.duplicate)return toast("This order is already in Measurement History.",1);
  toast(`Saved ${items.length} readings for ${cur.no} to Measurement History.`);await load();$("#order").value="";pick("")}
 catch(e){toast("Save failed: "+e.message,1);upd()}};

function tab(t){document.querySelectorAll("nav button").forEach(b=>b.classList.toggle("on",b.dataset.tab===t));$("#closing").hidden=t!=="closing"||!acct;$("#checklist").hidden=t!=="checklist"||!acct}
document.querySelectorAll("nav button").forEach(b=>b.onclick=()=>tab(b.dataset.tab));
$("#refresh").onclick=()=>load().then(()=>toast("Refreshed from Google Sheets.")).catch(e=>toast(e.message,1));

(async()=>{$("#refresh").hidden=false;$("#welcome").hidden=true;tab("closing");
 try{await load()}catch(e){toast("Could not read the Google Sheet: "+e.message,1)}
})();
})();
