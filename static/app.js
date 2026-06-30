const $ = (selector) => document.querySelector(selector);
const state = { files: [], rows: [], batchId: null, master: null };
const fields = ["particulars","bill_no","invoice_date","sb_no","sb_date","port_code","currency","foreign_amount","exchange_rate","inr","taxable_value","igst","fob","drawback","rodtep"];
const numeric = new Set(["foreign_amount","exchange_rate","taxable_value","igst","fob","drawback","rodtep"]);

function toast(message, error=false){ const el=$("#toast"); el.textContent=message; el.className=`toast show${error?" error":""}`; clearTimeout(el.timer); el.timer=setTimeout(()=>el.className="toast",3600); }
function busy(show, label="Reading the documents…"){ $("#busy strong").textContent=label; $("#busy").classList.toggle("hidden",!show); }
async function api(url, options={}){ const response=await fetch(url,options); const data=await response.json(); if(!response.ok||!data.ok) throw new Error(data.error||"Something went wrong."); return data; }

function renderMaster(master){
  state.master=master; const badge=$("#masterBadge"), info=$("#masterInfo");
  if(!master?.ready){ badge.className="badge muted"; badge.textContent="Not configured"; return; }
  badge.className="badge success"; badge.textContent="Ready";
  info.innerHTML=`<span class="status-orb">✓</span><div><strong>${escapeHtml(master.filename)}</strong><small>${escapeHtml(master.sheet)} · ${master.records} records · ${master.recognized_columns}/${master.total_columns} recognized columns</small></div>`;
}
function escapeHtml(value){ return String(value??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

async function loadStatus(){ try{ const data=await api("/api/status"); renderMaster(data.master); }catch(error){toast(error.message,true);} }

$("#useMaster").addEventListener("click",()=>$("#masterFile").click());
$("#masterFile").addEventListener("change",async event=>{
  const file=event.target.files[0]; if(!file)return;
  const form=new FormData(); form.append("action","existing"); form.append("master",file);
  busy(true,"Preparing your living master…");
  try{const data=await api("/api/master",{method:"POST",body:form});renderMaster(data.master);toast("Existing master is now active.");}
  catch(error){toast(error.message,true);}finally{busy(false);}
});
$("#newMaster").addEventListener("click",async()=>{
  const form=new FormData(); form.append("action","new"); busy(true,"Creating a polished master…");
  try{const data=await api("/api/master",{method:"POST",body:form});renderMaster(data.master);toast("New master created.");}
  catch(error){toast(error.message,true);}finally{busy(false);}
});

function setFiles(list){ state.files=[...list].filter(file=>file.name.toLowerCase().endsWith(".pdf")); $("#pdfCount").textContent=`${state.files.length} file${state.files.length===1?"":"s"}`; $("#extractButton").disabled=!state.files.length; $("#fileList").innerHTML=state.files.map(file=>`<span class="file-chip">◆ ${escapeHtml(file.name)} · ${(file.size/1024).toFixed(0)} KB</span>`).join(""); }
$("#pdfFiles").addEventListener("change",event=>setFiles(event.target.files));
const drop=$("#dropZone");
["dragenter","dragover"].forEach(name=>drop.addEventListener(name,event=>{event.preventDefault();drop.classList.add("dragging");}));
["dragleave","drop"].forEach(name=>drop.addEventListener(name,event=>{event.preventDefault();drop.classList.remove("dragging");}));
drop.addEventListener("drop",event=>setFiles(event.dataTransfer.files));

$("#extractButton").addEventListener("click",async()=>{
  if(!state.files.length)return; const form=new FormData(); state.files.forEach(file=>form.append("pdfs",file)); busy(true);
  try{const data=await api("/api/extract",{method:"POST",body:form});state.rows=data.rows;state.batchId=data.batch_id;renderRows();$("#reviewSection").classList.remove("hidden");$("#successPanel").classList.add("hidden");$("#reviewSection").scrollIntoView({behavior:"smooth",block:"start"});toast(`Extracted ${data.rows.length} record${data.rows.length===1?"":"s"}.`);}
  catch(error){toast(error.message,true);}finally{busy(false);}
});

function editor(field,row,index){
  const value=row[field]??"";
  if(field==="currency") return `<select data-row="${index}" data-field="currency">${["USD","EUR","GBP","JPY"].map(c=>`<option ${value===c?"selected":""}>${c}</option>`).join("")}</select>`;
  if(field==="inr") return `<td class="formula-cell" data-inr="${index}">${formatMoney(row.inr,"₹")}</td>`;
  const type=numeric.has(field)?"number":field.includes("date")?"date":"text"; const step=numeric.has(field)?"any":"";
  return `<input data-row="${index}" data-field="${field}" type="${type}" ${step?`step="${step}"`:""} value="${escapeHtml(value)}">`;
}
function formatMoney(value,symbol=""){const number=Number(value);return Number.isFinite(number)?`${symbol}${number.toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`:"—";}
function renderRows(){
  const body=$("#reviewTable tbody");
  body.innerHTML=state.rows.map((row,index)=>`<tr class="${row.warnings?.length?"needs-review":""}">${fields.map(field=>field==="inr"?editor(field,row,index):`<td class="${field==="particulars"?"particulars":""}">${editor(field,row,index)}</td>`).join("")}<td class="source-cell">${escapeHtml(row.source_file)}${row.warnings?.length?`<br><span title="${escapeHtml(row.warnings.join("; "))}" style="color:#a84a43">${row.warnings.length} field warning(s)</span>`:""}</td></tr>`).join("");
  const avg=Math.round(state.rows.reduce((sum,row)=>sum+(row.confidence||0),0)/state.rows.length); $("#confidenceBadge").textContent=`${avg}% field confidence`;
  body.querySelectorAll("input,select").forEach(input=>input.addEventListener("input",event=>{const row=state.rows[+event.target.dataset.row],field=event.target.dataset.field;row[field]=numeric.has(field)?(event.target.value===""?null:Number(event.target.value)):event.target.value;if(field==="foreign_amount"||field==="exchange_rate"){row.inr=(Number(row.foreign_amount)||0)*(Number(row.exchange_rate)||0);document.querySelector(`[data-inr="${event.target.dataset.row}"]`).textContent=formatMoney(row.inr,"₹");}}));
}
$("#clearReview").addEventListener("click",()=>{state.rows=[];state.batchId=null;$("#reviewSection").classList.add("hidden");});
$("#appendButton").addEventListener("click",async()=>{
  if(!state.master?.ready){toast("Choose or create a master workbook first.",true);return;} busy(true,"Appending without disturbing manual columns…");
  try{const data=await api("/api/append",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({batch_id:state.batchId,rows:state.rows})});renderMaster(data.master);const r=data.result;$("#successTitle").textContent=`${r.added+r.updated} record${r.added+r.updated===1?"":"s"} safely processed.`;$("#successText").textContent=`${r.added} added · ${r.updated} updated · manual columns preserved`;$("#successPanel").classList.remove("hidden");$("#successPanel").scrollIntoView({behavior:"smooth"});toast("Master workbook updated successfully.");}
  catch(error){toast(error.message,true);}finally{busy(false);}
});

loadStatus();
