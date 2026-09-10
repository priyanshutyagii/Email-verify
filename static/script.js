const fileInput=document.getElementById("fileInput");
const dropzone=document.getElementById("dropzone");
const startBtn=document.getElementById("startBtn");
const exportBtn=document.getElementById("exportBtn");
const exportValidBtn=document.getElementById("exportValidBtn");
const fileTitle=document.getElementById("fileTitle");
const fileSub=document.getElementById("fileSub");
const totalEl=document.getElementById("total");
const validEl=document.getElementById("valid");
const invalidEl=document.getElementById("invalid");
const riskyEl=document.getElementById("risky");
const unknownEl=document.getElementById("unknown");
const bar=document.getElementById("bar");
const percent=document.getElementById("percent");
const progressText=document.getElementById("progressText");
const tbody=document.getElementById("tbody");
const search=document.getElementById("search");

let selectedFile=null, jobId=null, allResults=[], pollTimer=null;

fileInput.addEventListener("change",()=>setFile(fileInput.files[0]));
["dragenter","dragover"].forEach(e=>dropzone.addEventListener(e,x=>{x.preventDefault();dropzone.classList.add("drag")}));
["dragleave","drop"].forEach(e=>dropzone.addEventListener(e,x=>{x.preventDefault();dropzone.classList.remove("drag")}));
dropzone.addEventListener("drop",e=>setFile(e.dataTransfer.files[0]));

function setFile(file){
  if(!file)return;
  const ok=/\.(xlsx|csv)$/i.test(file.name);
  if(!ok){alert("Please select an XLSX or CSV file.");return}
  if(file.size>500*1024*1024){alert("File is too large. Maximum upload size is 500 MB.");return}
  selectedFile=file;
  fileTitle.textContent=file.name;
  fileSub.textContent=(file.size/1024/1024).toFixed(2)+" MB";
  startBtn.disabled=false;
}

startBtn.addEventListener("click",async()=>{
  if(!selectedFile)return;
  if(pollTimer){clearTimeout(pollTimer);pollTimer=null}
  startBtn.disabled=true;
  exportBtn.disabled=true;
  exportValidBtn.disabled=true;
  allResults=[];
  updateCounts({VALID:0,INVALID:0,RISKY:0,UNKNOWN:0});
  render([]);
  bar.style.width="0%";
  percent.textContent="0%";
  progressText.textContent="Uploading and starting...";
  const fd=new FormData();
  fd.append("file",selectedFile);
  fd.append("smtp",document.getElementById("smtpCheck").checked?"true":"false");
  try{
    const r=await fetch("/api/verify",{method:"POST",body:fd});
    let data={};
    try{data=await r.json();}catch{}
    if(r.status===413)throw new Error(data.error||"File is too large. Maximum upload size is 500 MB.");
    if(!r.ok)throw new Error(data.error||"Upload failed");
    jobId=data.job_id;
    totalEl.textContent=data.total.toLocaleString();
    poll();
  }catch(e){
    alert(e.message);
    startBtn.disabled=false;
    progressText.textContent="Ready";
  }
});

async function poll(){
  try{
    const r=await fetch("/api/status/"+jobId);
    const data=await r.json();
    if(!r.ok)throw new Error(data.error||"Status error");
    const pct=data.total?Math.round(data.done/data.total*100):100;
    bar.style.width=pct+"%";
    percent.textContent=pct+"%";
    progressText.textContent=data.status==="completed"
      ?"Completed"
      :`Checking ${data.done.toLocaleString()} of ${data.total.toLocaleString()}`;

    // Backend may return only the latest batch; never include null placeholders.
    allResults=(data.results||[]).filter(x=>x&&x.email!=null);
    if(data.counts){
      updateCounts(data.counts);
    }else{
      updateCountsFromResults(allResults);
    }
    render(allResults);

    if(data.status==="completed"){
      exportBtn.disabled=false;
      exportValidBtn.disabled=false;
      startBtn.disabled=false;
      pollTimer=null;
      return;
    }
    pollTimer=setTimeout(poll,900);
  }catch(e){
    progressText.textContent="Error: "+e.message;
    startBtn.disabled=false;
    pollTimer=null;
  }
}

function updateCounts(counts){
  validEl.textContent=Number(counts.VALID||0).toLocaleString();
  invalidEl.textContent=Number(counts.INVALID||0).toLocaleString();
  riskyEl.textContent=Number(counts.RISKY||0).toLocaleString();
  unknownEl.textContent=Number(counts.UNKNOWN||0).toLocaleString();
}

function updateCountsFromResults(results){
  updateCounts({
    VALID:results.filter(x=>x.status==="VALID").length,
    INVALID:results.filter(x=>x.status==="INVALID").length,
    RISKY:results.filter(x=>x.status==="RISKY").length,
    UNKNOWN:results.filter(x=>x.status==="UNKNOWN").length
  });
}

function render(results){
  const q=search.value.trim().toLowerCase();
  const filtered=results.filter(x=>x&&(!q||String(x.email).toLowerCase().includes(q)));
  tbody.innerHTML=filtered.length?filtered.map(x=>`
    <tr><td>${esc(x.email)}</td><td><span class="status status-${esc(x.status).toLowerCase()}">${esc(x.status)}</span></td><td>${esc(x.reason)}</td></tr>
  `).join(""):`<tr><td colspan="3" class="empty">${results.length?"No matching records.":"No verification started yet."}</td></tr>`;
}
search.addEventListener("input",()=>render(allResults));

async function downloadFile(url, fallbackName){
  try{
    const r=await fetch(url);
    const type=(r.headers.get("content-type")||"").toLowerCase();
    if(!r.ok || type.includes("application/json")){
      let msg="Download failed";
      try{const d=await r.json(); msg=d.error||msg;}catch{}
      alert(msg);
      return;
    }
    const blob=await r.blob();
    const disp=r.headers.get("Content-Disposition")||"";
    const match=/filename\*?=(?:UTF-8''|")?([^\";]+)/i.exec(disp);
    const name=match?decodeURIComponent(match[1].replace(/"/g,"").trim()):fallbackName;
    const a=document.createElement("a");
    a.href=URL.createObjectURL(blob);
    a.download=name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  }catch(e){
    alert(e.message||"Download failed");
  }
}

exportBtn.addEventListener("click",()=>{if(jobId)downloadFile("/api/export/"+jobId,"verified_emails.xlsx")});
exportValidBtn.addEventListener("click",()=>{if(jobId)downloadFile("/api/export/"+jobId+"?only=valid","valid_emails.xlsx")});

function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
