const $ = (id) => document.getElementById(id);
const fmt = (n) => new Intl.NumberFormat('en-US').format(Math.round(Number(n || 0)));
const badgeClass = (txt='') => txt.includes('Priority 1') ? 'bad' : txt.includes('Priority 2') || txt.includes('High') ? 'warn' : txt.includes('Enabled') || txt.includes('Strong') ? 'good' : 'blue';
let DB = null;
let workbook = null;
let fileHandle = null;
let activeSection = 'summary';
let activeProspect = null;
let deferredPrompt = null;

function toast(msg){
  let t = document.querySelector('.toast');
  if(!t){ t = document.createElement('div'); t.className='toast'; document.body.appendChild(t); }
  t.textContent = msg; t.classList.add('show'); clearTimeout(window._toast); window._toast = setTimeout(()=>t.classList.remove('show'), 2200);
}

async function dbOpen(){
  return new Promise((resolve, reject)=>{
    const req = indexedDB.open('ascend-pwa', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function saveHandle(handle){
  try{
    const db = await dbOpen();
    await new Promise((resolve,reject)=>{
      const tx = db.transaction('handles','readwrite');
      tx.objectStore('handles').put(handle,'workbook');
      tx.oncomplete = resolve; tx.onerror = ()=>reject(tx.error);
    });
  }catch(e){ console.warn('Handle save failed', e); }
}
async function loadHandle(){
  try{
    const db = await dbOpen();
    return await new Promise((resolve,reject)=>{
      const tx = db.transaction('handles','readonly');
      const req = tx.objectStore('handles').get('workbook');
      req.onsuccess = ()=>resolve(req.result || null);
      req.onerror = ()=>reject(req.error);
    });
  }catch(e){ return null; }
}

function getSheetRows(name){
  return XLSX.utils.sheet_to_json(workbook.Sheets[name] || {}, {defval:''});
}
function mapKeyValue(rows, key, value){
  const out = {};
  rows.forEach(r => { out[r[key]] = r[value]; });
  return out;
}
function buildDBFromWorkbook(){
  DB = {
    design: {
      theme: mapKeyValue(getSheetRows('D01_Theme'),'ThemeKey','ThemeValue'),
      navigation: getSheetRows('D02_Navigation'),
      labels: mapKeyValue(getSheetRows('D04_Labels'),'LabelKey','LabelValue')
    },
    data: {
      prospects: getSheetRows('02_Prospects'),
      assessment_runs: getSheetRows('03_AssessmentRuns'),
      custom_code_inventory: getSheetRows('06_CustomCodeInventory'),
      atc_findings: getSheetRows('07_ATC_Findings'),
      roi_model: getSheetRows('14_ROI_Model'),
      roadmap_actions: getSheetRows('15_Roadmap_Actions')
    }
  };
}
function applyTheme(theme){
  if(!theme) return;
  const map = {'--primary': theme.PrimaryColor,'--secondary': theme.SecondaryColor,'--accent': theme.AccentColor,'--success': theme.SuccessColor,'--warning': theme.WarningColor,'--danger': theme.DangerColor};
  Object.entries(map).forEach(([k,v])=>{ if(v) document.documentElement.style.setProperty(k,v);});
  $('appTitle').textContent = theme.AppTitle || 'Ascend';
  $('appSubtitle').textContent = theme.AppSubtitle || 'SAP Transformation Agent';
}
async function readWorkbookFromHandle(handle){
  const file = await handle.getFile();
  const buf = await file.arrayBuffer();
  workbook = XLSX.read(buf, {type:'array'});
  buildDBFromWorkbook();
  applyTheme(DB.design.theme);
  activeProspect = DB.data.prospects[0]?.ProspectID || null;
  await saveHandle(handle);
  $('fileStatus').textContent = file.name;
  $('statusText').textContent = 'Workbook loaded';
  buildProspects();
  render();
}
function buildNav(){
  const visible = (DB?.design?.navigation || []).filter(x => String(x.Visible).toLowerCase() !== 'false');
  const groups = [...new Set(visible.map(x=>x.NavGroup))];
  $('nav').innerHTML = groups.map(g => {
    const items = visible.filter(x=>x.NavGroup===g).sort((a,b)=>Number(a.DisplayOrder)-Number(b.DisplayOrder)).map(s => `<button class="nav-btn ${s.SectionID===activeSection?'active':''}" data-sec="${s.SectionID}"><span>${s.Icon || '•'}</span><span>${s.SectionLabel}</span></button>`).join('');
    return `<div class="nav-group"><div class="nav-label">${g}</div>${items}</div>`;
  }).join('');
  document.querySelectorAll('.nav-btn').forEach(btn=>btn.onclick=()=>{activeSection=btn.dataset.sec; render();});
}
function buildProspects(){
  const items = DB?.data?.prospects || [];
  $('prospectSelect').innerHTML = items.map(p=>`<option value="${p.ProspectID}">${p.ProspectName} • ${p.Industry}</option>`).join('');
  $('prospectSelect').value = activeProspect;
  $('prospectSelect').onchange = (e)=>{ activeProspect = e.target.value; render(); };
}
function buildTabs(){
  const tabsBySection = {summary:['Clean Core','Outcomes','Risk'],analysis:['Strategy','Disposition','Impacts'],roadmap:['Now','Next','Scale']};
  $('tabs').innerHTML = (tabsBySection[activeSection] || ['Overview']).map((t,i)=>`<div class="tab">${t}</div>`).join('');
}
function getProspect(){ return (DB?.data?.prospects || []).find(x=>x.ProspectID===activeProspect) || {}; }
function getRun(){ return (DB?.data?.assessment_runs || []).filter(x=>x.ProspectID===activeProspect).slice(-1)[0] || {}; }
function byProspect(sheet){ return (DB?.data?.[sheet] || []).filter(x=>x.ProspectID===activeProspect); }
function panel(title, sub, body, right=''){ return `<div class="panel"><div class="panel-head"><div><div class="panel-title">${title}</div><div class="panel-sub">${sub || ''}</div></div>${right}</div><div class="panel-body">${body}</div></div>`; }
function renderSummary(){
  const p=getProspect(), run=getRun(), roi=byProspect('roi_model')[0]||{};
  return `<div class="grid">${panel('Prospect organization','demo company',`<div class="card"><h5>${p.ProspectName || ''}</h5><p>${p.Industry || ''} • ${p.TargetModel || ''}</p></div><div class="card"><h5>Owner <span class="badge blue">${p.OpportunityStage || ''}</span></h5><p>${p.Owner || ''}</p></div>`)}${panel('Executive summary','recommended direction',`<div class="hero-card"><div class="hero-top"><div><div class="hero-label">Recommended path</div><div class="hero-value">${run.RecommendedPath || 'N/A'}</div></div><div style="text-align:right"><div class="hero-label">Estimated program</div><div class="hero-value" style="font-size:24px">${run.EstimatedMonths || '-'} mo</div></div></div></div><div class="metric-row"><div class="kpi"><div class="label">Readiness</div><div class="value">${run.ReadinessScore || '-'}</div><div class="desc">Pre-signature posture</div></div><div class="kpi"><div class="label">Business case</div><div class="value">${run.BusinessCaseScore || '-'}</div><div class="desc">Executive urgency</div></div></div><div class="metric-row"><div class="kpi"><div class="label">Risk</div><div class="value">${run.RiskScore || '-'}</div><div class="desc">Lower is better</div></div><div class="kpi"><div class="label">Illustrative value</div><div class="value">${roi.IllustrativeValuePercent || '-'}%</div><div class="desc">ROI framing</div></div></div>`, `<span class="badge blue">${run.CleanCoreScore || '-'} clean core</span>`)}${panel('Quick facts','from workbook',`<div class="stack"><div class="card"><h5>Geography <span class="badge ${badgeClass(p.RegulatoryIntensity||'')} ">${p.RegulatoryIntensity || ''}</span></h5><p>${p.GeographicScope || ''} • customization ${p.CustomizationIntensity || ''}</p></div><div class="card"><h5>Source landscape <span class="badge blue">${p.SourceDB || ''}</span></h5><p>${p.SourceERPRelease || ''}</p></div></div>`)} </div>`;
}
function renderAnalysis(){
  const p=getProspect(), cci=byProspect('custom_code_inventory'), atc=byProspect('atc_findings');
  const total=cci.length*3625, dead=Math.round(total*.35), auto=Math.round(total*.4), ref=Math.round(total*.25);
  return `<div class="grid">${panel('Prospect organization','sample-style left context',`<div class="stack"><div class="card"><h5>${p.ProspectName || ''} <span class="badge blue">${p.Industry || ''}</span></h5><p>Demo dataset is prefilled so users start from a rich baseline.</p></div><div class="card"><h5>Customization <span class="badge ${badgeClass(p.CustomizationIntensity || '')}">${p.CustomizationIntensity || ''}</span></h5><p>Use this as the anchor for code-disposition conversations.</p></div></div>`)}${panel('Disposition strategy','spreadsheet-driven custom code plan',`<div class="hero-card"><div class="hero-top"><div><div class="hero-label">Total assessed objects</div><div class="hero-value">${fmt(total)}</div></div><div style="text-align:right"><div class="hero-label">Starting inventory</div><div class="hero-value" style="font-size:24px">${fmt(total)}</div></div></div></div><div class="flow-arrow">↓</div><div class="band green"><h4><span>Dead code decommissioning</span><span class="score">${fmt(dead)}</span></h4><div class="meta">identified via usage data / duplicate logic</div><div class="desc">35% of total</div></div><div class="flow-arrow">↓</div><div class="band blue"><h4><span>Auto-remediated code</span><span class="score">${fmt(auto)}</span></h4><div class="meta">successor APIs and quick fixes</div><div class="desc">40% of total</div></div><div class="flow-arrow">↓</div><div class="band amber"><h4><span>Manual refactoring / clean-core redesign</span><span class="score">${fmt(ref)}</span></h4><div class="meta">wrapper or architectural redesign required</div><div class="desc">25% of total</div></div>`)}${panel('Top simplification impacts','right-rail findings',`<div class="stack">${atc.slice(0,3).map(x=>`<div class="card"><h5>${x.FindingCategory} <span class="badge ${badgeClass(x.Priority || '')}">${x.Priority || ''}</span></h5><p>Action: ${x.RecommendedTreatment || ''}</p></div>`).join('')}</div>`)} </div>`;
}
function renderRoadmap(){
  const rows=byProspect('roadmap_actions');
  return `<div class="grid">${panel('Now','immediate controls',`<div class="stack">${rows.filter(x=>x.RoadmapPhase==='Now').map(x=>`<div class="card"><h5>${x.OwnerRole} <span class="badge blue">${x.Status}</span></h5><p>${x.ActionDescription}</p></div>`).join('')}</div>`)}${panel('Next / Scale','transition actions',`<div class="stack">${rows.filter(x=>x.RoadmapPhase!=='Now').map(x=>`<div class="card"><h5>${x.RoadmapPhase} <span class="badge warn">${x.Status}</span></h5><p>${x.ActionDescription} • ${x.OwnerRole}</p></div>`).join('')}</div>`)}${panel('Workbook-driven design','how this app works',`<div class="note">Design sheets control theme, labels, navigation, and rules. Data sheets control prospects, findings, ROI, and roadmap. Change Excel, then click Refresh from Workbook.</div>`)} </div>`;
}
function render(){
  if(!DB){ $('canvas').innerHTML = `<div class="note">Open the Ascend workbook to start. Use Microsoft Edge for the best experience.</div>`; return; }
  const navItem = DB.design.navigation.find(x=>x.SectionID===activeSection) || {};
  $('pageTitle').textContent = navItem.SectionLabel || 'Ascend';
  $('pageSubtitle').textContent = navItem.SectionSubtitle || 'Workbook-driven SaaS workspace';
  buildNav(); buildTabs();
  const html = activeSection==='analysis' ? renderAnalysis() : activeSection==='roadmap' ? renderRoadmap() : renderSummary();
  $('canvas').innerHTML = html;
}
async function openWorkbook(){
  if(!('showOpenFilePicker' in window)) return alert('Please use Microsoft Edge or Chrome on desktop.');
  const [handle] = await window.showOpenFilePicker({multiple:false,types:[{description:'Excel Workbook',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]});
  fileHandle = handle;
  await readWorkbookFromHandle(handle);
  toast('Workbook opened');
}
async function refreshWorkbook(){
  if(!fileHandle){ const remembered = await loadHandle(); if(remembered){ fileHandle = remembered; } }
  if(!fileHandle) return openWorkbook();
  await readWorkbookFromHandle(fileHandle);
  toast('Workbook refreshed');
}
async function saveWorkbook(){
  if(!fileHandle || !workbook) return toast('Open a workbook first');
  const wbout = XLSX.write(workbook, {bookType:'xlsx', type:'array'});
  const writable = await fileHandle.createWritable();
  await writable.write(wbout); await writable.close();
  $('statusText').textContent = 'Workbook saved';
  toast('Workbook saved');
}
async function restoreHandle(){
  const remembered = await loadHandle();
  if(remembered){
    fileHandle = remembered;
    try{ if((await fileHandle.queryPermission({mode:'readwrite'})) === 'granted'){ await readWorkbookFromHandle(fileHandle);} }catch(e){}
  }
}
window.addEventListener('beforeinstallprompt', (e)=>{ e.preventDefault(); deferredPrompt = e; $('installBtn').style.display='inline-flex'; });
async function installApp(){
  if(deferredPrompt){ deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt = null; $('installBtn').style.display='none'; }
  else toast('Use Edge address bar “App available” icon or Apps > Install this site as an app');
}
async function registerSW(){ if('serviceWorker' in navigator){ try{ await navigator.serviceWorker.register('./service-worker.js'); }catch(e){ console.warn(e);} } }
window.addEventListener('DOMContentLoaded', async ()=>{
  $('openBtn').onclick = openWorkbook; $('refreshBtn').onclick = refreshWorkbook; $('saveBtn').onclick = saveWorkbook; $('installBtn').onclick = installApp;
  await registerSW(); await restoreHandle(); render();
});
