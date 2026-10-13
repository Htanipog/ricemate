import {useEffect,useMemo,useRef,useState} from 'react';
import {FIELDS,autoMap,parseTranscript,validate} from './lib/parse.js';
import {open,rowsOf,detectHeader,riskReport,appendEntries,toBytes} from './lib/excel.js';
import {putEntry,allEntries,setKv,getKv} from './lib/db.js';
const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
const empty=Object.fromEntries(FIELDS.map(f=>[f.k,'']));
const stamp=()=>new Date().toISOString().replace(/[-:T]/g,'').slice(0,14).replace(/^(\d{8})/,'$1_');
const dupKey=f=>['name','phone','variety','quantity','date'].map(k=>String(f[k]||'').toLowerCase().trim()).join('|');
export default function App(){
 const [tab,setTab]=useState('home'),[buf,setBuf]=useState(null),[fname,setFname]=useState(''),[wb,setWb]=useState(null),[sheet,setSheet]=useState(''),[hdr,setHdr]=useState(-1),[map,setMap]=useState({}),[entries,setEntries]=useState([]),[form,setForm]=useState(empty),[tr,setTr]=useState(''),[lang,setLang]=useState('ta-IN'),[on,setOn]=useState(false),[msg,setMsg]=useState(null),[busy,setBusy]=useState(false),[online,setOnline]=useState(navigator.onLine);
 const rec=useRef(null),trRef=useRef('');
 const ok=m=>setMsg({t:'ok',m}),err=m=>setMsg({t:'err',m});
 const rows=useMemo(()=>wb?rowsOf(wb.Sheets[sheet]):[],[wb,sheet]);
 const headers=hdr>=0&&rows[hdr]?rows[hdr].map(String):[];
 useEffect(()=>{(async()=>{try{setEntries(await allEntries());const s=await getKv('file');if(s){const w=open(s.buf);setBuf(s.buf);setFname(s.name);setWb(w);setSheet(s.sheet);setHdr(s.hdr);setMap(s.map)}}catch(e){err('Could not restore saved data: '+e.message)}})();
  const u=()=>setOnline(navigator.onLine);addEventListener('online',u);addEventListener('offline',u);return()=>{removeEventListener('online',u);removeEventListener('offline',u)}},[]);
 const persist=p=>setKv('file',{buf,name:fname,sheet,hdr,map,...p});
 const pick=async(w,name,b,n)=>{const r=rowsOf(w.Sheets[name]),h=detectHeader(r),m=h>=0?autoMap(r[h].map(String)):{};setSheet(name);setHdr(h);setMap(m);await setKv('file',{buf:b,name:n,sheet:name,hdr:h,map:m});
  if(!r.length)err('This worksheet is empty. Choose another worksheet.');else if(h<0)err('No header row detected. Enter the header row number below.')};
 const onFile=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;if(!/\.(xlsx|xls)$/i.test(f.name))return err('Choose an .xlsx or .xls file.');
  try{const b=await f.arrayBuffer(),w=open(b);if(!w.SheetNames.length)return err('This workbook has no worksheets.');setBuf(b);setFname(f.name);setWb(w);ok('Workbook loaded. Your original file is not modified.');await pick(w,w.SheetNames[0],b,f.name)}catch(x){err('Cannot read this file: '+x.message)}};
 const setHeader=n=>{const h=n-1,m=autoMap((rows[h]||[]).map(String));setHdr(h);setMap(m);persist({hdr:h,map:m})};
 const setCol=(k,i)=>{const m={...map,[k]:i};setMap(m);persist({map:m})};
 const extract=()=>{const p=parseTranscript(trRef.current);setForm(f=>({...f,...p}));if(!Object.keys(p).length)err('Nothing could be extracted. Fill the fields manually.')};
 const listen=()=>{if(!SR)return err('Speech recognition is not supported in this browser. Use Chrome on Android, or type below.');
  if(on){rec.current?.stop();return}if(!online)err('Speech recognition needs internet. Type the details instead.');
  const r=new SR();r.lang=lang;r.interimResults=true;r.onresult=e=>{const t=[...e.results].map(x=>x[0].transcript).join(' ');trRef.current=t;setTr(t)};
  r.onerror=e=>err('Voice error: '+e.error+'. Use manual entry.');r.onend=()=>{setOn(false);extract()};rec.current=r;
  try{r.start();setOn(true);setMsg(null)}catch(x){err(x.message)}};
 const shown=FIELDS.filter(f=>!wb||map[f.k]>=0);
 const save=async(status,ex)=>{if(busy)return;const f=ex?ex.fields:form,v=validate(f,shown.map(x=>x.k));if(v.errors.length)return err(v.errors.join(' '));
  const key=dupKey(f);if(!ex&&entries.some(e=>e.key===key))return err('Duplicate: an identical entry already exists.');
  if(status==='Confirmed'&&!confirm('Confirm this entry?\n\n'+shown.map(x=>x.l+': '+(f[x.k]||'—')).join('\n')+(v.warnings.length?'\n\n⚠ '+v.warnings.join(' '):'')))return;
  setBusy(true);try{const e=ex?{...ex,status}:{id:crypto.randomUUID(),key,fields:f,status,createdAt:new Date().toISOString()};await putEntry(e);
   setEntries(p=>[e,...p.filter(x=>x.id!==e.id)]);if(!ex){setForm(empty);setTr('');trRef.current=''}
   ok(status==='Confirmed'?'Entry confirmed and stored on this device. Export to add it to Excel.':'Saved as Pending draft.')}catch(x){err('Save failed: '+x.message)}finally{setBusy(false)}};
 const exportX=async()=>{if(busy)return;const list=entries.filter(e=>e.status==='Confirmed').reverse();
  if(!buf||hdr<0)return err('Import a workbook and set its header row first.');if(!list.length)return err('No Confirmed entries to export.');
  const w=open(buf),risk=riskReport(w,fname);if(risk.block)return err(risk.block);
  if(!confirm(risk.warns.join('\n\n')+'\n\nExport a new copy with '+list.length+' new row(s)?'))return;
  setBusy(true);try{appendEntries(w,sheet,hdr,map,list);const out=toBytes(w),name=fname.replace(/\.[^.]+$/,'').replace(/_updated_\d{8}_\d{6}$/,'')+'_updated_'+stamp()+'.xlsx';
   const a=document.createElement('a'),u=URL.createObjectURL(new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));a.href=u;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),30000);
   const ex=list.map(e=>({...e,status:'Exported'}));for(const e of ex)await putEntry(e);setEntries(p=>p.map(e=>ex.find(x=>x.id===e.id)||e));
   setBuf(out);setFname(name);setWb(open(out));await setKv('file',{buf:out,name,sheet,hdr,map});ok('Download started: '+name+'. Check your Downloads folder. This copy is now the base for new entries.')}
  catch(x){err('Export failed: '+x.message)}finally{setBusy(false)}};
 const cnt=s=>entries.filter(e=>e.status===s).length;
 const T=[['home','Home'],['excel','Excel'],['voice','Voice / குரல்'],['list','Entries']];
 return <><main>
  <h1>RiceMate <span className="note">அரிசி உதவியாளர்</span></h1>
  {msg&&<div className={'msg '+msg.t} role="alert" onClick={()=>setMsg(null)}>{msg.m}</div>}
  {tab==='home'&&<>
   <div className="card"><b>{wb?fname:'No workbook loaded'}</b><div className="note">{wb?`Sheet: ${sheet} · ${rows.length} rows`:'Start on the Excel tab.'}</div>
    <p>Pending {cnt('Pending')} · Confirmed {cnt('Confirmed')} · Exported {cnt('Exported')}</p>
    <button className="b" onClick={()=>setTab('voice')}>🎤 New voice entry</button>
    <button className="b alt" onClick={()=>setTab('excel')}>Select Excel file</button>
    <button className="b alt" onClick={exportX} disabled={busy}>Export Updated Excel</button></div>
   <div className="card note">{online?'Online.':'Offline: voice recognition will not work; typing still does.'} Browser voice recognition usually sends audio to the browser vendor's server (internet needed). Entries stay on this phone in browser storage and can be lost if browser data is cleared, so export often. Nothing is uploaded by RiceMate.</div></>}
  {tab==='excel'&&<>
   <label htmlFor="f">Select Excel File (.xlsx / .xls)</label>
   <input id="f" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" onChange={onFile}/>
   {wb&&<><div className="card"><b>{fname}</b>
    <label>Worksheet</label><select value={sheet} onChange={e=>pick(wb,e.target.value,buf,fname)}>{wb.SheetNames.map(s=><option key={s}>{s}</option>)}</select>
    <label>Header row number</label><input type="number" min="1" value={hdr+1||''} onChange={e=>setHeader(+e.target.value)}/>
    {headers.length>0&&<><h2>Column mapping</h2>{FIELDS.map(f=><div key={f.k}><label>{f.l}</label><select value={map[f.k]??-1} onChange={e=>setCol(f.k,+e.target.value)}><option value={-1}>— not in workbook —</option>{headers.map((h,i)=>h&&<option key={i} value={i}>{h}</option>)}</select></div>)}</>}</div>
    <h2>Preview (last 8 rows)</h2><div className="scroll"><table><tbody>{rows.slice(Math.max(0,rows.length-8)).map((r,i)=><tr key={i}>{r.map((c,j)=><td key={j}>{String(c)}</td>)}</tr>)}</tbody></table></div></>}</>}
  {tab==='voice'&&<>
   <select aria-label="Language" value={lang} onChange={e=>setLang(e.target.value)}><option value="ta-IN">தமிழ் (Tamil)</option><option value="en-IN">English (India)</option></select>
   <button className={'mic'+(on?' on':'')} onClick={listen} aria-label="Microphone">{on?'■':'🎤'}</button>
   <p className="note" style={{textAlign:'center'}}>{SR?(on?'Listening… tap to stop':'Tap and speak. Needs internet.'):'Voice not supported here. Type below.'}</p>
   <label>Transcript (edit freely)</label><textarea rows="3" value={tr} onChange={e=>{setTr(e.target.value);trRef.current=e.target.value}}/>
   <button className="b alt" onClick={extract}>Extract fields from text</button>
   <div className="card"><h2>Confirm details</h2>
    {!wb&&<p className="note">No workbook loaded: all fields shown. Load your workbook so only its columns are used.</p>}
    {shown.map(f=><div key={f.k}><label htmlFor={'x'+f.k}>{f.l}</label><input id={'x'+f.k} type={f.k==='date'?'date':'text'} inputMode={['phone','quantity','price'].includes(f.k)?'decimal':'text'} value={form[f.k]} onChange={e=>setForm({...form,[f.k]:e.target.value})}/></div>)}
    <button className="b" disabled={busy} onClick={()=>save('Confirmed')}>Save Entry</button>
    <button className="b alt" disabled={busy} onClick={()=>save('Pending')}>Save as Pending draft</button>
    <button className="b alt" onClick={()=>{setForm(empty);setTr('');trRef.current=''}}>Clear</button></div></>}
  {tab==='list'&&<>
   <button className="b" onClick={exportX} disabled={busy}>Export Updated Excel</button>
   {!entries.length&&<p className="note">No entries yet. Record one on the Voice tab.</p>}
   {entries.map(e=><div className="card" key={e.id}><span className={'chip '+e.status}>{e.status}</span> <b>{e.fields.name}</b>
    <div className="note">{[e.fields.location,e.fields.variety,e.fields.quantity&&e.fields.quantity+' '+e.fields.unit,e.fields.price&&'₹'+e.fields.price,e.fields.payment].filter(Boolean).join(' · ')}</div>
    {e.status==='Pending'&&<button className="b alt" disabled={busy} onClick={()=>save('Confirmed',e)}>Review &amp; Confirm</button>}</div>)}</>}
 </main><nav>{T.map(([k,l])=><button key={k} aria-current={tab===k} onClick={()=>{setTab(k);setMsg(null)}}>{l}</button>)}</nav></>}
