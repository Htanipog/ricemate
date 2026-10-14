import * as XLSX from 'xlsx';
export const open=b=>XLSX.read(b,{type:'array'});
export const rowsOf=ws=>ws?XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,blankrows:true}):[];
export function detectHeader(rows){let best=-1,sc=0;rows.slice(0,20).forEach((r,i)=>{const n=r.filter(c=>typeof c==='string'&&c.trim()&&isNaN(c)).length;if(n>sc){sc=n;best=i}});return best}
export function riskReport(wb,name){
 if(wb.vbaraw||/\.(xlsm|xlsb)$/i.test(name))return{block:'This workbook has macros or a binary format. Exporting would destroy them, so export is blocked. Nothing was changed.',warns:[]};
 let f=0;for(const s of wb.SheetNames)for(const a in wb.Sheets[s])if(a[0]!=='!'&&wb.Sheets[s][a].f)f++;
 const w=['A NEW copy is created; your original file is never modified.','Colours, column widths, charts, images, conditional formatting and data validation may be lost (the free SheetJS build does not keep all formatting).'];
 if(f)w.push(f+' formula cells found. Formulas are kept, but formulas are NOT auto-filled for the new rows.');
 if(/\.xls$/i.test(name))w.push('.xls will be converted to .xlsx.');
 if(wb.Workbook&&wb.Workbook.Names&&wb.Workbook.Names.length)w.push('Defined names / external links may not survive.');
 return{block:'',warns:w}}
// Every header becomes a form column, except blanks, serial-number columns (auto-numbered) and formula columns (left alone).
export function colInfo(ws,rows,hdr){if(!ws)return[];
 return(rows[hdr]||[]).map((x,i)=>{const h=String(x).trim();let kind='field';
  if(!h)kind='skip';else if(/^(s\.?\s*no\.?|sl\.?\s*no\.?|sr\.?\s*no\.?|serial.*|#)$/i.test(h))kind='serial';
  else for(let r=hdr+1;r<rows.length;r++){const c=ws[XLSX.utils.encode_cell({r,c:i})];if(c&&c.f){kind='formula';break}}
  return{i,h,kind}})}
// Appends after the last non-empty row. entries[].vals = {columnIndex:value}. Existing cells untouched; null = no cell written.
export function appendEntries(wb,sheetName,hdr,map,entries){
 const ws=wb.Sheets[sheetName],rows=rowsOf(ws);let last=hdr;
 rows.forEach((r,i)=>{if(r.some(c=>String(c).trim()!==''))last=Math.max(last,i)});
 const start=last+1,ncols=Math.max(...rows.map(r=>r.length),1);
 const ser=colInfo(ws,rows,hdr).filter(c=>c.kind==='serial').map(c=>({i:c.i,n:Math.max(0,...rows.slice(hdr+1).map(r=>parseFloat(r[c.i])||0))}));
 entries.forEach((e,k)=>{const row=Array(ncols).fill(null);
  for(const i in e.vals){const v=e.vals[i],c=+i,t=String(v).trim();if(t==='')continue;
   const num=c===map.quantity||c===map.price||(c!==map.phone&&/^(0|[1-9]\d{0,8})(\.\d+)?$/.test(t));row[c]=num?Number(t):String(v)}
  ser.forEach(s=>{row[s.i]=++s.n});
  XLSX.utils.sheet_add_aoa(ws,[row],{origin:{r:start+k,c:0}})});
 return start}
export const toBytes=wb=>XLSX.write(wb,{bookType:'xlsx',type:'array'});
