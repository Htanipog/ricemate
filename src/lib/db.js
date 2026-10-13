const open=()=>new Promise((res,rej)=>{const r=indexedDB.open('ricemate',1);r.onupgradeneeded=()=>{r.result.createObjectStore('entries',{keyPath:'id'});r.result.createObjectStore('kv')};r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error)});
// Resolves only when the transaction has fully committed.
const tx=async(s,m,f)=>{const d=await open();return new Promise((res,rej)=>{const t=d.transaction(s,m);const q=f(t.objectStore(s));t.oncomplete=()=>res(q&&q.result);t.onerror=()=>rej(t.error);t.onabort=()=>rej(t.error)})};
export const putEntry=e=>tx('entries','readwrite',s=>s.put(e));
export const allEntries=async()=>(await tx('entries','readonly',s=>s.getAll())).sort((a,b)=>b.createdAt.localeCompare(a.createdAt));
export const setKv=(k,v)=>tx('kv','readwrite',s=>s.put(v,k));
export const getKv=k=>tx('kv','readonly',s=>s.get(k));
