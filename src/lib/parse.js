export const FIELDS=[
{k:'name',l:'Customer name',re:/name|party|shop|client|customer|பெயர்/i},
{k:'phone',l:'Phone',re:/phone|mobile|contact|cell|mob|போன்/i},
{k:'location',l:'Location',re:/location|city|place|area|town|village|address|ஊர்/i},
{k:'variety',l:'Rice variety',re:/variety|rice|product|item|brand|அரிசி/i},
{k:'quantity',l:'Quantity',re:/qty|quantity|bags|weight|எண்ணிக்கை/i},
{k:'unit',l:'Unit',re:/unit/i},
{k:'price',l:'Price',re:/price|rate|விலை/i},
{k:'payment',l:'Payment status',re:/payment|pay|status/i},
{k:'date',l:'Order date',re:/date|தேதி/i},
{k:'remarks',l:'Remarks',re:/remark|note|comment|குறிப்பு/i}];
const ORDER=['phone','date','unit','quantity','price','payment','variety','location','remarks','name'];
// Dynamic header mapping: header text -> field. Returns {field:columnIndex|-1}
export function autoMap(headers){const used=new Set(),m={};
 for(const k of ORDER){const re=FIELDS.find(f=>f.k===k).re;const i=headers.findIndex((h,i)=>!used.has(i)&&h&&re.test(h));m[k]=i;if(i>=0)used.add(i)}return m}
const VAR=['sona masoori','seeraga samba','jeera samba','karuppu kavuni','basmati','ponni','idli rice','matta','raw rice','boiled rice','பொன்னி','பாஸ்மதி','சீரக சம்பா','மட்டை'];
const UNIT=u=>/bag|மூட்டை/i.test(u)?'bags':/kg|kilo|கிலோ/i.test(u)?'kg':/quintal|குவிண்டால்/i.test(u)?'quintal':'tons';
const cap=s=>s.replace(/\b\w/g,c=>c.toUpperCase());
export const today=()=>new Date().toISOString().slice(0,10);
// Extracts ONLY what is present in the transcript. Missing values stay absent.
export function parseTranscript(t){const o={},s=(t||'').trim();if(!s)return o;
 const d=s.replace(/(\d)[\s-]+(?=\d)/g,'$1');
 const ph=d.match(/(?<!\d)(?:\+?91)?([6-9]\d{9})(?!\d)/);if(ph)o.phone=ph[1];
 const pr=d.match(/(\d+(?:\.\d+)?)\s*(?:rupees?|rs\.?|₹|ரூபாய்)\s*(?:per|\/|a)?\s*(kg|kilo|bag|கிலோ|மூட்டை)?/i);
 if(pr){o.price=pr[1];if(pr[2])o.remarks='Price per '+UNIT(pr[2])}
 const q=d.match(/(\d+(?:\.\d+)?)\s*(bags?|kgs?|kilos?|quintals?|tons?|tonnes?|மூட்டை|கிலோ|குவிண்டால்)/i);
 if(q){o.quantity=q[1];o.unit=UNIT(q[2])}
 if(/pending|unpaid|not paid|due|balance|பாக்கி|நிலுவை/i.test(s))o.payment='Pending';
 else if(/partial|advance|அட்வான்ஸ்/i.test(s))o.payment='Partial';
 else if(/paid|cash|gpay|upi|received|செலுத்தி|கொடுத்/i.test(s))o.payment='Paid';
 const low=s.toLowerCase(),v=VAR.find(x=>low.includes(x));if(v)o.variety=/^[a-z ]+$/.test(v)?cap(v):v;
 if(/today|இன்று/i.test(s))o.date=today();
 const dm=s.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);if(dm)o.date=`${dm[3]}-${dm[2].padStart(2,'0')}-${dm[1].padStart(2,'0')}`;
 const seg=s.split(/[,;،]/).map(x=>x.trim()).filter(Boolean);
 const plain=x=>x&&!/\d/.test(x)&&!VAR.some(w=>x.toLowerCase().includes(w))&&!/pending|paid|payment|phone/i.test(x);
 if(plain(seg[0]))o.name=seg[0];
 if(plain(seg[1]))o.location=seg[1];
 return o}
export function validate(f,keys){const errors=[],warnings=[],has=k=>keys.includes(k);
 if(has('name')&&!String(f.name||'').trim())errors.push('Customer name is required.');
 const n=x=>x!==''&&x!=null&&!/^\d+(\.\d+)?$/.test(String(x).trim());
 if(has('quantity')&&f.quantity!==''&&(n(f.quantity)||+f.quantity<=0))errors.push('Quantity must be a number above 0.');
 if(has('price')&&f.price!==''&&n(f.price))errors.push('Price must be a number (0 or more).');
 if(has('phone')){const p=String(f.phone||'');if(!p)warnings.push('Phone number is missing.');else if(!/^\d{10}$/.test(p))warnings.push('Phone is not 10 digits.')}
 return{errors,warnings}}
