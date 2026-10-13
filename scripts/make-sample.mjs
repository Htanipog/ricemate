import * as XLSX from 'xlsx';
const rows=[['S.No','Customer Name','Phone','Location','Rice Variety','Quantity','Unit','Price per Kg','Total','Payment Status','Order Date','Remarks'],
[1,'Murugan Traders','9843012345','Pollachi','Ponni',10,'bags',48,'=F2*H2*25','Paid','2026-10-01',''],
[2,'Annapoorna Stores','9944011223','Tiruppur','Sona Masoori',15,'bags',52,'=F3*H3*25','Pending','2026-10-05','']];
const ws=XLSX.utils.aoa_to_sheet(rows);ws['B2'].t='s';ws['C2']={t:'s',v:'9843012345'};ws['C3']={t:'s',v:'9944011223'};
const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Customers');
XLSX.writeFile(wb,'sample/RiceMate_sample.xlsx');console.log('Created sample/RiceMate_sample.xlsx (TEST FILE ONLY)');
