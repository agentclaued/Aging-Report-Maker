/* New Ageing: runs fully in the browser (ExcelJS). Adapted from Ageing_New_Entries_Finder.html. */
(function(){
const $=id=>document.getElementById('na-'+id);
const wbs={};
const msg=(t,e)=>{$('msg').textContent=t;$('msg').className=e?'err':''};
function setup(n){
  const d=$('d'+n),f=$('f'+n);
  d.onclick=()=>f.click();
  d.ondragover=e=>{e.preventDefault();d.classList.add('on')};d.ondragleave=()=>d.classList.remove('on');
  d.ondrop=e=>{e.preventDefault();d.classList.remove('on');if(e.dataTransfer.files[0])load(n,e.dataTransfer.files[0])};
  f.onchange=()=>{if(f.files[0])load(n,f.files[0]);f.value=''};
}
setup(1);setup(2);
async function load(n,file){
  msg('Reading '+file.name+' …');$('go').disabled=true;
  try{
    const wb=new ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());
    wbs[n]=wb;const d=$('d'+n);d.classList.add('has');d.querySelector('.fn').textContent=file.name+' ('+wb.worksheets.length+' sheets)';
    if(n===2){const s=$('sh'),cur=s.value;s.innerHTML=wb.worksheets.map(w=>'<option>'+esc(w.name)+'</option>').join('');
      const pick=wb.worksheets.find(w=>w.name===cur)||wb.worksheets.find(w=>/full ageing/i.test(w.name));if(pick)s.value=pick.name}
    msg('');
  }catch(e){msg('Could not read '+file.name+': '+e.message,1)}
  $('go').disabled=!(wbs[1]&&wbs[2]);
}
const val=v=>{if(v==null)return '';if(v instanceof Date)return v;if(typeof v==='object'){if(v.result!==undefined)return val(v.result);if(v.richText)return v.richText.map(t=>t.text).join('');if(v.text!==undefined)return v.text;return ''}return v};
const norm=v=>String(val(v)).trim().toUpperCase();
function findHeader(ws,name){
  const want=name.trim().toLowerCase();
  for(let r=1;r<=Math.min(ws.rowCount,15);r++){const row=ws.getRow(r);
    for(let c=1;c<=ws.columnCount;c++)if(String(val(row.getCell(c).value)).trim().toLowerCase()===want)return{r,c}}
  return null;
}
let R=null;
$('go').onclick=()=>{
  try{
    const sn=$('sh').value,col=$('col').value;
    const w2=wbs[2].getWorksheet(sn),w1=wbs[1].getWorksheet(sn);
    if(!w2)throw Error('Sheet "'+sn+'" not found in NEW file');if(!w1)throw Error('Sheet "'+sn+'" not found in OLD file');
    const h1=findHeader(w1,col),h2=findHeader(w2,col);
    if(!h2)throw Error('Column "'+col+'" not found in NEW file');if(!h1)throw Error('Column "'+col+'" not found in OLD file');
    const old=new Set();let oc=0;
    for(let r=h1.r+1;r<=w1.rowCount;r++){const k=norm(w1.getRow(r).getCell(h1.c).value);if(k){old.add(k);oc++}}
    const cur=new Set(),newRows=[];let nc=0;
    for(let r=h2.r+1;r<=w2.rowCount;r++){const k=norm(w2.getRow(r).getCell(h2.c).value);if(!k)continue;nc++;cur.add(k);if(!old.has(k))newRows.push(r)}
    const goneRows=[];for(let r=h1.r+1;r<=w1.rowCount;r++){const k=norm(w1.getRow(r).getCell(h1.c).value);if(k&&!cur.has(k))goneRows.push(r)}
    R={w2,w1,h1,h2,newRows,goneRows};R.pc2=findCol(w2,h2.r,'Assigned Person');R.pc1=findCol(w1,h1.r,'Assigned Person');P='';fillPeople();
    $('sO').textContent=oc.toLocaleString();$('sN').textContent=nc.toLocaleString();$('sW').textContent=newRows.length.toLocaleString();$('sG').textContent=goneRows.length.toLocaleString();
    $('res').hidden=false;$('dl').disabled=!newRows.length;$('dlg').disabled=!goneRows.length;
    msg('Done. Matching is by Ref# value (trimmed, case-insensitive).');
    render();
  }catch(e){msg(e.message,1)}
};
const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
function render(){
  const {w2,h2}=R,newRows=selRows('new'),q=$('q').value.trim().toLowerCase(),cols=w2.columnCount;
  let h='<thead><tr>';for(let c=1;c<=cols;c++)h+='<th>'+esc(val(w2.getRow(h2.r).getCell(c).value))+'</th>';h+='</tr></thead><tbody>';
  let shown=0,drawn=0;
  for(const r of newRows){const row=w2.getRow(r);const cells=[];let txt='';
    for(let c=1;c<=cols;c++){const v=val(row.getCell(c).value);const num=typeof v==='number';
      let t=v;if(v instanceof Date)t=v.toISOString().slice(0,10);else if(num)t=v.toLocaleString(undefined,{maximumFractionDigits:2});
      txt+=' '+t;cells.push('<td'+(num?' class="n"':'')+'>'+esc(t)+'</td>')}
    if(q&&!txt.toLowerCase().includes(q))continue;
    shown++;if(drawn<1000){drawn++;h+='<tr>'+cells.join('')+'</tr>'}}
  $('tb').innerHTML=h+'</tbody>';
  $('cnt').textContent=shown>1000?'showing first 1,000 of '+shown.toLocaleString()+(q?'':' (download has all)'):shown.toLocaleString()+' rows shown';
}
$('q').oninput=()=>R&&render();
// ---- person filter: pick a name from the Assigned Person column of the new entries
let P='';
function findCol(ws,hr,name){const want=name.toLowerCase();for(let c=1;c<=ws.columnCount;c++)if(String(val(ws.getRow(hr).getCell(c).value)).trim().toLowerCase()===want)return c;return 0}
const personOf=(ws,col,r)=>col?String(val(ws.getRow(r).getCell(col).value)).trim():'';
function selRows(which){
  const rows=which==='new'?R.newRows:R.goneRows,ws=which==='new'?R.w2:R.w1,col=which==='new'?R.pc2:R.pc1;
  return P&&col?rows.filter(r=>personOf(ws,col,r)===P):rows;
}
const personSuffix=()=>P?'_'+P.replace(/[^\w-]+/g,'_'):'';
function fillPeople(){
  const sel=$('person'),box=$('personbox');
  if(!R.pc2){box.hidden=true;sel.innerHTML='';return}
  const counts=new Map();for(const r of R.newRows){const n=personOf(R.w2,R.pc2,r)||'(blank)';counts.set(n,(counts.get(n)||0)+1)}
  const names=[...counts.keys()].sort((a,b)=>a.localeCompare(b));
  sel.innerHTML='<option value="">All people ('+R.newRows.length.toLocaleString()+')</option>'+names.map(n=>'<option value="'+esc(n==='(blank)'?'':n)+'"'+(n==='(blank)'?' disabled':'')+'>'+esc(n)+' ('+counts.get(n).toLocaleString()+')</option>').join('');
  box.hidden=false;
}
function syncButtons(){$('dl').disabled=!selRows('new').length;$('dlg').disabled=!selRows('gone').length}
$('person').onchange=()=>{P=$('person').value;if(R){render();syncButtons()}};
// new workbook with same title rows, header, column widths and cell styles as the source sheet
async function exportRows(ws,h,rows,label){
  const wb=new ExcelJS.Workbook(),out=wb.addWorksheet(ws.name);
  const cols=ws.columnCount;
  for(let c=1;c<=cols;c++){const src=ws.getColumn(c),oc=out.getColumn(c);if(src.width)oc.width=src.width;if(src.hidden)oc.hidden=true}
  const copy=(sr,dr)=>{const s=ws.getRow(sr),d=out.getRow(dr);if(s.height)d.height=s.height;
    for(let c=1;c<=cols;c++){const sc=s.getCell(c),dc=d.getCell(c);
      let v=sc.value;if(v&&typeof v==='object'&&!(v instanceof Date)&&(v.formula!==undefined||v.sharedFormula!==undefined))v=v.result!==undefined?v.result:null;
      dc.value=v;dc.style=JSON.parse(JSON.stringify(sc.style||{}))}};
  const strip=$('strip').checked,off=strip?h.r-1:0;
  for(let r=strip?h.r:1;r<=h.r;r++)copy(r,r-off);
  rows.forEach((r,i)=>copy(r,h.r+1+i-off));
  const last=Math.max(h.r+rows.length,h.r+1);
  if(!strip)for(let r=1;r<h.r;r++)for(let c=1;c<=cols;c++){const v=ws.getRow(r).getCell(c).value;
    if(v&&typeof v==='object'&&!(v instanceof Date)&&(v.sharedFormula||/^SUM\(/i.test(v.formula||''))){const L=out.getColumn(c).letter;out.getRow(r).getCell(c).value={formula:'SUM('+L+(h.r+1)+':'+L+last+')'}}}
  out.autoFilter={from:{row:h.r-off,column:1},to:{row:h.r-off,column:cols}};
  const buf=await wb.xlsx.writeBuffer();
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  a.download=label+'_'+new Date().toISOString().slice(0,10)+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),5000);if(window.processProgress&&window.processProgress.notify)window.processProgress.notify('\u2714 Completed',a.download,false);
}
async function dl(b,fn){b.disabled=true;const t=b.textContent;b.textContent='Building…';try{await fn()}catch(e){msg(e.message,1)}b.textContent=t;b.disabled=false}
$('dl').onclick=()=>dl($('dl'),()=>exportRows(R.w2,R.h2,selRows('new'),'New_Entries'+personSuffix()));
$('dlg').onclick=()=>dl($('dlg'),()=>exportRows(R.w1,R.h1,selRows('gone'),'Gone_Entries'+personSuffix()));
})();
