<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Conversor de planilha para Instabuy</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400&display=swap" rel="stylesheet">
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);
--bg:#eef3f4;--surface:#fff;--ink:#12262b;--muted:#5a6f75;--line:#cfdadd;--accent:#0a7c72;--accent-ink:#fff;
--err:#b42318;--err-bg:#fdeceb;--warn:#8a5a00;--warn-bg:#fff4dc;--ok:#1d6b3a;--ok-bg:#e5f4ea}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--bg:#0f1b1e;--surface:#16272b;--ink:#e6f0f1;--muted:#96aeb3;--line:#2b4146;--accent:#3cc4b6;--accent-ink:#06201d;--err:#ff8a80;--err-bg:#3a1c1a;--warn:#f2c46b;--warn-bg:#3a2e12;--ok:#7ad39a;--ok-bg:#16301f}}
:root[data-theme="dark"]{color-scheme:dark;--bg:#0f1b1e;--surface:#16272b;--ink:#e6f0f1;--muted:#96aeb3;--line:#2b4146;--accent:#3cc4b6;--accent-ink:#06201d;--err:#ff8a80;--err-bg:#3a1c1a;--warn:#f2c46b;--warn-bg:#3a2e12;--ok:#7ad39a;--ok-bg:#16301f}
*,*::before,*::after{box-sizing:inherit}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--bg);color:var(--ink);font:400 15px/1.5 "IBM Plex Sans",system-ui,sans-serif}
main{max-width:1100px;margin:0 auto;padding:28px 16px 60px}
h1{font-size:26px;line-height:1.2;margin:0 0 6px;font-weight:600}
h2{font-size:17px;margin:0 0 12px;font-weight:600}
p.lead{margin:0 0 24px;color:var(--muted);max-width:62ch}
section{background:var(--surface);border:1px solid var(--line);border-radius:6px;padding:18px;margin-bottom:16px}
.drop{border:2px dashed var(--line);border-radius:6px;padding:34px 16px;text-align:center;cursor:pointer;color:var(--muted)}
.drop:hover,.drop.on,.drop:focus-visible{border-color:var(--accent);color:var(--ink);outline:none}
.drop b{color:var(--ink);display:block;font-size:16px;margin-bottom:4px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:12px}
label.f{display:block;font-size:13px;color:var(--muted)}
select,input[type=text]{width:100%;font:inherit;color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:4px;padding:7px 8px}
select:focus-visible,input:focus-visible,button:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
button{font:inherit;font-weight:500;border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:4px;padding:8px 14px;cursor:pointer}
button.p{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}
button:disabled{opacity:.5;cursor:not-allowed}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:end;margin-top:14px}
.row>div{flex:1 1 180px}
.opts{display:flex;flex-wrap:wrap;gap:8px 22px}
.opts label{display:flex;gap:7px;align-items:center;font-size:14px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin:14px 0}
.chip{padding:4px 10px;border-radius:99px;font-size:13px;font-weight:500;background:var(--bg);border:1px solid var(--line)}
.chip.e{background:var(--err-bg);color:var(--err);border-color:transparent}
.chip.w{background:var(--warn-bg);color:var(--warn);border-color:transparent}
.chip.o{background:var(--ok-bg);color:var(--ok);border-color:transparent}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:4px;max-height:540px}
table{border-collapse:collapse;width:100%;min-width:1000px;font-size:13px}
th{position:sticky;top:0;background:var(--bg);text-align:left;padding:7px 6px;border-bottom:1px solid var(--line);font-weight:600;white-space:nowrap}
td{padding:2px;border-bottom:1px solid var(--line);vertical-align:middle}
td input{width:100%;min-width:70px;border:1px solid transparent;background:transparent;color:var(--ink);font:13px "IBM Plex Mono",monospace;padding:5px 5px;border-radius:3px}
td input:hover{border-color:var(--line)}
td.n{color:var(--muted);padding:2px 8px;text-align:right}
td.pr{font-size:12px;padding:2px 8px;min-width:200px}
tr.e td{background:var(--err-bg)}tr.w td{background:var(--warn-bg)}
tr.e .pr{color:var(--err)}tr.w .pr{color:var(--warn)}
.note{font-size:13px;color:var(--muted);margin:8px 0 0}
.hide{display:none}
#msg{font-size:13px;margin-top:10px;min-height:1.2em}
</style>
</head>
<body>
<main>
<h1>Conversor de planilha para Instabuy</h1>
<p class="lead">Envie a planilha que o cliente exportou do ERP. A ferramenta liga as colunas, limpa os dados, aponta os erros e entrega o arquivo no modelo de cadastro.</p>

<section>
<h2>1. Planilha do cliente</h2>
<div class="drop" id="drop" tabindex="0" role="button" aria-label="Escolher planilha">
<b>Arraste o arquivo aqui ou clique para escolher</b>CSV, XLS ou XLSX
</div>
<input type="file" id="file" class="hide" accept=".csv,.txt,.xls,.xlsx">
<div id="msg" role="status"></div>
</section>

<section id="s2" class="hide">
<h2>2. Colunas</h2>
<p class="note" style="margin:0 0 12px">Confira o que foi reconhecido. Campos sem coluna ficam vazios.</p>
<div class="grid" id="map"></div>
<div class="row">
<div><label class="f">Nome do perfil (ex.: Linx, Consinco, cliente X)<input type="text" id="pname" placeholder="Nome do ERP ou cliente"></label></div>
<button id="psave">Salvar este mapeamento</button>
<div><label class="f">Perfis salvos<select id="plist"></select></label></div>
<button id="pload">Aplicar perfil</button>
</div>
</section>

<section id="s3" class="hide">
<h2>3. Limpeza e revisão</h2>
<div class="opts">
<label><input type="checkbox" id="oCaps" checked> Corrigir nomes em CAIXA ALTA</label>
<label><input type="checkbox" id="oEst" checked> Estoque vazio ou negativo vira 0</label>
<label><input type="checkbox" id="oInc" checked> Incremento vazio vira 1</label>
<label><input type="checkbox" id="oEan"> Esvaziar código de barras inválido</label>
</div>
<div class="chips" id="chips"></div>
<div class="opts" style="margin-bottom:10px">
<label><input type="checkbox" id="oOnly"> Mostrar só linhas com problema</label>
</div>
<div class="tw"><table><thead id="thead"></thead><tbody id="tbody"></tbody></table></div>
<p class="note" id="cap"></p>
<p class="note">Clique em qualquer célula para corrigir. Linhas vermelhas têm erro e linhas amarelas têm aviso.</p>
</section>

<section id="s4" class="hide">
<h2>4. Baixar</h2>
<div class="row" style="margin-top:0">
<button class="p" id="dAll">Baixar planilha completa</button>
<button id="dOk">Baixar só linhas sem erro</button>
<button id="dRep">Baixar relatório de problemas</button>
</div>
<p class="note">O arquivo sai no padrão do modelo: separador ponto e vírgula, valores entre aspas, preço com ponto.</p>
</section>
</main>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
const F=[
{k:'plu',l:'PLU - Código Interno',s:['plu','codigo interno','cod interno','codigo','cod','sku','referencia','id produto']},
{k:'nome',l:'Nome',s:['nome','nome produto','descricao do produto','produto','descricao','desc produto','item']},
{k:'marca',l:'Marca',s:['marca','fabricante']},
{k:'desc',l:'Descrição',s:['descricao','descricao longa','descricao complementar','detalhe','observacao']},
{k:'ean',l:'Código de Barras',s:['codigo de barras','cod barras','ean','gtin','barras','barcode']},
{k:'estoque',l:'Estoque',s:['estoque','saldo','saldo estoque','qtd estoque','qtde estoque','quantidade em estoque','qtd']},
{k:'preco',l:'Preço',s:['preco','preco venda','preco de venda','valor venda','valor','pvenda']},
{k:'inc',l:'Quantidade de Incremento',s:['quantidade de incremento','incremento','passo','multiplo']}
];
const $=id=>document.getElementById(id);
let hdr=[],raw=[],map={},rows=[],edits={},fname='planilha';

const norm=s=>String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const say=(t,bad)=>{const m=$('msg');m.textContent=t;m.style.color=bad?'var(--err)':'var(--muted)'};

/* ---------- leitura ---------- */
function decode(buf){
 try{return new TextDecoder('utf-8',{fatal:true}).decode(buf)}catch(e){return new TextDecoder('windows-1252').decode(buf)}
}
function parseCSV(t){
 t=t.replace(/^\uFEFF/,'');
 const first=t.split(/\r?\n/)[0]||'';
 const cnt=c=>first.split(c).length;
 const d=[';','\t',','].sort((a,b)=>cnt(b)-cnt(a))[0];
 const out=[];let r=[],c='',q=false;
 for(let i=0;i<t.length;i++){
  const ch=t[i];
  if(q){if(ch==='"'){if(t[i+1]==='"'){c+='"';i++}else q=false}else c+=ch}
  else if(ch==='"')q=true;
  else if(ch===d){r.push(c);c=''}
  else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&t[i+1]==='\n')i++;r.push(c);c='';out.push(r);r=[]}
  else c+=ch;
 }
 if(c!==''||r.length){r.push(c);out.push(r)}
 return out;
}
async function load(file){
 try{
  say('Lendo '+file.name+'…');
  const buf=await file.arrayBuffer();
  let data;
  if(/\.xlsx?$/i.test(file.name)){
   if(!window.XLSX)throw new Error('Leitor de Excel indisponível. Salve a planilha como CSV e tente de novo.');
   const wb=XLSX.read(buf,{type:'array'});
   data=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:''}).map(r=>r.map(v=>typeof v==='number'?String(v):String(v)));
  }else data=parseCSV(decode(buf));
  data=data.filter(r=>r.some(v=>String(v).trim()!==''));
  if(data.length<2)throw new Error('A planilha precisa ter o cabeçalho e pelo menos uma linha de produto.');
  hdr=data[0].map((h,i)=>String(h).trim()||('Coluna '+(i+1)));
  raw=data.slice(1);edits={};
  fname=file.name.replace(/\.[^.]+$/,'');
  guess();drawMap();
  ['s2','s3','s4'].forEach(s=>$(s).classList.remove('hide'));
  recompute();
  say(raw.length+' linhas lidas de '+file.name+'.');
 }catch(e){say(e.message||'Não foi possível ler o arquivo.',true)}
}

/* ---------- mapeamento ---------- */
function guess(){
 const used=new Set();map={};
 F.forEach(f=>{
  let best=-1,bs=0;
  hdr.forEach((h,i)=>{
   if(used.has(i))return;
   const n=norm(h);let s=0;
   if(f.s.includes(n))s=10-Math.min(f.s.indexOf(n),9)*.1;
   else if(f.s.some(x=>x.length>3&&n.includes(x)))s=1;
   if(s>bs){bs=s;best=i}
  });
  map[f.k]=best;if(best>=0)used.add(best);
 });
}
function drawMap(){
 $('map').innerHTML='';
 F.forEach(f=>{
  const lab=document.createElement('label');lab.className='f';lab.textContent=f.l;
  const sel=document.createElement('select');
  sel.innerHTML='<option value="-1">— sem coluna —</option>'+hdr.map((h,i)=>'<option value="'+i+'">'+esc(h)+'</option>').join('');
  sel.value=map[f.k];
  sel.onchange=()=>{map[f.k]=+sel.value;recompute()};
  lab.appendChild(sel);$('map').appendChild(lab);
 });
}
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function profiles(){try{return JSON.parse(localStorage.getItem('ib_perfis')||'{}')}catch(e){return {}}}
function listProfiles(){
 const p=profiles(),n=Object.keys(p);
 $('plist').innerHTML=n.length?n.map(x=>'<option>'+esc(x)+'</option>').join(''):'<option value="">Nenhum salvo</option>';
}
$('psave').onclick=()=>{
 const name=$('pname').value.trim();if(!name){say('Dê um nome ao perfil antes de salvar.',true);return}
 const p=profiles(),m={};
 F.forEach(f=>{m[f.k]=map[f.k]>=0?norm(hdr[map[f.k]]):''});
 p[name]=m;
 try{localStorage.setItem('ib_perfis',JSON.stringify(p));say('Perfil “'+name+'” salvo neste navegador.')}catch(e){say('Este navegador não permitiu salvar o perfil.',true)}
 listProfiles();
};
$('pload').onclick=()=>{
 const m=profiles()[$('plist').value];if(!m)return;
 F.forEach(f=>{const i=hdr.findIndex(h=>norm(h)===m[f.k]);map[f.k]=m[f.k]&&i>=0?i:-1});
 drawMap();recompute();say('Perfil aplicado.');
};

/* ---------- limpeza ---------- */
const SMALL=['de','da','do','das','dos','e','com','sem','para','em','a','o','no','na'];
function fixCaps(s){
 const letters=s.replace(/[^A-Za-zÀ-ú]/g,'');
 if(letters.length<4||letters.replace(/[^A-ZÀ-Ý]/g,'').length/letters.length<.7)return s;
 return s.toLowerCase().split(' ').map((w,i)=>{
  const u=w.match(/^(\d+(?:[.,]\d+)?)(kg|g|mg|ml|l|un|und|cx|pct|m|cm)$/);
  if(u)return u[1]+(u[2]==='l'?'L':u[2]);
  if(i>0&&SMALL.includes(w))return w;
  return w.charAt(0).toUpperCase()+w.slice(1);
 }).join(' ');
}
function num(v){
 let s=String(v).replace(/R\$|\s/g,'');if(s==='')return NaN;
 if(s.includes(',')&&s.includes('.'))s=s.lastIndexOf(',')>s.lastIndexOf('.')?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');
 else if(s.includes(','))s=s.replace(',','.');
 return /^-?\d*\.?\d+$/.test(s)?parseFloat(s):NaN;
}
const fmt=n=>String(Math.round(n*1000)/1000);
function eanOk(d){
 if(![8,12,13,14].includes(d.length))return false;
 const a=d.split('').map(Number),c=a.pop();let s=0;
 a.reverse().forEach((x,i)=>s+=x*(i%2===0?3:1));
 return (10-s%10)%10===c;
}
function cleanEan(v){
 let s=String(v).trim();if(!s)return {v:'',ok:true};
 if(/^\d+([.,]\d+)?e\+?\d+$/i.test(s))s=Number(s.replace(',','.')).toFixed(0);
 s=s.replace(/\D/g,'');if(!s)return {v:'',ok:false,raw:String(v)};
 if(eanOk(s))return {v:s,ok:true};
 if(s.length>=9&&s.length<13){const p=s.padStart(13,'0');if(eanOk(p))return {v:p,ok:true}}
 return {v:s,ok:false};
}
function build(){
 const o={caps:$('oCaps').checked,est:$('oEst').checked,inc:$('oInc').checked,ean:$('oEan').checked};
 const g=(r,k)=>map[k]>=0?String(r[map[k]]??'').replace(/\s+/g,' ').trim():'';
 rows=raw.map((r,i)=>{
  const x={plu:g(r,'plu'),nome:g(r,'nome'),marca:g(r,'marca'),desc:g(r,'desc'),ean:'',estoque:'',preco:'',inc:''},iss=[];
  if(o.caps)x.nome=fixCaps(x.nome);
  const e=cleanEan(g(r,'ean'));x.ean=e.v;x._eanBad=!e.ok;
  if(!e.ok&&o.ean)x.ean='';
  const p=g(r,'preco');x._pRaw=p;const pn=num(p);x.preco=isNaN(pn)?p:fmt(pn);
  const es=g(r,'estoque');let en=num(es);
  x.estoque=isNaN(en)?(o.est?'0':es):(en<0&&o.est?'0':fmt(en));
  const ic=g(r,'inc');const inn=num(ic);
  x.inc=isNaN(inn)?(o.inc?'1':ic):fmt(inn);
  const ed=edits[i]||{};for(const k in ed)x[k]=ed[k];
  return x;
 });
}
function validate(){
 const seen={};rows.forEach(r=>{if(r.plu)seen[r.plu]=(seen[r.plu]||0)+1});
 rows.forEach(r=>{
  const p=[];
  if(!r.plu)p.push(['e','PLU vazio']);else if(seen[r.plu]>1)p.push(['e','PLU repetido']);
  if(!r.nome)p.push(['e','Nome vazio']);
  const pn=num(r.preco);
  if(r.preco===''||isNaN(pn))p.push(['e','Preço inválido']);else if(pn<=0)p.push(['e','Preço zerado']);
  if(r.estoque!==''&&(isNaN(num(r.estoque))||num(r.estoque)<0))p.push(['e','Estoque inválido']);
  if(r.inc===''||isNaN(num(r.inc))||num(r.inc)<=0)p.push(['e','Incremento inválido']);
  if(r.ean&&!cleanEan(r.ean).ok)p.push(['w','Código de barras inválido']);
  else if(!r.ean)p.push(['w','Sem código de barras']);
  if(!r.desc)p.push(['w','Sem descrição']);
  if(!r.marca)p.push(['w','Sem marca']);
  r._p=p;r._lv=p.some(x=>x[0]==='e')?'e':p.length?'w':'o';
 });
}

/* ---------- tela ---------- */
function recompute(){build();validate();drawTable();}
function chips(){
 const c={e:0,w:0,o:0};rows.forEach(r=>c[r._lv]++);
 $('chips').innerHTML='<span class="chip">'+rows.length+' produtos</span><span class="chip e">'+c.e+' com erro</span><span class="chip w">'+c.w+' com aviso</span><span class="chip o">'+c.o+' sem problemas</span>';
 $('dOk').disabled=c.e===rows.length;
}
function drawTable(){
 chips();
 $('thead').innerHTML='<tr><th>#</th>'+F.map(f=>'<th>'+f.l+'</th>').join('')+'<th>Problemas</th></tr>';
 const only=$('oOnly').checked,MAX=300;
 const idx=rows.map((r,i)=>i).filter(i=>!only||rows[i]._lv!=='o');
 const show=idx.slice(0,MAX);
 $('tbody').innerHTML=show.map(i=>{
  const r=rows[i];
  return '<tr data-i="'+i+'" class="'+r._lv+'"><td class="n">'+(i+2)+'</td>'+F.map(f=>'<td><input data-k="'+f.k+'" value="'+esc(r[f.k])+'"></td>').join('')+'<td class="pr">'+pt(r)+'</td></tr>';
 }).join('');
 $('cap').textContent=idx.length>MAX?'Mostrando as primeiras '+MAX+' de '+idx.length+' linhas. O arquivo baixado leva todas.':'';
}
const pt=r=>r._p.map(x=>x[1]).join(', ');
$('tbody').addEventListener('change',e=>{
 const inp=e.target;if(!inp.dataset.k)return;
 const i=+inp.closest('tr').dataset.i;
 (edits[i]=edits[i]||{})[inp.dataset.k]=inp.value.trim();
 rows[i][inp.dataset.k]=inp.value.trim();
 validate();chips();
 document.querySelectorAll('#tbody tr').forEach(tr=>{const r=rows[+tr.dataset.i];tr.className=r._lv;tr.querySelector('.pr').textContent=pt(r)});
});
['oCaps','oEst','oInc','oEan','oOnly'].forEach(id=>$(id).onchange=recompute);

/* ---------- saída ---------- */
const q=v=>'"'+String(v??'').replace(/"/g,'""')+'"';
function csv(list){
 const head=F.map(f=>f.l).join(';');
 return '\uFEFF'+[head].concat(list.map(r=>F.map(f=>q(r[f.k])).join(';'))).join('\r\n')+'\r\n';
}
function save(name,data){
 try{
  const url=URL.createObjectURL(new Blob([data],{type:'text/csv;charset=utf-8'}));
  const a=document.createElement('a');a.href=url;a.download=name;
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
  say('Arquivo '+name+' pronto.');
 }catch(e){say('Não foi possível baixar: '+((e&&e.message)||'erro desconhecido'),true)}
}
$('dAll').onclick=()=>save(fname+'_instabuy.csv',csv(rows));
$('dOk').onclick=()=>save(fname+'_instabuy_sem_erros.csv',csv(rows.filter(r=>r._lv!=='e')));
$('dRep').onclick=()=>{
 const l=['Linha;PLU;Nome;Tipo;Problema'];
 rows.forEach((r,i)=>r._p.forEach(p=>l.push([i+2,q(r.plu),q(r.nome),p[0]==='e'?'Erro':'Aviso',q(p[1])].join(';'))));
 save(fname+'_problemas.csv','\uFEFF'+l.join('\r\n')+'\r\n');
};

/* ---------- entrada ---------- */
const dz=$('drop');
dz.onclick=()=>$('file').click();
dz.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('file').click()}};
$('file').onchange=e=>{if(e.target.files[0])load(e.target.files[0])};
['dragover','dragenter'].forEach(t=>dz.addEventListener(t,e=>{e.preventDefault();dz.classList.add('on')}));
['dragleave','drop'].forEach(t=>dz.addEventListener(t,e=>{e.preventDefault();dz.classList.remove('on')}));
dz.addEventListener('drop',e=>{if(e.dataTransfer.files[0])load(e.dataTransfer.files[0])});
listProfiles();
</script>
</body>
</html>
