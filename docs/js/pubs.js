(function(){
var D=window.SYNTERA_PUBS||[],A=window.SYNTERA_AREAS||{},B=document.body.dataset.base||'';
function esc(t){return String(t).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
var T={journal:'Journal',conference:'Conference',book:'Book',chapter:'Chapter',thesis:'Thesis',preprint:'Preprint',other:'Other'};
function card(p){
 var au=p.authors.map(function(a){return a[1]?'<a class="au" href="'+B+'people/'+a[1]+'.html"><strong>'+esc(a[0])+'</strong></a>':esc(a[0])}).join('; ');
 var chips=p.areas.map(function(a){return '<a class="chip" style="--c:'+A[a].color+'" href="'+B+'research/'+a+'.html">'+A[a].short+'</a>'}).join('');
 var kw=(p.keywords||[]).map(function(k){return '<li>'+esc(k)+'</li>'}).join('');
 var li=document.createElement('li');li.className='pub';li.id='pub-'+p.id;
 li.innerHTML='<div class="pub__meta"><span class="pill pill--'+p.type+'">'+T[p.type]+'</span>'+(p.note?'<span class="pill">'+esc(p.note)+'</span>':'')+'<span class="pub__year">'+(p.year||'n.d.')+'</span></div><h3 class="pub__title">'+esc(p.title)+'</h3><p class="pub__au">'+au+'</p><p class="pub__venue"><em>'+esc(p.venue)+'</em></p>'
 +(kw?'<ul class="kw" aria-label="Keywords">'+kw+'</ul>':'')
 +(p.abstract?'<details class="abs"><summary>Abstract</summary><p>'+esc(p.abstract)+'</p></details>':'')
 +'<div class="pub__act">'+(p.doi?'<a class="pbtn" href="https://doi.org/'+esc(p.doi)+'" rel="noopener" target="_blank">DOI<span class="doi"> '+esc(p.doi)+'</span></a>':'')+'<button class="pbtn" type="button" data-cite="'+p.id+'">Cite (BibTeX)</button>'+chips+'</div>';
 return li}
function bib(p){var au=p.authors.filter(function(a){return a[0]!=='…'}).map(function(a){return a[0]}).join(' and ');
 var t={journal:'article',conference:'inproceedings',book:'book',chapter:'incollection',thesis:'phdthesis',preprint:'unpublished',other:'misc'}[p.type];
 var vf={journal:'journal',conference:'booktitle',book:'publisher',chapter:'booktitle',thesis:'school',preprint:'note',other:'howpublished'}[p.type];
 return '@'+t+'{'+p.id.replace(/-/g,'')+p.year+',\n  author = {'+au+'},\n  title = {'+p.title+'},\n  '+vf+' = {'+p.venue+'},\n  year = {'+p.year+'}'+(p.doi?',\n  doi = {'+p.doi+'}':'')+'\n}'}
document.addEventListener('click',function(e){var b=e.target.closest('[data-cite]');if(!b)return;var p=D.filter(function(x){return x.id===b.dataset.cite})[0],old=b.textContent;
 function done(){b.textContent='Copied';setTimeout(function(){b.textContent=old},1500)}
 if(navigator.clipboard)navigator.clipboard.writeText(bib(p)).then(done,function(){prompt('BibTeX',bib(p))});else prompt('BibTeX',bib(p))});
function fill(box,list){box.innerHTML='';list.forEach(function(p){box.appendChild(card(p))})}
document.querySelectorAll('[data-pubs]').forEach(function(box){
 if(box.id==='pubs')return;
 var l=D.filter(function(p){return (!box.dataset.member||p.authors.some(function(a){return a[1]===box.dataset.member}))&&(!box.dataset.area||p.areas.indexOf(box.dataset.area)>-1)});
 if(box.dataset.limit)l=l.slice(0,+box.dataset.limit);fill(box,l)});
var box=document.getElementById('pubs');if(!box)return;
var q=document.getElementById('q'),yr=document.getElementById('year'),st={type:'all',area:'all'};
var hay=D.map(function(p){return (p.title+' '+p.authors.map(function(a){return a[0]}).join(' ')+' '+p.venue+' '+(p.keywords||[]).join(' ')).toLowerCase()});
var tc=document.querySelectorAll('#tchips .fchip'),ac=document.querySelectorAll('#achips .fchip'),first=true;
function run(){var s=q.value.trim().toLowerCase();
 var l=D.filter(function(p,i){return (st.type==='all'||p.type===st.type)&&(st.area==='all'||p.areas.indexOf(st.area)>-1)&&(yr.value==='all'||String(p.year)===yr.value)&&(!s||hay[i].indexOf(s)>-1)});
 fill(box,l);document.getElementById('count').textContent=l.length+' publication'+(l.length===1?'':'s');document.getElementById('none').hidden=l.length>0;
 if(first&&location.hash){var t=document.getElementById(location.hash.slice(1));if(t)t.scrollIntoView()}first=false}
function bind(list,key,attr){list.forEach(function(b){b.addEventListener('click',function(){list.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');st[key]=b.dataset[attr];run()})})}
bind(tc,'type','type');bind(ac,'area','area');q.addEventListener('input',run);yr.addEventListener('change',run);
var u=new URLSearchParams(location.search);
if(u.get('area')){var b=document.querySelector('#achips [data-area="'+u.get('area')+'"]');if(b){ac.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');st.area=u.get('area')}}
if(u.get('type')){var t=document.querySelector('#tchips [data-type="'+u.get('type')+'"]');if(t){tc.forEach(function(x){x.classList.remove('is-on')});t.classList.add('is-on');st.type=u.get('type')}}
if(u.get('q'))q.value=u.get('q');
if(u.get('year'))yr.value=u.get('year');
run();
})();