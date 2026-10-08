(function(){
var d=document,r=d.documentElement;
var th=d.getElementById('theme');
if(th)th.addEventListener('click',function(){var dark=r.dataset.theme?r.dataset.theme==='dark':matchMedia('(prefers-color-scheme:dark)').matches;var n=dark?'light':'dark';r.dataset.theme=n;try{localStorage.setItem('syntera-theme',n)}catch(e){}});
var b=d.getElementById('burger'),nav=d.querySelector('.hdr nav');
if(b)b.addEventListener('click',function(){var o=nav.classList.toggle('open');b.setAttribute('aria-expanded',o)});
d.addEventListener('keydown',function(e){if(e.key==='Escape'&&nav){nav.classList.remove('open');b&&b.setAttribute('aria-expanded',false)}});
d.querySelectorAll('.pcard,.values li,.panel,.banner,.phead__row,.dir').forEach(function(x){x.classList.add('reveal')});
d.querySelectorAll('.grid,.values,.faces,.split,.pgroup').forEach(function(g){var i=0;g.querySelectorAll(':scope>.reveal,:scope>li.reveal').forEach(function(c){c.style.setProperty('--d',Math.min(i++,10)*.05+'s')})});
if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.1});d.querySelectorAll('.reveal').forEach(function(x){io.observe(x)})}else d.querySelectorAll('.reveal').forEach(function(x){x.classList.add('in')});
var reduce=matchMedia('(prefers-reduced-motion:reduce)').matches;
d.querySelectorAll('[data-count]').forEach(function(el){var n=+el.dataset.count;if(reduce)return;var t0=null;el.textContent='0';function f(t){t0=t0||t;var p=Math.min((t-t0)/900,1);el.textContent=Math.round(n*p);if(p<1)requestAnimationFrame(f)}requestAnimationFrame(f)});
var sn=d.querySelectorAll('.snav a');
if(sn.length&&'IntersectionObserver' in window){var map={};sn.forEach(function(a){map[a.getAttribute('href').slice(1)]=a});
var so=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){sn.forEach(function(a){a.classList.remove('on')});map[e.target.id].classList.add('on')}})},{rootMargin:'-140px 0px -65% 0px'});
d.querySelectorAll('.psec').forEach(function(x){so.observe(x)})}

var hd=d.querySelector('.hdr'),bar=d.createElement('div');bar.id='prog';d.body.appendChild(bar);
var tp=d.createElement('button');tp.id='top';tp.type='button';tp.setAttribute('aria-label','Back to top');tp.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5M5 12l7-7 7 7"/></svg>';d.body.appendChild(tp);
tp.addEventListener('click',function(){scrollTo({top:0,behavior:reduce?'auto':'smooth'})});
var tick=false;function onS(){var h=d.documentElement,m=h.scrollHeight-h.clientHeight,y=h.scrollTop;bar.style.transform='scaleX('+(m>0?y/m:0)+')';hd&&hd.classList.toggle('is-scrolled',y>10);tp.classList.toggle('show',y>600);tick=false}
addEventListener('scroll',function(){if(!tick){tick=true;requestAnimationFrame(onS)}},{passive:true});onS();
d.querySelectorAll('.area,.pcard>a,.pub,.panel,.values li').forEach(function(x){x.classList.add('spot');x.addEventListener('pointermove',function(e){var r=x.getBoundingClientRect();x.style.setProperty('--mx',(e.clientX-r.left)+'px');x.style.setProperty('--my',(e.clientY-r.top)+'px')})});


(function(){var b=d.querySelector('.hdr .brand');if(!b||reduce)return;var e=d.createElement('span');e.className='bf-logo';e.setAttribute('aria-hidden','true');
e.innerHTML='<svg viewBox="-20 -16 40 32"><g class="wl"><path d="M0 0C-6-14-18-15-19-7-20-1-10 2 0 0Z" fill="#FF8FC3"/><path d="M0 1C-8 3-15 9-11 13-7 16-1 9 0 1Z" fill="#E83E8C"/><circle cx="-11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><g class="wr"><path d="M0 0C6-14 18-15 19-7 20-1 10 2 0 0Z" fill="#FF8FC3"/><path d="M0 1C8 3 15 9 11 13 7 16 1 9 0 1Z" fill="#E83E8C"/><circle cx="11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><rect x="-1" y="-6" width="2" height="14" rx="1" fill="#fff"/></svg>';
b.appendChild(e)})();
var hero=d.querySelector('.hero');
if(!reduce&&hero){(function(){
var cols=[['#FF8FC3','#E83E8C'],['#8FB4FF','#3A7BFF'],['#C4A3FF','#8B5CF6'],['#FFC2DE','#E83E8C'],['#9ED0FF','#2D66E6']];
var n=innerWidth<700?3:5,L=d.createElement('div');L.id='bfly';L.setAttribute('aria-hidden','true');hero.appendChild(L);var W=0,H=0;function sz(){W=L.offsetWidth;H=L.offsetHeight}sz();addEventListener('resize',sz);
var mx=-999,my=-999;hero.addEventListener('pointermove',function(e){var r=L.getBoundingClientRect();mx=e.clientX-r.left;my=e.clientY-r.top},{passive:true});hero.addEventListener('pointerleave',function(){mx=my=-999});
function svg(c){return '<svg viewBox="-20 -16 40 32"><g class="wl"><path d="M0 0C-6-14-18-15-19-7-20-1-10 2 0 0Z" fill="'+c[0]+'"/><path d="M0 1C-8 3-15 9-11 13-7 16-1 9 0 1Z" fill="'+c[1]+'"/><circle cx="-11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><g class="wr"><path d="M0 0C6-14 18-15 19-7 20-1 10 2 0 0Z" fill="'+c[0]+'"/><path d="M0 1C8 3 15 9 11 13 7 16 1 9 0 1Z" fill="'+c[1]+'"/><circle cx="11" cy="-6" r="2.2" fill="#fff" opacity=".75"/></g><rect x="-1" y="-6" width="2" height="14" rx="1" fill="#14213D"/></svg>'}
var B=[];
for(var i=0;i<n;i++){var el=d.createElement('div');el.className='bf';el.style.setProperty('--f',(.26+Math.random()*.16)+'s');el.innerHTML=svg(cols[i%cols.length]);L.appendChild(el);
B.push({el:el,x:Math.random()*W,y:Math.random()*H,tx:0,ty:0,a:0,sp:.9+Math.random()*.9,ph:Math.random()*6,sc:.7+Math.random()*.6,rest:0})}
function pick(b){b.tx=Math.random()*W;b.ty=30+Math.random()*Math.max(H-80,10)}
B.forEach(pick);
var t=0,run=true,vis=true,inv=true;
function step(){if(!run)return;t+=.03;
B.forEach(function(b){
if(b.rest>0){b.rest--;b.el.style.setProperty('--f','.9s')}else{b.el.style.setProperty('--f','.3s')
var dx=b.tx-b.x,dy=b.ty-b.y,dist=Math.hypot(dx,dy);
var fx=b.x-mx,fy=b.y-my,fd=Math.hypot(fx,fy);
if(fd<130){b.x+=fx/fd*3.2;b.y+=fy/fd*3.2}
if(dist<30){pick(b);if(Math.random()<.35)b.rest=90+Math.random()*120}
else{var ang=Math.atan2(dy,dx)+Math.sin(t*2+b.ph)*.9;b.x+=Math.cos(ang)*b.sp;b.y+=Math.sin(ang)*b.sp+Math.sin(t*5+b.ph)*.5;b.a=ang}}
var rot=(Math.cos(b.a)*14);
b.el.style.transform='translate('+b.x.toFixed(1)+'px,'+b.y.toFixed(1)+'px) rotate('+rot.toFixed(1)+'deg) scale('+b.sc+')'});
requestAnimationFrame(step)}
function upd(){var r=vis&&inv;if(r&&!run){run=true;step()}else run=r}
d.addEventListener('visibilitychange',function(){vis=!d.hidden;upd()});
if('IntersectionObserver' in window)new IntersectionObserver(function(e){inv=e[0].isIntersecting;upd()}).observe(hero);
step();
})()}
})();