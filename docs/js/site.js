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
d.querySelectorAll('.area,.pcard a,.pub,.panel,.values li').forEach(function(x){x.classList.add('spot');x.addEventListener('pointermove',function(e){var r=x.getBoundingClientRect();x.style.setProperty('--mx',(e.clientX-r.left)+'px');x.style.setProperty('--my',(e.clientY-r.top)+'px')})});
})();