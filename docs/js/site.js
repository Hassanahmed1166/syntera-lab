(function(){
var d=document,r=d.documentElement;
var th=d.getElementById('theme');
if(th)th.addEventListener('click',function(){var dark=r.dataset.theme?r.dataset.theme==='dark':matchMedia('(prefers-color-scheme:dark)').matches;var n=dark?'light':'dark';r.dataset.theme=n;try{localStorage.setItem('syntera-theme',n)}catch(e){}});
var b=d.getElementById('burger'),nav=d.querySelector('.hdr nav');
if(b)b.addEventListener('click',function(){var o=nav.classList.toggle('open');b.setAttribute('aria-expanded',o)});
d.addEventListener('keydown',function(e){if(e.key==='Escape'&&nav){nav.classList.remove('open');b&&b.setAttribute('aria-expanded',false)}});
if('IntersectionObserver' in window){var io=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}})},{threshold:.1});d.querySelectorAll('.reveal').forEach(function(x){io.observe(x)})}else d.querySelectorAll('.reveal').forEach(function(x){x.classList.add('in')});
var reduce=matchMedia('(prefers-reduced-motion:reduce)').matches;
d.querySelectorAll('[data-count]').forEach(function(el){var n=+el.dataset.count;if(reduce)return;var t0=null;el.textContent='0';function f(t){t0=t0||t;var p=Math.min((t-t0)/900,1);el.textContent=Math.round(n*p);if(p<1)requestAnimationFrame(f)}requestAnimationFrame(f)});
var sn=d.querySelectorAll('.snav a');
if(sn.length&&'IntersectionObserver' in window){var map={};sn.forEach(function(a){map[a.getAttribute('href').slice(1)]=a});
var so=new IntersectionObserver(function(es){es.forEach(function(e){if(e.isIntersecting){sn.forEach(function(a){a.classList.remove('on')});map[e.target.id].classList.add('on')}})},{rootMargin:'-140px 0px -65% 0px'});
d.querySelectorAll('.psec').forEach(function(x){so.observe(x)})}
})();