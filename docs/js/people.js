(function(){
var q=document.getElementById('q'),chips=document.querySelectorAll('#fchips .fchip'),cards=document.querySelectorAll('.pcard'),area='all';
function run(){var s=q.value.trim().toLowerCase(),shown=0;
cards.forEach(function(c){var ok=(area==='all'||c.dataset.areas.split(' ').indexOf(area)>-1)&&(!s||c.dataset.name.indexOf(s)>-1);c.hidden=!ok;if(ok)shown++});
document.querySelectorAll('.pgroup').forEach(function(g){g.hidden=!g.querySelector('.pcard:not([hidden])')});
document.getElementById('none').hidden=shown>0}
q.addEventListener('input',run);
chips.forEach(function(b){b.addEventListener('click',function(){chips.forEach(function(x){x.classList.remove('is-on')});b.classList.add('is-on');area=b.dataset.area;run()})});
var p=new URLSearchParams(location.search).get('area');if(p){var b=document.querySelector('#fchips [data-area="'+p+'"]');if(b)b.click()}
})();