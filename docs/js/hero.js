(function(){
var c=document.getElementById('net');if(!c||matchMedia('(prefers-reduced-motion:reduce)').matches||innerWidth<700)return;
var x=c.getContext('2d'),W,H,P=[],mx=-999,my=-999,run=true;
function size(){W=c.width=c.offsetWidth;H=c.height=c.offsetHeight;P=[];var n=Math.round(W*H/16000);for(var i=0;i<n;i++)P.push({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.35,vy:(Math.random()-.5)*.35,k:i%2})}
size();addEventListener('resize',size);
c.parentNode.addEventListener('pointermove',function(e){var r=c.getBoundingClientRect();mx=e.clientX-r.left;my=e.clientY-r.top});
new IntersectionObserver(function(e){run=e[0].isIntersecting;if(run)loop()}).observe(c);
function loop(){if(!run)return;x.clearRect(0,0,W,H);
for(var i=0;i<P.length;i++){var a=P[i];a.x+=a.vx;a.y+=a.vy;if(a.x<0||a.x>W)a.vx*=-1;if(a.y<0||a.y>H)a.vy*=-1;
for(var j=i+1;j<P.length;j++){var b=P[j],dx=a.x-b.x,dy=a.y-b.y,d=dx*dx+dy*dy;if(d<17000){x.strokeStyle='rgba('+(a.k?'143,180,255':'255,143,195')+','+(1-d/17000)*.45+')';x.beginPath();x.moveTo(a.x,a.y);x.lineTo(b.x,b.y);x.stroke()}}
var near=(a.x-mx)*(a.x-mx)+(a.y-my)*(a.y-my)<9000;x.fillStyle=near?'#FF8FC3':a.k?'#8FB4FF':'#FF8FC3';x.beginPath();x.arc(a.x,a.y,near?3.2:2,0,6.3);x.fill()}
requestAnimationFrame(loop)}loop();
})();