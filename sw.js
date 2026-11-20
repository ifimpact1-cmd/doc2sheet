const C='d2s-v4';
self.addEventListener('install',e=>{self.skipWaiting()});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
 const q=e.request;if(q.method!=='GET')return;
 const same=new URL(q.url).origin===location.origin;
 const keep=r=>{if(r.ok||r.type==='opaque'){const cp=r.clone();caches.open(C).then(c=>c.put(q,cp))}return r};
 // file milik app: internet dulu (selalu terbaru), cache hanya untuk offline
 // pustaka CDN: cache dulu (hemat data & memori)
 e.respondWith(same
  ?fetch(q,{cache:'no-store'}).then(keep).catch(()=>caches.match(q).then(h=>h||caches.match('./index.html')))
  :caches.match(q).then(h=>h||fetch(q).then(keep)));
});
