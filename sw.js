const C='loto-v4';
self.addEventListener('install',e=>{self.skipWaiting()});
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method!=='GET'||new URL(r.url).origin!==location.origin)return;
  // la página se pide siempre recién hecha (saltando la caché del servidor) para no quedarse con una versión vieja
  const u=r.mode==='navigate'?r.url+(r.url.includes('?')?'&':'?')+'_sw='+Date.now():r.url;
  e.respondWith(fetch(u,{cache:'no-store',credentials:'same-origin'}).then(res=>{if(res.ok){const cp=res.clone();caches.open(C).then(c=>c.put(r.url,cp))}return res}).catch(()=>caches.match(r.url,{ignoreSearch:true}).then(m=>m||caches.match('./',{ignoreSearch:true}))));
});
