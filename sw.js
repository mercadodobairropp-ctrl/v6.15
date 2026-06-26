const CACHE_NAME="painel-operacional-v615-limpo-login-enter-20260626";
const ASSETS=["./","index.html","login.html","painel.html","checklist.html","monitor.html","styles.css","config.js","app.js","manifest.json","icon-192.png","icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(CACHE_NAME).then(c=>c.addAll(ASSETS)).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
async function cacheFirst(req){
  const cached=await caches.match(req);
  const update=fetch(req).then(r=>{const copy=r.clone();caches.open(CACHE_NAME).then(c=>c.put(req,copy));return r}).catch(()=>null);
  if(cached)return cached;
  const fresh=await update;
  return fresh||await caches.match("index.html");
}
async function navigation(req){
  try{
    const r=await Promise.race([fetch(req),new Promise((_,reject)=>setTimeout(()=>reject(new Error("timeout")),2500))]);
    const copy=r.clone();caches.open(CACHE_NAME).then(c=>c.put(req,copy));
    return r;
  }catch(e){
    return await caches.match(req)||await caches.match("login.html")||await caches.match("index.html");
  }
}
self.addEventListener("fetch",e=>{if(e.request.method!=="GET")return;const url=new URL(e.request.url);if(url.origin!==self.location.origin){e.respondWith(fetch(e.request));return}if(e.request.mode==="navigate"){e.respondWith(navigation(e.request));return}e.respondWith(cacheFirst(e.request))});

